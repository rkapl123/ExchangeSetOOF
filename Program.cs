using System;
using Microsoft.Exchange.WebServices.Data; // Exchange Web Service reference
using System.DirectoryServices.AccountManagement; // required to get the email address of the currently logged in user
using Microsoft.Win32; // required for registry handling
using System.IO; // required for log file
using System.Windows.Forms;
using System.Collections.Generic;

namespace ExchangeSetOOF {
    class Program {
        // Exchange web service that we're going to connect to
        static ExchangeService service;
        // prefix that defines OOF reply bodies to act as a template, ALWAYS uppercase (converted in code)
        static string templateSpec = "VORLAGE:";
        // placeholders and replacements for two languages (for replacement rules see ExchangeSetOOF.exe.cfg)
        public static string[] DateLang1 = { "!BisDatum!", "!VonDatumBisDatum!", "am", "von", "bis", "true" };
        public static string[] DateLang2 = { "!ToDate!", "!FromDateToDate!", "on", "from", "until", "true" };
        public static StreamWriter logfile;
        public static string EasterSetting;
        public static List<string> holidays;

        static void Main(string[] args) {
            try {
                logfile = new StreamWriter(System.Reflection.Assembly.GetExecutingAssembly().Location + ".log", false, System.Text.Encoding.GetEncoding(1252));
            } catch (Exception ex) {
                MessageBox.Show("Exception occurred when trying to write to log: " + ex.Message);
                return;
            }
            // reading configuration file for templateSpec and DateLang1 and DateLang2 placeholders
            logMsg("starting ExchangeSetOOF");
            try {
                StreamReader configfile = new StreamReader(System.Reflection.Assembly.GetExecutingAssembly().Location + ".cfg");
                templateSpec = configfile.ReadLine().ToUpper();
                string readline = configfile.ReadLine();
                DateLang1 = readline.Split('\t');
                readline = configfile.ReadLine();
                DateLang2 = readline.Split('\t');
                readline = configfile.ReadLine();
                holidays = new List<string>(readline.Split('\t'));
                EasterSetting = configfile.ReadLine().ToUpper();

                configfile.Close();
            } catch (Exception ex) {
                logFinal("Exception occurred when reading configuration file " + System.Reflection.Assembly.GetExecutingAssembly().Location + ".cfg" + " for ExchangeSetOOF.exe: " + ex.Message);
                return;
            }


            try {
                service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                service.UseDefaultCredentials = true;
                service.AutodiscoverUrl(UserPrincipal.Current.EmailAddress, RedirectionUrlValidationCallback);
            } catch (Exception ex) {
                logFinal("Exception occurred when setting service for EWS: " + ex.Message);
                return;
            }

            // find next "out of office" appointment being either today or on the next business day
            logMsg("getting OOF appointments");
            DateTime startDate = DateTime.Now;    //startDate = DateTime.Parse("2016-08-05"); //uncomment to test/debug
            DateTime endDate = startDate.AddBusinessDays(2, holidays, EasterSetting); // need to add 2 days because otherwise the endDate is <nextBDate> 00:00:00
            // Initialize the calendar folder object with only the folder ID.
            FindItemsResults<Appointment> appointments = null;
            try {
                CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
                // Set the start and end time and number of appointments to retrieve.
                CalendarView cView = new CalendarView(startDate, endDate, 20);
                // Limit the properties returned to the appointment's subject, start time, and end time.
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.LegacyFreeBusyStatus, AppointmentSchema.IsRecurring, AppointmentSchema.AppointmentType);
                // Retrieve a collection of appointments by using the calendar view.
                appointments = calendar.FindAppointments(cView);
            } catch (Exception ex) {
                logFinal("Exception occurred when searching OOF appointments in users calendar: " + ex.Message);
                return;
            }

            Appointment oofAppointment = null;
            DateTime myStartOOFDate = new DateTime();
            DateTime myEndOOFDate = new DateTime();
            foreach (Appointment a in appointments) {
                if (a.LegacyFreeBusyStatus == LegacyFreeBusyStatus.OOF) {
                    // search for longest OOF appointment
                    if (oofAppointment == null || oofAppointment.End < a.End) {
                        //  OOF end dates need to end in the future (otherwise results in an exception when setting the OOF schedule)
                        if (a.End > DateTime.Now) {
                            logMsg("oofAppointment " + a.Subject + " detected,Start: " + a.Start.ToString() + ",(later)End: " + a.End.ToString() + ",LegacyFreeBusyStatus: " + a.LegacyFreeBusyStatus.ToString() + ",IsRecurring: " + a.IsRecurring.ToString() + ",AppointmentType: " + a.AppointmentType.ToString());
                            // set the oofAppointment to control the OOF setting later...
                            oofAppointment = a;
                            myStartOOFDate = a.Start;
                            myEndOOFDate = a.End;
                        }
                    }
                }
            }

            // change automatic replies (out of office), first get the existing ones, used as templates for the actual ones.
            logMsg("getting users OOF settings");
            OofSettings myOOF = null;
            try {
                myOOF = service.GetUserOofSettings(UserPrincipal.Current.EmailAddress);
            } catch (Exception ex) {
                logFinal("Exception occurred when getting users OOF settings: " + ex.Message);
                return;
            }
            // templates for internal and external replies are stored in registry, if key doesn't exist, create it
            string keyName = "HKEY_CURRENT_USER\\Software\\RK\\ExchangeSetOOF";
            if (Registry.GetValue(keyName, "OOFtemplateInt", null) == null) {
                Registry.SetValue(keyName, "OOFtemplateInt", "", RegistryValueKind.String);
            }
            if (Registry.GetValue(keyName, "OOFtemplateExt", null) == null) {
                Registry.SetValue(keyName, "OOFtemplateExt", "", RegistryValueKind.String);
            }

            // templateSpec in both int and ext message as prefix -> save as Template
            if (myOOF.InternalReply.Message.ToUpper().Contains(templateSpec) && myOOF.ExternalReply.Message.ToUpper().Contains(templateSpec)) {
                Registry.SetValue(keyName, "OOFtemplateInt", myOOF.InternalReply.Message, RegistryValueKind.String);
                Registry.SetValue(keyName, "OOFtemplateExt", myOOF.ExternalReply.Message, RegistryValueKind.String);
                logMsg("Both internal and external replies contain templateSpec, so templates saved to registry");
                logMsg("=================================================== internal Reply Template:");
                logMsg(myOOF.InternalReply.Message);
                logMsg("=================================================== external Reply Template:");
                logMsg(myOOF.ExternalReply.Message);
                logMsg("===================================================");
                // OOF not enabled or scheduled -> restore Template (only if non empty!)
            } else if (myOOF.State == OofState.Disabled) {
                if (Registry.GetValue(keyName, "OOFtemplateInt", "").ToString() != "") {
                    myOOF.InternalReply.Message = Registry.GetValue(keyName, "OOFtemplateInt", "").ToString();
                }
                if (Registry.GetValue(keyName, "OOFtemplateExt", "").ToString() != "") {
                    myOOF.ExternalReply.Message = Registry.GetValue(keyName, "OOFtemplateExt", "").ToString();
                }
                logMsg("OOFstate disabled, so templates restored from registry:");
                logMsg("internal Reply:" + myOOF.InternalReply.Message);
                logMsg("external Reply:" + myOOF.ExternalReply.Message);
            } else {
                logMsg("nothing to do with templates: OOF.State = " + myOOF.State.ToString());
            }


            // out of office appointment today or on the next (business) day -> enable OOF, but only if not already enabled or scheduled (as this would reactivate OOF messages again, which is undesired)
            if (oofAppointment != null && myOOF.State != OofState.Enabled && myOOF.State != OofState.Scheduled) {
                string replyTextInt = "", replyTextExt = "";
                if (Registry.GetValue(keyName, "OOFtemplateInt", "").ToString() != "") {
                    replyTextInt = Registry.GetValue(keyName, "OOFtemplateInt", "").ToString();
                }
                if (Registry.GetValue(keyName, "OOFtemplateExt", "").ToString() != "") {
                    replyTextExt = Registry.GetValue(keyName, "OOFtemplateExt", "").ToString();
                }
                // remove templateSpec
                replyTextInt = replyTextInt.Replace(templateSpec, "");
                replyTextExt = replyTextExt.Replace(templateSpec, "");
                // convert end date to string for OOF Message
                string myEndOOFDateStr, myStartOOFDateStr, myEndOOFDateStrPresent;
                if (myEndOOFDate.TimeOfDay.ToString() == "00:00:00") { // modify whole dates
                    // end date is next day 00:00:00, which is too far, when truncated to date part only...
                    myEndOOFDateStr = myEndOOFDate.AddDays(-1).ToShortDateString();
                    // calc next business day to show the date when we're in the office present again (if last part of DateLang string is "true")
                    myEndOOFDateStrPresent = myEndOOFDate.AddBusinessDays(1, holidays, EasterSetting).AddDays(-1).ToShortDateString();
                } else {
                    myEndOOFDateStr = myEndOOFDate.ToString(); // incl. time part
                    myEndOOFDateStrPresent = myEndOOFDate.ToString(); // incl. time part
                }
                // notification ends one day before returning
                DateTime myEndOOFDateNotify = myEndOOFDate.AddBusinessDays(1, holidays, EasterSetting).AddDays(-2);
                // convert start date to string for OOF Message
                if (myStartOOFDate.TimeOfDay.ToString() == "00:00:00") { // modify whole dates
                    myStartOOFDateStr = myStartOOFDate.ToShortDateString();
                } else {
                    myStartOOFDateStr = myStartOOFDate.ToString(); // incl. time part
                }
                // replace template variables in two languages
                if (myEndOOFDate != myStartOOFDate) {
                    // in case there is only the end day to be shown ("!DatumBis!"/"!DateTo!") only show the Date
                    replyTextInt = replyTextInt.Replace(DateLang1[0], DateLang1[4] + " " + (DateLang1[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                    replyTextInt = replyTextInt.Replace(DateLang2[0], DateLang2[4] + " " + (DateLang2[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                    // in case there are both the start and the end day to be shown ("!Datum!"/"!Date!")
                    replyTextInt = replyTextInt.Replace(DateLang1[1], DateLang1[3] + " " + myStartOOFDateStr + " " + DateLang1[4] + " " + (DateLang1[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                    replyTextInt = replyTextInt.Replace(DateLang2[1], DateLang2[3] + " " + myStartOOFDateStr + " " + DateLang2[4] + " " + (DateLang2[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                    // the same for the external message
                    replyTextExt = replyTextExt.Replace(DateLang1[0], DateLang1[4] + " " + (DateLang1[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                    replyTextExt = replyTextExt.Replace(DateLang2[0], DateLang2[4] + " " + (DateLang2[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                    replyTextExt = replyTextExt.Replace(DateLang1[1], DateLang1[3] + " " + myStartOOFDateStr + " " + DateLang1[4] + " " + (DateLang1[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                    replyTextExt = replyTextExt.Replace(DateLang2[1], DateLang2[3] + " " + myStartOOFDateStr + " " + DateLang2[4] + " " + (DateLang2[5] == "true" ? myEndOOFDateStrPresent : myEndOOFDateStr));
                } else {
                    // special case: duration of OOF is exactly one day, always only show the start date
                    // in case there is only the end day to be shown ("!DatumBis!"/"!DateTo!")
                    replyTextInt = replyTextInt.Replace(DateLang1[0], DateLang1[2] + " " + myStartOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang2[0], DateLang2[2] + " " + myStartOOFDateStr);
                    // in case there are both the start and the end day to be shown ("!Datum!"/"!Date!")
                    replyTextInt = replyTextInt.Replace(DateLang1[1], DateLang1[2] + " " + myStartOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang2[1], DateLang2[2] + " " + myStartOOFDateStr);
                    // the same for the external message
                    replyTextExt = replyTextExt.Replace(DateLang1[0], DateLang1[2] + " " + myStartOOFDateStr);
                    replyTextExt = replyTextExt.Replace(DateLang2[0], DateLang2[2] + " " + myStartOOFDateStr);
                    replyTextExt = replyTextExt.Replace(DateLang1[1], DateLang1[2] + " " + myStartOOFDateStr);
                    replyTextExt = replyTextExt.Replace(DateLang2[1], DateLang2[2] + " " + myStartOOFDateStr);
                }

                // Set the OOF message for internal audience.
                myOOF.InternalReply.Message = replyTextInt;
                // Set the same OOF message for external audience.
                myOOF.ExternalReply.Message = replyTextExt;
                // Set the OOF status to scheduled time period.
                myOOF.State = OofState.Scheduled;
                // Select the scheduled time period to send OOF replies.
                myOOF.Duration = new TimeWindow(myStartOOFDate, myEndOOFDateNotify);
                logMsg("OOF appointment detected, so schedule set to " + myStartOOFDate.ToString() + " - " + myEndOOFDateNotify.ToString() + ", OOF state set to scheduled and int/ext replies set changed accordingly:");
                logMsg("=================================================== internal Reply:");
                logMsg(myOOF.InternalReply.Message);
                logMsg("=================================================== external Reply:");
                logMsg(myOOF.ExternalReply.Message);
                logMsg("===================================================");
                // Now send the OOF settings to Exchange server. This method will result in a call to EWS.
                try {
                    logMsg("sending OOF Settings and OOFState = Scheduled to EWS");
                    service.SetUserOofSettings(UserPrincipal.Current.EmailAddress, myOOF);
                } catch (Exception ex) {
                    logFinal("Exception occurred when sending to EWS: " + ex.Message);
                    return;
                }
            } else if (oofAppointment == null && myOOF.State != OofState.Disabled) {
                // just in case exchange server didn't disable OOF automatically.
                logMsg("no OOF appointment detected and OOFstate not disabled, so set OOFstate to disabled (just in case exchange didn't do this)");
                myOOF.State = OofState.Disabled;
                if (Registry.GetValue(keyName, "OOFtemplateInt", "").ToString() != "") {
                    myOOF.InternalReply.Message = Registry.GetValue(keyName, "OOFtemplateInt", "").ToString();
                }
                if (Registry.GetValue(keyName, "OOFtemplateExt", "").ToString() != "") {
                    myOOF.ExternalReply.Message = Registry.GetValue(keyName, "OOFtemplateExt", "").ToString();
                }
                logMsg("OOFstate set to disabled, templates restored from registry:");
                logMsg("internal Reply:" + myOOF.InternalReply.Message);
                logMsg("external Reply:" + myOOF.ExternalReply.Message);
                // Now send the OOF settings to Exchange server. This method will result in a call to EWS.
                try {
                    logMsg("sending OOFState = Disabled to EWS");
                    service.SetUserOofSettings(UserPrincipal.Current.EmailAddress, myOOF);
                } catch (Exception ex) {
                    logFinal("Exception occurred when sending to EWS: " + ex.Message);
                    return;
                }
            } else {
                logMsg("nothing to do with replacing/scheduling/setting state (OOF State: " + myOOF.State.ToString() + ")");
            }
            logFinal("finished ExchangeSetOOF");
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl) {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https") {
                result = true;
            }
            return result;
        }

        private static void logMsg(string msg) {
            logfile.WriteLine(msg);
            logfile.Flush();
        }

        private static void logFinal(string msg) {
            logMsg(msg);
            logfile.Close();
        }
    }

    public static class DateTimeExt {
        private static DateTime EasterSunday(int year) {
            int day = 0;
            int month = 0;

            int g = year % 19;
            int c = year / 100;
            int h = (c - (int)(c / 4) - (int)((8 * c + 13) / 25) + 19 * g + 15) % 30;
            int i = h - (int)(h / 28) * (1 - (int)(h / 28) * (int)(29 / (h + 1)) * (int)((21 - g) / 11));
            day = i - ((year + (int)(year / 4) + i + 2 - c + (int)(c / 4)) % 7) + 28;
            month = 3;
            if (day > 31) {
                month++;
                day -= 31;
            }
            return new DateTime(year, month, day);
        }

        public static bool isHoliday(this DateTime theDate, List<string> holidays, string EasterSetting) {
            string datechoice = theDate.Day + "." + theDate.Month;
            // fixed holidays
            if (holidays.Contains(datechoice)) return true;
            // weekends (leave empty the holiday line if they also should not be respected)
            if (holidays.Count > 0 && (theDate.DayOfWeek == DayOfWeek.Saturday || theDate.DayOfWeek == DayOfWeek.Sunday)) return true;
            // floating (EasterSunday dependent) holidays:
            // Good Friday, Easter Monday, Ascension day, Whit Monday, Corpus Christi day
            if (theDate == EasterSunday(theDate.Year).AddDays(-2) || theDate == EasterSunday(theDate.Year).AddDays(1) || theDate == EasterSunday(theDate.Year).AddDays(39) || theDate == EasterSunday(theDate.Year).AddDays(50) || theDate == EasterSunday(theDate.Year).AddDays(60)) {
                if (theDate == EasterSunday(theDate.Year).AddDays(-2)) {
                    if (EasterSetting == "EASTERGF") return true;
                } else {
                    if (EasterSetting.Contains("EASTER")) return true;
                }
            }
            return false;
        }

        public static DateTime AddBusinessDays(this DateTime date, int days, List<string> holidays, string EasterSetting) {
            if (days == 0) return date;
            while (days > 0) {
                date = date.AddDays(1);
                // add an additional day if it is a holiday or weekend (only regarded when any holiday is set)
                if (date.isHoliday(holidays, EasterSetting)) {
                    date = date.AddDays(1);
                } else {
                    days--;
                }
            }
            return date;
        }
    }

    public static class StringExt {
        public static string Truncate(this string value, int maxLength) {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }
}
