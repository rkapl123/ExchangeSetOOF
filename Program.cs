using System;
using Microsoft.Exchange.WebServices.Data; // Exchange Web Service reference
using System.DirectoryServices.AccountManagement; // required to get the email adress of the currently logged in user
using Microsoft.Win32; // required for registry handling

namespace ExchangeSetOOF
{
    class Program
    {
        static ExchangeService service;
        const string templateSpec = "VORLAGE:"; // prefix that defines OOF reply bodies to act as a template, ALWAYS uppercase (converted in code)
        // placeholders for two languages, being replaced using following rule (hardcoded):
        // DateLangX[0] changed to DateLangX[4] + " " + OOF_EndDate
        // DateLangX[1] changed to DateLangX[3] + " " + OOF_StartDate + " " + DateLangX[4] + " " + OOF_EndDate
        // in case of whole single day absences both DateLangX[0] and DateLangX[1] 
        // are being replaced by DateLangX[2] + " " + OOF_EndDate 
        public static readonly string[] DateLang1 = { "!DatumBis!", "!Datum!", "am", "von", "bis" };
        public static readonly string[] DateLang2 = { "!DateTo!", "!Date!", "on", "from", "until" };

        static void Main(string[] args) {
            Console.WriteLine("starting ExchangeSetOOF");
            try {
                service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                service.UseDefaultCredentials = true;
                service.AutodiscoverUrl(UserPrincipal.Current.EmailAddress, RedirectionUrlValidationCallback);
            } catch (Exception ex) {
                Console.WriteLine("Exception occured when setting service for EWS: " + ex.Message);
            }

            // find next "out of office" appointment being either today or on the next business day
            Console.WriteLine("getting oof appointments..");
            DateTime startDate = DateTime.Now;    //startDate = DateTime.Parse("2016-08-05"); //uncomment to test/debug
            DateTime endDate = startDate.AddBusinessDays(2); // need to add 2 days because otherwise the endDate is <nextBDate> 00:00:00
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
                Console.WriteLine("Exception occured when searching OOF appointments in users calendar: " + ex.Message);
            }
            Appointment oofAppointment = null;
            DateTime myStartOOFDate = new DateTime();
            DateTime myEndOOFDate = new DateTime();
            foreach (Appointment a in appointments) {
                if (a.LegacyFreeBusyStatus == LegacyFreeBusyStatus.OOF) {
                    // search for longest OOF appointment
                    if (oofAppointment == null  || oofAppointment.End < a.End) {
                        //  oof end dates need to end in the future (otherwise results in an exception when setting the OOF schedule)
                        if (a.End > DateTime.Now) {
                            Console.Write("oofAppointment " + a.Subject + " detected,Start: " + a.Start.ToString());
                            Console.Write(",(later)End: " + a.End.ToString());
                            Console.Write(",LegacyFreeBusyStatus: " + a.LegacyFreeBusyStatus.ToString());
                            Console.Write(",IsRecurring: " + a.IsRecurring.ToString());
                            Console.Write(",AppointmentType: " + a.AppointmentType.ToString());
                            Console.WriteLine();
                            // set the oofAppointment to control the OOF setting later...
                            oofAppointment = a;
                            myStartOOFDate = a.Start;
                            myEndOOFDate = a.End;
                        }
                    }
                }
            }

            // change automatic replies (out of office), first get the existing ones (template!)
            Console.WriteLine("getting users OOF settings");
            OofSettings myOOF = null;
            try {
                myOOF = service.GetUserOofSettings(UserPrincipal.Current.EmailAddress);
            } catch (Exception ex) {
                Console.WriteLine("Exception occured when getting users OOF settings: " + ex.Message);
            }
            // templates for internal and external replies are stored in registry, if key doesn't exist, create it
            string keyName = "HKEY_CURRENT_USER\\Software\\RK\\ExchangeSetOOF";
            if (Registry.GetValue(keyName, "OOFtemplateInt", null) == null) {
                Registry.SetValue(keyName, "OOFtemplateInt", "", RegistryValueKind.String);
            }
            if (Registry.GetValue(keyName, "OOFtemplateExt", null) == null) {
                Registry.SetValue(keyName, "OOFtemplateExt", "", RegistryValueKind.String);
            }
            // OOF not set and templateSpec as prefix -> save as Template
            if (myOOF.State == OofState.Disabled && myOOF.InternalReply.Message.ToUpper().Contains(templateSpec) && myOOF.ExternalReply.Message.ToUpper().Contains(templateSpec)) {
                Registry.SetValue(keyName, "OOFtemplateInt", myOOF.InternalReply.Message, RegistryValueKind.String);
                Registry.SetValue(keyName, "OOFtemplateExt", myOOF.ExternalReply.Message, RegistryValueKind.String);
                Console.WriteLine("OOFstate disabled and both internal and external replies contain templateSpec, so templates saved to registry");
                Console.WriteLine("=================================================== internal Reply Template:");
                Console.WriteLine(myOOF.InternalReply.Message);
                Console.WriteLine("=================================================== external Reply Template:");
                Console.WriteLine(myOOF.ExternalReply.Message);
                Console.WriteLine("===================================================");
            // OOF not set or no oof appointment and templateSpec not given -> restore Template
            } else if (myOOF.State == OofState.Disabled || oofAppointment == null) {
                myOOF.InternalReply.Message = Registry.GetValue(keyName, "OOFtemplateInt", "").ToString();
                myOOF.ExternalReply.Message = Registry.GetValue(keyName, "OOFtemplateExt", "").ToString();
                Console.WriteLine("either OOFstate disabled or no oof appointment found, so templates restored from registry:");
                Console.WriteLine("internal Reply:" + myOOF.InternalReply.Message);
                Console.WriteLine("external Reply:" + myOOF.ExternalReply.Message);
            } else {
                Console.WriteLine("nothing to do with templates: OOF.State = " + myOOF.State.ToString() + ", oofAppointment found: " + oofAppointment.Subject);
            }

            // out of office appointment today or on the next (business) day -> enable OOF
            if (!(oofAppointment == null) && myOOF.State == OofState.Disabled) {
                string replyTextInt = Registry.GetValue(keyName, "OOFtemplateInt", "").ToString();
                string replyTextExt = Registry.GetValue(keyName, "OOFtemplateExt", "").ToString();
                // remove templateSpec
                replyTextInt = replyTextInt.Replace(templateSpec, "");
                replyTextExt = replyTextExt.Replace(templateSpec, "");
                // convert end date to string for OOF Message
                string myEndOOFDateStr, myStartOOFDateStr;
                if (myEndOOFDate.TimeOfDay.ToString() == "00:00:00") { // modify whole dates
                    // end date is next day 00:00:00, which is too far, when truncated to date part only...
                    myEndOOFDateStr = myEndOOFDate.AddDays(-1).ToShortDateString();
                } else {
                    myEndOOFDateStr = myEndOOFDate.ToString(); // incl. time part
                }
                // convert start date to string for OOF Message
                if (myStartOOFDate.TimeOfDay.ToString() == "00:00:00") { // modify whole dates
                    myStartOOFDateStr = myStartOOFDate.ToShortDateString();
                } else {
                    myStartOOFDateStr = myStartOOFDate.ToString(); // incl. time part
                }
                // replace template variables in two languages
                if (myEndOOFDate != myStartOOFDate) {
                    replyTextInt = replyTextInt.Replace(DateLang1[0], DateLang1[4] + " " + myEndOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang2[0], DateLang2[4] + " " + myEndOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang1[1], DateLang1[3] + " " + myStartOOFDateStr + " " + DateLang1[4] + " " + myEndOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang2[1], DateLang2[3] + " " + myStartOOFDateStr + " " + DateLang2[4] + " " + myEndOOFDateStr);
                    replyTextExt = replyTextExt.Replace(DateLang1[0], DateLang1[4] + " " + myEndOOFDateStr);
                    replyTextExt = replyTextExt.Replace(DateLang2[0], DateLang2[4] + " " + myEndOOFDateStr);
                    replyTextExt = replyTextExt.Replace(DateLang1[1], DateLang1[3] + " " + myStartOOFDateStr + " " + DateLang1[4] + " " + myEndOOFDateStr);
                    replyTextExt = replyTextExt.Replace(DateLang2[1], DateLang2[3] + " " + myStartOOFDateStr + " " + DateLang2[4] + " " + myEndOOFDateStr);
                } else {
                    // special case: exactly one day
                    replyTextInt = replyTextInt.Replace(DateLang1[0], DateLang1[2] + " " + myStartOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang2[0], DateLang2[2] + " " + myStartOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang1[1], DateLang1[2] + " " + myStartOOFDateStr);
                    replyTextInt = replyTextInt.Replace(DateLang2[1], DateLang2[2] + " " + myStartOOFDateStr);
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
                myOOF.Duration = new TimeWindow(myStartOOFDate, myEndOOFDate);
                Console.WriteLine("oof appointment detected and OOFstate disabled, so schedule set, oof state set to scheduled and int/ext replies set changed accordingly:");
                Console.WriteLine("=================================================== internal Reply:");
                Console.WriteLine(myOOF.InternalReply.Message);
                Console.WriteLine("=================================================== external Reply:");
                Console.WriteLine(myOOF.ExternalReply.Message);
                Console.WriteLine("===================================================");
            } else if ((oofAppointment == null) && myOOF.State != OofState.Disabled) {
                // just in case exchange server didn't disable OOF automatically.
                myOOF.State = OofState.Disabled;
                Console.WriteLine("no oof appointment detected and OOFstate not disabled, so set OOFstate to disabled (just in case exchange didn't do this)");
            } else {
                Console.WriteLine("nothing to do with replacing/scheduling: OOF State: " + myOOF.State.ToString());
            }
            // Now send the OOF settings to Exchange server. This method will result in a call to EWS.
            try {
                Console.WriteLine("sending changed OOF Settings to EWS..");
                service.SetUserOofSettings(UserPrincipal.Current.EmailAddress, myOOF);
            } catch (Exception ex) {
                Console.WriteLine("Exception occured when sending User OOF Settings to EWS: " + ex.Message);
            }
            Console.WriteLine("finished ExchangeSetOOF");
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

        public static bool isHoliday(this DateTime theDate) {
            string datechoice = theDate.Day + "." + theDate.Month;
            // fixed austrian holidays
            switch (datechoice) {
                case "1.1":
                    return true;
                case "6.1":
                    return true;
                case "1.5":
                    return true;
                case "15.8":
                    return true;
                case "26.10":
                    return true;
                case "1.11":
                    return true;
                case "8.12":
                    return true;
                case "24.12":
                    return true;
                case "25.12":
                    return true;
                case "26.12":
                    return true;
            }
            // weekends
            if ((theDate.DayOfWeek == DayOfWeek.Saturday) || (theDate.DayOfWeek == DayOfWeek.Sunday)) {
                return true;
            }
            // floating (EasterSunday dependent) austrian holidays:
            // Easter Monday (Good Friday would be -2), ascension day (Christi Himmelfahrt), whit monday (Pfingstmontag), corpus christi day (Fronleichnam)
            if ((theDate == EasterSunday(theDate.Year).AddDays(1)) || (theDate == EasterSunday(theDate.Year).AddDays(39)) || (theDate == EasterSunday(theDate.Year).AddDays(50)) || (theDate == EasterSunday(theDate.Year).AddDays(60))) {
                return true;
            }
            return false;
        }

            //Console.WriteLine("{0}", DateTime.Parse("2012-01-01").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-01-06").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-05-01").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-08-15").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-10-26").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-11-01").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-12-08").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-12-24").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-12-25").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-12-26").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-04-05").isHoliday()); //false - maundy thursday
            //Console.WriteLine("{0}", DateTime.Parse("2012-04-06").isHoliday()); //false - good friday
            //Console.WriteLine("{0}", DateTime.Parse("2012-04-07").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-04-08").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-04-09").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-05-17").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-05-28").isHoliday());
            //Console.WriteLine("{0}", DateTime.Parse("2012-06-07").isHoliday());

        public static DateTime AddBusinessDays(this DateTime date, int days) {
            if (days == 0) return date;
            while (days > 0) {
                date = date.AddDays(1);
                if (date.isHoliday()) {
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
