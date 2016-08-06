using System;
using System.Xml.Linq;
using Microsoft.Exchange.WebServices.Data;
using System.DirectoryServices.AccountManagement;
using Microsoft.Win32;

namespace ExchangeSetOOF
{
    class Program
    {
        static ExchangeService service;
        const string templateSpec = "VORLAGE:"; // prefix that defines OOF bodies to act as a template, ALWAYS uppercase (converted in code)
        public static readonly string[] DateLang1 = { "<DatumBis>", "<Datum>", "am", "von", "bis" };
        public static readonly string[] DateLang2 = { "<DateTo>", "<Date>", "on", "from", "until" };

        static void Main(string[] args) {
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            service.UseDefaultCredentials = true;
            service.AutodiscoverUrl(UserPrincipal.Current.EmailAddress, RedirectionUrlValidationCallback);

            // find next "out of office" appointment being either today or on the next business day
            DateTime startDate = DateTime.Now;            
            startDate = DateTime.Parse("2016-08-03");
            DateTime endDate = startDate.AddBusinessDays(2); // need to add 2 days because otherwise the endDate is <nextBDate> 00:00:00
            // Initialize the calendar folder object with only the folder ID. 
            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
            // Set the start and end time and number of appointments to retrieve.
            CalendarView cView = new CalendarView(startDate, endDate, 20);
            // Limit the properties returned to the appointment's subject, start time, and end time.
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.LegacyFreeBusyStatus, AppointmentSchema.IsRecurring, AppointmentSchema.AppointmentType);
            // Retrieve a collection of appointments by using the calendar view.
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);
            Appointment oofAppointment = null;
            DateTime myStartOOFDate = new DateTime();
            DateTime myEndOOFDate = new DateTime();
            foreach (Appointment a in appointments) {
                if (a.LegacyFreeBusyStatus == LegacyFreeBusyStatus.OOF) {
                    // search for longest OOF appointment
                    if (oofAppointment == null || oofAppointment.End < a.End) {
                        Console.Write("oofAppointment "+a.Subject+" detected,Start: " + a.Start.ToString());
                        Console.Write(",(later)End: " + a.End.ToString());
                        Console.Write(",LegacyFreeBusyStatus: " + a.LegacyFreeBusyStatus.ToString());
                        Console.Write(",IsRecurring: " + a.IsRecurring.ToString());
                        Console.Write(",AppointmentType: " + a.AppointmentType.ToString());
                        Console.WriteLine();

                        oofAppointment = a;
                        myStartOOFDate = a.Start;
                        myEndOOFDate = a.End;
                    }
                }
            }

            // set automatic replies (out of office)
            OofSettings myOOF = service.GetUserOofSettings(UserPrincipal.Current.EmailAddress);
            // templates are stored in registry, if key doesn't exist, create it
            string keyName = "HKEY_CURRENT_USER\\Software\\RK\\ExchangeSetOOF";
            if (Registry.GetValue(keyName, "OOFtemplate", null) == null) {
                Registry.SetValue(keyName, "OOFtemplate", "", RegistryValueKind.String);
            }
            // OOF not set and templateSpec as prefix -> save as Template
            XDocument myReply = XDocument.Parse(myOOF.InternalReply.Message,);
            if (myOOF.State == OofState.Disabled && myOOF.InternalReply.ToString().ToUpper().StartsWith(templateSpec)) {
                Registry.SetValue(keyName, "OOFtemplate", myOOF.InternalReply.Message, RegistryValueKind.String);
                Console.WriteLine("template saved to registry");
            // OOF not set and templateSpec not given -> restore Template
            } else if (myOOF.State == OofState.Disabled) {
                myOOF.InternalReply.Message = Registry.GetValue(keyName, "OOFtemplate", "").ToString();
                Console.WriteLine("template restored from registry");
            }

            // out of office appointment today or on the next (business) day -> enable OOF
            if (!(oofAppointment == null) && myOOF.State == OofState.Disabled) {
                string replyText = Registry.GetValue(keyName, "OOFtemplate", "").ToString().Substring(templateSpec.Length); //cut prefix defined in templateSpec

                // replace template variables
                if (myEndOOFDate != myStartOOFDate) {
                    replyText.Replace("<DatumBis>", "bis " + myEndOOFDate.ToString());
                    replyText.Replace("<DateTo>", "until " + myEndOOFDate.ToString());
                    replyText.Replace("<Datum>", "von " + myStartOOFDate.ToString() + " bis " + myEndOOFDate.ToString());
                    replyText.Replace("<Date>", "from " + myStartOOFDate.ToString() + " until " + myEndOOFDate.ToString());
                } else {
                    // special case: exactly one day
                    replyText.Replace("<DatumBis>", "am " + myEndOOFDate.ToString());
                    replyText.Replace("<DateTo>", "on " + myEndOOFDate.ToString());
                    replyText.Replace("<Datum>", "am " + myEndOOFDate.ToString());
                    replyText.Replace("<Date>", "on " + myEndOOFDate.ToString());
                }

                // Set the OOF message for your internal audience.
                myOOF.InternalReply.Message = replyText;
                // Set the OOF message for your external audience.
                myOOF.ExternalReply.Message = replyText;
                // Set the OOF status to be a scheduled time period.
                myOOF.State = OofState.Scheduled;
                // Select the time period to be OOF.
                myOOF.Duration = new TimeWindow(myStartOOFDate, myEndOOFDate);
                // Select the external audience that will receive OOF messages.
                myOOF.ExternalAudience = OofExternalAudience.All;
                // Set the selected values. This method will result in a call to the Exchange server.
                service.SetUserOofSettings(UserPrincipal.Current.EmailAddress, myOOF);
                Console.WriteLine("oofUserSettings saved");
            } else if ((oofAppointment == null) && myOOF.State == OofState.Enabled) {
                // just in case exchange server didn't disable OOF automatically.
                myOOF.State = OofState.Disabled;
                service.SetUserOofSettings(UserPrincipal.Current.EmailAddress, myOOF);
            }
        }

        static void errorHandler(string errtext, Exception Ex, string ErrContext) {
            if (Ex != null) {
                errtext += " Exception: " + Ex.Message + ", at: " + Ex.StackTrace;
            }
            Console.WriteLine(ErrContext + ": " + errtext);
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
            // floating (EasterSunday dependent) austrian holidays
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
            //Console.WriteLine("{0}", DateTime.Parse("2012-04-05").isHoliday()); //falsch - donnerstag
            //Console.WriteLine("{0}", DateTime.Parse("2012-04-06").isHoliday()); //falsch - karfreitag
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
