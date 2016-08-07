# ExchangeSetOOF
programmatic setting of automatic replies (out of office) in an exchange environment based on OOF appointments.

ExchangeSetOOF logs in to the currently logged in users account (using EWS AutoDiscover with users account Emailaddress using System.DirectoryServices.AccountManagement)
and searches the appointments between today and the next business day (based on only Austrian Holidays, this is currently hardcoded) for appointments being set "away".

If any such appointment is found, ExchangeSetOOF replaces the template's date placeholder with the respective end date and (if wanted) also start date.
The languages used for the replacement of the date placeholders is hardcoded german and english (at the top of the program).

If no such appointment is found and the internal and external replies both contain a template specification (being "VORLAGE:", hardcoded at the top of the program),
the replies are stored in the registry as a template, which is being restored after the OOF period has passed.

The settings for templates can be changed in the code accordingly:

const string templateSpec = "VORLAGE:"; // prefix that defines OOF reply bodies to act as a template, ALWAYS uppercase (converted in code)
public static readonly string[] DateLang1 = { "!DatumBis!", "!Datum!", "am", "von", "bis" };
public static readonly string[] DateLang2 = { "!DateTo!", "!Date!", "on", "from", "until" };

Placeholders are being replaced using following rule (hardcoded):
!DatumBis!/!DateTo! changed to DateLang1/DateLang2[4] + " " + OOF_EndDate
!Datum!/!Date! changed to DateLang1/DateLang2[3] + " " + OOF_StartDate + " " + DateLang1/DateLang2[4] + " " + OOF_EndDate
in case of whole single day absences both !DatumBis/DateTo! and !Datum/Date! are being replaced by DateLang1/DateLang2[2] + " " + OOF_EndDate 


Installation: After compilation, copy ExchangeSetOOF.exe and both Managed EWS assemblies (Microsoft.Exchange.WebServices.Auth.dll and 
and Microsoft.Exchange.WebServices.dll) anywhere you want (no config required), 
and start on a regular basis (e.g. using task scheduler), execution hints are sent to stdout for problem determination.

Building: to compile succesfully, you need to download Managed EWS (used and tested here: 2.2) and set references to above assemblies accordingly.