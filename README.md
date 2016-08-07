# ExchangeSetOOF
programmatic setting of automatic replies (out of office) in an exchange environment based on OOF appointments.

ExchangeSetOOF logs in to the currently logged in users account (using EWS AutoDiscover with users account Emailaddress using System.DirectoryServices.AccountManagement)
and searches the appointments between today and the next business day (based on only austrian holidays, this is currently hardcoded in function isHoliday) for appointments being set "away".

If any such appointment is found, ExchangeSetOOF replaces the template's date placeholder with the respective end date and (if wanted) also start date.
The languages used for the replacement of the date placeholders is hardcoded german and english (at the top of the program).
The automatic reply (out of office) is being scheduled to start from the Start Date of the OOF appointment and end on the End Date of the OOF appointment.

If no such appointment is found and the internal and external replies both contain a template specification (being "VORLAGE:", hardcoded at the top of the program),
the replies are stored in the registry as a template, which is being restored after the OOF period has passed.

The settings for templates can be changed in the code accordingly:

const string templateSpec = "VORLAGE:"; // prefix defining OOF reply bodies as a template, ALWAYS uppercase!  
public static readonly string[] DateLang1 = { "!DatumBis!", "!Datum!", "am", "von", "bis" };  
public static readonly string[] DateLang2 = { "!DateTo!", "!Date!", "on", "from", "until" };  

Placeholders are being replaced using following rule (hardcoded):  
DateLangX[0] changed to DateLangX[4] + " " + OOF_EndDate  
--> "bis/until dd.mm.yyyy", in case of time component in OOF_EndDate: "bis/until dd.mm.yyyy hh:mm:ss"   
DateLangX[1] changed to DateLangX[3] + " " + OOF_StartDate + " " + DateLangX[4] + " " + OOF_EndDate  
--> "von/from dd.mm.yyyy bis/until dd.mm.yyyy", in case of time component in OOF_EndDate/StartDate: "bis/until dd.mm.yyyy hh:mm:ss"  
in case of whole single day absences both DateLangX[0] and DateLangX[1] are being replaced by DateLangX[2] + " " + OOF_StartDate  
--> "am/on dd.mm.yyyy", there can be no time component for whole day absences!  

Install: copy ExchangeSetOOF.exe and both Managed EWS assemblies (Microsoft.Exchange.WebServices.Auth.dll
and Microsoft.Exchange.WebServices.dll) anywhere you want (no config required),  
and start on a regular basis (e.g. using task scheduler), execution hints are sent to stdout for problem determination.

Build: Download/clone repository to a folder named ExchangeSetOOF.  
To compile succesfully, you also need to download Managed EWS (used/tested version 2.2: https://www.microsoft.com/en-us/download/details.aspx?id=42951) and set references to Microsoft.Exchange.WebServices.Auth.dll
and Microsoft.Exchange.WebServices.dll accordingly.