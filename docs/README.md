# ExchangeSetOOF

ExchangeSetOOF logs in to the currently logged in users account (using EWS AutoDiscover with users account Email-address using System.DirectoryServices.AccountManagement)
and searches the appointments between today and the next business day (based on holidays that are configured in ExchangeSetOOF.exe.cfg) for appointments being set "away".

If any such appointment is found, ExchangeSetOOF replaces the template's date placeholder with the respective end date and (if wanted) also start date.
The languages used for the replacement of the date placeholders is hardcoded german and english (this can be configured in ExchangeSetOOF.exe.cfg, see below).
The automatic reply (out of office) is being scheduled to start from the Start Date of the absence appointment (`OOF_StartDate`) and end on the End Date of the OOF appointment (`OOF_EndDate`).

If no such appointment is found and the internal and external replies both contain a template specification (default being hardcoded to "VORLAGE:", this can be configured in ExchangeSetOOF.exe.cfg, see below),
the replies are stored in the registry as a template, which is being restored after the absence period has passed.

## Configuration (language placeholder and calendar):
Two language placeholder settings can be configured in the accompanying file ExchangeSetOOF.exe.cfg together for filling the templates with the actual dates:  

Settings in Line
1. template specification, a string that indicates the following OOF reply body acts as a template if given as a prefix, case insensitive (converted to uppercase in code)  
2. tab separated array of placeholders and replacements for DateLang1, being replaced using the rules described further below.  
3. same for DateLang2  
4. tab separated array of fixed holiday dates to be regarded when going forward for the last day of absence (weekends are also regarded if the holiday dates are not empty. In order to ONLY regard weekends, just fill any value in the holiday dates), in the format day.month  
5. if filled with EASTER then regard Easter Monday, Ascension day, Whit Monday, Corpus Christi day as holidays, if filled with EASTERGF then also regard Good Friday as a holiday.  

The default settings for above (if no cfg file is found) can be changed in the code:

```VB
string templateSpec = "VORLAGE:"; // prefix defining OOF reply bodies as a template, case insensitive as converted to uppercase in code!  
public static string[] DateLang1 = { "!BisDatum!", "!VonDatumBisDatum!", "am", "von", "bis", "true" };  
public static string[] DateLang2 = { "!ToDate!", "!FromDateToDate!", "on", "from", "until", "true" };  
```

Example config:
```VB
VORLAGE:
!BisDatum!	!VonDatumBisDatum!	am	von	bis	true
!ToDate!	!FromDateToDate!	on	from	until	true
1.1	6.1	1.5	15.8	26.10	1.11	8.12	24.12	25.12	26.12
EASTER
```

## Placeholders are being replaced using following rules (hardcoded):
```VB
Array Idx:	0		1			2	3	4	5
DateLang1:	!BisDatum!	!VonDatumBisDatum!	am	von	bis	true
DateLang2:	!ToDate!	!FromDateToDate!	on	from	until	true
```

The `OOF_EndDate` is set to the next business day after the absence in case `DateLangX[5]` (last parameter) is "true" thus allowing ".. until my return on dd/mm/yyyy".  
If this is not the case, `OOF_EndDate` is set to the first day of return.  

Any appearance of `DateLangX[0]` (`!ToDate!/!BisDatum!`) is changed to `DateLangX[4] + " " + OOF_EndDate`  
--> "until dd/mm/yyyy", in case of an additional time component in `OOF_EndDate`: "until dd/mm/yyyy hh:mm:ss"   

Any appearance of `DateLangX[1]` is changed to `DateLangX[3] + " " + OOF_StartDate + " " + DateLangX[4] + " " + OOF_EndDate`  
--> "from dd/mm/yyyy until dd/mm/yyyy", in case of an additional time component in `OOF_EndDate/StartDate`: "from dd/mm/yyyy hh:mm:ss until dd/mm/yyyy hh:mm:ss"  

in case of whole single day absences both `DateLangX[0]` and `DateLangX[1]` are being replaced by `DateLangX[2] + " " + OOF_StartDate`  
--> "on dd/mm/yyyy", there can be no time component for whole day absences!  

Examples for the above:  
"I'm on holiday !FromDateToDate!" would be changed to "I'm on holiday from 01/12/2005 until 05/12/2005" (for an absence that was found in the calendar from 01/12/2005 to 05/12/2005).
For a single day absence only on 01/12/2005 it would be "I'm on holiday on 01/12/2005".
"I'm on holiday !ToDate!" would be changed to "I'm on holiday until 05/12/2005" (for an absence that was found in the calendar from any day up to 05/12/2005).

If you want to completely control the from/to/until indicators in your template, this is just possible for the return date only (`!ToDate!/!BisDatum!`) setting. In this case remove both "on" and "until", then you'll get the date/time only.

The date format is using the current locale, so it might be different from "dd/mm/yyyy" or "dd/mm/yyyy hh:mm:ss" !

## Install
copy ExchangeSetOOF.exe (optionally ExchangeSetOOF.exe.cfg for different template specifications/replacements) and both Managed EWS assemblies (Microsoft.Exchange.WebServices.Auth.dll
and Microsoft.Exchange.WebServices.dll) anywhere you want and start on a regular basis (e.g. using task scheduler, the vb script "setTask.vbs" does this automatically), execution hints/exceptions are sent to ExchangeSetOOF.exe.log for problem determination.

## Build
Download/clone repository to a folder named ExchangeSetOOF.  
To compile successfully, you also need to download Managed EWS (ExchangeSetOOF was used/tested with version 2.2, can be downloaded from nuget: [https://www.nuget.org/packages/Microsoft.Exchange.WebServices/](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/)) and set references to Microsoft.Exchange.WebServices.Auth.dll
and Microsoft.Exchange.WebServices.dll accordingly.