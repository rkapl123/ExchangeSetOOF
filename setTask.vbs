' For all enumeration documentation see - http://msdn.microsoft.com/en-us/library/windows/desktop/aa383602%28v=vs.85%29.aspx
Option Explicit

'  TASK_ACTION_TYPE
Const TASK_ACTION_EXEC = 0
Const TASK_ACTION_COM_HANDLER = 5
Const TASK_ACTION_SEND_EMAIL = 6
Const TASK_ACTION_SHOW_MESSAGE = 7

' TASK_COMPATIBILITY
Const TASK_COMPATIBILITY_AT = 0
Const TASK_COMPATIBILITY_V1 = 1
Const TASK_COMPATIBILITY_V2 = 2

' TASK_CREATION
Const TASK_VALIDATE_ONLY = &H01&
Const TASK_CREATE = &H02&
Const TASK_UPDATE = &H04&
Const TASK_CREATE_OR_UPDATE = &H06&
Const TASK_DISABLE = &H08&
Const TASK_DONT_ADD_PRINCIPAL_ACE = &H10&
Const TASK_IGNORE_REGISTRATION_TRIGGERS = &H20&

' TASK_INSTANCES_POLICY
Const TASK_INSTANCES_PARALLEL = 0
Const TASK_INSTANCES_QUEUE = 1
Const TASK_INSTANCES_IGNORE_NEW = 2
Const TASK_INSTANCES_STOP_EXISTING = 3

' TASK_LOGON_TYPE
Const TASK_LOGON_NONE = 0
Const TASK_LOGON_PASSWORD = 1
Const TASK_LOGON_S4U = 2
Const TASK_LOGON_INTERACTIVE_TOKEN = 3
Const TASK_LOGON_GROUP = 4
Const TASK_LOGON_SERVICE_ACCOUNT = 5
Const TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD = 6

' TASK_RUNLEVEL_TYPE
Const TASK_RUNLEVEL_LUA = 0
Const TASK_RUNLEVEL_HIGHEST = 1

' TASK_TRIGGER_TYPE2
Const TASK_TRIGGER_EVENT = 0
Const TASK_TRIGGER_TIME  = 1
Const TASK_TRIGGER_DAILY = 2
Const TASK_TRIGGER_WEEKLY = 3
Const TASK_TRIGGER_MONTHLY = 4
Const TASK_TRIGGER_MONTHLYDOW = 5
Const TASK_TRIGGER_IDLE = 6
Const TASK_TRIGGER_REGISTRATION = 7
Const TASK_TRIGGER_BOOT = 8
Const TASK_TRIGGER_LOGON = 9
Const TASK_TRIGGER_SESSION_STATE_CHANGE = 11
' -------------------------------------------------------------------------------

Dim objTaskService, objRootFolder, objTaskFolder, objNewTaskDefinition
Dim objTaskTrigger, objTaskAction, objTaskTriggers, blnFoundTask
Dim oFSO
Dim taskCommand

taskCommand = replace(Wscript.ScriptFullName, Wscript.ScriptName, "") & "ExchangeSetOOF.exe"
msgbox taskCommand

' Create the TaskService object and connect
Set objTaskService = CreateObject("Schedule.Service")
call objTaskService.Connect()

' Get the Root Folder where we will place this task
Set objTaskFolder = objTaskService.GetFolder("\")
dim xmlText
xmlText = "<?xml version=""1.0"" encoding=""UTF-16""?><Task version=""1.3"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task""><RegistrationInfo><Date>2015-09-18T21:21:38.8028563</Date><Author></Author></RegistrationInfo><Triggers><CalendarTrigger><Repetition><Interval>PT1H</Interval><Duration>PT10H</Duration><StopAtDurationEnd>false</StopAtDurationEnd></Repetition><StartBoundary>2015-09-18T07:00:00</StartBoundary><Enabled>true</Enabled><ScheduleByWeek><DaysOfWeek><Monday /><Tuesday /><Wednesday /><Thursday /><Friday /></DaysOfWeek><WeeksInterval>1</WeeksInterval></ScheduleByWeek></CalendarTrigger></Triggers><Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries><StopIfGoingOnBatteries>true</StopIfGoingOnBatteries><AllowHardTerminate>true</AllowHardTerminate><StartWhenAvailable>false</StartWhenAvailable><RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable><IdleSettings><StopOnIdleEnd>true</StopOnIdleEnd><RestartOnIdle>false</RestartOnIdle></IdleSettings><AllowStartOnDemand>true</AllowStartOnDemand><Enabled>true</Enabled><Hidden>false</Hidden><RunOnlyIfIdle>false</RunOnlyIfIdle><DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession><UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine><WakeToRun>false</WakeToRun><ExecutionTimeLimit>P3D</ExecutionTimeLimit><Priority>7</Priority></Settings><Actions Context=""Author""><Exec><Command>" & taskCommand & "</Command></Exec></Actions></Task>"
dim pTask, strUser, strPwd
strUser = CreateObject("WScript.Network").UserName
strPwd = GetPassword("Bitte Ihr Windows Passwort zum Starten des Tasks eingeben:" )
if strPwd = "" then Wscript.Quit
' set TASK_LOGON_PASSWORD if user can logon as a batch job (privilege)
objTaskFolder.RegisterTask "ExchangeSetOOF", xmlText, TASK_CREATE_OR_UPDATE, strUser, strPwd, TASK_LOGON_INTERACTIVE_TOKEN, pTask

Set oFSO = CreateObject("Scripting.FileSystemObject")
' Create temp folder, ignore error if already exists
on error resume next
oFSO.CreateFolder "C:\temp"
msgbox "Der Task wurde in der Aufgabenplanung eingetragen, nicht vergessen: Abwesenheitsnotiz anpassen!"


Function GetPassword( myPrompt )
dim objIE
    Set objIE = CreateObject( "InternetExplorer.Application" )
    objIE.Navigate "about:blank"
    objIE.Document.Title = "Password " & String( 100, "." )
    objIE.ToolBar        = False
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 400
    objIE.Height         = 220
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
    ' Insert the HTML code to prompt for a password
    objIE.Document.Body.InnerHTML = "<div align=""center""><p>" & myPrompt _
                                  & "</p><p><input type=""password"" size=""20"" " _
                                  & "id=""Password""></p><p><input type=" _
                                  & """hidden"" id=""OK"" name=""OK"" value=""0"">" _
                                  & "<input type=""submit"" value="" OK "" " _
                                  & "onclick=""VBScript:OK.Value=1""></p></div>"
    
    objIE.Document.Body.Style.overflow = "auto"
    objIE.Visible = True
    objIE.Document.All.Password.Focus

    On Error Resume Next
    Do While objIE.Document.All.OK.Value = 0
        WScript.Sleep 200
        If Err Then    'user clicked red X (or alt-F4) to close IE window
            IELogin = Array( "", "" )
            objIE.Quit
            Set objIE = Nothing
            Exit Function
        End if
    Loop
    On Error Goto 0

    ' Read the password from the dialog window
    GetPassword = objIE.Document.All.Password.Value

    ' Close and release the object
    objIE.Quit
    Set objIE = Nothing
End Function 