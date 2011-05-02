'Windows Update Agent Script
'Created by Andras TIM @ 2010
'It based on http://community.spiceworks.com/scripts/show_download/82 from Rob Dunn

'RETURN CODE:
'0  All OK
'1  Errors / fail occured
'2  Aborted

'STDOUT:
'<numberOfUpdates: 0-... > <restartRequired: 0/1 >


'***********************************************************************************************************************' Declare
'***********************************************************************************************************************
Const scriptVersion = "1.2"

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const cdoSendUsingPort = 2
Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const cdoSMTPServerport = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const cdoSMTPconnectiontimeout = "http://schemas.microsoft.com/cdo/configuration/Connectiontimeout"

'Public Objects
Dim wshshell, wshsysenv, fso, objADInfo, updateAgentSession, autoUpdateClient, searchResult, logFile, objWMIService, objReg
'Settings
Dim update_criteria, wuaInstallerPath, wuLogPath, wuErrorlistPath, logfilePath, strWUAgentVersion, strLocaleVerDelim, boolEmailReportEnabled, boolEmailIfAllOK, boolFullDNSName, iTimeFormat
Dim strMailFrom, strMail_to, strMail_subject, strMail_smtpHost, strMail_smtpPort, strMail_smtpAuthType, strMail_smtpAuthID, strMail_smtpAuthPassword

'State variables
Dim boolCScript, regWSUSServer, boolUpdatesInstalled, boolRebootRequired
Dim statInProgress, statInstalled, statCompleteWithErrors, statFailed, statAborted, intLinesBefore
Dim lastTryedHotfix, wuErrorlist

'*********************************************************************************************************************** PreInit
'***********************************************************************************************************************
'Get instances
Set wshshell = wscript.CreateObject("WScript.Shell")
Set wshsysenv = wshshell.Environment("PROCESS")
Set fso = CreateObject("Scripting.FileSystemObject")
Set objADInfo = CreateObject("ADSystemInfo")

'Get authentication information
strDomain = wshsysenv("userdomain")
strUser = wshsysenv("username")
strComputer = wshshell.ExpandEnvironmentStrings("%Computername%")

'Get other instances
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'*********************************************************************************************************************** User variables
'***********************************************************************************************************************
'Updates filter for updating
'update_criteria = "IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software' or Type='Driver'"
update_criteria = "IsInstalled=0 and DeploymentAction='Installation' or IsPresent=1 and DeploymentAction='Uninstallation' or IsInstalled=1 and DeploymentAction='Installation' and RebootRequired=1 or IsInstalled=0 and DeploymentAction='Uninstallation' and RebootRequired=1"

'Full EXE path to Windows Update Agent installation exe. It will install it slently if the PC needs it
wuaInstallerPath = """\\FIXME.SERVER\SHARE\WindowsUpdate\WindowsUpdateAgent30-x86.exe"""

'Windows Update log file path
wuLogPath = wshsysenv("WINDIR") & "\WindowsUpdate.log"

'Windows Update's error description file
wuErrorlistPath = "wua-errorlist.csv"

'Logfile path
'logfilePath = wshsysenv("SYSTEMDRIVE") & "\" & "wua-last.log"
logfilePath = "wua-last.log"

'Mail settings
strMailFrom = LCase(strComputer)
strMail_to = "FIXME@EMAIL"
strMail_subject = "WUA Script - WSUS Update log file from" 'computer name
strMail_smtpHost = "FIXME.SMTP.SERVER"
strMail_smtpPort = 25
'set your SMTP server authentication type.
' Possible values:cdoAnonymous|cdoBasic|cdoNTLM
' You do not need to configure an id/pass combo with cdoAnonymous
strMail_smtpAuthType = "cdoAnonymous"
strMail_smtpAuthID = ""
strMail_smtpAuthPassword = ""

'Version number of the Windows Update agent you wish to compare installed
' version against.  If the version installed is not equal to this version, then
' it will install the exe referred to in var 'wuaInstallerPath' above.
strWUAgentVersion = "7.0.6000.374"
'version 2.0 SP1 is 5.8.0.2469
'version 3.0 is 7.0.6000.374
strLocaleVerDelim = "."

'Turns email function on/off.  If an email address is specified in the
' command-line arguments, then this will automatically turn on ('1').
' False = off, don't email
' True = on, email using default address defined in the var 'strMail_to' above.
boolEmailReportEnabled = True

'boolEmailIfAllOK Determines if email always sent or only if updates or reboot
' needed.
' False = off, don't send email if no updates needed and no reboot needed
' True = on always send email
boolEmailIfAllOK = True

'boolFullDNSName Determines if the email subject contains the full dns name of
' the server or just the computer name.
' False = off, just use computer name
' True = on,  use full dns name
boolFullDNSName = False

'Timestamp format in log
' 0 = vbGeneralDate - Default. Returns time: hh:mm:ss PM/AM.
' 3 = vbLongTime - Returns time: hh:mm:ss PM/AM
' 4 = vbShortTime - Return time: hh:mm
iTimeFormat = 4

'*********************************************************************************************************************** Init
'***********************************************************************************************************************

'Try to pick up computername via AD'
strComputer1 = objADInfo.ComputerName
'Get computer OU
On Error Resume Next
strOU = "Computer OU: Not detected"
Set objComputer = GetObject("LDAP://" & strComputer1)
If objComputer.Parent <> "" Then strOU = "Computer OU: " & Replace(objComputer.Parent, "LDAP://", "")
On Error GoTo 0

'Open logfile
Set logFile = fso.OpenTextFile(logfilePath, ForAppending, True)

'Print header (version info)
print_debug "ScriptInit", "Windows Update Agent Script v" & scriptVersion

'Check the start environment
boolCScript = (InStr(UCase(wscript.FullName), "\CSCRIPT.EXE") > 0)
If Not boolCScript Then
    print_debug "ScriptInit", "WARNING: Use the ""cscript.exe //nologo"" command console output"
End If

'Print data
print_debug "ScriptInit", ">>> Environment info <<<" & vbCrLf & _
    "Command: " & wscript.FullName & vbCrLf & _
    "Computer: " & strComputer & vbCrLf & _
    strOU & vbCrLf & _
    "Executed by: " & strDomain & "\" & strUser

'Reset counters
statInProgress = 0
statInstalled = 0
statCompleteWithErrors = 0
statFailed = 0
statAborted = 0

'Store start time
print_debug "ScriptMain", "Script started"

'Init the applyed hotfix list
lastTryedHotfix = ""

'Init Windows Update's errorlist
wuErrorlist = ""

'*********************************************************************************************************************** Common functions
'***********************************************************************************************************************

Sub print(strMsg)
    aMsg = strMsg
    If (Right(aMsg, 2) = vbCrLf) Then aMsg = Left(strMsg, Len(strMsg) - 2)
    print_debug "STDOUT", aMsg
    If boolCScript Then
        wscript.echo aMsg
    Else
        MsgBox strMsg, vbOKOnly
    End If
End Sub

'***********************************************************************************************************************
Sub print_debug(strObj, strMsg)
    aMsg = strMsg
    aTime = FormatDateTime(Time, iTimeFormat)
    If (Right(aMsg, 2) = vbCrLf) Then aMsg = Left(strMsg, Len(strMsg) - 2)
    aMsg = Replace(aMsg, vbCrLf, vbCrLf & vbTab & vbTab)
    logFile.writeline "[" & aTime & "] " & strObj & vbTab & aMsg
End Sub

'***********************************************************************************************************************
Function getLineNumber(strText)
    strObjID = "getLineNumber"
    ret = 0

    If strText <> "" Then
        'Append a line end > the UBound-1 equal with the lines
        arrText = Split(Replace(strText & vbLf, vbCrLf, vbLf), vbLf)
        ret = UBound(arrText)
    End If

    getLineNumber = ret
End Function

'***********************************************************************************************************************
Function getLineRange(strText,intStart,intEnd)
    strObjID = "getLineRange"
    ret = ""

    If strText <> "" Then
        'Append a line end > the UBound-1 equal with the lines
        arrText = Split(Replace(strText & vbLf, vbCrLf, vbLf), vbLf)

        i = intStart
        Do Until i < LBound(arrText) Or i > UBound(arrText) Or i > IntEnd
            ret = ret & arrText(i) & vbCrLf
            i = i +1
        Loop
    End If

    getLineRange = ret
End Function

'***********************************************************************************************************************
Function findLine(strText,strFind)
    strObjID = "findLine"
    ret = ""

    i = InStr(1, strText, strFind)
    If i > 0 Then
        j = InStr(i + 1, strText, vbCrLf, vbTextCompare)
        If j < i Then j = Len(strText) + 1

        ret = Mid(strText, i, j - i)
    End If

    findLine = ret
End Function

'***********************************************************************************************************************
Sub commonErrorHandler(strObjID, errNum, errDesc, boolFatal)
    print_debug strObjID, "Error 0x" & Hex(errNum) & " has occured.  Description: " & errDesc
    If boolFatal Then exitScript 1
End Sub

'***********************************************************************************************************************
Sub exitScript(intErrCode)
    print_debug "ScriptMain", "Script ended"
    logFile.Close
    wscript.quit intErrCode
End Sub

'***********************************************************************************************************************
Function sendMail(strFrom, strTo, strMail_subject, strMessage)
    strObjID = "sendMail"
    Dim iMsg, Flds

    print_debug strObjID, ">>> Calling sendMail routine <<<" & vbCrLf & _
        "To: " & strMail_to & vbCrLf & _
        "From: " & strMailFrom & vbCrLf & _
        "Subject: " & strMail_subject & vbCrLf & _
        "SMTP Server: " & strMail_smtpHost

    '//  Create the CDO connections.
    On Error Resume Next
    Set iMsg = CreateObject("CDO.Message")
    Set Flds = iMsg.Configuration.Fields

    'Error handles
    If Err.Number <> 0 Then
        en = Err.Number: ed = "Init error, the mail will be not send!' (" & Err.Description & ")"
        On Error GoTo 0
        commonErrorHandler strObjID, en, ed, False
        Exit Function
    End If
    On Error GoTo 0

    With Flds
        '// SMTP protocol init
        If LCase(strMail_smtpAuthType) <> "cdoanonymous" Then
            'Type of authentication, NONE, Basic (Base64 encoded), NTLM
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = strMail_smtpAuthType

            'Your UserID on the SMTP server
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strMail_smtpAuthID

            'Your password on the SMTP server
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strMail_smtpAuthPassword
        End If

        '// SMTP server configuration.
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = strMail_smtpHost
        .Item(cdoSMTPServerport) = strMail_smtpPort
        .Item(cdoSMTPconnectiontimeout) = 60
        .Update
    End With

    Dim r
    Set r = fso.OpenTextFile(logfilePath, ForReading, False, TristateUseDefault)
    strMessage = "<pre>" & r.readall & "</pre>"

    '//  Set the message properties.
    With iMsg
        .To = strMail_to
        .From = strMailFrom
        .Subject = strMail_subject
    End With

    'iMsg.AddAttachment wsuslog
    strMessage = Replace(strMessage, vbTab, "&#09;")
    iMsg.HTMLBody = Replace(strMessage, vbNewLine, "<br>")
    '//  Send the message.

    On Error Resume Next
    iMsg.Send ' send the message.
    Set iMsg = Nothing

    'Error handles
    If Err.Number <> 0 Then
        en = Err.Number: ed = "Problem sending mail to '" & strMail_smtpHost & "' (" & Err.Description & ")"
        On Error GoTo 0
        commonErrorHandler strObjID, en, ed, False
    Else
        On Error GoTo 0
        print_debug strObjID, "The email has been sent to " & strMail_smtpHost
    End If

End Function

'***********************************************************************************************************************
Sub endOfScript()
    If boolRebootRequired Then
        intRebootReq = 1
    Else
        intRebootReq = 0
    End If

    Print statInstalled & " " & intRebootReq

    If statCompleteWithErrors > 0 Or statFailed > 0 Then
        'RETURN: Errors / fail occured
        exitScript 1
    ElseIf statAborted > 0 Then
        'RETURN: Aborted
        exitScript 2
    End If

    'RETURN: All ok
    exitScript 0
End Sub

'*********************************************************************************************************************** Service functions
'***********************************************************************************************************************

Function serviceGetState(strService)
    strObjID = "serviceGetState"
    Dim colServiceList, objService, ret
    On Error Goto 0

    'Filtering for service
    Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service where Name='" & strService & "'")
    If Not colServiceList.Count = 1 Then
        ret = "Bad results for get service: " & strService
    Else
        For Each objService In colServiceList
            ret = objService.State
        Next
    End If
    serviceGetState = ret
End Function

'***********************************************************************************************************************
Function serviceStart(strService)
    strObjID = "serviceStart"
    Dim colServiceList, objService, intTimeout, strState, ret
    On Error Goto 0

    'Filtering for service
    Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service where Name='" & strService & "'")
    If Not colServiceList.Count = 1 Then
        print_debug strObjID, "Bad results for get service: " & strService
        serviceStart = False
        Exit Function
    End If

    'Control the result services
    ret = True
    intTimeout=30'sec
    'Get Service
    For Each objService In colServiceList
        print_debug strObjID, "Starting service: " & strService & " (" & objService.DisplayName & ")"
        'Start service
        errReturn = objService.StartService()
        If errReturn = 0 Then
            Do
                Wscript.Sleep 1000
                strState = serviceGetState(strService)
                print_debug strObjID, "State (" & intTimeout & "): " & strState
                intTimeout = intTimeout -1
            Loop Until strState = "Running" Or intTimeout = 0

            If intTimeout = 0 Then
                print_debug strObjID, "Starting service timeout: " & strService & " (" & objService.DisplayName & ")"
                ret = False
            End If
        Else
            print_debug strObjID, "Starting service error: " & strService & " (" & objService.DisplayName & "); return: " & errReturn
        End If
    Next

    serviceStart = ret
End Function

'***********************************************************************************************************************
Function serviceStop(strService)
    strObjID = strObjID & "serviceStop"
    Dim colServiceList, objService, intTimeout, strState, ret
    On Error Goto 0

    'Filtering for service
    Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service where Name='" & strService & "'")
    If Not colServiceList.Count = 1 Then
        print_debug strObjID, "Bad results for get service: " & strService
        serviceStop = False
        Exit Function
    End If

    'Control the result services
    ret = True
    intTimeout=30'sec
    'Get Service
    For Each objService In colServiceList
        print_debug strObjID, "Stoping service: " & strService & " (" & objService.DisplayName & ")"
        'Stop service
        errReturn = objService.StopService()
        If errReturn = 0 Then
            Do
                Wscript.Sleep 1000
                strState = serviceGetState(strService)
                print_debug strObjID, "State (" & intTimeout & "): " & strState
                intTimeout = intTimeout -1
            Loop Until strState = "Stopped" Or intTimeout = 0

            If intTimeout = 0 Then
                print_debug strObjID, "Stoping service timeout: " & strService & " (" & objService.DisplayName & ")"
                ret = False
            End If
        Else
            print_debug strObjID, "Stoping service error: " & strService & " (" & objService.DisplayName & "); return: " & errReturn
        End If
    Next

    serviceStop = ret
End Function

'*********************************************************************************************************************** FS functions
'***********************************************************************************************************************

Function delSubItems(strPath)
    strObjID = "delSubItems"
    Dim ret, objFolder, objItem
    ret = False

    'Get object
    print_debug strObjID, "Delete contents in: " & strPath
    Set objFolder = fso.GetFolder(strPath)

    On Error Resume Next
    'Folders
    For Each objItem in objFolder.SubFolders
        objItem.Delete
    Next
    'Files
    For Each objItem in objFolder.Files
        objItem.Delete
    Next

    'Complete?
    If Err.Number = 0 Then
        If objFolder.SubFolders.Count = 0 And objFolder.Files.Count = 0 Then
            print_debug strObjID, "Delete completed"
            ret = True
        Else
            print_debug strObjID, "Delete semicompleted"
        End if
    Else
        print_debug strObjID, "Error occured on delete"
    End If
    On Error Goto 0

    delSubItems = ret
End Function


'***********************************************************************************************************************
Function getFileToText(fn)
    strObjID = "getFileToText"
    ret = ""

    'Open file, read and close
    On Error Resume Next
    Set objReadFile = fso.OpenTextFile(fn, ForReading)
    ret = objReadFile.ReadAll
    objReadFile.Close

    'Error handles
    If Err.Number <> 0 Then
        en = Err.Number: ed = "Can't read the  '" & fn & "' file (" & Err.Description & ")"
        On Error GoTo 0
        commonErrorHandler strObjID, en, ed, False
    End If
    On Error GoTo 0

    'Return the file contents
    getFileToText = ret
End Function

'*********************************************************************************************************************** Run functions
'***********************************************************************************************************************

Function runCommand(strCmd)
    strObjID = "runCommand"
    dim ret

    'Run command
    print_debug strObjID, "Run command: " & strCmd
    ret = wshshell.Run(strCmd, 0, true)

    'Return
    print_debug strObjID, "Return code: " & ret
    runCommand = ret
End Function

'*********************************************************************************************************************** WUA functions
'***********************************************************************************************************************

Sub chkMailSets()
    strObjID = "chkMailSets"
    If boolEmailReportEnabled = False Then Exit Sub

    If LCase(strMail_smtpAuthType) <> "cdoanonymous" Then
        If strMail_smtpAuthType = "" Then
            strMail_smtpAuthType = "cdoAnonymous"
        Else
            print_debug strObjID, "SMTP Auth User ID: " & sAuthID
            If SMTPUserID = "" Then
                print_debug strObjID, "No SMTP user ID was specified, even though SMTP Authentication was configured " & _
                    "for " & strMail_smtpAuthType & "." & vbCrLf & "Attempting to switch to anonymous authentication..."
                strMail_smtpAuthType = "cdoAnonymous"
                If strMail_smtpAuthPassword <> "" Then
                    print_debug strObjID, "You have specified a SMTP password, but no user ID has been configured for " & _
                        "authentication. Check the INI file (" & sINI & ") again and re-run the script."
                End If
            Else
                If strMail_smtpAuthPassword = "" Then
                    print_debug strObjID, "You have specified a SMTP user ID, but have not specified a password." & _
                        vbCrLf & "Switching to anonymous authentication."
                End If
                strMail_smtpAuthType = "cdoAnonymous"
            End If
            If strMail_smtpAuthPassword <> "" Then print_debug strObjID, "SMTP password configured, but hidden..."
        End If
    End If
    print_debug strObjID, "SMTP Authentication type: " & strMail_smtpAuthType
End Sub

'***********************************************************************************************************************
Function chkAgentVer()
    strObjID = "chkAgentVer"

    'Check Service State
    If Not serviceStart("wuauserv") Then
        en = 0: ed = "Can't start the 'wuauserv' service"
        commonErrorHandler strObjID, en, ed, True
    End If

    On Error Resume Next
    bUpdateNeeded = True ' init value
    print_debug strObjID, "Checking version of Windows Update agent against version " & strWUAgentVersion & "..."
    Set updateAgentSession = CreateObject("Microsoft.Update.AgentInfo")
    If Err.Number = 0 Then
        updateinfo = updateAgentSession.GetInfo("ProductVersionString")
        If Replace(updateinfo, strLocaleVerDelim, "") = Replace(strWUAgentVersion, strLocaleVerDelim, "") Then
            bUpdateNeeded = False
        ElseIf updateinfo > strWUAgentVersion Then
            print_debug strObjID, "Your installed version of the Windows Update Agent (" & updateinfo & ") is " & _
                "newer than the referenced version (" & strWUAgentVersion & ")."
            bUpdateNeeded = False
        End If
    End If
    On Error Goto 0

    If bUpdateNeeded Then
        print_debug strObjID, "File version (" & updateinfo & ") does not match, WUA udapte required."

        'stop the Automatic Updates service
        If Not serviceStop("wuauserv") Then
            en = 0: ed = "Can't stop the 'wuauserv' service"
            commonErrorHandler strObjID, en, ed, True
        End If

        'Install the newer WUA
        On Error Resume Next
        Set oEnv = wshshell.Environment("PROCESS")
        oEnv("SEE_MASK_NOZONECHECKS") = 1
        rCmd = wuaInstallerPath & " /quiet /norestart"
        WriteLog ("Attempting to install WUA: " & rCmd)
        wshshell.Run rCmd, 1, True

        'Error handles
        If Err.Number <> 0 Then
            en = Err.Number: ed = "Error executing '" & wuaInstallerPath & "' Agent EXE (" & Err.Description & ")"
            On Error GoTo 0
            commonErrorHandler strObjID, en, ed, False
        End If
        On Error GoTo 0
    End If

    'All OK
    chkAgentVer = True
End Function

'***********************************************************************************************************************
Function chkAgentSets()
    strObjID = "chkAgentSets"
    strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"

    Set autoUpdateClient = CreateObject("Microsoft.Update.AutoUpdate")

    'Server
    strValueName = "WUServer"
    objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, regWSUSServer
    If IsNull(regWSUSServer) Or Trim(regWSUSServer) = "" Then regWSUSServer = "Microsoft Windows Update"

    'Scheduled
    retScheduled = chkAgentSet_getSchedule

    'TargetGroup
    retTargetGroup = chkAgentSet_getTargetGroup(strKeyPath)

    'NotificationLevel
    retWUAmode = chkAgentSet_getretWUAmode

    'Debug print
    print_debug strObjID, ">>> WUA Settings <<<" & vbCrLf & _
        "WUA mode: " & retWUAmode  & vbCrLf & _
        "Server: " & regWSUSServer & vbCrLf & _
        "Scheduled: " & retScheduled & vbCrLf & _
        "TargetGroup: " & retTargetGroup

    'All OK
    chkAgentSets = True
End Function

'***********************************************************************************************************************
Function chkAgentSet_getSchedule()
    Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
    Set objSettings = objAutoUpdate.Settings

    Select Case objSettings.ScheduledInstallationDay
        Case 0:     strDay = "every day"
        Case 1:     strDay = "sunday"
        Case 2:     strDay = "monday"
        Case 3:     strDay = "tuesday"
        Case 4:     strDay = "wednesday"
        Case 5:     strDay = "thursday"
        Case 6:     strDay = "friday"
        Case 7:     strDay = "saturday"
        Case Else:  strDay = "The scheduled installation day is could not be determined."
    End Select

    If Len(objSettings.ScheduledInstallationTime) = 1 Then
        strScheduledTime  = "0" & objSettings.ScheduledInstallationTime
    Else
        strScheduledTime  = objSettings.ScheduledInstallationTime
    End If
    chkAgentSet_getSchedule = strDay & " at " & strScheduledTime & ":00"
End Function

'***********************************************************************************************************************
Function chkAgentSet_getTargetGroup(strKeyPath)
    strValueName = "TargetGroup"
    ret = "Not specified"
    objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, regTargetGroup
    If regTargetGroup <> "" And Not IsNull(regTargetGroup) Then ret = regTargetGroup
    chkAgentSet_getTargetGroup = ret
End Function

'***********************************************************************************************************************
Function chkAgentSet_getretWUAmode()
    Set updateinfo = autoUpdateClient.Settings

    ret = "-"
    Select Case updateinfo.notificationlevel
        Case 0: ret = "WU agent is not configured."
        Case 1: ret = "WU agent is DISABLED."
        Case 2: ret = "Users are prompted to approve updates prior to installing"
        Case 3: ret = "Updates are downloaded automatically, and users are prompted to install."
        Case 4: ret = "Updates are downloaded and installed automatically at a pre-determined time."
    End Select
    chkAgentSet_getretWUAmode = ret
End Function

'***********************************************************************************************************************
Function updateSearcher()
    strObjID = "updateSearcher"

    On Error Resume Next
    Set updateAgentSession = CreateObject("Microsoft.Update.Session")
    Set updateSearcher = updateAgentSession.CreateupdateSearcher()

    print_debug strObjID, "Filtering updates: " & update_criteria
    Set searchResult = updateSearcher.Search(update_criteria)

    'Handle some common errors here
    If Err.Number <> 0 Then
        en = Err.Number: ed = Err.Description
        On Error GoTo 0
        If Not wuaErrorHandler(strObjID, en, ed, True) Then
            updateSearcher = False
            Exit Function
        End If
    End If
    On Error GoTo 0

    chkReboot ("beginning")
    updateSearcher = True
End Function

'***********************************************************************************************************************' be kell kotni oket
Function wuaGetErrorDescription(errNum)' Array :: [""] if we don't know description; ["ID", "Desc"] if we know the error
    hexErrNum = "0x" & UCase(Hex(errNum))
    ret = Null

    'Init errorlist
    If wuErrorlist = "" Then
        wuErrorlist = getFileToText(wuErrorlistPath)
    End If

    'Exist errorlist
    If Not wuErrorlist = "" Then
        txt = findLine(wuErrorlist, vbCrLf & hexErrNum & vbTab)
        If Not txt = "" Then
            arr = Split(Right(txt, Len(txt) - 2), vbTab)
            ReDim Preserve arr(2)
            ret = Array(arr(1), arr(2))
        End If
    End If

    wuaGetErrorDescription = ret
End Function

'***********************************************************************************************************************' be kell kotni oket
Function wuaErrorHandler(strObjID, errNum, errDesc, ifUnhandledBeFatal)' Boolean :: true=all ok; false=check
    strObjID = "wuaErrorHandler"

    'Check fatal
    Select Case "0x" & UCase(Hex(errNum))
        Case "0x80072F78", "0x80072EFD", "0x8024002B", "0x7", "0x8024400D", "0x8024A000", "0x80072F8F"
            boolFatal = True
        Case Else
            boolFatal = ifUnhandledBeFatal
    End Select

    'Get error description
    arrError = wuaGetErrorDescription(errNum)
    If IsNull(arrError) Then arrError = Array("-","Unknown error")

    'Print error
    en = errNum: ed = arrError(0) & vbCrLf & "Details: " & arrError(1)
    commonErrorHandler strObjID, en, ed, False

    'Try hotfix it
    print_debug strObjID, "Checking hotfixes for 0x" & Hex(errNum) & " update error..."
    res = errorHotfixes(errNum)
    If Not res Then
        en = errNum: ed = "Could not find solution for previous error"
        commonErrorHandler strObjID, en, ed, boolFatal
    Else
        print_debug strObjID, "Hotfix applied, restart the update process"
        wuaErrorHandler = False
        Exit Function
    End if

    wuaErrorHandler = True
End Function

'***********************************************************************************************************************' be kell kotni oket
Function errorHotfixes(errNum)'boolean :: true if we have hotfix for it
    strObjID = "errorHotfixes"
    hexErrNum = "0x" & UCase(Hex(errNum))

    'Check the last applyed hotfix ID (for recursive hotfixapply check)
    If lastTryedHotfix = hexErrNum Then
        en = errNum: ed = "The known hotfixes could not fix the problem!"
        commonErrorHandler strObjID, en, ed, False
        errorHotfixes = False
        Exit Function
    End If

    Select Case hexErrNum
        Case "0x8024400D"
            ret = True
            ret = ret Or serviceStop("wuauserv")
            ret = ret Or delSubItems("C:\WINDOWS\SoftwareDistribution\DataStore")
            ret = ret Or delSubItems("C:\Windows\SoftwareDistribution\Download")
            ret = ret Or runCommand("reg delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"" /v SusClientId /f")
            ret = ret Or runCommand("reg delete ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"" /v SusClientIdValidation /f")
            ret = ret Or serviceStart("wuauserv")

        Case "0x8024A000"
            ret = True
            ret = ret Or serviceStop("wuauserv")
            ret = ret Or runCommand("reg add ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"" /v AUOptions /t REG_DWORD /d 2 /f") 'Set: Auto check, never donload
            ret = ret Or serviceStart("wuauserv")

        Case "0x80072F8F"
            ret = True
            ret = ret Or runCommand("regsvr32 /s Mssip32.dll")
            ret = ret Or runCommand("regsvr32 /s Initpki.dll")

        Case Else
            ret = False'Hotfix not fpund
    End Select

    'Store the last applyed hotfix ID (for recursive hotfixapply check)
    If ret Then lastTryedHotfix = hexErrNum

    errorHotfixes = ret
End Function

'***********************************************************************************************************************
Sub chkReboot(beforeorafter)
    strObjID = "chkReboot"
    Set ComputerStatus = CreateObject("Microsoft.Update.SystemInfo")
    Select Case beforeorafter
        Case "beginning": strCheck = "Pre-check"
        Case "end":       strCheck = "Post-check"
    End Select

    boolRebootRequired = ComputerStatus.RebootRequired
    If boolRebootRequired Then
        strMsg = "Computer has a pending reboot (" & strCheck & ")"
    Else
        strMsg = "Computer does not have any pending reboots (" & strCheck & ")."
    End If
    If strMsg <> "" Then print_debug strObjID, strMsg
   'wscript.sleep 4000
End Sub

'***********************************************************************************************************************
Function detectNow()
    strObjID = "DetectNow"

    'SEARCHING UPDATES
    print_debug strObjID, "Searching updates..."
    On Error Resume Next
    autoUpdateClient.detectnow()

    'Handle some common errors here
    If Err.Number <> 0 Then
        en = Err.Number: ed = Err.Description
        On Error GoTo 0
        If Not wuaErrorHandler(strObjID, en, ed, True) Then
            detectNow = False
            Exit Function
        End If
    End If
    On Error GoTo 0

    'LIST NEW UPDATES
    strUpdates = ""
    For i = 0 To searchResult.updates.Count - 1
        Set Update = searchResult.updates.Item(i)
        strUpdates = strUpdates & Update.Title & vbCrLf
    Next
    print_debug strObjID, ">>> Required updates (" & searchResult.updates.Count & ") <<< " & vbCrLf & strUpdates

    'All OK
    detectNow = True
End Function

'***********************************************************************************************************************
Function dwlUpdates()
    strObjID = "dwlUpdates"

    'CATALOGING
    print_debug strObjID, "Cataloging updates..."
    Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")
    For i = 0 To searchResult.updates.Count - 1
        Set Update = searchResult.updates.Item(i)
        If Not Update.EulaAccepted Then Update.AcceptEula
        updatesToDownload.Add Update
    Next

    'DOWNLOADING
    print_debug strObjID, "Downloading updates..."
    On Error Resume Next
    Set downloader = updateAgentSession.CreateUpdateDownloader()
    downloader.updates = updatesToDownload
    downloader.Download()

    'Error handles
    If Err.Number <> 0 Then
        en = Err.Number: ed = Err.Description
        On Error GoTo 0
        commonErrorHandler strObjID, en, ed, False
    End If
    On Error GoTo 0

    'All OK
    dwlUpdates = True
End Function

'***********************************************************************************************************************
Function instUpdates()
    strObjID = "instUpdates"

    'COLLECTING
    print_debug strObjID, "Creating collection of updates needed to install..."
    Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
    For i = 0 To searchResult.updates.Count - 1
        Set Update = searchResult.updates.Item(i)
        If Update.IsDownloaded Then updatesToInstall.Add Update
    Next

    'INSTALLER INIT
    print_debug strObjID, "Installing updates..."
    On Error Resume Next
    Set installer = updateAgentSession.CreateUpdateInstaller()
    installer.AllowSourcePrompts = False
    installer.ForceQuiet = True
    If Err.Number <> 0 Then WriteLog ("Error " & Err.Number & " has occured.  Error description: " & Err.Description)
    On Error GoTo 0

    'INSTALLING
    installer.updates = updatesToInstall
    boolUpdatesInstalled = True
    On Error Resume Next
    Set installationResult = installer.Install()

    'Error handles
    If Err.Number <> 0 Then
        en = Err.Number: ed = "Error installing updates!' (" & Err.Description & ")"
        On Error GoTo 0
        commonErrorHandler strObjID, en, ed, False
    End If
    On Error GoTo 0

    'RESULT
    strUpdates = ""
    For i = 0 To updatesToInstall.Count - 1
        Select Case installationResult.GetUpdateResult(i).ResultCode
            Case 1: strResult = "In progress                        ": statInProgress = statInProgress + 1
            Case 2: strResult = "Installed                          ": statInstalled = statInstalled + 1
            Case 3: strResult = "Operation complete, but with errors": statCompleteWithErrors = statCompleteWithErrors + 1
            Case 4: strResult = "Operation failed                   ": statFailed = statFailed + 1
            Case 5: strResult = "Operation aborted                  ": statAborted = statAborted + 1
        End Select
        strUpdates = strUpdates & strResult & " : " & updatesToInstall.Item(i).Title & vbCrLf
    Next
    print_debug strObjID, ">>> Installation results (" & updatesToInstall.Count & " updates) <<<" & vbCrLf & strUpdates

    print_debug strObjID, ">>> Installation Summary <<<" & vbCrLf & _
        "Result: " & installationResult.ResultCode & vbCrLf & _
        "Reboot Required: " & installationResult.RebootRequired & vbCrLf & _
        "In progress: " & statInProgress & vbCrLf & _
        "Installed: " & statInstalled & vbCrLf & _
        "Operation complete, but with errors: " & statCompleteWithErrors & vbCrLf & _
        "Operation failed: " & statFailed & vbCrLf & _
        "Operation aborted: " & statAborted

    'All OK
    instUpdates = True
End Function

'***********************************************************************************************************************
Sub sendReport()
    strObjID = "sendReport"

    If boolUpdatesInstalled Then chkReboot "end"
    If boolEmailReportEnabled Then
        If searchResult.updates.Count = 0 And Not boolRebootRequired And boolEmailIfAllOK = False Then
            print_debug strObjID, "No updates required, no pending reboot, therefore not sending email"
        Else
            If boolFullDNSName Then
                StrDomainName = wshshell.ExpandEnvironmentStrings("%USERDNSDOMAIN%")
                strOutputComputerName = strComputer & "." & StrDomainName
            Else
                strOutputComputerName = strComputer
            End If
            If emailifallok = 0 Or emailifallok = 1 Then
                If Not InStr(strMail_smtpHost, "x") Then
                    sendMail strFrom, strTo, strMail_subject & " " & strOutputComputerName, strMessage
                End If
            End If
        End If
    End If
End Sub

'***********************************************************************************************************************
Sub initUpdateLog()
    strObjID = "initUpdateLog"

    strLog = getFileToText(wuLogPath)
    'Exist logfile
    If strLog <> "" Then
        intLinesBefore = getLineNumber(strLog)
    Else
        intLinesBefore = 0
    End If
End Sub

'***********************************************************************************************************************
Sub getUpdateLog()
    strObjID = "getUpdateLog"

    'Read update logs
    strLog = getFileToText(wuLogPath)
    'Exist logfile
    If strLog = "" Then Exit Sub
    intLinesNow = getLineNumber(strLog)

    'Filter to last update
    print_debug strObjID, ">>> Windows Update logfile <<<" & vbCrLf & getLineRange(strLog, intLinesBefore, intLinesNow-1)
End Sub

'*********************************************************************************************************************** WUA main
'***********************************************************************************************************************

'Init script
chkMailSets

'Init Updater logfile
initUpdateLog

Do
    aOK = True 'All OK

    'Check and update WUA
    If aOK Then aOK = chkAgentVer

    'Init WUA
    If aOK Then aOK = chkAgentSets

    'Init updateSearcher
    If aOK Then aOK = updateSearcher

    'Searching updates
    If aOK Then aOK = detectNow

    'Check update count
    If aOK Then
        If searchResult.updates.Count = 0 Then
            print_debug "ScriptMain", "There's no new update"
            'Print results
            getUpdateLog
            sendReport
            endOfScript
        End If
    End If

    'Downloading updates
    If aOK Then aOK = dwlUpdates

    'Installing updates
    If aOK Then aOK = instUpdates

    'Print results
    If aOK Then
        getUpdateLog
        sendReport
        endOfScript
    End If

    If Not aOK Then print_debug "ScriptMain", "===RESTART==="
Loop Until aOK
