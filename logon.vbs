Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: Logon Script
' Version: 4.00k
' Author: Egil Hansen, egilhansen.com
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Constants
Const DEFAULT_XMLFILE = "logonrules.xml"
Const ALWAYS_RUN_ON_SERVER = false
Const WRITE_EVENT_LOG = true
Const DEBUG_MODE = false
Const DELAY_BEFORE_CLOSE = 10000

' Global script objects
Dim settings: Set settings = new ScriptSettings
Set settings.ObjFSO = CreateObject("Scripting.FileSystemObject")

' Initialize script settings class
settings.StrCurrentDirectory = settings.ObjFSO.GetParentFolderName(WScript.ScriptFullName)
settings.StrIncludeDirectory = settings.ObjFSO.BuildPath(settings.StrCurrentDirectory, "Includes")
settings.BlnSilent = true
settings.StrXMLFile = settings.ObjFSO.BuildPath(settings.StrCurrentDirectory, DEFAULT_XMLFILE)
settings.BlnDebug = DEBUG_MODE

' Initialize store for groups
Set settings.ObjGroups = CreateObject("Scripting.Dictionary")
settings.ObjGroups.CompareMode = vbTextCompare

' Initialize logger
Set settings.ObjLog = new Logger
settings.ObjLog.ShowIEWindow = settings.BlnSilent
settings.ObjLog.PrintDebug = settings.BlnDebug
settings.ObjLog.WaitBeforeClose = DELAY_BEFORE_CLOSE

' Read argumennts, end script if ReadArguments return false
If ReadArguments(settings) Then
    Call FinalizeAndCleanup(0, settings)
End If

' Test if XML rules file exists
If Not settings.ObjFSO.FileExists(settings.StrXMLFile) Then
    Msgbox "The specified rules XML file do not exists.", vbCritical, "XML file not found!"
    settings.ObjLog.ShowIEWindow = false
    Call FinalizeAndCleanup(1, settings)
End If

' Start logging
Call settings.Log("Running logon script, date is " & Now, 0, 0)
Call settings.Info("", 0, 0)

' Get dynamic setting
If Not GetDynamicSettings(settings) Then
    Call FinalizeAndCleanup(1, settings)
End If

' Print friendly info about user and domain etc.
Call settings.Info("Detected <strong>" & settings.StrUserName & "</strong> logging on to <strong>" & settings.StrDomain & "</strong> through <strong>" & settings.StrComputerName & "</strong> authenticated by <strong>" & settings.StrDC & "</strong>.", 0, 0)
Call settings.Info("", 0, 0)

' Respond to computer possibly being a server
If Not ALWAYS_RUN_ON_SERVER Then
    If Not settings.BlnIsICASession And Not settings.BlnIsRDPSession Then
        If settings.BlnIsServer Then
            Dim intAnswer
            intAnswer = Msgbox("You are logging on to a server, do you want to continue running this logon script?", vbYesNo, "Continue running logon script?")
            If intAnswer = vbNo Then
                Call FinalizeAndCleanup(0, settings)
            End If
        End If
    End If
End If

' Test password age for user
Call TestPasswordAge(settings)
Call settings.Info("", 0, 0)

' Get users group membership
Call GetGroupMembership(settings)

' Parse XML files and execute actions
Set settings.ObjRules = new LogonRules
Call settings.ObjRules.Settings(settings)
Call settings.Info("Applying logon rules:", 0, 0)
settings.ObjRules.Execute()

' Finalize and cleanup
Call FinalizeAndCleanup(0, settings)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This will finalize the script, writing eventlog and cleaning up
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FinalizeAndCleanup(ByVal intErrorCode, ByRef s)
    ' If an error is logged, update global var s.BlnErrorDuringExecution
    Dim errorDuringExecution: errorDuringExecution = s.ObjLog.ErrorEventLogged
    If intErrorCode = 1 Then
        s.BlnErrorDuringExecution = true
    End If

    ' Finalize logging
    If Not errorDuringExecution Then
        Call s.Info("", 0, 0)
        Call s.Log("Finished running logon script, date is " & Now, 0, 0)
        Call s.ObjLog.FinalizeLogging(0)
        WScript.Quit(0)
    Else
        Call s.Info("", 0, 0)
        Call s.Log("Finished running logon script, date is " & Now, 0, 0)
        Call s.Info("", 0, 0)
        Call s.Log("There were one or more error's during execution.", 0, 1)
        Call s.Log("Please contact the IT Department for assistance.", 0, 1)
        Call s.ObjLog.FinalizeLogging(1)
        WScript.Quit(1)
    End If
End Sub

' Container class, used to pass all the script settings and objects around
Class ScriptSettings

    Dim ObjComputer
    Dim ObjDomain
    Dim ObjFSO
    Dim ObjGroups
    Dim ObjLog
    Dim ObjUser
    Dim ObjRules
    Dim StrIncludeDirectory
    Dim StrCurrentDirectory
    Dim StrComputerNameDN
    Dim StrComputerName
    Dim StrDomain
    Dim StrDC
    Dim StrSiteName
    Dim StrUserName
    Dim StrUserDN
    Dim StrXMLFile
    Dim BlnSilent
    Dim BlnDebug
    Dim BlnIsServer
    Dim BlnIsICASession
    Dim BlnIsRDPSession
    Dim BlnErrorDuringExecution

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
        Set ObjComputer = Nothing
        Set ObjDomain = Nothing
        Set ObjFSO = Nothing
        Set ObjGroups = Nothing
        Set ObjLog = Nothing
        Set ObjUser = Nothing
        Set ObjRules = Nothing
    End Sub

    Public Sub Info(ByVal strEvent, ByVal intIdent, ByVal intType)
        If Not IsNull(ObjLog) Then
            Call ObjLog.Info(strEvent, intIdent, intType)
        End If
    End Sub

    Public Sub Debug(ByVal strEvent, ByVal intIdent, ByVal intType)
        If Not IsNull(ObjLog) Then
            Call ObjLog.Debug(strEvent, intIdent, intType)
        End If
    End Sub

    Public Sub Log(ByVal strEvent, ByVal intIdent, ByVal intType)
        If Not IsNull(ObjLog) Then
            Call ObjLog.Log(strEvent, intIdent, intType)
        End If
    End Sub

End Class


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Include another script and execute its content
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IncludeSubScript(ByVal strIncludeDirectory, ByVal strScriptName)
    ExecuteGlobal sObjFSO.OpenTextFile(sObjFSO.BuildPath(strIncludeDirectory, strScriptName)).ReadAll
End Sub

Private Function ReadArguments(ByRef s)
    Dim arg, blnSilent
    Dim objArgs: Set objArgs = WScript.Arguments.Named

    ' Default values
    s.ObjLog.ShowIEWindow = false
    ReadArguments = false
    blnSilent = false

    ' Log
    Call s.Debug("Reading script arguments:", 0, 0)

    ' Check the command line for arguments
    For Each arg in objArgs
        Select Case LCase(arg)
            Case "silent"
                blnSilent = true
                Call s.Debug("/silent switch found", 1, 0)
            Case "rules"
                s.StrXMLFile = settings.ObjFSO.BuildPath(settings.StrCurrentDirectory, WScript.Arguments.Named.Item("rules"))
                Call s.Debug("/rules switch found. xml file is set to " & s.StrXMLFile, 1, 0)
            Case "forcedebug"
                s.BlnDebug = true
                Call s.Debug("/forcedebug switch found.", 1, 0)
            Case "/?"
                MsgBox "Logon Script" & vbCRLF & vbCRLF & _
                "   Logon.vbs /silent" & vbCRLF & vbCRLF & _
                "   /silent" & vbCRLF & _
                "            No screen display - non interactive"  & vbCRLF & _
                "   /forcedebug" & vbCRLF & _
                "            Force logon script into debug mode"  & vbCRLF & _
                "   /rules" & vbCRLF & _
                "            Force logon script to use a specific xml files"  & vbCRLF & _
                "   /?"  & vbCRLF & _
                "            This help screen"  & vbCRLF & vbCRLF
                ReadArguments = true
        End Select
    Next

    ' update ShowIEWindow and silent
    s.BlnSilent = blnSilent
    s.ObjLog.ShowIEWindow = not s.BlnSilent

    ' update PrintDebug
    s.ObjLog.PrintDebug = s.ObjLog.PrintDebug or s.BlnDebug

    ' Clean up
    Set objArgs = Nothing
End Function

Private Function GetDynamicSettings(ByRef s)
    Dim objWMIService, objOperatingSystem, objComp, objDomain
    Dim colComputers, colOperatingSystems
    Dim objADSystemInfo: Set objADSystemInfo = CreateObject("ADSystemInfo")
    Dim objShell: Set objShell = WScript.CreateObject("WScript.Shell")
    Dim objProcEnv: Set objProcEnv = objShell.Environment("Process")

    ' set default return value (dynamic settings found OK)
    GetDynamicSettings = true

    ' Log
    Call s.Debug("Getting dynamic settings:", 0, 0)

    ' enable error handling
    On Error Resume Next

    ' AD info
    s.StrDomain = objADSystemInfo.DomainDNSName

    ' test if there was an error when contacting domain
    If Err.number <> 0 Then
        Call s.Log("A domain either does not exist or could not be contacted.", 0, 1)
        Err.Clear
        GetDynamicSettings = false
        Exit Function
    End If
    s.StrUserDN = objADSystemInfo.UserName
    s.StrComputerNameDN = objADSystemInfo.ComputerName
    Call s.Debug("UserDN = " & s.StrUserDN, 1, 0)
    Call s.Debug("ComputerDN = " & s.StrComputerNameDN, 1, 0)
    Call s.Debug("Domain = " & s.StrDomain, 1, 0)

    ' User account
    Set s.ObjUser = GetObject("LDAP://" & s.StrDomain & "/" & s.StrUserDN)
    s.StrUserName = LCase(s.ObjUser.sAMAccountName)
    Call s.Debug("User = " & s.StrUserName, 1, 0)

    ' Computer account
    Set s.ObjComputer = GetObject("LDAP://" & s.StrDomain & "/" & s.StrComputerNameDN)
    s.StrComputerName = LCase(Replace(s.ObjComputer.Name,"CN=","",1,1,1))
    Call s.Debug("Computer = " & s.StrComputerName, 1, 0)

    ' Get site name
    s.StrSiteName = LCase(objADSystemInfo.SiteName)
    Call s.Debug("SiteName = " & s.StrSiteName, 1, 0)

    ' Authenticating domain controller
    Set s.ObjDomain = GetObject("LDAP://rootDSE")
    s.StrDC = s.objDomain.Get("dnsHostName")
    Call s.Debug("Authenticating domain controller = " & s.StrDC, 1, 0)

    ' test if there was an error when contacting domain and getting user/computer objects
    If Err.number <> 0 Then
        Call s.Log("Unable to read get all dynamic user, computer and domain settings.", 0, 1)
        Err.Clear
        GetDynamicSettings = false
        Exit Function
    End If

    ' Find out if its a server or workstation based on OS type
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colComputers = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each objComp in colComputers
        ' Values of DomainRole:
        ' 0 = Standalone Workstation
        ' 1 = Member Workstation
        ' 2 = Standalone Server
        ' 3 = Member Server
        ' 4 = Backup Domain Controller
        ' 5 = Primary Domain Controlle
        If objComp.DomainRole < 2 Then
            s.BlnIsServer = false
        Else
            s.BlnIsServer = true
        End If
    Next
    Call s.Debug("Is server = " & s.BlnIsServer, 1, 0)

    ' Is script being run on a from a terminal or citrix server session
    Dim strSessionName: strSessionName = objProcEnv("SESSIONNAME")
    If Left(strSessionName, 3) = "ICA" Then
        s.BlnIsICASession = true
    Else
        s.BlnIsICASession = False
    End If
    If Left(strSessionName, 3) = "RDP" Then
        s.BlnIsRDPSession = true
    Else
        s.BlnIsRDPSession = False
    End If
    Call s.Debug("Is ICA session = " & s.BlnIsICASession, 1, 0)
    Call s.Debug("Is RDP session = " & s.BlnIsRDPSession, 1, 0)

    If Err.number <> 0 Then
        wscript.echo Err.Description
        Call s.Log("Unable to detected computer and/or session type.", 0, 1)
        Err.Clear
        GetDynamicSettings = false
    End If

    ' reset error handling
    On Error Goto 0

    ' Clean up
    Set objShell = Nothing
    Set objProcEnv = Nothing
    Set objADSystemInfo = Nothing
    Set objWMIService = Nothing
    Set colOperatingSystems = Nothing
    Set objOperatingSystem = Nothing
    Set objDomain = Nothing
    Set colComputers = Nothing
    Set objComp = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test if/when the current users password will expire
'
' More info: http://www.cruto.com/resources/vbscript/vbscript%2Dexamples/ad/users/pwds/List-When-a-Password-Expires.asp
'            http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnclinic/html/scripting09102002.asp
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TestPasswordAge(ByVal s)

    Const SEC_IN_DAY = 86400
    Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000

    Dim objUserLDAP: Set objUserLDAP = s.ObjUser
    Dim intCurrentValue: intCurrentValue = objUserLDAP.Get("userAccountControl")

    Call s.Log("Retriving password information:", 0, 0)

    If intCurrentValue and ADS_UF_DONT_EXPIRE_PASSWD Then
        Call s.Log("The password does not expire.", 1, 0)
    Else
        Dim dtmValue: dtmValue = objUserLDAP.PasswordLastChanged
        s.Log "The password was last changed on " & DateValue(dtmValue) & " at " & TimeValue(dtmValue), 1, 0
        s.Log "The difference between when the password was last set and today is " & int(now - dtmValue) & " days", 1, 0
        Dim intTimeInterval: intTimeInterval = int(now - dtmValue)

        Dim objDomainNT: Set objDomainNT = GetObject("WinNT://" & s.StrDomain)
        Dim intMaxPwdAge: intMaxPwdAge = objDomainNT.Get("MaxPasswordAge")
        If intMaxPwdAge < 0 Then
            s.Log "The Maximum Password Age is set to 0 in the domain. Therefore, the password does not expire.", 1, 0
        Else
            intMaxPwdAge = (intMaxPwdAge/SEC_IN_DAY)
            s.Log "The maximum password age is " & Round(intMaxPwdAge) & " days", 1, 0
            If intTimeInterval >= intMaxPwdAge Then
              s.Log "The password has expired.", 1, 1
            Else
              s.Log "The password will expire on " & _
                  DateValue(dtmValue + intMaxPwdAge) & " (" & _
                      int((dtmValue + intMaxPwdAge) - now) & " days from today" & _
                          ").", 1, 0
            End If
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This will retrive the current users group memberships
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function IsMember(ByVal objADObject, ByVal strGroupNTName, ByVal objGroups)
    ' Return True if this user or computer is a member of the group.
    IsMember = objGroups.Exists(objADObject.sAMAccountName & "\" & strGroupNTName)
End Function

Private Sub GetGroupMembership(ByRef s)
    Call s.Debug("Getting group memberships:", 0, 0)

    ' Create new ADO connection
    Dim objAdoConnection: Set objAdoConnection = CreateObject("ADODB.Connection")
    objAdoConnection.Provider = "ADsDSOObject"
    objAdoConnection.Open "Active Directory Provider"

    ' Create new ADO command
    Dim objAdoCommand: Set objAdoCommand = CreateObject("ADODB.Command")
    objAdoCommand.ActiveConnection = objAdoConnection
    objAdoCommand.Properties("Page Size") = 100
    objAdoCommand.Properties("Timeout") = 30
    objAdoCommand.Properties("Cache Results") = False

    ' Search entire domain.
    Dim strBase: strBase = "<LDAP://" & s.StrDomain & ">"

    ' Retrieve NT name of each group.
    Dim strAttributes: strAttributes = "sAMAccountName"

    ' Add user name to dictionary object, so LoadGroups need only be
    ' called once for each user or computer.
    s.ObjGroups.Add s.StrUserName & "\", True

    ' Retrieve tokenGroups array, a calculated attribute.
    s.ObjUser.GetInfoEx Array("tokenGroups"), 0
    Dim arrByteGroups: arrByteGroups = s.ObjUser.Get("tokenGroups")

    ' Create a filter to search for groups with objectSid equal to each
    ' value in tokenGroups array.
    Dim strFilter: strFilter = "(|"
    If (TypeName(arrByteGroups) = "Byte()") Then
        ' tokenGroups has one entry.
        strFilter = strFilter & "(objectSid=" & OctetToHexStr(arrByteGroups) & ")"
    ElseIf (UBound(arrByteGroups) > -1) Then
        ' TokenGroups is an array of two or more objectSid's.
        Dim k
        For k = 0 To UBound(arrByteGroups)
            strFilter = strFilter & "(objectSid=" & OctetToHexStr(arrByteGroups(k)) & ")"
        Next
    Else
        ' tokenGroups has no objectSid's.
        Exit Sub
    End If
    strFilter = strFilter & ")"

    ' Use ADO to search for groups whose objectSid matches any of the
    ' tokenGroups values for this user or computer.
    objAdoCommand.CommandText = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

    ' Execute search in AD
    Dim objAdoRecordset: Set objAdoRecordset = objAdoCommand.Execute

    ' Enumerate groups and add NT name to dictionary object.
    Dim strGroupName
    Do Until objAdoRecordset.EOF
        ' Read group name
        strGroupName = objAdoRecordset.Fields("sAMAccountName").Value
        ' Add group to dictionary
        s.ObjGroups.Add s.ObjUser.sAMAccountName & "\" & strGroupName, True
        ' Log
        Call s.Debug(strGroupName, 1, 0)

        objAdoRecordset.MoveNext
    Loop
    objAdoRecordset.Close

    ' Clean up
    Set objAdoRecordset = Nothing
    Set objAdoConnection = Nothing
    Set objAdoCommand = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function to convert OctetString (byte array) to Hex string,
' with bytes delimited by \ for an ADO filter.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function OctetToHexStr(ByVal arrbytOctet)
    Dim k
    OctetToHexStr = ""
    For k = 1 To Lenb(arrbytOctet)
        OctetToHexStr = OctetToHexStr & "\" & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
    Next
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compares two files' date last modified value
'
' Returns -1 if objFile1 < objFile2
' Returns 1 if objFile1 > objFile2
' Returns 0 if objFile1 = objFile2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CompareFileDateLastModified(ByRef objFile1, ByRef objFile2)
	If objFile1.DateLastModified < objFile2.DateLastModified Then
	  CompareFileDateLastModified = -1
	ElseIf objFile1.DateLastModified > objFile2.DateLastModified Then
	  CompareFileDateLastModified = 1
	Else
	  CompareFileDateLastModified = 0
	End If
End Function

Sub E(ByVal Str)
	WScript.Echo Str
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Replaces wildes in string
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function ReplaceWildCards(ByVal strInput, ByRef s)
    Dim tmp
    Dim objShell: Set objShell = WScript.CreateObject("WScript.Shell")
    Dim objProcessEnv: Set objProcessEnv = objShell.Environment("Process")

    tmp = strInput
    tmp = Replace(tmp, "%computername%", s.StrComputerName, 1, -1, 1)
    tmp = Replace(tmp, "%username%", s.StrUserName, 1, -1, 1)
    tmp = Replace(tmp, "%domain%", s.StrDomain, 1, -1, 1)
    tmp = Replace(tmp, "%logonserver%", s.StrDC, 1, -1, 1)
    tmp = Replace(tmp, "%allusersdesktop%", objShell.SpecialFolders("AllUsersDesktop"), 1, -1, 1)
    tmp = Replace(tmp, "%allusersstartmenu%", objShell.SpecialFolders("AllUsersStartMenu"), 1, -1, 1)
    tmp = Replace(tmp, "%allusersprograms%", objShell.SpecialFolders("AllUsersPrograms"), 1, -1, 1)
    tmp = Replace(tmp, "%allusersstartup%", objShell.SpecialFolders("AllUsersStartup"), 1, -1, 1)
    tmp = Replace(tmp, "%desktop%", objShell.SpecialFolders("Desktop"), 1, -1, 1)
    tmp = Replace(tmp, "%favorites%", objShell.SpecialFolders("Favorites"), 1, -1, 1)
    tmp = Replace(tmp, "%fonts%", objShell.SpecialFolders("Fonts"), 1, -1, 1)
    tmp = Replace(tmp, "%mydocuments%", objShell.SpecialFolders("MyDocuments"), 1, -1, 1)
    tmp = Replace(tmp, "%nethood%", objShell.SpecialFolders("NetHood"), 1, -1, 1)
    tmp = Replace(tmp, "%programs%", objShell.SpecialFolders("Programs"), 1, -1, 1)
    tmp = Replace(tmp, "%fonts%", objShell.SpecialFolders("Fonts"), 1, -1, 1)
    tmp = Replace(tmp, "%recent%", objShell.SpecialFolders("Recent"), 1, -1, 1)
    tmp = Replace(tmp, "%sendto%", objShell.SpecialFolders("SendTo"), 1, -1, 1)
    tmp = Replace(tmp, "%startmenu%", objShell.SpecialFolders("StartMenu"), 1, -1, 1)
    tmp = Replace(tmp, "%templates%", objShell.SpecialFolders("Templates"), 1, -1, 1)
    tmp = Replace(tmp, "%systemdrive%", objProcessEnv("SYSTEMDRIVE"), 1, -1, 1)
    tmp = Replace(tmp, "%systemroot%", objProcessEnv("SYSTEMROOT"), 1, -1, 1)
    ReplaceWildCards = tmp

    Set objShell = Nothing
    Set objProcessEnv = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Opens IEWindow, provides methods to print to it
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class IEWindow
    Private objIntExplorer
    Dim CloseIEWindowOnExit, WaitBeforeClose

    Private Sub Class_Initialize()
        CloseIEWindowOnExit = true
        WaitBeforeClose = 0
    End Sub

    Private Sub Class_Terminate()
        Call CloseIE()
        Set objIntExplorer = Nothing
    End Sub

    Private Sub SetupIE()
        Dim strTitle
	    strTitle = "Logon Script"

	    ' Create reference to objIntExplorer
	    ' This will be used for the user messages. Also set IE display attributes
	    With objIntExplorer
		    .Navigate "about:blank"
		    .ToolBar   = 0
		    .Menubar   = 0
		    .StatusBar = 0
		    .Width     = 650
		    .Height    = 500
		    .Left      = 50
		    .Top       = 20
	    End With

	    ' Set some formating
	    With objIntExplorer.Document
		    .WriteLn ("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">")
		    .WriteLn ("<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"">")
		    .WriteLn   ("<head>")
		    .WriteLn    ("<title>" & strTitle & "</title>")
		    .WriteLn     ("<style type=""text/css"">")
		    .WriteLn      ("body {border: 0; margin: 5; text-align: left; font-family: verdana; font-size: 9pt; } pre {font-family: verdana; font-size: 9pt;}")
		    .WriteLn     ("</style>")
		    .WriteLn   ("</head>")
		    .WriteLn   ("<body>")
		    .WriteLn   ("<div>")
	    End With

	    ' Wait for IE to finish
	    Do While (objIntExplorer.Busy)
		    Wscript.Sleep 100
	    Loop

	    ' Show IE
	    objIntExplorer.Visible = 1
    End Sub

    Private Sub CloseIE()
        If IsObject(objIntExplorer) Then
            ' Insert wait period before close
            wscript.sleep(WaitBeforeClose)
            ' Write standard html end-tags
            On Error Resume Next
            objIntExplorer.Document.WriteLn("</div>")
            objIntExplorer.Document.WriteLn("</body>")
            objIntExplorer.Document.WriteLn("</html>")
            If CloseIEWindowOnExit Then
                objIntExplorer.Quit()
            End If
            Set objIntExplorer = Nothing
        End If
    End Sub

    Public Sub WriteLn(ByVal str)
        If Not IsObject(objIntExplorer) Then
            Set objIntExplorer = Wscript.CreateObject("InternetExplorer.Application")
            Call SetupIE()
        End If
        On Error Resume Next
        objIntExplorer.Document.WriteLn(str & "<br />")
        objIntExplorer.Document.parentWindow.scrollTo 0, objIntExplorer.Document.body.offsetHeight
    End Sub
End Class

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logs event to IEWindow and eventlog
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class Logger
    Dim WriteEventLog, ErrorEventLogged, PrintDebug, ShowIEWindow
    Private Events, HasTerminated, IE

    Public Property Get WaitBeforeClose
        WaitBeforeClose = IE.WaitBeforeClose
    End Property

    Public Property Let WaitBeforeClose(intMilisecs)
        IE.WaitBeforeClose = intMilisecs
    End Property

    Private Sub Class_Initialize()
        Events = ""
        HasTerminated = false
        ErrorEventLogged = false
        WriteEventLog = true
        PrintDebug = false
        ShowIEWindow = false
        Set IE = new IEWindow
    End Sub

    Private Sub Class_Terminate()
        If Not HasTerminated Then
            Call FinalizeLogging(0)
        End If
        Set IE = Nothing
    End Sub

    Public Sub FinalizeLogging(ByVal intErrorCode)
        If Not HasTerminated Then
            ' Write to event log
            If WriteEventLog Then
    	        Dim objWshShell: Set objWshShell = CreateObject("WScript.Shell")
	            On Error Resume Next
	            objWshShell.LogEvent intErrorCode, Events
	            On Error Goto 0
	            Set objWshShell = Nothing
            End If
            IE.CloseIEWindowOnExit = intErrorCode = 0
            HasTerminated = true
        End If
    End Sub

    Public Sub Info(ByVal strEvent, ByVal intIdent, ByVal intType)
        If ShowIEWindow Then
            ' If an error is logged, update global var sBlnErrorDuringExecution
            If intType = 1 Then
                ErrorEventLogged = true
            End If

            ' Add tabs
            Dim strWebTabs: strWebTabs = ""
            Dim intCounter
            For intCounter = 0 To intIdent
                strWebTabs = strWebTabs & "&nbsp;&nbsp;&nbsp;&nbsp;"
            Next

            Call WriteToIE(strWebTabs & strEvent, intType)
        End If
    End Sub

    Public Sub Debug(ByVal strEvent, ByVal intIdent, ByVal intType)
        ' If an error is logged, update global var sBlnErrorDuringExecution
        If intType = 1 Then
            ErrorEventLogged = true
        End If

        ' Add tabs
        Dim strTabs: strTabs = ""
        Dim strWebTabs: strWebTabs = ""
        Dim intCounter
        For intCounter = 0 To intIdent
            strTabs = strTabs & vbTab
            strWebTabs = strWebTabs & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Next

        Call WriteToEventLog(strTabs & strEvent)

        If ShowIEWindow And PrintDebug Then
            Call WriteToIE(strWebTabs & strEvent, intType)
        End If
    End Sub

    Public Sub Log(ByVal strEvent, ByVal intIdent, ByVal intType)
        ' If an error is logged, update global var sBlnErrorDuringExecution
        If intType = 1 Then
            ErrorEventLogged = true
        End If

        ' Add tabs
        Dim strTabs: strTabs = ""
        Dim strWebTabs: strWebTabs = ""
        Dim intCounter
        For intCounter = 0 To intIdent
            strTabs = strTabs & vbTab
            strWebTabs = strWebTabs & "&nbsp;&nbsp;&nbsp;&nbsp;"
        Next

        Call WriteToEventLog(strTabs & strEvent)

        If ShowIEWindow Then
            Call WriteToIE(strWebTabs & strEvent, intType)
        End If
    End Sub

    Private Sub WriteToIE(ByVal strEvent, ByVal intType)
        ' Write to IE Window
        If intType = 0 Then
            Call IE.WriteLn(strEvent)
        Else
            Call IE.WriteLn("<span style=""color: red;"">" & strEvent & "</span>")
        End If
    End Sub

    Private Sub WriteToEventLog(ByVal strEvent)
        ' Write to event log
        If WriteEventLog Then
            ' Add event to string events.
            Events = Events & strEvent & vbCrLf
        End If
    End Sub

End Class

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MapDrive class, handles mapping of network drives
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Class MapDrive
	Dim Drive, Unc, Persistent
    Private s

    Public Sub Settings(ByRef objSettings)
        Set s = objSettings
    End Sub

    Public Sub Execute()
        ' If a drive range is specified, iterate over it an check for available drivers
	    ' otherwise just try to map the single network drive
	    If Len(Drive) > 1 Then
		    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
	        Dim strCurrentDrive, intCounter
            Dim blnFreeDrive: blnFreeDrive = false

            For intCounter = 1 To Len(Drive)
                strCurrentDrive = UCase(Trim(Mid(Drive,intCounter,1)))

                ' Make sure drive letter string is not empty
                If Len(strCurrentDrive) = 1 Then

                    ' Make sure drive letters in indeed a letter between A to Z
                    If Asc(strCurrentDrive) >= 65 AND Asc(strCurrentDrive) <= 90 Then

                        ' Map drive if it doesn't already exists
                        If Not objFSO.DriveExists(strCurrentDrive) Then
                    	    ' Add : to strDrive
	                        strCurrentDrive = strCurrentDrive & ":"

	                        ' Map drive
                            Call MapNetworkDrive(strCurrentDrive, Unc, Persistent)

                            ' Set blnFreeDrive
                            blnFreeDrive = true

                            ' Exit loop
                            Exit For
                        End If
                    End If
                End If
            Next

            ' Print warning if no free drive letters was found
            If Not blnFreeDrive Then
                Call s.Log("Error: Unable to map " & Unc, 1, 1)
		        Call s.Log("Error description: Unable to find any available drive letters within the specified range (" & Drive & ").", 1, 1)
		    End If

	        ' Clean up
	        Set objFSO = Nothing
	    Else
	        ' Add : to strDrive
	        Drive = Drive & ":"
	        ' Map drive
            Call MapNetworkDrive(Drive, Unc, Persistent)
        End If
    End Sub

    ' Maps a network path to a drive
    Private Sub MapNetworkDrive(ByVal strDrive, ByVal strPath, ByVal blnPersistent)
        Dim objWshNetwork: Set objWshNetwork = CreateObject("WScript.Network")
        Dim intErrNumber, strErrDescription

        ' Switch on error handling
        On Error Resume Next

        ' Since a drive might already be mapped, unmap it.
	    ' This assures everyone has the same drive mappings
	    Err.Clear
	    objWshNetwork.RemoveNetworkDrive strDrive, True, True
	    intErrNumber = Err.number
	    strErrDescription = Err.Description

	    ' Error codes:
	    ' -2147022646 = Drive not mapped.
        ' 0 = drive succesfully unmapped.
	    If intErrNumber = 0 Or intErrNumber = -2147022646 Then
	        Err.Clear
		    objWshNetwork.MapNetworkDrive strDrive, strPath, blnPersistent
		    intErrNumber = Err.number
		    strErrDescription = Err.Description
        End If

        If blnPersistent Then
	        ' Check error condition and output appropriate user message
	        If intErrNumber <> 0 Then
	            Call s.Log("Error: Unable to create a persistent mapping of " & strDrive & " to " & strPath, 1, 1)
		        Call s.Log("Error description: " & strErrDescription, 1, 1)
	        Else
	            Call s.Log("Created a persistent mapping of " & strPath & " to " & strDrive, 1, 0)
	        End If
	    Else
	        ' Check error condition and output appropriate user message
	        If intErrNumber <> 0 Then
	            Call s.Log("Error: Unable to map " & strDrive & " to " & strPath, 1, 1)
		        Call s.Log("Error description: " & strErrDescription, 1, 1)
	        Else
	            Call s.Log("Mapped " & strPath & " to " & strDrive, 1, 0)
	        End If
	    End If

	    ' Clean up
	    Set objWshNetwork = Nothing
    End Sub
End Class

Class AddPrinter
    Dim Unc
    Private s

    Public Sub Settings(ByRef objSettings)
        Set s = objSettings
    End Sub

    Public Sub Execute()
        Dim objWshNetwork, objPrinters, strErrDescription, intErrNumber
	    Set objWshNetwork = CreateObject("WScript.Network")

	    ' Handle errors
	    On Error Resume Next

	    ' Map printer
        objWshNetwork.AddWindowsPrinterConnection Unc
        strErrDescription = Err.Description
        intErrNumber = Err.number

	    ' Check error condition and output appropriate user message
	    If intErrNumber <> 0 Then
		    Call s.Log("Error: Unable to connect to network printer " & Unc, 1, 1)
		    ' Seems that the AddWindowsPrinterConnection doesn't return a error description.
		    'Call LogEvent("Error description: " & strErrDescription)
	    Else
		    Call s.Log("Added printer connection to " & Unc, 1, 0)
	    End If

	    ' Clean up
	    Set objWshNetwork = Nothing
	    Set objPrinter = Nothing
    End Sub
End Class

Class SetDefaultPrinter
    Dim Unc
    Private s

    Public Sub Settings(ByRef objSettings)
        Set s = objSettings
    End Sub

    Public Sub Execute()
        Dim objWshNetwork
        Set objWshNetwork = CreateObject("WScript.Network")
        On Error Resume Next
        objWshNetwork.SetDefaultPrinter Unc
        Dim strErrDescription
        strErrDescription = Err.Description
        If Err.number = 0 Then
	        Call s.Log("Setting default printer to " & Unc, 1, 0)
        Else
            Call s.Log("Error: Unable to set the default printer to """ & Unc & """", 1, 1)
	        Call s.Log("Error description: " & strErrDescription, 1, 1)
        End If
        ' Clean up
        Set objWshNetwork = Nothing
    End Sub
End Class

Class CopyFile
    Dim Source, Destination, CopyOption
    Private s, objSource, objDestination

    Public Sub Settings(ByRef objSettings)
        Set s = objSettings
    End Sub

    Public Sub Execute()
        Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")

	    ' Test if source file exsits
	    If objFSO.FileExists(Source) Then
		    ' Copy if true
		    Dim blnCopy: blnCopy = False

		    ' Get handler for source file
		    Set objSource = objFSO.GetFile(Source)

		    Select Case LCase(CopyOption)
		        Case "ifdifferent"
		            If objFSO.FileExists(Destination) Then
		                Set objDestination = objFSO.GetFile(Destination)
			            If CompareFileDateLastModified(objSource, objDestination) <> 0 Then
				            blnCopy = True
			            Else
			                Call s.Log("Skipped copying " & Source & " to " & Destination & ". Destination file is up to date.", 1, 0)
			            End If
		            Else
		                blnCopy = True
		            End If
		        case "onlyifnotexists"
		            If not objFSO.FileExists(Destination) Then
		                blnCopy = True
		            Else
		                Call s.Log("Skipped copying " & Source & " to " & Destination & ". Destination file already exists.", 1, 0)
		            End If
		        Case "alwayscopy"
		        	blnCopy = True
		    End Select

		    ' Copy File if needed
		    If blnCopy Then
			    On Error Resume Next
			    Call s.Log("Copying " & Source & " to " & Destination, 1, 0)
			    objSource.Copy(Destination)
			    If Err.Number <> 0 Then
				    Call s.Log("Error copying file: " & Err.Description, 1, 1)
				    On Error GoTo 0
			    End If
		    End If

	    Else
		    Call s.Log("Source file does not exists.", 1, 1)
	    End If

	    ' Clean up
	    Set objDestination = Nothing
	    Set objSource = Nothing
	    Set objFSO = Nothing
    End Sub
End Class

Class RunProgramInBackground
    Dim Executable
    Private s

    Public Sub Settings(ByRef objSettings)
        Set s = objSettings
    End Sub

    Public Sub Execute()
	    Dim objWshShell
	    Set objWshShell = CreateObject("WScript.Shell")
	    On Error Resume Next
	    objWshShell.Run Executable, 0, false
	    ' If there was an error (finding the program ect.) print an error to the user.
	    If Err.number = 0 Then
	        Call s.Log("Running """ & Executable & """", 1, 0)
	    Else
		    Call s.Log("Error: Unable to run program """ & Executable & """", 1, 1)
	    End If
	    ' Clean up
	    Set objWshShell = Nothing
    End Sub
End Class

Class RunProgramAndWait
    Dim Executable
    Private s

    Public Sub Settings(ByRef objSettings)
        Set s = objSettings
    End Sub

    Public Sub Execute()
	    Dim objWshShell
	    Set objWshShell = CreateObject("WScript.Shell")
	    Call s.Log("Running """ & Executable & """", 1, 0)
	    On Error Resume Next
	    objWshShell.Run Executable, 0, true
	    ' If there was an error (finding the program ect.) print an error to the user.
	    If Err.number = 0 Then
	        Call s.Log("Finished running """ & Executable & """", 1, 0)
	    Else
		    Call s.Log("Error: Unable to run program """ & Executable & """", 1, 1)
	    End If
	    ' Clean up
	    Set objWshShell = Nothing
    End Sub
End Class

Class LogonRules
    Private s
    Private strComputerName, strUserName, m_objGroups, m_objComputer, m_objUser

    Public Sub Settings(ByRef objSettings)
        Set s = objSettings
    End Sub

    Public Sub Execute
        Dim objXMLDoc, blnSuccess
        Set objXMLDoc = CreateObject("Microsoft.XMLDOM")

        ' Log event
        Call s.Debug("Parsing logon rules:", 0, 0)

        objXMLDoc.async = "false"
        ' Try loading the xmlfile
        blnSuccess = objXMLDoc.load(s.StrXMLFile)

        If blnSuccess Then
	        Dim strXPath, colNodes, objNode, colActions, colFilters, objAction
	        Dim obj
	        Dim i, j

	        ' Make a collection of the actionSets
	        strXPath = "//logonRules/actionSet"
	        Set colNodes = objXMLDoc.selectNodes(strXPath)

	        ' Loop through collection of actionSets
	        For i = 0 To colNodes.length - 1
		        Set objNode = colNodes(i)
		        Set colFilters = objNode.selectNodes("filterSet/filter")

		        ' if actionSet's filterSet is true, execute actions
		        If TestConditions(colFilters) Then
		            Set colActions = objNode.selectNodes("actions/*")

		            ' iterate over each action
                    For j = 0 To colActions.length - 1
                        Set objAction = colActions(j)
		                Select Case objAction.nodeName
		                    Case "addPrinter"
		                        Set obj = new AddPrinter
		                        obj.Settings(s)
		                        obj.Unc = ReplaceWildCards(objAction.getAttribute("unc"), s)
		                        obj.Execute()
		                    Case "copyFile"
		                        Set obj = new CopyFile
		                        obj.Settings(s)
		                        obj.Source = ReplaceWildCards(objAction.getAttribute("source"), s)
		                        obj.Destination = ReplaceWildCards(objAction.getAttribute("destination"), s)
		                        obj.CopyOption = objAction.getAttribute("option")
		                        obj.Execute()
		                    Case "executeInBackground"
		                        Set obj = new RunProgramInBackground
		                        obj.Settings(s)
		                        obj.Executable = ReplaceWildCards(objAction.getAttribute("executable"), s)
		                        obj.Execute()
		                    Case "executeAndWait"
		                        Set obj = new RunProgramAndWait
		                        obj.Settings(s)
		                        obj.Executable = ReplaceWildCards(objAction.getAttribute("executable"), s)
		                        obj.Execute()
		                    Case "mapDrive"
		                        Set obj = new MapDrive
		                        obj.Settings(s)
		                        obj.Drive = objAction.getAttribute("drive")
		                        obj.Unc = ReplaceWildCards(objAction.getAttribute("unc"), s)
		                        obj.Persistent = CBool(objAction.getAttribute("persistent") = "true")
		                        obj.Execute()
		                    Case "setDefaultPrinter"
		                        Set obj = new SetDefaultPrinter
		                        obj.Settings(s)
		                        obj.Unc = ReplaceWildCards(objAction.getAttribute("unc"), s)
		                        obj.Execute()
		                End Select
		                Set obj = Nothing
    		        Next
		        End If
            Next

	        ' Clean up memory
	        Set colNodes = Nothing
	        Set objNode = Nothing
	        Set colActions = Nothing
	        Set colFilters = Nothing
	        Set objAction = Nothing
        Else
	        ' The document failed to load.
	        Dim strErrText, xPE
	        ' Obtain the ParseError object
	        Set xPE = objXMLDoc.parseError
	        With xPE
		        strErrText = "Your XML Document failed to load " & _
			        "due the following error." & vbCrLf & _
			        "Error #: " & .errorCode & ": " & xPE.reason & _
			        "Line #: " & .Line & vbCrLf & _
			        "Line Position: " & .linepos & vbCrLf & _
			        "Position In File: " & .filepos & vbCrLf & _
			        "Source Text: " & .srcText & vbCrLf & _
			        "Document URL: " & .url
	        End With
	        Call s.Debug(strErrText,0,1)
	        MsgBox strErrText, vbExclamation

	        ' Clean up used memory
	        Set xPE = Nothing
	        s.ObjLog.FinalizeLogging(1)
        End If

        ' Clean up
        Set objXMLDoc = Nothing
    End Sub

    Private Function TestConditions(ByRef objFilters)
        Dim objFilter, i
        ' initialize test condition to true
        TestConditions = true

        ' iterate over each filter
        For i = 0 To objFilters.length - 1
            Set objFilter = objFilters(i)

            ' Test condition
            Select Case LCase(objFilter.getAttribute("condition"))
                Case "and"
                    TestConditions = TestConditions And TestFilter(objFilter)
                Case "or"
                    TestConditions = TestConditions or TestFilter(objFilter)
                Case "not"
                    TestConditions = TestConditions And Not TestFilter(objFilter)
                    If Not TestConditions Then
                        Exit For
                    End If
            End Select
        Next

        Set objFilter = Nothing
    End Function

    Private Function TestFilter(ByRef objFilter)
        Dim filterType
        filterType = objFilter.nodeName

        If filterType = "filter" Then
            Select Case LCase(objFilter.getAttribute("type"))
                case "group"
                    TestFilter = IsMember(s.ObjUser, objFilter.getAttribute("name"), s.ObjGroups)
                case "user"
                    TestFilter = LCase(objFilter.getAttribute("name")) = s.StrUserName
                case "computer"
                    TestFilter = LCase(objFilter.getAttribute("name")) = s.StrComputerName
                case "domaincontroller"
                    TestFilter = LCase(objFilter.getAttribute("name")) = s.StrDC
                case "site"
                    TestFilter = LCase(objFilter.getAttribute("name")) = s.StrSiteName
            End Select
        End If

        If filterType = "boolFilter" Then
            Select Case LCase(objFilter.getAttribute("type"))
                case "isICASession"
                    TestFilter = s.BlnIsICASession
                case "isRDPSession"
                    TestFilter = s.BlnIsRDPSession
                case "isServer"
                    TestFilter = s.BlnIsServer
            End Select
        End If

    End Function

End Class