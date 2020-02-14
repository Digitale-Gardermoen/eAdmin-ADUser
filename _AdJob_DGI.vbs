Option Explicit
'On Error Resume Next

' Dim and set all global variables
Dim DomainName, ouDC, ouGroups, ouCustomers, ouDeleted, serverPrefix, domainController, Password
GetConfig()
' This is the script directory, for use when calling the powershell scripts.
Dim scriptdir
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Main
Dim Op
Op = WScript.Arguments.Named.Item("op")
Op = LCase(Op)
if Op = "au" then
  AddUser
elseif Op = "ru" then
  RemoveUser
elseif Op = "am" then
  AddMail
elseif Op = "rm" then
  RemoveMail
elseif Op = "aug" or Op = "rug" then
  AddRemoveGroupMembership (Op)
else
  WScript.Echo "Argument /op:?"
end if

WScript.Quit ' Stop the script if no operation was called
' End Main

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Errorhandling, used when Err.Number <> 0
Sub CheckError(ByVal Location)
  ' DO NOT ADD 'On Error Resume Next' HERE
  if Err.Number <> 0 then
    WScript.Echo Location & " [" & Err.Number & "] " & Err.Description
    WScript.Quit
  end if
end Sub

' This will print the error and quit the script
Sub QuitWithError(ByVal ErrorMessage)
  ' DO NOT ADD 'On Error Resume Next' HERE
  WScript.Echo ErrorMessage
  WScript.Quit
end Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Config Ops
' This reads a INI file from the current cwd
' Important to set the correct Dir if cwd is wrong
Function ReadIniFile(sFSpec)
  Dim goFS   : Set goFS   = CreateObject("Scripting.FileSystemObject")
  Dim dicTmp : Set dicTmp = CreateObject("Scripting.Dictionary")
  Dim tsIn   : Set tsIn   = goFS.OpenTextFile(sFSpec)
  Dim sLine, sSec, aKV
  Do Until tsIn.AtEndOfStream
     sLine = Trim(tsIn.ReadLine())
     If "[" = Left(sLine, 1) Then
        sSec = Mid(sLine, 2, Len(sLine) - 2)
        Set dicTmp(sSEc) = CreateObject("Scripting.Dictionary")
     Else
        If "" <> sLine Then
           aKV = Split(sLine, "==")
           If 1 = UBound(aKV) Then
              dicTmp(sSec)(Trim(aKV(0))) = Trim(aKV(1))
           End If
        End If
     End If
  Loop
  tsIn.Close
  Set ReadIniFile = dicTmp
End Function

Function GetConfig()
  Dim dicIni : Set dicIni = ReadIniFile(".\_AdJob_DGI.ini")
  Dim sSec
  For Each sSec In dicIni.Keys()
    DomainName       = dicIni(sSec)("DomainName")
    ouDC             = dicIni(sSec)("ouDC")
    ouGroups         = dicIni(sSec)("ouGroups")
    ouCustomers      = dicIni(sSec)("ouCustomers")
    ouDeleted        = dicIni(sSec)("ouDeleted")
    serverPrefix     = dicIni(sSec)("serverPrefix")
    domainController = dicIni(sSec)("domainController")
    Password         = dicIni(sSec)("Password")
  Next
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' User operations
Sub AddToGroup(ByVal GroupName)
  On Error Resume Next
  Err.Clear
  if GroupName <> "" then
    Dim oGrp, GroupDN, usrName, uGrp
    usrName = GetUserDN(LoginName)
    Set uGrp = GetObject(serverPrefix & usrName)
    uGrp = uGrp.MemberOf
    if InStr(uGrp, GroupName) then
      ' Do nothing
    else
      GroupDN = GetGroupDN(GroupName)
      Set oGrp = GetObject(serverPrefix & GroupDN)
      CheckError("AddToGroup/1")
      oGrp.Add (serverPrefix & usrName)
      CheckError("AddToGroup/2")
      set oGrp = Nothing
    end if
  end if
end Sub

Sub RemoveFromGroup(ByVal GroupName)
  On Error Resume Next
  Err.Clear
  if GroupName <> "" then
    Dim oGrp
    Set oGrp = GetObject("WinNT://" & DomainName & "/" & GroupName)
    CheckError("RemoveFromGroup/1")
    oGrp.Remove("WinNT://" & DomainName & "/" & LoginName)
    CheckError("RemoveFromGroup/2")
    set oGrp = Nothing
  end if
end Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' AddUser
Sub AddUser
  On Error Resume Next
  Err.Clear
  ' Declare local variables
  Dim LoginName, OrgUnit, ou2, FirstName, LastName, FullName
  Dim DepartmentName, usrName, objOU, objUser, userOU
  LoginName      = WScript.Arguments.Named.Item("login")
  LoginName      = LCase(LoginName)
  OrgUnit        = WScript.Arguments.Named.Item("ou")
  ou2            = WScript.Arguments.Named.Item("ou2")
  FirstName      = WScript.Arguments.Named.Item("firstn")
  LastName       = WScript.Arguments.Named.Item("lastn")
  FullName       = WScript.Arguments.Named.Item("fulln")
  DepartmentName = WScript.Arguments.Named.Item("cust")
  usrName        = GetUserDN(LoginName)
  userOU         = "OU=Users,OU=" & ou2 & "," & ouCustomers
  
  CheckError("AU-Start")
  if LoginName = "" then
    QuitWithError("Scriptet fikk ikke argument /login:")
  end if

  if Password = "" then
    QuitWithError("Scriptet fikk ikke argument /pw:")
  end if

  Set objOU = GetObject(serverPrefix & userOU & "," & ouDC)
  CheckError("AU-GetUsersOU")

  'Check if the user exists, if it does, move the user
  if usrName <> "" then
    if InStr(1, usrName, ouDeleted, vbTextCompare) > 0 then
      objOU.MoveHere serverPrefix & usrName, vbNullString       'Move the user
    end if
    CheckError("AU-MoveUser")
    usrName = GetUserDN(LoginName)
    Set objUser = GetObject(serverPrefix & usrName)
    CheckError("AU-GetMovedUser")
    'unflag the user
    UnflagUser
  else
    'Create the user
    Set objUser = objOU.Create("User", "CN=" & FullName)  
    objUser.Put "samAccountName", LoginName
    objUser.SetInfo
    objUser.Put "userPrincipalName", LoginName & "@" & DomainName
    objUser.SetInfo
    CheckError("AU-CreateUser")

    'Set Name
    if Firstname <> "" then 
      objUser.Put "givenName", FirstName
    end if
    objUser.put "sn", LastName
    objUser.Put "displayName", FullName
    objUser.SetInfo
    CheckError("AU-SetName")

    'Set Password
    objUser.SetPassword Password
    objUser.SetInfo
    CheckError("AU-Set Password")

    '512 = NORMAL_ACCOUNT
    '544 = PASSWD_NOTREQD | NORMAL_ACCOUNT
    'Setting UAC for the user, VBS seems to create users with 544 UAC.
    objUser.Put "userAccountControl", 512
    objUser.SetInfo
    CheckError("AU-Set UAC")
  
    objUser.put "department", DepartmentName  'We set the DepartmentName even tho this gets synced to the user, this is to fix a bug when creating DGI users.
    objUser.Put "pwdLastSet", 0
    objUser.SetInfo
    CheckError("AU-SetCompany")
  end if

  'Enable user
  if objUser.AccountDisabled = true then
    objUser.AccountDisabled = False
    objUser.SetInfo
  end if
  CheckError("AU-EnableUser")

  'Add user to defualt groups
  Dim groupNameArray, group, trimmedGroup, grpObj, GroupDN, uGrp
  usrName = GetUserDN(LoginName)                          'get this if its a new user since the value is ""
  groupNameArray = Split(WScript.Arguments.Named.Item("defgs"), ",")
  For Each group In groupNameArray
    trimmedGroup = Trim(group)                            'Trim the groupname incase of stray spaces.
    if trimmedGroup <> "" then
	  GroupDN = GetGroupDN(trimmedGroup)
	  Set grpObj = GetObject(serverPrefix & GroupDN)
	  CheckError("AU-AddToGroup/1")
	  if (grpObj.IsMember(serverPrefix & usrName) = 0) then
	    grpObj.Add(serverPrefix & usrName)
	  end if
	  CheckError("AU-AddToGroup/2")
	  set grpObj = Nothing
    end if
  Next
  CheckError("AU-AddGroups")

  'SetCompany
  'Must be set before adding a mail
  SetCompany
  CheckError("AU-AfterSC")

  'Create or restore the users mailbox
  'then set UPN
  AddMail
  CheckError("AU-AfterAM")

  WScript.Echo "OK, bruker " & LoginName & " ble opprettet."
  CheckError("AU-end")
  WScript.Quit
end Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Remove User
Sub RemoveUser
  On Error Resume Next
  Err.Clear
  Dim UserDN, objUser, objOU, LoginName, arrMemberOf, strGroupName, objGroup, Group, FullName
  LoginName = WScript.Arguments.Named.Item("login")
  CheckError("ru-Begin")

  UserDN = GetUserDN(LoginName)
  CheckError("ru-GetUserDN")
  if Not ObjectExists(serverPrefix & UserDN) then
    WScript.Echo "Bruker " & LoginName & " eksisterer ikke."  'Echo the error so the user gets feedback
    WScript.Quit
  end if

  Set objUser = GetObject(serverPrefix & UserDN)          'Get the User object
  FullName = objUser.displayName
  CheckError("ru-SetobjUser")

  arrMemberOf = objUser.GetEx("MemberOf")                 'Get groups
  if Not Err.Number <> 0 then                             'Check if user is in any groups
    For Each Group in arrMemberOf                         'loop over groups and remove user from them
      Set objGroup = GetObject(serverPrefix & Group)      'Create a Group object
      CheckError("ru-GetGroupObj")
      strGroupName = replace(objGroup.Name, "CN=", "")    'Create a string so we can match the group with "Domain Users"
      if (strGroupName <> "Domain Users") then            'See if the groupname is anything else than "Domain Users"
        objGroup.Remove(serverPrefix & UserDN)            'Remove the group
        objGroup.SetInfo
      end if
    CheckError("ru-InsideGroupLoop")
    Next
  else
    Err.Clear                                             'Clear the error if error number is anything other than 0(= no error)
  end if
  CheckError("ru-GroupLoop")

  if Not objUser.AccountDisabled then                     'Check if the account is disabled, if enabled -> continue
    objUser.AccountDisabled = True                        'Disable the user
    objUser.SetInfo                                       'Commit changes(Disable user)
  end if
  CheckError("ru-DisableUser")

  Set objOU = GetObject(serverPrefix & ouDeleted & "," & ouCustomers & "," & ouDC)  'Set the OU for where the user should be moved
  objOU.MoveHere serverPrefix & UserDN, vbNullString                                'Move the user
  if Err.Number = -2147019886 then                                                  'check if the action returned the "-2147019886" error number.
    Err.Clear                                                                       'clear the error since we caught it. This means there already exists a user with that name in the OU.
    objOU.MoveHere serverPrefix & UserDN, "CN=" & FullName & "2"                    'Rename the user, do this part of the code last so we dont have the get the user object again.
    UserDN = GetUserDN(LoginName)                                                   'get the userdn again since we renamed the user
    objOU.MoveHere serverPrefix & UserDN, vbNullString                              'Then move the user.
  end if
  CheckError("ru-MoveobjToOU")
  
  'Call the removemail method
  RemoveMail
  CheckError("ru-afterRM")

  'flag the user for removal
  FlagUser
  CheckError("ru-afterFU")
  
  'Return to the UI
  WScript.Echo "OK, bruker " & LoginName & " deaktivert og flyttet til ny OU"
  CheckError("ru-end")
end Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Group Operations

Function ObjectExists(ByVal ObjectName)
  On Error Resume Next
  Err.Clear
  Dim grp
  Set grp = GetObject(ObjectName)
  if Err.Number = 0 then
  ObjectExists = True
  else
  ObjectExists = False
  end if
  Err.Clear
end Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Group Operations

Function GetUserDN(ByVal LoginName)
  On Error Resume Next
  Err.Clear
  GetUserDN = ""
  Const ADS_SCOPE_SUBTREE = 2
  Dim objConnection, objCommand, objRecordSet
  Set objConnection = CreateObject("ADODB.Connection")
  Set objCommand =   CreateObject("ADODB.Command")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  Set objCommand.ActiveConnection = objConnection
  objCommand.Properties("Page Size") = 1000
  objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
  objCommand.CommandText = "SELECT distinguishedName FROM '" & serverPrefix & ouDC & "' WHERE objectCategory = 'user' AND sAMAccountName = '" & LoginName & "'"

  Set objRecordSet = objCommand.Execute

  Do While Not objRecordSet.EOF
  GetUserDN = objRecordSet.Fields("distinguishedName").Value
  objRecordSet.MoveNext
  Loop

  objRecordSet.Close
end Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Function GetGroupDN(ByVal GroupName)
  On Error Resume Next
  Err.Clear
  GetGroupDN = ""
  Const ADS_SCOPE_SUBTREE = 2
  Dim objConnection, objCommand, objRecordSet
  Set objConnection = CreateObject("ADODB.Connection")
  Set objCommand =   CreateObject("ADODB.Command")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  Set objCommand.ActiveConnection = objConnection
  objCommand.Properties("Page Size") = 1000
  objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
  objCommand.CommandText = "SELECT distinguishedName FROM 'LDAP://" & ouDC & "' WHERE objectCategory = 'group' AND sAMAccountName = '" & GroupName & "'"

  Set objRecordSet = objCommand.Execute

  Do While Not objRecordSet.EOF
    GetGroupDN = objRecordSet.Fields("distinguishedName").Value
    objRecordSet.MoveNext
  Loop

  objRecordSet.Close
end Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' AddRemoveGroupMembership
Sub   AddRemoveGroupMembership (byVal Op)
  
  On Error Resume Next
  Err.Clear
  
  LoginName = WScript.Arguments.Named.Item("login")
  GroupName = WScript.Arguments.Named.Item("group")
  if LoginName = "" or GroupName = "" then
    QuitWithError("Scriptet fikk ikke alle argument.")
  end if
  
  Dim grpName
  grpName = GetGroupDN(GroupName)
  if Not ObjectExists(serverPrefix & grpName) then
    WScript.Echo "Gruppe " & grpName & " eksisterer ikke."
    WScript.Quit
  end if
  
  Dim usrName
  
  usrName = GetUserDN(LoginName)

  if Not ObjectExists(serverPrefix & usrName) then
    WScript.Echo "Bruker " & LoginName & " eksisterer ikke."
    WScript.Quit
  end if

  Dim objGroup
  Dim objUser 
  Dim grpList 
  Dim grpMember 
  Set objGroup = GetObject(serverPrefix & grpName)
  Set objUser = GetObject(serverPrefix & usrName)
  const ADS_PROPERTY_APPend = 3 
  Const ADS_PROPERTY_DELETE = 4  

  'grpList = LCase(Join(objUser.MemberOf))
  
  Dim vt
  vt = VarType(objUser.MemberOf)
  
  if vt = 0 then ' Empty
    grpList = ""
  elseif vt = 8 then ' String
    grpList = LCase(objUser.MemberOf)
  elseif vt = 8204 then ' Array of strings
    grpList = LCase(Join(objUser.MemberOf, "/"))
  else ' unknown type
    wscript.echo "Group error"
    wscript.quit
  end if

  grpMember = LCase(GroupName)

  if InStr(grpList, grpMember) then
    if Op = "rug" then

      objGroup.PutEx ADS_PROPERTY_DELETE, "member", Array(usrName)
      objGroup.SetInfo
      CheckError("rug")
      WScript.Echo "OK, bruker " & LoginName& " meldt UT av gruppe " & GroupName & "."
    elseif op = "aug" then
      WScript.Echo "OK, men bruker " & LoginName & " var med i gruppa " & GroupName & " fra før."
    end if
  else
    if Op = "aug" then

      objGroup.PutEx ADS_PROPERTY_APPend, "member", Array(usrName)
      objGroup.SetInfo
      CheckError("aug")
      WScript.Echo "OK, bruker " & LoginName & " meldt INN i gruppe " & GroupName & "."
    elseif op = "rug" then
       WScript.Echo "OK, men bruker " & LoginName & " var ikke med i gruppa " & GroupName & " fra før."
    end if
  end if



  Set objGroup = Nothing


end Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Sub AddMail
  'On Error Resume Next
  Err.Clear
  CheckError("am-Start")
  Dim LoginName, sMailDB, PSCommand, PSLoc, PSCommand2, wsh, wshRun, Ident
  LoginName = WScript.Arguments.Named.Item("login")
  sMailDB = WScript.Arguments.Named.Item("mdb")
  Ident = WScript.Arguments.Named.Item("fulln")
  CheckError("am-Setarg")

  if LoginName = "" or sMailDB = "" then
  QuitWithError("am-CKArgs")
  end if

  Set wsh = WScript.CreateObject("WScript.Shell")
  wsh.CurrentDirectory = scriptdir
  PSLoc = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
  CheckError("am-SetShell")
  
  'Create mailbox
  PSCommand = PSLoc & " " & scriptdir & "\Powershell\EnableMailbox.ps1"" -User '" & LoginName & "' -MailDB '" & sMailDB & "' -DC '" & domainController & "' -Ident '" & Ident & "'"
  wshRun = wsh.run(PSCommand, 0, True)
  CheckError("am-Enable")

  'SetUPN
  PSCommand2 = PSLoc & " " & scriptdir & "\Powershell\SetUPN.ps1"" -User '" & LoginName & "' -DC '" & domainController & "'"
  wshRun = wsh.run(PSCommand2, 0, True)
  CheckError("am-UPN")
  
  'Cleanup variables
  Set wshRun = nothing
  Set wsh = nothing
  CheckError("am-cleanup")
  'Report to Interface
  WScript.Echo "OK, brukerens postboks er klargjort."
  CheckError("am-echo")
  Err.Clear
end Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Sub RemoveMail
  CheckError("rm-begin")
  Dim PSCommand, PSLoc, wshRun, wshShell, LoginName
  LoginName = WScript.Arguments.Named.Item("login")

  Set wshShell = WScript.CreateObject("WScript.Shell")
  wshShell.CurrentDirectory = scriptdir
  PSLoc = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
  'Disable mailbox
  PSCommand = PSLoc & " " & scriptdir & "\Powershell\DisableMailbox.ps1"" -User " & LoginName & " -DC " & domainController & ""
  CheckError("rm-Build PowerShell Command")
  wshRun = wshShell.run(PSCommand, 0, True)
  CheckError("rm-Execute PowerShell Command")
  Set wshRun = nothing
  Set wshShell = nothing
end Sub

Sub FlagUser
  CheckError("fu-begin")
  Dim PSCommand, PSLoc, wshRun, wshShell, LoginName
  LoginName = WScript.Arguments.Named.Item("login")                     'Get the loginname
  Set wshShell = WScript.CreateObject("WScript.Shell")                  'Create shell object to run powershell
  wshShell.CurrentDirectory = scriptdir
  PSLoc = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"   'Set powershell location
  'Build the powershell command
  PSCommand = PSLoc & " " & scriptdir & "\Powershell\FlagUser.ps1"" -username " & LoginName & ""
  CheckError("fu-Build PowerShell Command")
  wshRun = wshShell.run(PSCommand, 0, True)
  CheckError("fu-Execute PowerShell Command")
end Sub

Sub UnflagUser
  CheckError("ufu-begin")
  Dim PSCommand, PSLoc, wshRun, wshShell, LoginName
  LoginName = WScript.Arguments.Named.Item("login")                     'Get the loginname
  Set wshShell = WScript.CreateObject("WScript.Shell")                  'Create shell object to run powershell
  wshShell.CurrentDirectory = scriptdir
  PSLoc = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"   'Set powershell location
  'Build the powershell command
  PSCommand = PSLoc & " " & scriptdir & "\Powershell\UnflagUser.ps1"" -username " & LoginName & ""
  CheckError("ufu-Build PowerShell Command")
  wshRun = wshShell.run(PSCommand, 0, True)
  CheckError("ufu-Execute PowerShell Command")
end Sub

Sub SetCompany
  'On Error Resume Next
  Err.Clear
  Dim LoginName, wsh, PSLoc, PSCommand, wshRun
  LoginName = WScript.Arguments.Named.Item("login")
  CheckError("sc-Start")

  Set wsh = WScript.CreateObject("WScript.Shell")
  wsh.CurrentDirectory = scriptdir
  PSLoc = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
  CheckError("sc-SetShell")

  'SetCompany
  PSCommand = PSLoc & " " & scriptdir & "\Powershell\SetCompany.ps1"" -User '" & LoginName & "' -DC '" & domainController & "'"
  wshRun = wsh.run(PSCommand, 0, True)
  CheckError("sc-Company")
  
  'Cleanup variables
  Set wshRun = nothing
  Set wsh = nothing
  CheckError("sc-cleanup")
  Err.Clear
end Sub
