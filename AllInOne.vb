'This file is part of File Server Audit.

'File Server Audit is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'File Server Audit is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with Foobar.  If not, see <http://www.gnu.org/licenses/>.


'Class and Tread Framework 
'written by Margus Martsepp AKA m2s87
'http://www.dreamincode.net/code/snippet875.htm

'GetUser from: http://www.codeproject.com/KB/system/active_directory_in_vbnet.aspx?display=Print

Imports Str = Microsoft.VisualBasic.Strings
Public Class AllInOne : Inherits Main
    Dim blnShowOutput As Boolean = True 'Used to show output
    Dim blnInherited As Boolean = False 'Used to remove inherited ACLs
    Dim setRS As Boolean = True 'Recursive_Search Not used? (Public Property Recursive_Search)
    Dim blnSQL As Boolean = False ' Used to send output to SQL Connection

    'Dim intError As Integer = 0
    Dim iniMaxCount As Integer = 0 'Used to show Total Max folder Count
    Dim iniPCount As Integer = 0 'Used to show current folder Count 

    Dim strLog As String 'Used to hold Log Path
    'Dim StrSCN As String
    Dim strSearch As String = Nothing 'Used To Search AD? (Public Property Check_AD)
    Dim strDate As String = TodayFileDate() 'Holds Date in string format
    Dim StrGlSQLCommand As String
    Dim intGLSQLCount As Integer = 0
    Dim intGLSQLMaxCount As Integer = 500
    Dim arrDrives() As String 'Holds selected Drive

    Dim dicADUserType As New Dictionary(Of String, String) 'Hold Cache for AD Users/Groups (AD sAMAccountName, AD Object Type)
    Dim dicADGroup As New Dictionary(Of String, Array) 'Holds Group sAMAccountName and Array of Users sAMAccountName
    Dim dicADGroupManager As New Dictionary(Of String, String) 'Holds Group sAMAccountName and Managed By name
    Dim dicADGroupSubGroup As New Dictionary(Of String, Array) 'Holds Group sAMAccountName and Sub-Group sAMAccountName

    Public objLogFile As System.IO.TextWriter

    Public Event CurrentRecord(ByVal CurrentRecordProcess As String)
    Public Event CurrentACLCountMax(ByVal CurrentACLMaxNumber As Integer)
    Public Event CurrentACLCount(ByVal CurrentACLNumber As Integer)
    Public Event BGDone(ByVal intDone As Integer)

    Private EndedAt As String, StartedAt As String
    Private StartTime As DateTime, EndTime As DateTime, ElapsedTime As TimeSpan
    Private tegutseb As Boolean = False, Notifieonend As Boolean

    Private WithEvents BackgroundWorker1 As New System.ComponentModel.BackgroundWorker

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, _
    ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) _
    Handles BackgroundWorker1.RunWorkerCompleted ' Compleated
        tegutseb = False
        EndTime = Now
        If Notifieonend = True Then MsgBox(Timestamp, , "Process Done")
    End Sub
    Public ReadOnly Property Timestamp() As String
        Get
            StartedAt = "Started     :" & vbTab & Format(StartTime, "HH:mm:ss.fff")
            EndedAt = "Ended       :" & vbTab & Format(EndTime, "HH:mm:ss.fff")
            ElapsedTime = EndTime.Subtract(StartTime)
            Return StartedAt & Chr(13) & EndedAt & Chr(13) _
            & "Total Time :" & vbTab & ElapsedTime.Hours & ":" & ElapsedTime.Minutes & ":" & ElapsedTime.Seconds & "." & ElapsedTime.Milliseconds
        End Get
    End Property
    Public Property IsWorking() As Boolean
        Get
            Return tegutseb
        End Get
        Set(ByVal value As Boolean)
            If value = False Then
                BackgroundWorker1.WorkerSupportsCancellation = True
                BackgroundWorker1.CancelAsync()
            End If
        End Set
    End Property
    Public Property Notifie_on_end() As Boolean
        Get
            Return Notifieonend
        End Get
        Set(ByVal value As Boolean)
            Notifieonend = value
        End Set
    End Property
    Public Property Show_Output() As Boolean
        Get
            Return blnShowOutput
        End Get
        Set(ByVal value As Boolean)
            blnShowOutput = value
        End Set
    End Property
    Public Property Remove_Inherited() As Boolean
        Get
            Return blnInherited
        End Get
        Set(ByVal value As Boolean)
            blnInherited = value
        End Set
    End Property
    Public Property Use_SQL() As Boolean
        Get
            Return blnSQL
        End Get
        Set(ByVal value As Boolean)
            blnSQL = value
        End Set
    End Property
    Public Property Recursive_Search() As Boolean
        Get
            Return setRS
        End Get
        Set(ByVal value As Boolean)
            setRS = value
        End Set
    End Property
    Public Property Check_Drives() As Array
        Get
            Return arrDrives
        End Get
        Set(ByVal value As Array)
            arrDrives = value
        End Set
    End Property
    Public Property Check_AD() As String
        Get
            Return strSearch
        End Get
        Set(ByVal value As String)
            strSearch = value
        End Set
    End Property
    Public Property Log_File() As String
        Get
            Return strLog
        End Get
        Set(ByVal value As String)
            strLog = value
        End Set
    End Property
    Public Sub startBackgroundTask() ' This will start the backgroundworker
        tegutseb = True
        StartTime = Now
        'Check to make sure a log file is passed
        If Not strLog = "" Or blnSQL Then
            'Check to make sure that a drive was selected 
            If arrDrives.Length > 0 Then
                If Not blnSQL Then
                    objLogFile = New System.IO.StreamWriter(strLog, True)
                Else
                    Try
                        'objSQLConnection.Open()
                        SQL_Int()
                    Catch ex As Exception
                        MsgBox("Cannot connect to database" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "SQL Error")
                    End Try

                End If
                BackgroundWorker1.WorkerSupportsCancellation = True
                BackgroundWorker1.RunWorkerAsync()
            Else
                RaiseEvent BGDone(1)
                Threading.Thread.Sleep(0)
                Exit Sub
            End If
        Else
            RaiseEvent BGDone(2)
            Threading.Thread.Sleep(0)
            Exit Sub
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, _
    ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ' Add your code here
        Dim strDrive As String
        'Prints Header Information
        If Not blnSQL Then LogWrite(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", "Folder Path", "Top Level SAMAccountName", "Group SAMAccountName", "Managed By", "Inheritance", "Rights", "Owner", "Run Date", "Computer", "IsInherited"))
        'Loops through all drives selected 
        If arrDrives.Length > 0 Then
            For Each strDrive In arrDrives
                If Not strDrive = "" Then
                    'Dim ObjFolder As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(strDrive)
                    'Reads Discretionary Access Control List of Drive to LogWrite
                    ReadACLs(strDrive)
                    'Reads Discretionary Access Control List of sub-folders to Logwrite
                    RecursiveSearch(strDrive)
                End If
            Next
        End If
        If blnSQL Then
            If objSQLConnection.state = ConnectionState.Open Then objSQLConnection.Close()
        Else
            'Closes Log File
            objLogFile.Flush()
            objLogFile.Close()
            objLogFile.Dispose()
        End If


        RaiseEvent CurrentRecord("Done Running Thru Permissions")
        Threading.Thread.Sleep(0)
        RaiseEvent BGDone(0)
        Threading.Thread.Sleep(0)
        Main.Start.Text = "Start"

    End Sub
    'Loops through all folders and Subfolders
    Public Sub RecursiveSearch(ByVal strFolderFullName As String)
        'Dim ObjSubFolders As System.IO.DirectoryInfo
        Dim iniLocalMaxCount As Integer
        Dim SkipSubFolders As Boolean = False
        'Dim strFolderFullName As String = ObjFolder.FullName
        Dim strSubFolderFullName As String = ""



        On Error GoTo ErrorHandle

        'Sends numbers back to status bar
        'This is buggy and results in status bar jumping all over the place. 
        'Need and easy way to get count of total amount of folders on drive.
        iniLocalMaxCount = System.IO.Directory.GetDirectories(strFolderFullName).GetUpperBound(0)
        If iniLocalMaxCount >= 0 Then
            iniMaxCount = iniMaxCount + iniLocalMaxCount
            RaiseEvent CurrentACLCount(iniMaxCount)
            Threading.Thread.Sleep(0)
            'Read and loops through all subfolders
            For Each strSubFolderFullName In System.IO.Directory.GetDirectories(strFolderFullName)
                'ObjSubFolders = New System.IO.DirectoryInfo(strSubFolderFullName)
                'Puts Full path in Varible for Debuging
                'strSubFolderFullName = ObjSubFolders.FullName
                'Reads Discretionary Access Control List of Folder
                ReadACLs(strSubFolderFullName)
                'If and Error occurs then recursion is stop
                'SkipSubFolders set in error handler 
                If SkipSubFolders = False Then
                    iniPCount = iniPCount + 1
                    RaiseEvent CurrentACLCount(iniPCount)
                    Threading.Thread.Sleep(0)
                    'Calls itself to perform a recurse into subdirectories
                    RecursiveSearch(strSubFolderFullName)
                End If
                SkipSubFolders = False
            Next
        End If

ErrorHandle:
        'If error happens the log error and resume next.
        Select Case Err.Number
            Case 0
                CurrentRecord_Event("ERROR: " & strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                Threading.Thread.Sleep(0)
                Err.Clear()
            Case 5 'Access Denied
                If Not strSubFolderFullName = "" Then
                    LogWrite(strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Threading.Thread.Sleep(0)
                    Err.Clear()
                    SkipSubFolders = True
                    Resume Next
                Else
                    LogWrite(strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Threading.Thread.Sleep(0)
                    Err.Clear()
                    SkipSubFolders = True
                    Resume Next
                End If
            Case 9 'Index was Out Side of bounds
                If Not strSubFolderFullName = "" Then
                    LogWrite(strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Threading.Thread.Sleep(0)
                    Err.Clear()
                    'SkipSubFolders = True
                    'Resume Next
                Else
                    LogWrite(strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Threading.Thread.Sleep(0)
                    Err.Clear()
                    'SkipSubFolders = True
                    'Resume Next
                End If
            Case 20 'Resume without error
                CurrentRecord_Event("ERROR: " & strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                Threading.Thread.Sleep(0)
                Err.Clear()
            Case 91 'Object reference not set to an instance of an object.
                If Not strSubFolderFullName = "" Then
                    LogWrite(strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Threading.Thread.Sleep(0)
                    Err.Clear()
                    'SkipSubFolders = True
                    'Resume Next
                Else
                    LogWrite(strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Err.Clear()
                    'SkipSubFolders = True
                    'Resume Next
                End If
            Case Else
                If Not strSubFolderFullName = "" Then
                    LogWrite(strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strSubFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Threading.Thread.Sleep(0)
                    Err.Clear()
                    SkipSubFolders = True
                    Resume Next
                Else
                    LogWrite(strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    CurrentRecord_Event("ERROR: " & strFolderFullName & "; " & Err.Number & ";" & Err.Description & ";" & Err.Source & ";" & New System.Diagnostics.StackTrace().GetFrame(0).GetMethod.ToString() & ";;;" & strDate & ";")
                    Threading.Thread.Sleep(0)
                    Err.Clear()
                    SkipSubFolders = True
                    Resume Next
                End If
        End Select
    End Sub
    'Reads Discretionary Access Control Lists of Folder
    'Has issues With Foreign Security Principal (Trusted Domains)
    Public Sub ReadACLs(ByVal strInput As String, Optional ByVal blnRemoveInherited As Boolean = True)
        'strInput is the full path to a folder
        'blnRemoveInherited is used to remove inherited ACLs to reduce output


        'Bindings:

        'Binds to the Directory using the string provided 
        Dim DirInfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(strInput) 'Directory is already binded
        'Binds to Directory Security
        'This allows user to read ACLs
        Dim DirSec As System.Security.AccessControl.DirectorySecurity = DirInfo.GetAccessControl()
        'Gets Directory Access Control Lists
        'Collection is returned with ACLs so we can enumerate later
        Dim ACLs As System.Security.AccessControl.AuthorizationRuleCollection = DirSec.GetAccessRules(True, True, GetType(System.Security.Principal.NTAccount))
        'Access Contol List Entry
        'Varible to hold a ACL
        Dim ACL As System.Security.AccessControl.FileSystemAccessRule
        'Gets Folder owner ID
        Dim Owner As System.Security.Principal.IdentityReference = DirSec.GetOwner(GetType(System.Security.Principal.NTAccount))
        'Binds to the default domain
        'For this to work this will have to run under a domain account
        Dim DomainContext As New System.DirectoryServices.ActiveDirectory.DirectoryContext(System.DirectoryServices.ActiveDirectory.DirectoryContextType.Domain)
        'Binds to AD to get account and group info
        Dim objMember As New System.DirectoryServices.DirectoryEntry

        'Declarations:

        Dim arrDUSplit(1) As String 'Used to hold domain and user ex domain\user
        Dim strTemp As String = "" 'Holds working String
        Dim strSAMA As String = "" 'Users Username ("sAMAccountName")
        Dim strClass As String = "" ' AD Object type Group or User
        Dim strGroupSAMA As String = "" 'The Group Users is in
        Dim strManagedBy As String = "" 'Person that manages that AD Group.
        Dim strOwner As String = Owner.ToString 'Holds Owners name as string. Converts folders Owners Object to String
        'Dim strDate As String = TodayFileDate() 'Holds todays date in correct format 'Set as globe option.
        Dim StrACLIdentityReference As String = "" 'Holds ACL.IdentityReference.ToString Properties 
        Dim StrInheritanceFlags As String = "" 'Holds ACL.InheritanceFlags.ToString Properties
        Dim StrFileSystemRights As String = "" 'Holds ACL.FileSystemRights.ToString Properties
        Dim StrIsInherited As String = "" 'Holds ACL.IsInherited.ToString Properties
        'Dim strInput As String = DirInfo.FullName 'Holds Current Folder Full Name.
        Dim StrArrayLoopSubGroup As String 'Allows Looping thru all sub-group
        Dim StrArrayLoopUser As String 'Allows Looping thru all Users

        On Error GoTo ErrorHandle

        'Main Section of sub

        'Loops through all the Access Control Lists For the Folder
        For Each ACL In ACLs
            'Setup Most Used ACL Properties
            StrACLIdentityReference = ACL.IdentityReference.ToString
            StrInheritanceFlags = ACL.InheritanceFlags.ToString
            StrFileSystemRights = ACL.FileSystemRights.ToString
            StrIsInherited = ACL.IsInherited.ToString

            If Str.Left(StrACLIdentityReference, 2) = "S-" Then
                'Used to catch Deleted Accounts. 
                strSAMA = "Unknown"
                strGroupSAMA = StrACLIdentityReference
                strManagedBy = "Unknown"
                strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                    strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                If blnRemoveInherited Then
                    If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                Else
                    LogWrite(strTemp)
                End If
            Else
                'Splits Domain and User or Group a part.
                'This cleans up the output alittle.
                arrDUSplit = Split(StrACLIdentityReference, "\")
                'Formats the output differently depending on the domain. 
                Select Case arrDUSplit(0)
                    'Built-in security principals need special treatment   
                    Case "BUILTIN"
                        strSAMA = arrDUSplit(1)
                        strGroupSAMA = "System"
                        strManagedBy = "Systems Administrators"
                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                        If blnRemoveInherited Then
                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                        Else
                            LogWrite(strTemp)
                        End If
                    Case "NT AUTHORITY"
                        strSAMA = arrDUSplit(1)
                        strGroupSAMA = "System"
                        strManagedBy = "Systems Administrators"
                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                        If blnRemoveInherited Then
                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                        Else
                            LogWrite(strTemp)
                        End If
                    Case "CREATOR OWNER"
                        strSAMA = arrDUSplit(0)
                        strGroupSAMA = "System"
                        strManagedBy = "Systems Administrators"
                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                        If blnRemoveInherited Then
                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                        Else
                            LogWrite(strTemp)
                        End If
                    Case "Everyone"
                        strSAMA = arrDUSplit(0)
                        strGroupSAMA = "Everyone"
                        strManagedBy = "Systems Administrators"
                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                        If blnRemoveInherited Then
                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                        Else
                            LogWrite(strTemp)
                        End If

                    Case Environment.MachineName.ToString
                        'Used for Local Users that have the Domain of the computer name.
                        strSAMA = arrDUSplit(1)
                        strGroupSAMA = "Local"
                        strManagedBy = "Systems Administrators"
                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                        If blnRemoveInherited Then
                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                        Else
                            LogWrite(strTemp)
                        End If
                    Case Environment.UserDomainName.ToString
                        'Uses Currently login domain to enumerate groups
                        'Checks to see if the account type is is in the Cache 
                        If dicADUserType.ContainsKey(arrDUSplit(1)) Then
                            strSAMA = arrDUSplit(1)
                            strClass = dicADUserType(arrDUSplit(1)).ToString
                        Else
                            'If AD Users are not in the Cache Dictionary Check AD and Put them in the Cache
                            objMember = GetUser(arrDUSplit(1))
                            If objMember Is Nothing Then
                                CurrentRecord_Event("Error Number: " & Err.Number & " " & Err.Description & vbCrLf _
                                    & " Sub GetUser: " & vbCrLf _
                                    & vbTab & "DirInfo.FullName: " & DirInfo.FullName & vbCrLf _
                                    & vbTab & "NTFS ACL: " & StrACLIdentityReference & vbCrLf _
                                    & vbTab & "Domain: " & arrDUSplit(0) & vbCrLf _
                                    & vbTab & "User: " & arrDUSplit(1) & vbCrLf)
                            Else
                                strClass = objMember.SchemaClassName.ToString 'Get Record AD Type User or Group
                                strSAMA = objMember.Properties("sAMAccountName").Value.ToString 'Get AD UserName
                                dicADUserType.Add(strSAMA, strClass) 'sets username in Dictionary,'sets class in Dictionary
                            End If
                        End If

                        'Do different actions based on AD Class
                        Select Case strClass
                            'Write out Each User
                            'Each of this Users has explicit permissions
                            Case "user", "computer"
                                strGroupSAMA = "Direct"
                                strManagedBy = "Systems Administrators " & Environment.UserDomainName.ToString & " Domain Accounts"
                                strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                                    strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                                If blnRemoveInherited Then
                                    If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                                Else
                                    LogWrite(strTemp)
                                End If

                            Case "group"
                                'For Groups do a recursive search for all sub groups and users
                                'Check to see of Group is in Dictionary.
                                If dicADGroup.ContainsKey(arrDUSplit(1)) Then
                                    'Set Manager Of Group 
                                    strManagedBy = ""
                                    If dicADGroupManager.ContainsKey(arrDUSplit(1)) Then strManagedBy = dicADGroupManager(arrDUSplit(1))
                                    'Loop thru all users in that group.
                                    For Each StrArrayLoopUser In dicADGroup(arrDUSplit(1)) 'Loops thru all user in a group
                                        'AD Group was Found in Cache.
                                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                                        If blnRemoveInherited Then
                                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                                        Else
                                            LogWrite(strTemp)
                                        End If
                                    Next
                                    'Loop thru all sub Groups
                                    If dicADGroupSubGroup.ContainsKey(arrDUSplit(1)) Then
                                        For Each StrArrayLoopSubGroup In dicADGroupSubGroup(arrDUSplit(1)) 'Loops thru all sub-group
                                            strManagedBy = ""
                                            If dicADGroupManager.ContainsKey(StrArrayLoopSubGroup) Then strManagedBy = dicADGroupManager(arrDUSplit(1))
                                            'Test SubGroup to see if the are in cache
                                            If Not dicADGroup.ContainsKey(StrArrayLoopSubGroup) Then
                                                If Not RecursiveGroupSub(objMember) = StrArrayLoopSubGroup Then GoTo ErrorHandle
                                            End If
                                            For Each StrArrayLoopUser In dicADGroup(StrArrayLoopSubGroup) 'Loops thru all user in a group
                                                'AD Group was Found in Cache.
                                                strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                                                    strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                                                If blnRemoveInherited Then
                                                    If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                                                Else
                                                    LogWrite(strTemp)
                                                End If
                                            Next
                                        Next
                                    Else

                                    End If
                                Else
                                    'Group is not is Cache so start RecursiveGroupSub to put it in Cache
                                    'RecursiveGroupSub should return the group it is worked on
                                    If RecursiveGroupSub(objMember) = arrDUSplit(1) Then
                                        If Not dicADGroup.ContainsKey(arrDUSplit(1)) Then GoTo ErrorHandle
                                    Else
                                        If Not dicADGroup.ContainsKey(RecursiveGroupSub(objMember)) Then GoTo ErrorHandle
                                    End If
                                    'Set the Manager of the group
                                    If dicADGroupManager.ContainsKey(arrDUSplit(1)) Then strManagedBy = dicADGroupManager(objMember.Properties("sAMAccountName").Value.ToString)
                                    'Loop thru all user on the Group
                                    For Each StrArrayLoopUser In dicADGroup(arrDUSplit(1)) 'Loops thru all user in a group
                                        'AD Group was Found in Cache.
                                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                                        If blnRemoveInherited Then
                                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                                        Else
                                            LogWrite(strTemp)
                                        End If
                                    Next
                                    'Loop thru all sub Groups
                                    If dicADGroupSubGroup.ContainsKey(arrDUSplit(1)) Then
                                        For Each StrArrayLoopSubGroup In dicADGroupSubGroup(arrDUSplit(1)) 'Loops thru all sub-group
                                            strManagedBy = ""
                                            If dicADGroupManager.ContainsKey(StrArrayLoopSubGroup) Then strManagedBy = dicADGroupManager(StrArrayLoopSubGroup)
                                            For Each StrArrayLoopUser In dicADGroup(StrArrayLoopSubGroup) 'Loops thru all user in a group
                                                'AD Group was Found in Cache.
                                                strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                                                    strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                                                If blnRemoveInherited Then
                                                    If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                                                Else
                                                    LogWrite(strTemp)
                                                End If
                                            Next
                                        Next
                                    Else

                                    End If
                                End If
                            Case "contact"

                            Case Else
                                MsgBox("Unexpected strClass: " & strClass)
                        End Select
                    Case Else
                        'Catch all just Write out Entries without group enumeration
                        'If you have multiple domains in your forest the domain you are not logon on too will show up here.
                        strSAMA = arrDUSplit(1)
                        strGroupSAMA = arrDUSplit(0)
                        strManagedBy = arrDUSplit(1) & " Administrators"
                        strTemp = String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9}", Str.Replace(DirInfo.FullName, ";", ","), strSAMA, strGroupSAMA, _
                                            strManagedBy, StrInheritanceFlags, StrFileSystemRights, strOwner, strDate, Environment.MachineName.ToString, StrIsInherited)
                        If blnRemoveInherited Then
                            If ACL.IsInherited = False Or Str.Len(strInput) >= 3 Then LogWrite(strTemp)
                        Else
                            LogWrite(strTemp)
                        End If
                End Select
            End If
        Next
        Exit Sub
ErrorHandle:
        Select Case Err.Number
            Case 0
                strTemp = "Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                    & " Sub ReadACLs: " & vbCrLf _
                    & vbTab & "DirInfo.FullName: " & strInput & vbCrLf _
                    & vbTab & "NTFS ACL:" & StrACLIdentityReference & vbCrLf _
                    & vbTab & "User:" & strSAMA & vbCrLf _
                    & vbTab & "arrDUSplit(0): " & arrDUSplit(0) & vbCrLf _
                    & vbTab & "arrDUSplit(1): " & arrDUSplit(1) & vbCrLf
                CurrentRecord_Event(strTemp)
                Threading.Thread.Sleep(0)
                Err.Clear()
            Case 5 'The given key was not present in the dictionary.
                'Log error in the GUI
                strTemp = "Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                    & " Sub ReadACLs: " & vbCrLf _
                    & vbTab & "DirInfo.FullName: " & strInput & vbCrLf _
                    & vbTab & "NTFS ACL:" & StrACLIdentityReference & vbCrLf _
                    & vbTab & "User:" & strSAMA & vbCrLf _
                    & vbTab & "arrDUSplit(0): " & arrDUSplit(0) & vbCrLf _
                    & vbTab & "arrDUSplit(1): " & arrDUSplit(1) & vbCrLf
                CurrentRecord_Event(strTemp)
                Threading.Thread.Sleep(0)
                Err.Clear()
                Resume Next
            Case 91 'Object reference not set to an instance of an object
                'Log error in the GUI
                strTemp = "Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                    & " Sub ReadACLs: " & vbCrLf _
                    & vbTab & "DirInfo.FullName: " & strInput & vbCrLf _
                    & vbTab & "NTFS ACL:" & StrACLIdentityReference & vbCrLf _
                    & vbTab & "User:" & strSAMA & vbCrLf _
                    & vbTab & "arrDUSplit(0): " & arrDUSplit(0) & vbCrLf _
                    & vbTab & "arrDUSplit(1): " & arrDUSplit(1) & vbCrLf
                CurrentRecord_Event(strTemp)
                Threading.Thread.Sleep(0)
                Err.Clear()
                Resume Next
            Case Else
                'Log error in the GUI
                strTemp = "Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                    & " Sub ReadACLs: " & vbCrLf _
                    & vbTab & "DirInfo.FullName: " & strInput & vbCrLf _
                    & vbTab & "NTFS ACL:" & StrACLIdentityReference & vbCrLf _
                    & vbTab & "User:" & strSAMA & vbCrLf _
                    & vbTab & "arrDUSplit(0): " & arrDUSplit(0) & vbCrLf _
                    & vbTab & "arrDUSplit(1): " & arrDUSplit(1) & vbCrLf
                CurrentRecord_Event(strTemp)
                Threading.Thread.Sleep(0)
                'Message Box about the error click ok to quit or X to continue
                If MsgBox("Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                    & " Sub ReadACLs: " & vbCrLf _
                    & vbTab & "Input Directory: " & strInput & vbCrLf _
                    & vbTab & "NTFS ACL:" & StrACLIdentityReference & vbCrLf _
                    & vbTab & "User:" & strSAMA & vbCrLf _
                    & vbTab & "arrDUSplit(0): " & arrDUSplit(0) & vbCrLf _
                    & vbTab & "arrDUSplit(1): " & arrDUSplit(1) & vbCrLf _
                    , MsgBoxStyle.Critical, "Critical Error File Server Audit 2 Quitting") Then
                    Err.Clear()
                    Main.Close()
                Else
                    Err.Clear()
                    Exit Sub
                End If
        End Select
    End Sub
    'Recursively goes through all subgroups and gets all user in all sub groups
    'If error happens function will return nothing.
    Public Function RecursiveGroupSub(ByVal ADGroup As System.DirectoryServices.DirectoryEntry, Optional ByVal lisinfiniteloop As System.Collections.Generic.List(Of String) = Nothing) As String
        Dim objMember As New System.DirectoryServices.DirectoryEntry
        Dim arrUsers(0) As String
        Dim arrSubGroups(0) As String
        Dim strsAMAccountName As String
        Dim strMembersAMAccountName As String
        Dim strMember As String
        Dim strManagedBy As String
        Dim lisWorking As New List(Of String)
        Dim strTemp As String = ""
        Dim strLoop As String

        On Error GoTo ErrorHandle

        If Not lisinfiniteloop Is Nothing Then lisWorking = lisinfiniteloop

        'Checks to make sure the account has a username and Checks to make sure the account is a group
        If ADGroup.Properties("sAMAccountName").Value Is Nothing Then
            CurrentRecord_Event("Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                & "Function RGS: " & vbCrLf _
                & vbTab & "AD SAM Name: " & ADGroup.Properties("sAMAccountName").Value & vbCrLf _
                & vbTab & "AD Class: " & ADGroup.SchemaClassName.ToString & vbCrLf)
            Threading.Thread.Sleep(100)
            Return Nothing
        Else
            strsAMAccountName = ADGroup.Properties("sAMAccountName").Value.ToString
            'Add current Group to circular dependency list
            lisWorking.Add(strsAMAccountName)
            If Not dicADGroup.ContainsKey(strsAMAccountName) Then
                If ADGroup.SchemaClassName.ToString = "group" Then
                    'Gets Managed By Info for Group
                    If Not ADGroup.Properties("managedby").Value Is Nothing Then
                        objMember = New System.DirectoryServices.DirectoryEntry("LDAP://" & ADGroup.Properties("managedby").Value.ToString)
                        strManagedBy = objMember.Properties("DisplayName").Value.ToString
                        If Not strManagedBy = "" Then
                            If Not dicADGroupManager.ContainsKey(strsAMAccountName) Then dicADGroupManager.Add(strsAMAccountName, strManagedBy)
                        End If
                    End If

                    'Loops through all members of the group
                    For Each strMember In ADGroup.Properties("member")
                        strMembersAMAccountName = ""
                        'Binds to each Object in the member list 
                        objMember = New System.DirectoryServices.DirectoryEntry("LDAP://" & strMember)
                        'Gets the members AD class
                        Select Case LCase(objMember.SchemaClassName.ToString)
                            Case "user", "computer"
                                'Sets Username and User Class
                                strMembersAMAccountName = objMember.Properties("sAMAccountName").Value.ToString
                                If Not dicADUserType.ContainsKey(strMembersAMAccountName) Then
                                    dicADUserType.Add(strMembersAMAccountName, LCase(objMember.SchemaClassName.ToString))
                                End If
                                'Adds user to array
                                If arrUsers(0) = "" Then
                                    arrUsers(0) = strMembersAMAccountName
                                Else
                                    ReDim Preserve arrUsers(arrUsers.GetUpperBound(0) + 1)
                                    arrUsers(arrUsers.GetUpperBound(0)) = strMembersAMAccountName
                                End If
                            Case "group"
                                'Sets Group and Group Class
                                strMembersAMAccountName = objMember.Properties("sAMAccountName").Value.ToString
                                If Not dicADUserType.ContainsKey(strMembersAMAccountName) Then
                                    dicADUserType.Add(strMembersAMAccountName, LCase(objMember.SchemaClassName.ToString))
                                End If
                                'Adds Group to Array
                                If dicADGroup.ContainsKey(strMembersAMAccountName) Then
                                    If arrSubGroups(0) = "" Then
                                        arrSubGroups(0) = strMembersAMAccountName
                                    Else
                                        ReDim Preserve arrSubGroups(arrSubGroups.GetUpperBound(0) + 1)
                                        arrSubGroups(arrSubGroups.GetUpperBound(0)) = strMembersAMAccountName
                                    End If
                                Else
                                    'Function calls it self to populate sub-Groups
                                    'Add Group to Array
                                    'Checks list to see if there is a circular dependency
                                    If Not lisWorking.Contains(strMembersAMAccountName) Then
                                        If arrSubGroups(0) = "" Then
                                            arrSubGroups(0) = RecursiveGroupSub(objMember, lisWorking) 'Calls itself for recussive search
                                        Else
                                            ReDim Preserve arrSubGroups(arrSubGroups.GetUpperBound(0) + 1)
                                            arrSubGroups(arrSubGroups.GetUpperBound(0)) = RecursiveGroupSub(objMember, lisWorking) 'Calls itself for recussive search
                                        End If
                                    Else
                                        'Show all Groups for Error.
                                        For Each strLoop In lisWorking
                                            strTemp = strTemp & " " & strLoop
                                        Next
                                        'Writes Error that has happend
                                        LogWrite("Error: Circular Group Dependency; " & strMembersAMAccountName & " ; " & strTemp & "; ; ; ; ;" & strDate & ";")
                                        CurrentRecord_Event("Error: Circular Group Dependency; " & strMembersAMAccountName & " ; " & strTemp & "; ; ; ; ;" & strDate & ";")
                                    End If
                                End If
                            Case "contact"
                                CurrentRecord_Event("Error: Contact; " & strMembersAMAccountName & " ; " & strTemp & ";This error is not logged since contacts cant have ntfs rights. ; ; ;" & strDate & "; ;")
                            Case Else
                        End Select

                    Next

                    'Add Key to Cache
                    If Not arrUsers(0) = "" Then dicADGroup.Add(strsAMAccountName, arrUsers)

                    'Add key for Sub Groups
                    If Not arrSubGroups(0) = "" Then
                        If Not dicADGroupSubGroup.ContainsKey(strsAMAccountName) Then dicADGroupSubGroup.Add(strsAMAccountName, arrSubGroups)
                    End If
                    'Return Group name to calling sub
                    Return strsAMAccountName
                Else
                    Return Nothing
                End If
            Else
                Return strsAMAccountName
            End If
        End If

        strMember = Nothing
        objMember.Close()

ErrorHandle:
        Select Case Err.Number
            Case 0
                Err.Clear()
                'Case 91 'Null Object
                'Err.Clear()
                'Resume Next
            Case Else
                CurrentRecord_Event("Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                    & "Function RecursiveGroupSub: " & vbCrLf _
                    & vbTab & "AD SAM Name: " & ADGroup.Properties("sAMAccountName").Value & vbCrLf _
                    & vbTab & "AD Class: " & ADGroup.SchemaClassName.ToString & vbCrLf _
                    & vbTab & "Sub AD SAM Name:" & objMember.Properties("sAMAccountName").Value & vbCrLf _
                    & vbTab & "Sub AD Class: " & objMember.SchemaClassName.ToString & vbCrLf)
                Threading.Thread.Sleep(0)
                If MsgBox("Error Number:" & Err.Number & " " & Err.Description & vbCrLf _
                    & "Function RecursiveGroupSub: " & vbCrLf _
                    & vbTab & "AD SAM Name: " & ADGroup.Properties("sAMAccountName").Value & vbCrLf _
                    & vbTab & "AD Class: " & ADGroup.SchemaClassName.ToString & vbCrLf _
                    & vbTab & "Sub AD SAM Name:" & objMember.Properties("sAMAccountName").Value & vbCrLf _
                    & vbTab & "Sub AD Class: " & objMember.SchemaClassName.ToString & vbCrLf _
                    , MsgBoxStyle.YesNo, "Critical Error File Server Audit 2 Quitting") Then
                    Err.Clear()
                    strMember = Nothing
                    objMember.Close()
                    Return Nothing
                    Main.Close()
                Else
                    Err.Clear()
                    strMember = Nothing
                    objMember.Close()
                    Exit Function
                End If
        End Select
    End Function

    'Following Code from: http://www.codeproject.com/KB/system/active_directory_in_vbnet.aspx?display=Print
    'The Code Project Open License (CPOL) 
    '  http://www.codeproject.com/info/cpol10.aspx
    ''' <summary>
    ''' This will return a DirectoryEntry object if the user does exist
    ''' </summary>
    ''' <param name="UserName"></param>
    ''' <returns></returns>

    Public Function GetUser(ByVal UserName As String) As System.DirectoryServices.DirectoryEntry
        'create an instance of the DirectoryEntry
        Dim dirEntry As System.DirectoryServices.DirectoryEntry = Nothing

        'create instance fo the direcory searcher
        Dim dirSearch As New System.DirectoryServices.DirectorySearcher(dirEntry)

        dirSearch.SearchRoot = dirEntry
        'set the search filter
        'Changed the Filter From :
        'dirSearch.Filter = "(&(objectCategory=user)(cn=" + UserName + "))"
        'To:
        dirSearch.Filter = "(&(SAMAccountName=" + UserName + "))"

        'deSearch.SearchScope = SearchScope.Subtree;

        'find the first instance
        Dim searchResults As System.DirectoryServices.SearchResult = dirSearch.FindOne()

        'if found then return, otherwise return Null
        If Not searchResults Is Nothing Then
            'de= new DirectoryEntry(results.Path,ADAdminUser, _
            'ADAdminPassword,AuthenticationTypes.Secure);
            'if so then return the DirectoryEntry object
            Return searchResults.GetDirectoryEntry()
        Else
            Return Nothing
        End If
    End Function
    'Formats Date to be used in File output
    Private Function TodayFileDate() As String
        Dim strDate As String
        strDate = Now.Year
        If Now.Month < 10 Then
            strDate = strDate & "0" & Now.Month
        Else
            strDate = strDate & Now.Month
        End If
        If Now.Day < 10 Then
            strDate = strDate & "0" & Now.Day
        Else
            strDate = strDate & Now.Day
        End If
        Return strDate
    End Function
    Private Sub SQL_Int()
        If objSQLConnection = Nothing Then
            Select Case UCase(My.Settings.SQLServerType)
                Case "MSSQL"
                    objSQLConnection = New System.Data.SqlClient.SqlConnection(My.Settings.SQLConnection)
                    'StrDatabase = dicSQLInfo.Item("Initial Catalog")
                    StrDatabase = objSQLConnection.Database
                Case "MYSQL"
                    objSQLConnection = New MySql.Data.MySqlClient.MySqlConnection(My.Settings.SQLConnection)

                    'StrDatabase = dicSQLInfo.Item("database")
                    StrDatabase = objSQLConnection.Database
                Case Else
                    MsgBox("SQLServerType setting not understood: SQLServerType =" & My.Settings.SQLServerType, MsgBoxStyle.Critical, "SQL Server Type")
                    Main.Close()
            End Select
        End If
        If IsNothing(objSQLConnection) Then Throw New Exception("No Database Connection")
    End Sub
    Private Function SQL_Command(StrSQL_Command As String, Optional blnNoRecursive As Boolean = False) As String
        Dim IntOutput As Integer
        Dim objSQLCommand As Object
        Dim StrSingleCommand As String
        If Not objSQLConnection.state = ConnectionState.Open Then objSQLConnection.Open()
        Try
            Select Case UCase(My.Settings.SQLServerType)
                Case "MSSQL"
                    objSQLCommand = New System.Data.SqlClient.SqlCommand(StrSQL_Command, objSQLConnection)
                    IntOutput = objSQLCommand.ExecuteNonQuery()
                Case "MYSQL"
                    StrSQL_Command = StrSQL_Command.Replace("\", "\\")
                    StrSQL_Command = StrSQL_Command.Replace(":", "\:")
                    StrSQL_Command = StrSQL_Command.Replace(";", "\;")
                    StrSQL_Command = StrSQL_Command.Replace("')\; ", "'); ")
                    'objSQLCommand = New MySql.Data.MySqlClient.MySqlCommand(StrSQL_Command, objSQLConnection)
                    If objSQLCommand = Nothing Then objSQLCommand = New MySql.Data.MySqlClient.MySqlCommand
                    objSQLCommand.Connection = objSQLConnection
                    objSQLCommand.CommandText = StrSQL_Command
                    objSQLCommand.CommandType = CommandType.Text
                    IntOutput = objSQLCommand.ExecuteNonQuery()

                Case Else
                    MsgBox("SQLServerType setting not understood: SQLServerType =" & My.Settings.SQLServerType, MsgBoxStyle.Critical, "SQL Server Type")
                    Main.Close()
            End Select

        Catch ex As Exception
            If blnNoRecursive = False Then
                For Each StrSingleCommand In StrSQL_Command.Split(";")
                    IntOutput = SQL_Command(StrSingleCommand & ";", True)
                Next
            Else
                RaiseEvent CurrentRecord("Error Inserting SQL: " & StrSQL_Command)
            End If
        End Try
        Return IntOutput
    End Function
    'Writes to Log file, to Output Box in GUI and/or to SQL
    Private Sub LogWrite(ByVal strInput As String)
        If blnSQL Then
            'Logging to SQL Enabled
            'Dim objSQLCommand As System.Data.SqlClient.SqlCommand
            Dim StrSQLCommand As String
            Dim argTemp() As String
            Dim StrTemp As String
            'Dim strline As String
            Dim intRows As Integer

            argTemp = Split(Replace(strInput, "'", Chr(&H22)), ";")

            Try
                'Inserting records is buggy
                StrSQLCommand = "Insert into " & StrTableName & " (FolderPath, AccountSAMAccountName, GroupSAMAccountName," _
                        & " ManagedBy, Inheritance, Rights, Owner, RunDate, Computer, IsInherited" _
                        & " ) Values ( '" _
                        & argTemp(0) & "','" & argTemp(1) & "','" & argTemp(2) & "','" & argTemp(3) _
                        & "','" & argTemp(4) & "','" & argTemp(5) & "','" & argTemp(6) & "','" & argTemp(7) _
                        & "','" & argTemp(8) & "','" & argTemp(9) & "'); "
                ' StrGlSQLCommand As String
                ' intGLSQLCount As Integer
                ' intGLSQLMaxCount
                If intGLSQLMaxCount <= intGLSQLCount Then
                    intRows = SQL_Command(StrGlSQLCommand)

                    If intRows = 0 Then
                        RaiseEvent CurrentRecord("Error Inserting: " & StrGlSQLCommand & vbCrLf)
                        Threading.Thread.Sleep(0)
                    Else
                        If Me.Show_Output Then
                            RaiseEvent CurrentRecord("Cached SQL: " & StrGlSQLCommand & vbCrLf)
                            Threading.Thread.Sleep(0)
                        End If

                    End If
                    intGLSQLCount = 0
                    StrGlSQLCommand = ""
                Else
                    StrGlSQLCommand = StrGlSQLCommand & StrSQLCommand
                    intGLSQLCount = intGLSQLCount + 1
                End If



            Catch ex As Exception
                'Logs Error to GUI
                StrTemp = ("Cannot Insert Record Array has " & argTemp.GetUpperBound(0) & " Records. Values: " & vbCrLf)
                'For Each strline In argTemp
                '    StrTemp = StrTemp & strline & vbCrLf
                'Next
                'RaiseEvent CurrentRecord(StrTemp)
                Threading.Thread.Sleep(0)
                'StrTemp = "Cannot Insert Record" & vbCrLf & ex.Message & vbCrLf _
                '                    & "Insert into [" & StrDatabase & "] (FolderPath, AccountSAMAccountName, GroupSAMAccountName," _
                '                    & " ManagedBy, Inheritance, Rights, Owner, RunDate, IsInherited" _
                '                    & " ) Values ( '" _
                '                    & argTemp(0) & "','" & argTemp(1) & "','" & argTemp(2) & "','" & argTemp(3) _
                '                    & "','" & argTemp(4) & "','" & argTemp(5) & "','" & argTemp(6) & "','" & argTemp(7) _
                '                    & "','" & argTemp(8) & "')" & vbCrLf _
                '                    & "Input String: " & strInput & vbCrLf
                'RaiseEvent CurrentRecord(StrTemp)
                Threading.Thread.Sleep(0)
            End Try
        Else
            objLogFile.WriteLine(strInput)
            objLogFile.Flush()
            If Me.Show_Output Then
                RaiseEvent CurrentRecord(strInput & vbCrLf)
                Threading.Thread.Sleep(0)
            End If
        End If
    End Sub

    Private Function Format(EndTime As Date, p2 As String) As String
        Throw New NotImplementedException
    End Function

End Class
