Imports Str = Microsoft.VisualBasic.Strings
Imports MySql.Data.MySqlClient
Public Class Main
    Dim myBuildInfo As System.Diagnostics.FileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
    Public objSQLConnection As Object
    'Public dicSQLInfo As New Dictionary(Of String, String)
    Public StrDatabase As String
    Public StrTableName As String = My.Settings.SQLTableName
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim objDrive As System.IO.DriveInfo
        For Each objDrive In System.IO.DriveInfo.GetDrives()
            Try
                If objDrive.IsReady Then
                    If objDrive.DriveFormat = "NTFS" Then
                        Me.InputFA.Items.Add(New AvalibleDrives(objDrive.DriveType.ToString, objDrive.Name.ToString, objDrive.VolumeLabel))


                    End If
                End If
            Catch
                Output.Text = Output.Text & vbCrLf & Err.Description & " " & objDrive.DriveType.ToString & " " & objDrive.Name.ToString & " " & objDrive.VolumeLabel
            End Try
        Next
        'Setting up SQL Connection.
        SQL_Int()
        'Me.Output.AppendText("Default connection: " & My.Settings.SQLConnection & vbCrLf)



    End Sub
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
                    Me.Close()
            End Select
        End If
    End Sub
    Private Function SQL_Command(StrSQL_Command As String) As String
        Dim IntOutput As Integer
        Dim objSQLCommand As Object
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
                    'objSQLCommand = New MySql.Data.MySqlClient.MySqlCommand(StrSQL_Command, objSQLConnection)
                    If objSQLCommand = Nothing Then objSQLCommand = New MySql.Data.MySqlClient.MySqlCommand
                    objSQLCommand.Connection = objSQLConnection
                    objSQLCommand.CommandText = StrSQL_Command
                    objSQLCommand.CommandType = CommandType.Text
                    IntOutput = objSQLCommand.ExecuteNonQuery()

                Case Else
                    MsgBox("SQLServerType setting not understood: SQLServerType =" & My.Settings.SQLServerType, MsgBoxStyle.Critical, "SQL Server Type")
                    Me.Close()
            End Select

        Catch ex As Exception
            Output.Text = Output.Text & vbCrLf & ex.Message & vbCrLf & "Cannot Insert Record" & StrSQL_Command & vbCrLf
            Threading.Thread.Sleep(0)
        End Try
        Return IntOutput
    End Function
    Private Sub SQLRemoveTodaysRecords()
        Dim strSQL As String
        strSQL = "DELETE FROM " & StrTableName & " WHERE RunDate like '%" & TodayFileDate() & "%'"
        'Shows how many rows were removed from table.
        Output.Text = "Removed " & SQL_Command(strSQL) & " row(s) that were inserted on " & TodayFileDate() & "." & vbCrLf
        'Shows how many rows remain in Table
        strSQL = "Select Count(ID) From " & StrDatabase & " "

        Output.Text = Output.Text & vbCrLf & SQL_Command(strSQL) & " row(s) still exist." & vbCrLf
    End Sub
    Private Sub SQLRemoveALLRecords()
        Dim strSQL As String
        strSQL = "DELETE FROM " & StrTableName & " "
        'Shows how many rows were removed from table.
        Output.Text = "Removed " & SQL_Command(strSQL) & " row(s) that were removed." & vbCrLf
        'Shows how many rows remain in Table
        strSQL = "Select Count(ID) From " & StrTableName & " "
        Output.Text = Output.Text & vbCrLf & SQL_Command(strSQL) & " row(s) still exist." & vbCrLf
    End Sub
    Private Delegate Sub AppendOutputDelegate(ByVal message As String)

    Public Sub CurrentRecord_Event(ByVal CurrentRecordProcess As String)
        If Me.Output.InvokeRequired Then
            Dim caller As AppendOutputDelegate = AddressOf CurrentRecord_Event
            Me.Output.Invoke(caller, New Object() {CurrentRecordProcess})
        Else
            Me.Output.AppendText(CurrentRecordProcess)
        End If

    End Sub

    Private Sub CurrentCountMax_Event(ByVal CurrentACLMaxNumber As Integer)
        Me.ToolStripProgressBar.Maximum = CurrentACLMaxNumber
        Me.ToolStripProgressBar.Value = 0
    End Sub

    Private Sub CurrentCount_Event(ByVal CurrentACLNumber As Integer)
        On Error Resume Next
        If Me.ToolStripProgressBar.Maximum < CurrentACLNumber Then
            Me.ToolStripProgressBar.Maximum = CurrentACLNumber
            Me.ToolStripProgressBar.Value = CurrentACLNumber
        Else
            Me.ToolStripProgressBar.Value = CurrentACLNumber
        End If

    End Sub
    Private Delegate Sub BGDone_EventDelegate(ByVal message As Integer)
    Private Sub BGDone_Event(ByVal input As Integer)
        If Me.Output.InvokeRequired Then
            Dim caller As BGDone_EventDelegate = AddressOf BGDone_Event
            Me.Output.Invoke(caller, New Object() {input})
        Else
            Select Case input
                Case 0
                    Me.ToolStripProgressBar.Value = 100
                    Me.ToolStripProgressBar.Maximum = 100
                    Me.ToolStripStatusLabel.Text = "Done"
                Case 1
                    Me.Output.AppendText("Error: No Input Directory Given")
                    Me.ToolStripStatusLabel.Text = "Error"
                    Me.ToolStripProgressBar.Value = 0
                Case 2
                    Me.Output.AppendText("Error: No Log File Given")
                    Me.ToolStripStatusLabel.Text = "Error"
                    Me.ToolStripProgressBar.Value = 0
                Case Else
                    Me.Output.AppendText("Error: ")
                    Me.ToolStripStatusLabel.Text = "Error"
                    Me.ToolStripProgressBar.Value = 0
            End Select
            'fixes bug with Start button
            Me.Start.Text = "Start"
        End If
    End Sub
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
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox(myBuildInfo.ProductName & " Version: " & myBuildInfo.ProductMajorPart & "." & myBuildInfo.ProductMinorPart & "." & myBuildInfo.ProductBuildPart & vbCrLf _
        & myBuildInfo.LegalCopyright.ToString & vbCrLf _
        & myBuildInfo.LegalTrademarks.ToString & vbCrLf _
        & myBuildInfo.Comments _
        , MsgBoxStyle.ApplicationModal, "Help About")
    End Sub



    Private Sub Start_Click(sender As System.Object, e As System.EventArgs) Handles Start.Click
        Dim AO As New AllInOne
        Select Case Start.Text
            Case "Start"
                If CheckBoxSQLRemove.Checked Then SQLRemoveTodaysRecords()
                If RARBS.Checked Then SQLRemoveALLRecords()
                Dim objDrive As AvalibleDrives
                Dim arrTemp(InputFA.CheckedItems.Count - 1) As String
                Dim x As Integer

                If Not InputFA.CheckedItems.Count > 0 Then
                    MsgBox("Error: No Drives Selected", MsgBoxStyle.Critical, "Error: No Drives Selected")
                    Exit Sub
                End If
                'Updating Status
                Start.Text = "Cancel"
                Me.ToolStripStatusLabel.Text = "Working"
                'Start Setup of Backround Process
                AO.Use_SQL = True
                AO.Recursive_Search = True
                AO.Notifie_on_end = CheckBoxNoE.Checked
                AO.Remove_Inherited = CheckBoxRH.Checked
                AO.Show_Output = ShowOutput.Checked
                AddHandler AO.CurrentRecord, AddressOf CurrentRecord_Event
                AddHandler AO.CurrentACLCountMax, AddressOf CurrentCountMax_Event
                AddHandler AO.CurrentACLCount, AddressOf CurrentCount_Event
                AddHandler AO.BGDone, AddressOf BGDone_Event
                For Each objDrive In InputFA.CheckedItems
                    If Not objDrive.DriveLetter = "" Then arrTemp(x) = objDrive.DriveLetter
                    x = x + 1
                Next
                AO.Check_Drives = arrTemp
                AO.startBackgroundTask()
            Case "Cancel"
                AO.IsWorking = False
                Me.Close()
                Start.Text = "Start"
                ToolStripStatusLabel.Text = "Ready"
                'ToolStripProgressBar.Value = 0
        End Select
    End Sub

    'Shared Sub Close()
    'Throw New System.NotImplementedException
    'End Sub

End Class

Public Class AvalibleDrives
    Public DriveType As String
    Public DriveLetter As String
    Public DriveName As String
    'Public DriveIcon As Image

    Public Sub New()
    End Sub

    Public Sub New(ByVal strDriveType, ByVal strDriveLetter, ByVal strDriveName)
        Me.DriveType = strDriveType
        Me.DriveLetter = strDriveLetter
        Me.DriveName = strDriveName
        'Me.DriveIcon = Form1.ImageList.Images.Item(7)
    End Sub

    Public Overrides Function ToString() As String
        Select Case LCase(DriveType)
            Case "fixed"
                Return DriveLetter & vbTab & DriveType & vbTab & vbTab & DriveName
            Case "network"
                Return DriveLetter & vbTab & DriveType & vbTab & DriveName
            Case Else
                Return DriveLetter & vbTab & DriveType & vbTab & DriveName
        End Select
    End Function
End Class
