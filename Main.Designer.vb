<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main))
        Me.RARBS = New System.Windows.Forms.CheckBox()
        Me.CheckBoxSQLRemove = New System.Windows.Forms.CheckBox()
        Me.CheckBoxRH = New System.Windows.Forms.CheckBox()
        Me.ShowOutput = New System.Windows.Forms.CheckBox()
        Me.Start = New System.Windows.Forms.Button()
        Me.GroupBoxFA = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.InputFA = New System.Windows.Forms.CheckedListBox()
        Me.Output = New System.Windows.Forms.TextBox()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar = New System.Windows.Forms.ToolStripProgressBar()
        Me.Help = New System.Windows.Forms.ToolStripDropDownButton()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxNoE = New System.Windows.Forms.CheckBox()
        Me.ImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.GroupBoxFA.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'RARBS
        '
        Me.RARBS.AutoSize = True
        Me.RARBS.Location = New System.Drawing.Point(274, 35)
        Me.RARBS.Name = "RARBS"
        Me.RARBS.Size = New System.Drawing.Size(206, 17)
        Me.RARBS.TabIndex = 4
        Me.RARBS.Text = "Remove All SQL Records Before Start"
        Me.RARBS.UseVisualStyleBackColor = True
        '
        'CheckBoxSQLRemove
        '
        Me.CheckBoxSQLRemove.AutoSize = True
        Me.CheckBoxSQLRemove.Location = New System.Drawing.Point(12, 35)
        Me.CheckBoxSQLRemove.Name = "CheckBoxSQLRemove"
        Me.CheckBoxSQLRemove.Size = New System.Drawing.Size(230, 17)
        Me.CheckBoxSQLRemove.TabIndex = 2
        Me.CheckBoxSQLRemove.Text = "Remove Todays SQL Records Before Start"
        Me.CheckBoxSQLRemove.UseVisualStyleBackColor = True
        '
        'CheckBoxRH
        '
        Me.CheckBoxRH.AutoSize = True
        Me.CheckBoxRH.Location = New System.Drawing.Point(274, 12)
        Me.CheckBoxRH.Name = "CheckBoxRH"
        Me.CheckBoxRH.Size = New System.Drawing.Size(113, 17)
        Me.CheckBoxRH.TabIndex = 3
        Me.CheckBoxRH.Text = "Remove Inherited "
        Me.CheckBoxRH.UseVisualStyleBackColor = True
        '
        'ShowOutput
        '
        Me.ShowOutput.AutoSize = True
        Me.ShowOutput.Checked = True
        Me.ShowOutput.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ShowOutput.Location = New System.Drawing.Point(12, 12)
        Me.ShowOutput.Name = "ShowOutput"
        Me.ShowOutput.Size = New System.Drawing.Size(88, 17)
        Me.ShowOutput.TabIndex = 1
        Me.ShowOutput.Text = "Show Output"
        Me.ShowOutput.UseVisualStyleBackColor = True
        '
        'Start
        '
        Me.Start.Location = New System.Drawing.Point(521, 29)
        Me.Start.Name = "Start"
        Me.Start.Size = New System.Drawing.Size(75, 23)
        Me.Start.TabIndex = 5
        Me.Start.Text = "Start"
        Me.Start.UseVisualStyleBackColor = True
        '
        'GroupBoxFA
        '
        Me.GroupBoxFA.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxFA.AutoSize = True
        Me.GroupBoxFA.Controls.Add(Me.Label3)
        Me.GroupBoxFA.Controls.Add(Me.InputFA)
        Me.GroupBoxFA.Location = New System.Drawing.Point(0, 58)
        Me.GroupBoxFA.Name = "GroupBoxFA"
        Me.GroupBoxFA.Size = New System.Drawing.Size(684, 145)
        Me.GroupBoxFA.TabIndex = 45
        Me.GroupBoxFA.TabStop = False
        Me.GroupBoxFA.Text = "Folder Audit"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(159, 13)
        Me.Label3.TabIndex = 33
        Me.Label3.Text = "Select Drives To Report Against"
        '
        'InputFA
        '
        Me.InputFA.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.InputFA.CheckOnClick = True
        Me.InputFA.ColumnWidth = 250
        Me.InputFA.Location = New System.Drawing.Point(12, 32)
        Me.InputFA.MultiColumn = True
        Me.InputFA.Name = "InputFA"
        Me.InputFA.Size = New System.Drawing.Size(660, 94)
        Me.InputFA.TabIndex = 8
        '
        'Output
        '
        Me.Output.AcceptsReturn = True
        Me.Output.AcceptsTab = True
        Me.Output.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Output.Location = New System.Drawing.Point(0, 17)
        Me.Output.Multiline = True
        Me.Output.Name = "Output"
        Me.Output.ReadOnly = True
        Me.Output.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.Output.Size = New System.Drawing.Size(684, 246)
        Me.Output.TabIndex = 10
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel, Me.ToolStripProgressBar, Me.Help})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 481)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(684, 22)
        Me.StatusStrip.TabIndex = 48
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(39, 17)
        Me.ToolStripStatusLabel.Text = "Ready"
        '
        'ToolStripProgressBar
        '
        Me.ToolStripProgressBar.Name = "ToolStripProgressBar"
        Me.ToolStripProgressBar.Size = New System.Drawing.Size(100, 16)
        '
        'Help
        '
        Me.Help.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.Help.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.Help.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.Help.Image = CType(resources.GetObject("Help.Image"), System.Drawing.Image)
        Me.Help.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Help.Name = "Help"
        Me.Help.Size = New System.Drawing.Size(45, 20)
        Me.Help.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.AutoSize = True
        Me.GroupBox1.Controls.Add(Me.Output)
        Me.GroupBox1.Location = New System.Drawing.Point(0, 209)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(684, 269)
        Me.GroupBox1.TabIndex = 49
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Output"
        '
        'CheckBoxNoE
        '
        Me.CheckBoxNoE.AutoSize = True
        Me.CheckBoxNoE.Location = New System.Drawing.Point(521, 6)
        Me.CheckBoxNoE.Name = "CheckBoxNoE"
        Me.CheckBoxNoE.Size = New System.Drawing.Size(92, 17)
        Me.CheckBoxNoE.TabIndex = 50
        Me.CheckBoxNoE.Text = "Notify On End"
        Me.CheckBoxNoE.UseVisualStyleBackColor = True
        '
        'ImageList
        '
        Me.ImageList.ImageStream = CType(resources.GetObject("ImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList.Images.SetKeyName(0, "Folder.ico")
        Me.ImageList.Images.SetKeyName(1, "AD.ico")
        Me.ImageList.Images.SetKeyName(2, "ADOU.ico")
        Me.ImageList.Images.SetKeyName(3, "ADAccount.ico")
        Me.ImageList.Images.SetKeyName(4, "ADComputer.ico")
        Me.ImageList.Images.SetKeyName(5, "ADGroup.ico")
        Me.ImageList.Images.SetKeyName(6, "LogonKeys.ico")
        Me.ImageList.Images.SetKeyName(7, "HardDrive.ico")
        Me.ImageList.Images.SetKeyName(8, "NetworkDrive.ico")
        Me.ImageList.Images.SetKeyName(9, "Floppy.ico")
        Me.ImageList.Images.SetKeyName(10, "CDROM.ico")
        Me.ImageList.Images.SetKeyName(11, "Removable.ico")
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(684, 503)
        Me.Controls.Add(Me.CheckBoxNoE)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.GroupBoxFA)
        Me.Controls.Add(Me.Start)
        Me.Controls.Add(Me.RARBS)
        Me.Controls.Add(Me.CheckBoxSQLRemove)
        Me.Controls.Add(Me.CheckBoxRH)
        Me.Controls.Add(Me.ShowOutput)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(700, 541)
        Me.Name = "Main"
        Me.Text = "File Server Audit"
        Me.GroupBoxFA.ResumeLayout(False)
        Me.GroupBoxFA.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RARBS As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxSQLRemove As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxRH As System.Windows.Forms.CheckBox
    Friend WithEvents ShowOutput As System.Windows.Forms.CheckBox
    Friend WithEvents Start As System.Windows.Forms.Button
    Friend WithEvents GroupBoxFA As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents InputFA As System.Windows.Forms.CheckedListBox
    Friend WithEvents Output As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Public WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripProgressBar As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents Help As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxNoE As System.Windows.Forms.CheckBox
    Friend WithEvents ImageList As System.Windows.Forms.ImageList

End Class
