'hospital 
'696, 689
'form size=880, 780

Imports System.Data.OleDb

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region "Public variable"
    Public Shared flag As Integer
#End Region


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents pnlMain As System.Windows.Forms.Panel
    Friend WithEvents pnlContent As System.Windows.Forms.Panel
    Friend WithEvents btnRoutineCheckup As System.Windows.Forms.Button
    Friend WithEvents btnAdmitpatient As System.Windows.Forms.Button
    Friend WithEvents btnAddDoctor As System.Windows.Forms.Button
    Friend WithEvents btnAddsisboy As System.Windows.Forms.Button
    Friend WithEvents btnDischargePatient As System.Windows.Forms.Button
    Friend WithEvents pnlShowlist As System.Windows.Forms.Panel
    Friend WithEvents btnBill As System.Windows.Forms.Button
    Friend WithEvents mnuF As System.Windows.Forms.MainMenu
    Friend WithEvents mnuF_File As System.Windows.Forms.MenuItem
    Friend WithEvents mnuH As System.Windows.Forms.MenuItem
    Friend WithEvents mnuH_Help As System.Windows.Forms.MenuItem
    Friend WithEvents mnuH_About As System.Windows.Forms.MenuItem
    Friend WithEvents tmr_SB As System.Windows.Forms.Timer
    Friend WithEvents btnCalculator As System.Windows.Forms.Button
    Friend WithEvents btnNotepad As System.Windows.Forms.Button
    Friend WithEvents imgList_Icon As System.Windows.Forms.ImageList
    Friend WithEvents btnEmail As System.Windows.Forms.Button
    Friend WithEvents mnuT As System.Windows.Forms.MenuItem
    Friend WithEvents mnuT_Calculator As System.Windows.Forms.MenuItem
    Friend WithEvents mnuT_Notepad As System.Windows.Forms.MenuItem
    Friend WithEvents mnuT_Email As System.Windows.Forms.MenuItem
    Friend WithEvents btnBed As System.Windows.Forms.Button
    Friend WithEvents lblDetails As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents mnuF_Exit As System.Windows.Forms.MenuItem
    Friend WithEvents SB_Main As System.Windows.Forms.StatusBar
    Friend WithEvents SBP1_Status As System.Windows.Forms.StatusBarPanel
    Friend WithEvents SBP2_Today As System.Windows.Forms.StatusBarPanel
    Friend WithEvents SBP3_Date As System.Windows.Forms.StatusBarPanel
    Friend WithEvents SBP4_Time As System.Windows.Forms.StatusBarPanel
    Friend WithEvents SBP5_Showtime As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.pnlMain = New System.Windows.Forms.Panel
        Me.mnuF = New System.Windows.Forms.MainMenu
        Me.mnuF_File = New System.Windows.Forms.MenuItem
        Me.mnuF_Exit = New System.Windows.Forms.MenuItem
        Me.mnuT = New System.Windows.Forms.MenuItem
        Me.mnuT_Calculator = New System.Windows.Forms.MenuItem
        Me.mnuT_Notepad = New System.Windows.Forms.MenuItem
        Me.mnuT_Email = New System.Windows.Forms.MenuItem
        Me.mnuH = New System.Windows.Forms.MenuItem
        Me.mnuH_Help = New System.Windows.Forms.MenuItem
        Me.mnuH_About = New System.Windows.Forms.MenuItem
        Me.pnlContent = New System.Windows.Forms.Panel
        Me.btnBed = New System.Windows.Forms.Button
        Me.btnBill = New System.Windows.Forms.Button
        Me.btnDischargePatient = New System.Windows.Forms.Button
        Me.btnAddsisboy = New System.Windows.Forms.Button
        Me.btnAddDoctor = New System.Windows.Forms.Button
        Me.btnAdmitpatient = New System.Windows.Forms.Button
        Me.btnRoutineCheckup = New System.Windows.Forms.Button
        Me.lblDetails = New System.Windows.Forms.Label
        Me.imgList_Icon = New System.Windows.Forms.ImageList(Me.components)
        Me.pnlShowlist = New System.Windows.Forms.Panel
        Me.btnEmail = New System.Windows.Forms.Button
        Me.btnNotepad = New System.Windows.Forms.Button
        Me.btnCalculator = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.tmr_SB = New System.Windows.Forms.Timer(Me.components)
        Me.SB_Main = New System.Windows.Forms.StatusBar
        Me.SBP1_Status = New System.Windows.Forms.StatusBarPanel
        Me.SBP2_Today = New System.Windows.Forms.StatusBarPanel
        Me.SBP3_Date = New System.Windows.Forms.StatusBarPanel
        Me.SBP4_Time = New System.Windows.Forms.StatusBarPanel
        Me.SBP5_Showtime = New System.Windows.Forms.StatusBarPanel
        Me.pnlContent.SuspendLayout()
        Me.pnlShowlist.SuspendLayout()
        CType(Me.SBP1_Status, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SBP2_Today, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SBP3_Date, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SBP4_Time, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SBP5_Showtime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlMain
        '
        Me.pnlMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlMain.Location = New System.Drawing.Point(168, 8)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(704, 624)
        Me.pnlMain.TabIndex = 0
        '
        'mnuF
        '
        Me.mnuF.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuF_File, Me.mnuT, Me.mnuH})
        '
        'mnuF_File
        '
        Me.mnuF_File.Index = 0
        Me.mnuF_File.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuF_Exit})
        Me.mnuF_File.Text = "File"
        '
        'mnuF_Exit
        '
        Me.mnuF_Exit.Index = 0
        Me.mnuF_Exit.Text = "Exit"
        '
        'mnuT
        '
        Me.mnuT.Index = 1
        Me.mnuT.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuT_Calculator, Me.mnuT_Notepad, Me.mnuT_Email})
        Me.mnuT.Text = "Tool"
        '
        'mnuT_Calculator
        '
        Me.mnuT_Calculator.Index = 0
        Me.mnuT_Calculator.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuT_Calculator.Text = "Calculator"
        '
        'mnuT_Notepad
        '
        Me.mnuT_Notepad.Index = 1
        Me.mnuT_Notepad.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuT_Notepad.Text = "Notepad"
        '
        'mnuT_Email
        '
        Me.mnuT_Email.Index = 2
        Me.mnuT_Email.Shortcut = System.Windows.Forms.Shortcut.CtrlE
        Me.mnuT_Email.Text = "Email"
        '
        'mnuH
        '
        Me.mnuH.Index = 2
        Me.mnuH.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuH_Help, Me.mnuH_About})
        Me.mnuH.Text = "Help"
        '
        'mnuH_Help
        '
        Me.mnuH_Help.Index = 0
        Me.mnuH_Help.Shortcut = System.Windows.Forms.Shortcut.F1
        Me.mnuH_Help.Text = "Help"
        '
        'mnuH_About
        '
        Me.mnuH_About.Index = 1
        Me.mnuH_About.Shortcut = System.Windows.Forms.Shortcut.F2
        Me.mnuH_About.Text = "About"
        '
        'pnlContent
        '
        Me.pnlContent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlContent.Controls.Add(Me.btnBed)
        Me.pnlContent.Controls.Add(Me.btnBill)
        Me.pnlContent.Controls.Add(Me.btnDischargePatient)
        Me.pnlContent.Controls.Add(Me.btnAddsisboy)
        Me.pnlContent.Controls.Add(Me.btnAddDoctor)
        Me.pnlContent.Controls.Add(Me.btnAdmitpatient)
        Me.pnlContent.Controls.Add(Me.btnRoutineCheckup)
        Me.pnlContent.Controls.Add(Me.lblDetails)
        Me.pnlContent.Location = New System.Drawing.Point(8, 8)
        Me.pnlContent.Name = "pnlContent"
        Me.pnlContent.Size = New System.Drawing.Size(152, 392)
        Me.pnlContent.TabIndex = 2
        '
        'btnBed
        '
        Me.btnBed.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnBed.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBed.Location = New System.Drawing.Point(24, 336)
        Me.btnBed.Name = "btnBed"
        Me.btnBed.Size = New System.Drawing.Size(104, 32)
        Me.btnBed.TabIndex = 7
        Me.btnBed.Text = "Bed"
        '
        'btnBill
        '
        Me.btnBill.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnBill.Location = New System.Drawing.Point(24, 288)
        Me.btnBill.Name = "btnBill"
        Me.btnBill.Size = New System.Drawing.Size(104, 32)
        Me.btnBill.TabIndex = 6
        Me.btnBill.Text = "Bill"
        '
        'btnDischargePatient
        '
        Me.btnDischargePatient.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDischargePatient.Location = New System.Drawing.Point(24, 240)
        Me.btnDischargePatient.Name = "btnDischargePatient"
        Me.btnDischargePatient.Size = New System.Drawing.Size(104, 32)
        Me.btnDischargePatient.TabIndex = 5
        Me.btnDischargePatient.Text = "Discharge Patient"
        '
        'btnAddsisboy
        '
        Me.btnAddsisboy.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddsisboy.Location = New System.Drawing.Point(24, 192)
        Me.btnAddsisboy.Name = "btnAddsisboy"
        Me.btnAddsisboy.Size = New System.Drawing.Size(104, 32)
        Me.btnAddsisboy.TabIndex = 4
        Me.btnAddsisboy.Text = "Add Sister/Wardboy"
        '
        'btnAddDoctor
        '
        Me.btnAddDoctor.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddDoctor.Location = New System.Drawing.Point(24, 144)
        Me.btnAddDoctor.Name = "btnAddDoctor"
        Me.btnAddDoctor.Size = New System.Drawing.Size(104, 32)
        Me.btnAddDoctor.TabIndex = 3
        Me.btnAddDoctor.Text = "Add Doctor"
        '
        'btnAdmitpatient
        '
        Me.btnAdmitpatient.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAdmitpatient.Location = New System.Drawing.Point(24, 96)
        Me.btnAdmitpatient.Name = "btnAdmitpatient"
        Me.btnAdmitpatient.Size = New System.Drawing.Size(104, 32)
        Me.btnAdmitpatient.TabIndex = 2
        Me.btnAdmitpatient.Text = "Admit Patient"
        '
        'btnRoutineCheckup
        '
        Me.btnRoutineCheckup.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnRoutineCheckup.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRoutineCheckup.Location = New System.Drawing.Point(24, 48)
        Me.btnRoutineCheckup.Name = "btnRoutineCheckup"
        Me.btnRoutineCheckup.Size = New System.Drawing.Size(104, 32)
        Me.btnRoutineCheckup.TabIndex = 1
        Me.btnRoutineCheckup.Text = "Routine Checkup"
        '
        'lblDetails
        '
        Me.lblDetails.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold)
        Me.lblDetails.ForeColor = System.Drawing.Color.Teal
        Me.lblDetails.Location = New System.Drawing.Point(16, 8)
        Me.lblDetails.Name = "lblDetails"
        Me.lblDetails.Size = New System.Drawing.Size(120, 32)
        Me.lblDetails.TabIndex = 7
        Me.lblDetails.Text = "Records"
        Me.lblDetails.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'imgList_Icon
        '
        Me.imgList_Icon.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgList_Icon.ImageStream = CType(resources.GetObject("imgList_Icon.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgList_Icon.TransparentColor = System.Drawing.Color.Transparent
        '
        'pnlShowlist
        '
        Me.pnlShowlist.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlShowlist.Controls.Add(Me.btnEmail)
        Me.pnlShowlist.Controls.Add(Me.btnNotepad)
        Me.pnlShowlist.Controls.Add(Me.btnCalculator)
        Me.pnlShowlist.Controls.Add(Me.Label1)
        Me.pnlShowlist.Location = New System.Drawing.Point(8, 424)
        Me.pnlShowlist.Name = "pnlShowlist"
        Me.pnlShowlist.Size = New System.Drawing.Size(152, 208)
        Me.pnlShowlist.TabIndex = 4
        '
        'btnEmail
        '
        Me.btnEmail.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnEmail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEmail.Location = New System.Drawing.Point(24, 160)
        Me.btnEmail.Name = "btnEmail"
        Me.btnEmail.Size = New System.Drawing.Size(104, 32)
        Me.btnEmail.TabIndex = 10
        Me.btnEmail.Text = "Email"
        '
        'btnNotepad
        '
        Me.btnNotepad.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnNotepad.Location = New System.Drawing.Point(24, 112)
        Me.btnNotepad.Name = "btnNotepad"
        Me.btnNotepad.Size = New System.Drawing.Size(104, 32)
        Me.btnNotepad.TabIndex = 9
        Me.btnNotepad.Text = "Notepad"
        '
        'btnCalculator
        '
        Me.btnCalculator.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCalculator.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCalculator.Location = New System.Drawing.Point(24, 64)
        Me.btnCalculator.Name = "btnCalculator"
        Me.btnCalculator.Size = New System.Drawing.Size(104, 32)
        Me.btnCalculator.TabIndex = 8
        Me.btnCalculator.Text = "Calculator"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold)
        Me.Label1.ForeColor = System.Drawing.Color.Teal
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 32)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Tools"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tmr_SB
        '
        Me.tmr_SB.Enabled = True
        '
        'SB_Main
        '
        Me.SB_Main.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SB_Main.Location = New System.Drawing.Point(0, 633)
        Me.SB_Main.Name = "SB_Main"
        Me.SB_Main.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.SBP1_Status, Me.SBP2_Today, Me.SBP3_Date, Me.SBP4_Time, Me.SBP5_Showtime})
        Me.SB_Main.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SB_Main.ShowPanels = True
        Me.SB_Main.Size = New System.Drawing.Size(872, 32)
        Me.SB_Main.TabIndex = 5
        '
        'SBP1_Status
        '
        Me.SBP1_Status.Text = "Status : "
        Me.SBP1_Status.Width = 450
        '
        'SBP2_Today
        '
        Me.SBP2_Today.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.SBP2_Today.Text = "Today"
        Me.SBP2_Today.Width = 60
        '
        'SBP3_Date
        '
        Me.SBP3_Date.Icon = CType(resources.GetObject("SBP3_Date.Icon"), System.Drawing.Icon)
        '
        'SBP4_Time
        '
        Me.SBP4_Time.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.SBP4_Time.Text = "Time"
        Me.SBP4_Time.Width = 60
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(872, 665)
        Me.Controls.Add(Me.SB_Main)
        Me.Controls.Add(Me.pnlShowlist)
        Me.Controls.Add(Me.pnlContent)
        Me.Controls.Add(Me.pnlMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.mnuF
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Hospital Management (v2.0)"
        Me.pnlContent.ResumeLayout(False)
        Me.pnlShowlist.ResumeLayout(False)
        CType(Me.SBP1_Status, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SBP2_Today, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SBP3_Date, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SBP4_Time, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SBP5_Showtime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "All Button Event"

    Private Sub btnRoutineCheckup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRoutineCheckup.Click
        'counter_routine = 1

        Dim r As DialogResult

        Dim obj As New frmMain()
        Dim getflag As Integer
        getflag = obj.setFlag

        'check for record is saved or not
        If getflag = 1 Then
            r = MessageBox.Show("The record is not saved. Do you want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ElseIf getflag = 0 Then
            btnRoutineCheckup.Enabled = False
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Clinical Checkup Record"
            SBP1_Status.Text = "Clinical Checkup Record"
            'lbl_Title.Visible = True
            Load_frmClinicalCheckup_IntoPanel()
        End If 'end

        'check for response
        If r = DialogResult.No Then
            Exit Sub
        ElseIf r = DialogResult.Yes Then
            obj.setFlag = 0

            btnRoutineCheckup.Enabled = False
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Clinical Checkup Record"
            SBP1_Status.Text = "Clinical Checkup Record"
            'lbl_Title.Visible = True
            Load_frmClinicalCheckup_IntoPanel()
        End If 'end

    End Sub

    Private Sub btnAdmitpatient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdmitpatient.Click

        Dim admit_result As DialogResult
        Dim obj As New frmMain()
        Dim getflag As Integer
        getflag = obj.setFlag

        'check for record is saved or not
        If getflag = 1 Then
            admit_result = MessageBox.Show("The record is not saved. Do you want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ElseIf getflag = 0 Then
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = False
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Admit Record"
            SBP1_Status.Text = "Admit Record"
            'lbl_Title.Visible = True
            Load_frmAdmitPatient_IntoPanel()
        End If 'end

        'check for response
        If admit_result = DialogResult.No Then
            Exit Sub
        ElseIf admit_result = DialogResult.Yes Then
            obj.setFlag = 0
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = False
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Admit Record"
            SBP1_Status.Text = "Admit Record"
            'lbl_Title.Visible = True
            Load_frmAdmitPatient_IntoPanel()
        End If 'end

    End Sub

    Private Sub btnAddDoctor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddDoctor.Click

        Dim doctor_result As DialogResult

        Dim obj As New frmMain()
        Dim getflag As Integer
        getflag = obj.setFlag

        'check for record is saved or not
        If getflag = 1 Then
            doctor_result = MessageBox.Show("The record is not saved. Do you want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ElseIf getflag = 0 Then
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = False
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Doctor Record"
            SBP1_Status.Text = "Doctor Record"
            'lbl_Title.Visible = True
            Load_frmDoctor_IntoPanel()
        End If 'end

        'check for response
        If doctor_result = DialogResult.No Then
            Exit Sub
        ElseIf doctor_result = DialogResult.Yes Then
            obj.setFlag = 0
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = False
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Doctor Record"
            SBP1_Status.Text = "Doctor Record"
            'lbl_Title.Visible = True
            Load_frmDoctor_IntoPanel()
        End If 'end

    End Sub

    Private Sub btnAddsisboy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddsisboy.Click
        Dim sisboy_result As DialogResult

        Dim obj As New frmMain()
        Dim getflag As Integer
        getflag = obj.setFlag

        'check for record is saved or not
        If getflag = 1 Then
            sisboy_result = MessageBox.Show("The record is not saved. Do you want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ElseIf getflag = 0 Then
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = False
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Sister/Wardboy Record"
            SBP1_Status.Text = "Sister/Wardboy Record"
            'lbl_Title.Visible = True
            Load_frmSisterWardboy_IntoPanel()
        End If 'end

        'check for response
        If sisboy_result = DialogResult.No Then
            Exit Sub
        ElseIf sisboy_result = DialogResult.Yes Then
            obj.setFlag = 0
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = False
            btnDischargePatient.Enabled = True
            btnBill.Enabled = True

            'lbl_Title.Text = "Sister/Wardboy Record"
            SBP1_Status.Text = "Sister/Wardboy Record"
            'lbl_Title.Visible = True
            Load_frmSisterWardboy_IntoPanel()
        End If 'end

    End Sub

    Private Sub btnDischargePatient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDischargePatient.Click
        Dim discharge_result As DialogResult

        Dim obj As New frmMain()
        Dim getflag As Integer
        getflag = obj.setFlag

        'check for record is saved or not
        If getflag = 1 Then
            discharge_result = MessageBox.Show("The record is not saved. Do you want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ElseIf getflag = 0 Then
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = False
            btnBill.Enabled = True

            'lbl_Title.Text = "Discharge Record"
            SBP1_Status.Text = "Discharge Record"
            'lbl_Title.Visible = True
            Load_frmDischarge_IntoPanel()
        End If 'end

        'check for response
        If discharge_result = DialogResult.No Then

            Exit Sub
        ElseIf discharge_result = DialogResult.Yes Then
            obj.setFlag = 0
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = False
            btnBill.Enabled = True

            'lbl_Title.Text = "Discharge Record"
            SBP1_Status.Text = "Discharge Record"
            'lbl_Title.Visible = True
            Load_frmDischarge_IntoPanel()
        End If 'end
    End Sub

    Private Sub btnBill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBill.Click
        Dim bill_result As DialogResult

        Dim obj As New frmMain()
        Dim getflag As Integer
        getflag = obj.setFlag

        'check for record is saved or not
        If getflag = 1 Then
            bill_result = MessageBox.Show("The record is not saved. Do you want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ElseIf getflag = 0 Then
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = False

            'lbl_Title.Text = "Bill Record"
            SBP1_Status.Text = "Bill Record"
            'lbl_Title.Visible = True
            Load_frmBill_IntoPanel()
        End If 'end

        'check for response
        If bill_result = DialogResult.No Then

            Exit Sub
        ElseIf bill_result = DialogResult.Yes Then
            obj.setFlag = 0
            btnRoutineCheckup.Enabled = True
            btnAdmitpatient.Enabled = True
            btnAddDoctor.Enabled = True
            btnAddsisboy.Enabled = True
            btnDischargePatient.Enabled = True
            btnBill.Enabled = False

            'lbl_Title.Text = "Bill Record"
            SBP1_Status.Text = "Bill Record"
            'lbl_Title.Visible = True
            Load_frmBill_IntoPanel()
        End If 'end

    End Sub

    Private Sub btnCalculator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculator.Click
        Shell("C:\WINDOWS\system32\calc")
    End Sub

    Private Sub btnNotepad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNotepad.Click
        Shell("C:\WINDOWS\notepad")
    End Sub

    Private Sub btnEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEmail.Click
        Dim objfrmSenEMail As New frmSenEMail()
        objfrmSenEMail.Show()
    End Sub

    Private Sub btnBed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBed.Click

        Dim counter_Generalward As Integer
        Dim counter_Specialward As Integer

        Dim insertCB_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
        Dim insertCB_cmdstr As String = "select * from Generalward"
        Dim insertCB_con As New OleDbConnection(insertCB_constr)
        Dim insertCB_com As New OleDbCommand(insertCB_cmdstr, insertCB_con)
        insertCB_con.Open()
        Dim insertCB_DR As OleDbDataReader = insertCB_com.ExecuteReader
        While insertCB_DR.Read
            counter_Generalward = counter_Generalward + 1
        End While
        insertCB_con.Close()


        Dim insertCB_constr1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
        Dim insertCB_cmdstr1 As String = "select * from Specialward"
        Dim insertCB_con1 As New OleDbConnection(insertCB_constr1)
        Dim insertCB_com1 As New OleDbCommand(insertCB_cmdstr1, insertCB_con1)
        insertCB_con1.Open()
        Dim insertCB_DR1 As OleDbDataReader = insertCB_com1.ExecuteReader
        While insertCB_DR1.Read
            counter_Specialward = counter_Specialward + 1
        End While
        insertCB_con1.Close()

        Dim str1 As String = counter_Generalward & " beds are available for Generalward"
        Dim str2 As String = counter_Specialward & " beds are available for Specialward"
        Dim str As String = str1 & vbCrLf & str2

        If counter_Generalward = 0 And counter_Specialward = 0 Then
            MessageBox.Show("All beds are allocated.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show(str, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

#End Region

#Region "Form Load Function"

    Sub Load_frmClinicalCheckup_IntoPanel()
        pnlMain.Controls.Clear()
        Dim objfrmClinicalCheckup As New frmClinicalCheckup()
        objfrmClinicalCheckup.Size = pnlMain.Size
        objfrmClinicalCheckup.TopLevel = False
        objfrmClinicalCheckup.Parent = pnlMain
        objfrmClinicalCheckup.Show()
    End Sub

    Private Sub Load_frmAdmitPatient_IntoPanel()
        pnlMain.Controls.Clear()
        Dim objfrmAdmitPatient As New frmAdmitPatient()
        objfrmAdmitPatient.Size = pnlMain.Size
        objfrmAdmitPatient.TopLevel = False
        objfrmAdmitPatient.Parent = pnlMain
        objfrmAdmitPatient.Show()
    End Sub

    Private Sub Load_frmDoctor_IntoPanel()
        pnlMain.Controls.Clear()
        Dim objfrmDoctor As New frmDoctor()
        objfrmDoctor.Size = pnlMain.Size
        objfrmDoctor.TopLevel = False
        objfrmDoctor.Parent = pnlMain
        objfrmDoctor.Show()
    End Sub

    Private Sub Load_frmSisterWardboy_IntoPanel()
        pnlMain.Controls.Clear()
        Dim objfrmSisterWardboy As New frmSisterWardboy()
        objfrmSisterWardboy.Size = pnlMain.Size
        objfrmSisterWardboy.TopLevel = False
        objfrmSisterWardboy.Parent = pnlMain
        objfrmSisterWardboy.Show()
    End Sub

    Private Sub Load_frmDischarge_IntoPanel()
        pnlMain.Controls.Clear()
        Dim objfrmDischarge As New frmDischarge()
        objfrmDischarge.Size = pnlMain.Size
        objfrmDischarge.TopLevel = False
        objfrmDischarge.Parent = pnlMain
        objfrmDischarge.Show()
    End Sub

    Private Sub Load_frmBill_IntoPanel()
        pnlMain.Controls.Clear()
        Dim objfrmBill As New frmBill()
        objfrmBill.Size = pnlMain.Size
        objfrmBill.TopLevel = False
        objfrmBill.Parent = pnlMain
        objfrmBill.Show()
    End Sub

#End Region

#Region "Property For Setting Save Flag"

    Property setFlag() As Integer
        Get
            Return flag
        End Get
        Set(ByVal Value As Integer)
            flag = Value
        End Set
    End Property

#End Region

#Region "Status Bar Timer Event"
    Private Sub tmr_SB_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr_SB.Tick
        SBP3_Date.Text = System.DateTime.Today
        SBP5_Showtime.Text = System.DateTime.Now.ToLongTimeString
    End Sub
#End Region

#Region "frmMain close event"
    'Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    '    Dim frmMain_result As DialogResult

    '    Dim obj As New frmMain()
    '    Dim getflag As Integer
    '    getflag = obj.setFlag

    '    'check for record is saved or not
    '    If getflag = 1 Then
    '        frmMain_result = MessageBox.Show("The record is not saved. Do you want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    '    ElseIf getflag = 0 Then
    '        Try
    '            Dim x As Long
    '            For x = 0 To 249
    '                Do Until Me.Height = 0
    '                    Me.Height -= 50
    '                    Exit Do
    '                Loop
    '            Next
    '            Dim y As Long
    '            For y = 0 To 249
    '                Do Until Me.Width = 0
    '                    Me.Width -= 50
    '                    Exit Do
    '                Loop
    '            Next
    '            Application.Exit()
    '        Catch ex As Exception
    '            Application.Exit()
    '        Finally
    '            Application.Exit()
    '        End Try
    '    End If 'end

    '    'check for response
    '    If frmMain_result = DialogResult.No Then
    '        Exit Sub

    '    ElseIf frmMain_result = DialogResult.Yes Then
    '        obj.setFlag = 0

    '        Try
    '            Dim x As Long
    '            For x = 0 To 249
    '                Do Until Me.Height = 0
    '                    Me.Height -= 20 '7
    '                    Exit Do
    '                Loop
    '            Next
    '            Dim y As Long
    '            For y = 0 To 249
    '                Do Until Me.Width = 0
    '                    Me.Width -= 20 '7
    '                    Exit Do
    '                Loop
    '            Next
    '            Application.Exit()
    '        Catch ex As Exception
    '            Application.Exit()
    '        Finally
    '            Application.Exit()
    '        End Try

    '    End If 'end
    'End Sub
    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Application.Exit()
    End Sub
#End Region

#Region "All File Menu item"

    Private Sub mnuF_Exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuF_Exit.Click
        Application.Exit()
    End Sub

#End Region

#Region "All Help menu items"
    Private Sub mnuH_About_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuH_About.Click
        Dim objfrmAbout As New frmAbout()
        objfrmAbout.Show()
    End Sub

    Private Sub mnuH_Help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuH_Help.Click
        Dim objfrmHelp As New frmHelp()
        objfrmHelp.Show()
    End Sub
#End Region

#Region "All Tool menu items"

    Private Sub mnuT_Calculator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuT_Calculator.Click
        Shell("C:\WINDOWS\system32\calc")
    End Sub

    Private Sub mnuT_Email_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuT_Email.Click
        Dim objfrmSenEMail As New frmSenEMail()
        objfrmSenEMail.Show()
    End Sub

    Private Sub mnuT_Notepad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuT_Notepad.Click
        Shell("C:\WINDOWS\notepad")
    End Sub

#End Region


    
End Class
