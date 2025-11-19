Imports System.Data
Imports System.Data.OleDb

Public Class frmAdmitPatient
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call


        'code for filling the combobox of doctors
        Dim r1 As String
        Dim FillcbDoctor_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim FillcbDoctor_cmdstr As String = "select * from Doctors"
        'Dim Fill_Textbox_cmdstr As String = "select * from clinicpatients"
        Dim FillcbDoctor_con As New OleDbConnection(FillcbDoctor_constr)
        Dim FillcbDoctor_com As New OleDbCommand(FillcbDoctor_cmdstr, FillcbDoctor_con)
        FillcbDoctor_con.Open()
        Dim FillcbDoctor_DR As OleDbDataReader = FillcbDoctor_com.ExecuteReader
        While FillcbDoctor_DR.Read
            'r0 = insert_DR(0)
            r1 = FillcbDoctor_DR(1)
            cbDoctor.Items.Add(r1)
        End While
        FillcbDoctor_con.Close() 'code end

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
    Friend WithEvents GrpBox_PatiInfo As System.Windows.Forms.GroupBox
    Friend WithEvents cbDoctor As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtDisease As System.Windows.Forms.TextBox
    Friend WithEvents txtphone As System.Windows.Forms.TextBox
    Friend WithEvents txtAge As System.Windows.Forms.TextBox
    Friend WithEvents txtPatiname As System.Windows.Forms.TextBox
    Friend WithEvents txtRegno As System.Windows.Forms.TextBox
    Friend WithEvents GB_PatientView As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents LV_PatiView As System.Windows.Forms.ListView
    Friend WithEvents DTP_date As System.Windows.Forms.DateTimePicker
    Friend WithEvents CB_Bloodgp As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CB_Bedno As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents CB_BedCategory As System.Windows.Forms.ComboBox
    Friend WithEvents txtBedCharge As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCustomprint As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnDischarge As System.Windows.Forms.Button
    Friend WithEvents btnSaveRegistration As System.Windows.Forms.Button
    Friend WithEvents btnCancelRegistration As System.Windows.Forms.Button
    Friend WithEvents btnEditRegistration As System.Windows.Forms.Button
    Friend WithEvents btnRegisterPatient As System.Windows.Forms.Button
    Friend WithEvents RtxtAddress As System.Windows.Forms.TextBox
    Friend WithEvents Rtxt_Remark As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GrpBox_PatiInfo = New System.Windows.Forms.GroupBox()
        Me.Rtxt_Remark = New System.Windows.Forms.TextBox()
        Me.RtxtAddress = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnDischarge = New System.Windows.Forms.Button()
        Me.btnSaveRegistration = New System.Windows.Forms.Button()
        Me.btnCancelRegistration = New System.Windows.Forms.Button()
        Me.btnEditRegistration = New System.Windows.Forms.Button()
        Me.btnRegisterPatient = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtBedCharge = New System.Windows.Forms.TextBox()
        Me.DTP_date = New System.Windows.Forms.DateTimePicker()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.cbDoctor = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDisease = New System.Windows.Forms.TextBox()
        Me.txtphone = New System.Windows.Forms.TextBox()
        Me.txtAge = New System.Windows.Forms.TextBox()
        Me.txtPatiname = New System.Windows.Forms.TextBox()
        Me.txtRegno = New System.Windows.Forms.TextBox()
        Me.CB_Bloodgp = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CB_BedCategory = New System.Windows.Forms.ComboBox()
        Me.CB_Bedno = New System.Windows.Forms.ComboBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.GB_PatientView = New System.Windows.Forms.GroupBox()
        Me.LV_PatiView = New System.Windows.Forms.ListView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnCustomprint = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GrpBox_PatiInfo.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GB_PatientView.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GrpBox_PatiInfo
        '
        Me.GrpBox_PatiInfo.Controls.AddRange(New System.Windows.Forms.Control() {Me.Rtxt_Remark, Me.RtxtAddress, Me.GroupBox3, Me.Label13, Me.txtBedCharge, Me.DTP_date, Me.Label10, Me.Label9, Me.cbDoctor, Me.Label8, Me.Label7, Me.Label6, Me.Label5, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.txtDisease, Me.txtphone, Me.txtAge, Me.txtPatiname, Me.txtRegno, Me.CB_Bloodgp, Me.Label11, Me.Label12, Me.CB_BedCategory, Me.CB_Bedno})
        Me.GrpBox_PatiInfo.Location = New System.Drawing.Point(16, 8)
        Me.GrpBox_PatiInfo.Name = "GrpBox_PatiInfo"
        Me.GrpBox_PatiInfo.Size = New System.Drawing.Size(512, 480)
        Me.GrpBox_PatiInfo.TabIndex = 4
        Me.GrpBox_PatiInfo.TabStop = False
        Me.GrpBox_PatiInfo.Text = "Patient Detail"
        '
        'Rtxt_Remark
        '
        Me.Rtxt_Remark.Location = New System.Drawing.Point(144, 432)
        Me.Rtxt_Remark.Multiline = True
        Me.Rtxt_Remark.Name = "Rtxt_Remark"
        Me.Rtxt_Remark.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.Rtxt_Remark.Size = New System.Drawing.Size(100, 40)
        Me.Rtxt_Remark.TabIndex = 12
        Me.Rtxt_Remark.Text = ""
        '
        'RtxtAddress
        '
        Me.RtxtAddress.Location = New System.Drawing.Point(144, 224)
        Me.RtxtAddress.Multiline = True
        Me.RtxtAddress.Name = "RtxtAddress"
        Me.RtxtAddress.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.RtxtAddress.Size = New System.Drawing.Size(104, 40)
        Me.RtxtAddress.TabIndex = 6
        Me.RtxtAddress.Text = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnDischarge, Me.btnSaveRegistration, Me.btnCancelRegistration, Me.btnEditRegistration, Me.btnRegisterPatient})
        Me.GroupBox3.Location = New System.Drawing.Point(320, 40)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(176, 280)
        Me.GroupBox3.TabIndex = 69
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Transaction"
        '
        'btnDischarge
        '
        Me.btnDischarge.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDischarge.Location = New System.Drawing.Point(32, 224)
        Me.btnDischarge.Name = "btnDischarge"
        Me.btnDischarge.Size = New System.Drawing.Size(112, 32)
        Me.btnDischarge.TabIndex = 18
        Me.btnDischarge.Text = "Discharge Patient"
        '
        'btnSaveRegistration
        '
        Me.btnSaveRegistration.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSaveRegistration.Location = New System.Drawing.Point(32, 128)
        Me.btnSaveRegistration.Name = "btnSaveRegistration"
        Me.btnSaveRegistration.Size = New System.Drawing.Size(112, 32)
        Me.btnSaveRegistration.TabIndex = 16
        Me.btnSaveRegistration.Text = "Save  Registration"
        '
        'btnCancelRegistration
        '
        Me.btnCancelRegistration.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancelRegistration.Location = New System.Drawing.Point(32, 176)
        Me.btnCancelRegistration.Name = "btnCancelRegistration"
        Me.btnCancelRegistration.Size = New System.Drawing.Size(112, 32)
        Me.btnCancelRegistration.TabIndex = 17
        Me.btnCancelRegistration.Text = "Cancel  Registration"
        '
        'btnEditRegistration
        '
        Me.btnEditRegistration.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnEditRegistration.Location = New System.Drawing.Point(32, 80)
        Me.btnEditRegistration.Name = "btnEditRegistration"
        Me.btnEditRegistration.Size = New System.Drawing.Size(112, 32)
        Me.btnEditRegistration.TabIndex = 14
        Me.btnEditRegistration.Text = "Edit Registration"
        '
        'btnRegisterPatient
        '
        Me.btnRegisterPatient.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnRegisterPatient.Location = New System.Drawing.Point(32, 32)
        Me.btnRegisterPatient.Name = "btnRegisterPatient"
        Me.btnRegisterPatient.Size = New System.Drawing.Size(112, 32)
        Me.btnRegisterPatient.TabIndex = 13
        Me.btnRegisterPatient.Text = "Register Patient"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(24, 128)
        Me.Label13.Name = "Label13"
        Me.Label13.TabIndex = 68
        Me.Label13.Text = "Bed Charge/Day"
        '
        'txtBedCharge
        '
        Me.txtBedCharge.Location = New System.Drawing.Point(144, 128)
        Me.txtBedCharge.Name = "txtBedCharge"
        Me.txtBedCharge.Size = New System.Drawing.Size(104, 20)
        Me.txtBedCharge.TabIndex = 3
        Me.txtBedCharge.Text = ""
        '
        'DTP_date
        '
        Me.DTP_date.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DTP_date.Location = New System.Drawing.Point(144, 304)
        Me.DTP_date.Name = "DTP_date"
        Me.DTP_date.Size = New System.Drawing.Size(104, 20)
        Me.DTP_date.TabIndex = 8
        Me.DTP_date.Value = New Date(2008, 1, 1, 0, 0, 0, 0)
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(24, 440)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 59
        Me.Label10.Text = "Remarks"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(24, 368)
        Me.Label9.Name = "Label9"
        Me.Label9.TabIndex = 58
        Me.Label9.Text = "Blood Group"
        '
        'cbDoctor
        '
        Me.cbDoctor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbDoctor.Location = New System.Drawing.Point(144, 400)
        Me.cbDoctor.Name = "cbDoctor"
        Me.cbDoctor.Size = New System.Drawing.Size(104, 21)
        Me.cbDoctor.TabIndex = 11
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 400)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Doctor"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 336)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Disease"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 304)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Date"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 272)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Phone No"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 224)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Address"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 192)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Age"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 160)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Patient Name"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Registration No"
        '
        'txtDisease
        '
        Me.txtDisease.Location = New System.Drawing.Point(144, 336)
        Me.txtDisease.Name = "txtDisease"
        Me.txtDisease.Size = New System.Drawing.Size(104, 20)
        Me.txtDisease.TabIndex = 9
        Me.txtDisease.Text = ""
        '
        'txtphone
        '
        Me.txtphone.Location = New System.Drawing.Point(144, 272)
        Me.txtphone.Name = "txtphone"
        Me.txtphone.Size = New System.Drawing.Size(104, 20)
        Me.txtphone.TabIndex = 7
        Me.txtphone.Text = ""
        '
        'txtAge
        '
        Me.txtAge.Location = New System.Drawing.Point(144, 192)
        Me.txtAge.Name = "txtAge"
        Me.txtAge.Size = New System.Drawing.Size(104, 20)
        Me.txtAge.TabIndex = 5
        Me.txtAge.Text = ""
        '
        'txtPatiname
        '
        Me.txtPatiname.Location = New System.Drawing.Point(144, 160)
        Me.txtPatiname.Name = "txtPatiname"
        Me.txtPatiname.Size = New System.Drawing.Size(104, 20)
        Me.txtPatiname.TabIndex = 4
        Me.txtPatiname.Text = ""
        '
        'txtRegno
        '
        Me.txtRegno.Location = New System.Drawing.Point(144, 32)
        Me.txtRegno.Name = "txtRegno"
        Me.txtRegno.ReadOnly = True
        Me.txtRegno.Size = New System.Drawing.Size(104, 20)
        Me.txtRegno.TabIndex = 7
        Me.txtRegno.Text = ""
        '
        'CB_Bloodgp
        '
        Me.CB_Bloodgp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_Bloodgp.Items.AddRange(New Object() {"A+", "B+", "AB+", "O+", "A-", "B-", "AB-", "O-"})
        Me.CB_Bloodgp.Location = New System.Drawing.Point(144, 368)
        Me.CB_Bloodgp.Name = "CB_Bloodgp"
        Me.CB_Bloodgp.Size = New System.Drawing.Size(104, 21)
        Me.CB_Bloodgp.TabIndex = 10
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(24, 64)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 23)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Bed Type"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(24, 96)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 24)
        Me.Label12.TabIndex = 64
        Me.Label12.Text = "Bed No"
        '
        'CB_BedCategory
        '
        Me.CB_BedCategory.Items.AddRange(New Object() {"General", "Special"})
        Me.CB_BedCategory.Location = New System.Drawing.Point(144, 64)
        Me.CB_BedCategory.Name = "CB_BedCategory"
        Me.CB_BedCategory.Size = New System.Drawing.Size(104, 21)
        Me.CB_BedCategory.TabIndex = 1
        '
        'CB_Bedno
        '
        Me.CB_Bedno.Enabled = False
        Me.CB_Bedno.Location = New System.Drawing.Point(144, 96)
        Me.CB_Bedno.Name = "CB_Bedno"
        Me.CB_Bedno.Size = New System.Drawing.Size(104, 21)
        Me.CB_Bedno.TabIndex = 2
        '
        'btnSearch
        '
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSearch.Location = New System.Drawing.Point(32, 32)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 32)
        Me.btnSearch.TabIndex = 20
        Me.btnSearch.Text = "Search"
        '
        'btnExcel
        '
        Me.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExcel.Location = New System.Drawing.Point(32, 80)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(75, 32)
        Me.btnExcel.TabIndex = 21
        Me.btnExcel.Text = "Excel"
        '
        'GB_PatientView
        '
        Me.GB_PatientView.Controls.AddRange(New System.Windows.Forms.Control() {Me.LV_PatiView})
        Me.GB_PatientView.Location = New System.Drawing.Point(16, 496)
        Me.GB_PatientView.Name = "GB_PatientView"
        Me.GB_PatientView.Size = New System.Drawing.Size(664, 128)
        Me.GB_PatientView.TabIndex = 5
        Me.GB_PatientView.TabStop = False
        Me.GB_PatientView.Text = "Patient List View"
        '
        'LV_PatiView
        '
        Me.LV_PatiView.FullRowSelect = True
        Me.LV_PatiView.GridLines = True
        Me.LV_PatiView.Location = New System.Drawing.Point(16, 16)
        Me.LV_PatiView.Name = "LV_PatiView"
        Me.LV_PatiView.Size = New System.Drawing.Size(632, 104)
        Me.LV_PatiView.TabIndex = 0
        Me.LV_PatiView.View = System.Windows.Forms.View.Details
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCustomprint})
        Me.GroupBox1.Location = New System.Drawing.Point(544, 48)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(136, 104)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print"
        '
        'btnCustomprint
        '
        Me.btnCustomprint.BackColor = System.Drawing.Color.LightGray
        Me.btnCustomprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCustomprint.Location = New System.Drawing.Point(32, 40)
        Me.btnCustomprint.Name = "btnCustomprint"
        Me.btnCustomprint.Size = New System.Drawing.Size(72, 32)
        Me.btnCustomprint.TabIndex = 19
        Me.btnCustomprint.Text = "List"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSearch, Me.btnExcel})
        Me.GroupBox2.Location = New System.Drawing.Point(544, 248)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(136, 128)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Search/Export"
        '
        'frmAdmitPatient
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(696, 632)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.GroupBox1, Me.GrpBox_PatiInfo, Me.GB_PatientView})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmAdmitPatient"
        Me.Text = "frmClinicalCheckup"
        Me.GrpBox_PatiInfo.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GB_PatientView.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Function for binding textbox with record"
    'code for binding text box with 1st record of table
    Sub Fill_Textbox_Admit()
        Dim r0, r1, r2, r3, r4, r5, r6, r7, r8, r9, r10, r11, r12 As String

        Dim Fill_Textbox_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim Fill_Textbox_cmdstr As String = "select Regno,Name,Age,Address,Phone,Checkupdate,Checkuptime,Disease,Bloodgroup,Doctor,Remarks from clinicpatients where Regno=1"
        Dim Fill_Textbox_cmdstr As String = "select * from admit"
        Dim Fill_Textbox_con As New OleDbConnection(Fill_Textbox_constr)
        Dim insert_com As New OleDbCommand(Fill_Textbox_cmdstr, Fill_Textbox_con)
        Fill_Textbox_con.Open()
        Dim insert_DR As OleDbDataReader = insert_com.ExecuteReader
        While insert_DR.Read
            r0 = insert_DR(0)
            r1 = insert_DR(1)
            r2 = insert_DR(2)
            r3 = insert_DR(3)
            r4 = insert_DR(4)
            r5 = insert_DR(5)
            r6 = insert_DR(6)
            r7 = insert_DR(7)
            r8 = insert_DR(8)
            r9 = insert_DR(9)
            r10 = insert_DR(10)
            r11 = insert_DR(11)
            r12 = insert_DR(12)
            Exit While
        End While
        Fill_Textbox_con.Close()

        txtRegno.Text = r0
        CB_BedCategory.Text = r1
        CB_Bedno.Text = r2
        txtBedCharge.Text = r3
        txtPatiname.Text = r4
        txtAge.Text = r5
        RtxtAddress.Text = r6
        txtphone.Text = r7
        If r8 = Nothing Then
            DTP_date.Text = ""
        Else
            DTP_date.Value = r8
        End If

        txtDisease.Text = r9
        CB_Bloodgp.Text = r10
        cbDoctor.Text = r11
        Rtxt_Remark.Text = r12
    End Sub
#End Region

#Region "All button events"
    'add record     
    'Private Sub btnRegisterPatient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim r0, r1, r2, r3 As String

    '    Dim insert_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
    '    'Dim insert_cmdstr As String = "select * from admit" 'admitpatients"
    '    Dim insert_cmdstr As String = "select * from admitID"
    '    Dim insert_con As New OleDbConnection(insert_constr)
    '    Dim insert_com As New OleDbCommand(insert_cmdstr, insert_con)
    '    insert_con.Open()
    '    Dim insert_DR As OleDbDataReader = insert_com.ExecuteReader
    '    While insert_DR.Read
    '        r0 = insert_DR(0)
    '    End While
    '    insert_con.Close()

    '    CB_BedCategory.Text = ""
    '    CB_Bedno.Text = ""
    '    txtBedCharge.Text = ""
    '    txtPatiname.Text = ""
    '    txtAge.Text = ""
    '    RtxtAddress.Text = ""
    '    txtphone.Text = ""
    '    'DTP_date
    '    'DTP_time 
    '    txtDisease.Text = ""
    '    CB_Bloodgp.Text = ""
    '    cbDoctor.Text = ""
    '    Rtxt_Remark.Text = ""

    '    txtRegno.Text = r0 + 1

    '    Dim obj As New frmMain()
    '    obj.setFlag = 1
    'End Sub

    'save record
    'Private Sub btnSaveRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    'Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    'Dim save_Comstr As String
    '    'Dim save_Con As OleDbConnection
    '    'Dim save_Com As OleDbCommand
    '    'Dim save_DA As OleDbDataAdapter
    '    'Dim save_DS As DataSet
    '    'Dim save_DT As DataTable
    '    'Dim txt As String
    '    'txt = CType(RtxtAddress.Text, String)
    '    'save_Comstr = "insert into admitpatients(Regno,Bedno,Name,Age,Address,Phone,Admitdatetime,Disease,Bloodgroup,Doctor,Remarks)values" & _
    '    '("(" + "'" + txtRegno.Text + "'" + "," + "'" + CB_Bedno.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" + txtAge.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + txtphone.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + txtDisease.Text + "'" + "," + "'" + CB_Bloodgp.Text + "'" + "," + "'" + cbDoctor.Text + "'" + "," + "'" + Rtxt_Remark.Text + "'" + ")")

    '    'save_Con = New OleDbConnection(save_Constr)
    '    'save_Com = New OleDbCommand(save_Comstr, save_Con)

    '    'save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
    '    'Dim save_builder As New OleDbCommandBuilder(save_DA)

    '    'save_DS = New DataSet()

    '    'Try
    '    '    save_DA.Fill(save_DS, "hm")
    '    '    MsgBox("Patient record inserted successfully.", MsgBoxStyle.Information, "Register Patient")
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message)
    '    'End Try
    '    'save_Con.Close()
    '    'fill_listview()
    '    Check_Textbox()
    'End Sub

    'update record
    'Private Sub btnEditRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim update_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    Dim update_Comstr As String
    '    Dim update_Con As OleDbConnection
    '    Dim update_Com As OleDbCommand

    '    Dim ID As String

    '    ID = txtRegno.Text

    '    update_Comstr = "UPDATE admit SET Name=" & "'" & txtPatiname.Text & "'" & "," & _
    '    "Bedcategory=" & "'" & CB_BedCategory.Text & "'" & "," & _
    '    "Bedno=" & "'" & CB_Bedno.Text & "'" & "," & _
    '    "Charge=" & "'" & txtBedCharge.Text & "'" & "," & _
    '    "Age=" & "'" & txtAge.Text & "'" & "," & _
    '    "Address=" & "'" & RtxtAddress.Text & "'" & "," & _
    '    "Phone=" & "'" & txtphone.Text & "'" & "," & _
    '    "Admitdatetime=" & "'" & DTP_date.Value & "'" & "," & _
    '    "Disease=" & "'" & txtDisease.Text & "'" & "," & _
    '    "Bloodgroup=" & "'" & CB_Bloodgp.Text & "'" & "," & _
    '    "Doctor=" & "'" & cbDoctor.Text & "'" & "," & _
    '     "Remarks=" & "'" & Rtxt_Remark.Text & "'" & " WHERE Regno=" & ID & " ; "

    '    update_Con = New OleDbConnection(update_Constr)
    '    update_Com = New OleDbCommand(update_Comstr, update_Con)

    '    update_Con.Open()
    '    Try
    '        update_Com.ExecuteNonQuery()
    '        MsgBox("Patient record updated Successfully.", MsgBoxStyle.Information, "Edit Registration")
    '    Catch ex As OleDbException
    '        MsgBox(ex.Message)
    '    End Try
    '    update_Con.Close()
    '    fill_listview()
    'End Sub

    'delete record
    'Private Sub btnDeleteRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    'Dim delete_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    'Dim delete_Comstr As String
    '    'Dim delete_Con As OleDbConnection
    '    'Dim delete_Com As OleDbCommand

    '    'Dim ID As String

    '    'ID = txtRegno.Text
    '    'delete_Comstr = "Delete from admit WHERE Regno=" & ID

    '    'delete_Con = New OleDbConnection(delete_Constr)
    '    'delete_Com = New OleDbCommand(delete_Comstr, delete_Con)

    '    'delete_Con.Open()
    '    'Try
    '    '    delete_Com.ExecuteNonQuery()
    '    '    MsgBox("Patient record deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
    '    'Catch ex As OleDbException
    '    '    MsgBox(ex.Message)
    '    'End Try
    '    'delete_Con.Close()
    '    'fill_listview()
    '    delete_Record()
    'End Sub

    'Private Sub btnCancelRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    txtRegno.Text = ""
    '    CB_BedCategory.Text = ""
    '    CB_Bedno.Text = ""
    '    txtBedCharge.Text = ""
    '    txtPatiname.Text = ""
    '    txtAge.Text = ""
    '    RtxtAddress.Text = ""
    '    txtphone.Text = ""
    '    'DTP_date.Value = 
    '    'DTP_time.Value = ""
    '    txtDisease.Text = ""
    '    CB_Bloodgp.Text = ""
    '    cbDoctor.Text = ""
    '    Rtxt_Remark.Text = ""

    '    Fill_Textbox_Admit()
    'End Sub

    'Private Sub btnDischarge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    ''code for saving record into discharge table
    '    'Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    'Dim save_Comstr As String
    '    'Dim save_Con As OleDbConnection
    '    'Dim save_Com As OleDbCommand
    '    'Dim save_DA As OleDbDataAdapter
    '    'Dim save_DS As DataSet
    '    'Dim save_DT As DataTable

    '    Dim currentdate As Date
    '    currentdate = currentdate.Now

    '    'save_Comstr = "insert into discharge(Regno,Bedcategory,Bedno,Charge,Name,Age,Address,Phone,Admitdatetime,Disease,Bloodgroup,Doctor,Remarks,Dischargedate)values" & _
    '    '("(" + "'" + txtRegno.Text + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + "'" + CB_Bedno.Text + "'" + "," + "'" + txtBedCharge.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" _
    '    '+ txtAge.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + txtphone.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + txtDisease.Text + "'" _
    '    '+ "," + "'" + CB_Bloodgp.Text + "'" + "," + "'" + cbDoctor.Text + "'" + "," + "'" + Rtxt_Remark.Text + "'" + "," + "'" + currentdate + "'" + ")")

    '    'save_Con = New OleDbConnection(save_Constr)
    '    'save_Com = New OleDbCommand(save_Comstr, save_Con)

    '    'save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
    '    'Dim save_builder As New OleDbCommandBuilder(save_DA)

    '    'save_DS = New DataSet()

    '    'Try
    '    '    save_DA.Fill(save_DS, "hm")
    '    '    MsgBox("Patient discharged successfully.", MsgBoxStyle.Information, "Register Patient")
    '    '    'save_Con.Close()

    '    '    'call this function for deleting the patient from admit list after discharging
    '    '    delete_Record()
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message)
    '    'End Try 'code end for saving record into discharge table


    '    ''''''''''''''''''''''''''
    '    'code for saving data into bill table
    '    Dim savebill_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    Dim savebill_Comstr As String
    '    Dim savebill_Con As OleDbConnection
    '    Dim savebill_Com As OleDbCommand
    '    Dim savebill_DA As OleDbDataAdapter
    '    Dim savebill_DS As DataSet
    '    Dim savebill_DT As DataTable

    '    Dim bill_currentdate As Date
    '    bill_currentdate = currentdate.Today


    '    Dim totBedcharge As Integer
    '    Dim totDay As Integer
    '    Dim bedchargeperday As Integer

    '    totDay = bill_currentdate.Today.Day - DTP_date.Value.Day
    '    bedchargeperday = Val(txtBedCharge.Text)
    '    totBedcharge = totDay * bedchargeperday

    '    Dim txtTotBedcharge As New TextBox()
    '    Dim txtOperationother As New TextBox()
    '    Dim txtTotal As New TextBox()
    '    txtTotBedcharge.Text = totBedcharge
    '    txtOperationother.Text = "0"
    '    txtTotal.Text = "0"

    '    'Dim Total_Bill As Integer
    '    'Total_Bill = totBedcharge + Othercharge
    '    'Try
    '    '    savebill_Comstr = "insert into Bill(Regno,Name,Admitdate,Bedtype,Dischargedate,Chargeperday,Bedcharge,Operationother,Total)values" & _
    '    '("(" + "'" + txtRegno.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + "'" + currentdate + "'" + "," + _
    '    '"'" + txtBedCharge.Text + "'" + "," + "'" + "100" + "'" + "," + "'" + "" + "'" + "," + "'" + "" + "'" + ")")
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message)
    '    'End Try

    '    savebill_Comstr = "insert into Bill(Regno,Name,Admitdate,Dischargedate,Bedtype,Chargeperday,Bedcharge,Operationother,Total)values" & _
    '    ("(" + "'" + txtRegno.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + currentdate + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + _
    '    "'" + txtBedCharge.Text + "'" + "," + "'" + txtTotBedcharge.Text + "'" + "," + "'" + txtOperationother.Text + "'" + "," + "'" + txtTotal.Text + "'" + ")")

    '    savebill_Con = New OleDbConnection(savebill_Constr)
    '    savebill_Com = New OleDbCommand(savebill_Comstr, savebill_Con)

    '    savebill_DA = New OleDbDataAdapter(savebill_Comstr, savebill_Constr)
    '    Dim savebill_builder As New OleDbCommandBuilder(savebill_DA)

    '    savebill_DS = New DataSet()

    '    Try
    '        savebill_DA.Fill(savebill_DS, "hm")
    '        MsgBox("Patient bill amount filled successfully in bill record.", MsgBoxStyle.Information, "Register Patient")
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    '    savebill_Con.Close() 'code end for saving data into bill table


    '    'code for refilling the bed occupied by the discharged patient
    '    Dim saveBed_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    Dim saveBed_Comstr As String
    '    Dim saveBed_Con As OleDbConnection
    '    Dim saveBed_Com As OleDbCommand
    '    Dim saveBed_DA As OleDbDataAdapter
    '    Dim saveBed_DS As DataSet
    '    Dim saveBed_DT As DataTable

    '    Dim saveBed_currentdate As Date

    '    'check the bed type

    '    If CB_BedCategory.Text = "General" Then

    '        saveBed_Comstr = "insert into Generalward(Bedno)values" & _
    '               ("(" + "'" + CB_Bedno.Text + "'" + ")")

    '    ElseIf CB_BedCategory.Text = "Special" Then

    '        saveBed_Comstr = "insert into Specialward(Bedno)values" & _
    '               ("(" + "'" + CB_Bedno.Text + "'" + ")")

    '    Else
    '        Exit Sub

    '    End If


    '    saveBed_Con = New OleDbConnection(saveBed_Constr)
    '    saveBed_Com = New OleDbCommand(saveBed_Comstr, saveBed_Con)

    '    saveBed_DA = New OleDbDataAdapter(saveBed_Comstr, saveBed_Constr)
    '    Dim saveBed_builder As New OleDbCommandBuilder(saveBed_DA)

    '    saveBed_DS = New DataSet()

    '    Try
    '        saveBed_DA.Fill(saveBed_DS, "hm")
    '        MsgBox("Bed refilled in bed record.", MsgBoxStyle.Information, "Register Patient")
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    '    savebill_Con.Close() 'code ends for refilling the bed occupied by the discharged patient

    '    'code for saving record into discharge table
    '    Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
    '    Dim save_Comstr As String
    '    Dim save_Con As OleDbConnection
    '    Dim save_Com As OleDbCommand
    '    Dim save_DA As OleDbDataAdapter
    '    Dim save_DS As DataSet
    '    Dim save_DT As DataTable

    '    'Dim currentdate As Date
    '    'currentdate = currentdate.Now

    '    save_Comstr = "insert into discharge(Regno,Bedcategory,Bedno,Charge,Name,Age,Address,Phone,Admitdatetime,Disease,Bloodgroup,Doctor,Remarks,Dischargedate)values" & _
    '    ("(" + "'" + txtRegno.Text + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + "'" + CB_Bedno.Text + "'" + "," + "'" + txtBedCharge.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" _
    '    + txtAge.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + txtphone.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + txtDisease.Text + "'" _
    '    + "," + "'" + CB_Bloodgp.Text + "'" + "," + "'" + cbDoctor.Text + "'" + "," + "'" + Rtxt_Remark.Text + "'" + "," + "'" + currentdate + "'" + ")")

    '    save_Con = New OleDbConnection(save_Constr)
    '    save_Com = New OleDbCommand(save_Comstr, save_Con)

    '    save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
    '    Dim save_builder As New OleDbCommandBuilder(save_DA)

    '    save_DS = New DataSet()

    '    Try
    '        save_DA.Fill(save_DS, "hm")
    '        MsgBox("Patient discharged successfully.", MsgBoxStyle.Information, "Register Patient")
    '        'save_Con.Close()

    '        'call this function for deleting the patient from admit list after discharging
    '        delete_Record()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try 'code end for saving record into discharge table
    'End Sub


#End Region

#Region "Function for deleting record"

    Sub delete_Record()
        Dim delete_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim delete_Comstr As String
        Dim delete_Con As OleDbConnection
        Dim delete_Com As OleDbCommand

        Dim ID As String

        ID = txtRegno.Text
        delete_Comstr = "Delete from admit WHERE Regno=" & ID

        delete_Con = New OleDbConnection(delete_Constr)
        delete_Com = New OleDbCommand(delete_Comstr, delete_Con)

        delete_Con.Open()
        Try
            delete_Com.ExecuteNonQuery()
            MsgBox("Patient record deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
        delete_Con.Close()

        Fill_Textbox_Admit()
        fill_listview()
    End Sub
#End Region

#Region "Load event of frmAdmitPatient"
    Private Sub frmAdmitPatient_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Fill_Textbox_Admit()
        fill_listview()

        If LV_PatiView.Items.Count = 0 Then
            'MessageBox.Show("There is no record to show.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnRegisterPatient.Enabled = True
            btnEditRegistration.Enabled = False
            'btnDeleteRegistration.Enabled = False
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False
            btnDischarge.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False

            CB_BedCategory.Enabled = False
            CB_Bedno.Enabled = False


            'code when listview is empty
            txtBedCharge.ReadOnly = True
            txtPatiname.ReadOnly = True
            txtAge.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtphone.ReadOnly = True
            DTP_date.Enabled = False
            txtDisease.ReadOnly = True
            CB_Bloodgp.Enabled = False
            cbDoctor.Enabled = False
            Rtxt_Remark.ReadOnly = True
        Else
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False

            CB_BedCategory.Enabled = False
            CB_Bedno.Enabled = False

            'code when listview is empty
            txtBedCharge.ReadOnly = False
            txtPatiname.ReadOnly = False
            txtAge.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtphone.ReadOnly = False
            DTP_date.Enabled = True
            txtDisease.ReadOnly = False
            CB_Bloodgp.Enabled = True
            cbDoctor.Enabled = True
            Rtxt_Remark.ReadOnly = False
        End If
    End Sub
#End Region

#Region "Fuction for filling Listview"

    'fill listview
    Sub fill_listview()
        ' Clear list view column headers and items
        LV_PatiView.Columns.Clear()
        LV_PatiView.Items.Clear()

        'If LV.Items.Count > 0 Then
        'LV.Items.Clear()
        'End If

        ' Get SQL Query from textbox
        Dim fillLV_comstr As String = "select * from admit"
        Dim fillLV_Con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb")

        ' Create Command object
        Dim NewQuery As New OleDbCommand(fillLV_comstr, fillLV_Con)

        Try
            ' Open Connection
            Try
                fillLV_Con.Open()
            Catch e As Exception
                MsgBox(e.Message)
            End Try


            ' Execute Command and Get Data 
            Dim NewReader As OleDbDataReader = NewQuery.ExecuteReader()

            ' Get column names for list view from data reader

            Dim i As Integer
            For i = 0 To NewReader.FieldCount - 1
                Dim header As New ColumnHeader()
                header.Text = NewReader.GetName(i)
                LV_PatiView.Columns.Add(header)
            Next

            ' Get rows of data and show in list view
            While NewReader.Read()
                ' Create list view item
                Dim NewItem As New ListViewItem()

                ' Specify text and subitems of list view
                NewItem.Text = NewReader.GetValue(0).ToString()
                For i = 1 To NewReader.FieldCount - 1
                    NewItem.SubItems.Add(NewReader.GetValue(i).ToString())
                Next

                ' Add item to list view items collection
                LV_PatiView.Items.Add(NewItem)
            End While

            ' Close data reader
            NewReader.Close()

        Catch ex As OleDbException
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "oledb Error"
            LV_PatiView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem(ex.Message)
            LV_PatiView.Items.Add(ErrorItem)

        Catch ex As Exception
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "Error"
            LV_PatiView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem("An error has occurred")
            LV_PatiView.Items.Add(ErrorItem)

        Finally
            fillLV_Con.Close()
        End Try
    End Sub
#End Region

#Region "Listview event"

    Private Sub LV_PatiView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LV_PatiView.SelectedIndexChanged
        Dim SelectedItems As ListView.SelectedListViewItemCollection = _
      CType(sender, ListView).SelectedItems
        If (SelectedItems.Count > 0) Then
            'MessageBox.Show(SelectedItems(0).SubItems(0).Text & _
            '    " " & SelectedItems(0).SubItems(1).Text & _
            '    " " & SelectedItems(0).SubItems(2).Text & _
            '    " " & SelectedItems(0).SubItems(3).Text)

            txtRegno.Text = SelectedItems(0).SubItems(0).Text
            CB_BedCategory.Text = SelectedItems(0).SubItems(1).Text
            CB_Bedno.Text = SelectedItems(0).SubItems(2).Text
            txtBedCharge.Text = SelectedItems(0).SubItems(3).Text
            txtPatiname.Text = SelectedItems(0).SubItems(4).Text
            txtAge.Text = SelectedItems(0).SubItems(5).Text
            RtxtAddress.Text = SelectedItems(0).SubItems(6).Text
            txtphone.Text = SelectedItems(0).SubItems(7).Text
            DTP_date.Value = SelectedItems(0).SubItems(8).Text
            txtDisease.Text = SelectedItems(0).SubItems(9).Text
            CB_Bloodgp.Text = SelectedItems(0).SubItems(10).Text
            cbDoctor.Text = SelectedItems(0).SubItems(11).Text
            Rtxt_Remark.Text = SelectedItems(0).SubItems(12).Text

        End If
    End Sub
#End Region

#Region "Function for saving record , remove bed , adding ID"
    Sub save_record()
        Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_Comstr As String
        Dim save_Con As OleDbConnection
        Dim save_Com As OleDbCommand
        Dim save_DA As OleDbDataAdapter
        Dim save_DS As DataSet
        Dim save_DT As DataTable
        Dim txt As String
        txt = CType(RtxtAddress.Text, String)
        save_Comstr = "insert into admit(Regno,Bedcategory,Bedno,Charge,Name,Age,Address,Phone,Admitdatetime,Disease,Bloodgroup,Doctor,Remarks)values" & _
        ("(" + "'" + txtRegno.Text + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," _
        + "'" + CB_Bedno.Text + "'" + "," + "'" + txtBedCharge.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," _
        + "'" + txtAge.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," _
        + "'" + txtphone.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," _
        + "'" + txtDisease.Text + "'" + "," + "'" + CB_Bloodgp.Text + "'" + "," _
        + "'" + cbDoctor.Text + "'" + "," + "'" + Rtxt_Remark.Text + "'" + ")")

        save_Con = New OleDbConnection(save_Constr)
        save_Com = New OleDbCommand(save_Comstr, save_Con)

        save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
        Dim save_builder As New OleDbCommandBuilder(save_DA)

        save_DS = New DataSet()

        Try
            save_DA.Fill(save_DS, "hm")
            MsgBox("Patient record inserted successfully.", MsgBoxStyle.Information, "Register Patient")

            CB_BedCategory.Enabled = False
            CB_Bedno.Enabled = False

            Dim obj As New frmMain()
            obj.setFlag = 0
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        save_Con.Close()

        '''''''''''''''''''''''''''''
        'code for removing the bed no which is selected above
        'from the general ward.
        If CB_BedCategory.Text = "General" Then
            Dim remove_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            Dim remove_Comstr As String
            Dim remove_Con As OleDbConnection
            Dim remove_Com As OleDbCommand

            Dim ID As String

            ID = CB_Bedno.Text
            remove_Comstr = "Delete from Generalward WHERE Bedno=" & ID

            remove_Con = New OleDbConnection(remove_Constr)
            remove_Com = New OleDbCommand(remove_Comstr, remove_Con)

            remove_Con.Open()
            Try
                remove_Com.ExecuteNonQuery()
                'MsgBox("Bed deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
            Catch ex As OleDbException
                'MsgBox(ex.Message)
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            remove_Con.Close()

        ElseIf CB_BedCategory.Text = "Special" Then
            Dim remove_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            Dim remove_Comstr As String
            Dim remove_Con As OleDbConnection
            Dim remove_Com As OleDbCommand

            Dim ID As String

            ID = CB_Bedno.Text
            remove_Comstr = "Delete from Specialward WHERE Bedno=" & ID

            remove_Con = New OleDbConnection(remove_Constr)
            remove_Com = New OleDbCommand(remove_Comstr, remove_Con)

            remove_Con.Open()
            Try
                remove_Com.ExecuteNonQuery()
                'MsgBox("Bed deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
            Catch ex As OleDbException
                'MsgBox(ex.Message)
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            remove_Con.Close()
        End If

        'code for adding ID in admitID table
        Dim save_admitID_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_admitID_Comstr As String
        Dim save_admitID_Con As OleDbConnection
        Dim save_admitID_Com As OleDbCommand
        Dim save_admitID_DA As OleDbDataAdapter
        Dim save_admitID_DS As DataSet
        Dim save_admitID_DT As DataTable

        save_admitID_Comstr = "insert into admitID(aID)values" & _
        ("(" + "'" + txtRegno.Text + "'" + ")")

        save_admitID_Con = New OleDbConnection(save_admitID_Constr)
        save_admitID_Com = New OleDbCommand(save_admitID_Comstr, save_admitID_Con)

        save_admitID_DA = New OleDbDataAdapter(save_admitID_Comstr, save_admitID_Constr)
        Dim save_admitID_builder As New OleDbCommandBuilder(save_admitID_DA)

        save_admitID_DS = New DataSet()

        Try
            save_admitID_DA.Fill(save_admitID_DS, "hm")
            ' MsgBox("ID inserted in admitID table successfully.", MsgBoxStyle.Information, "Register Patient")
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        save_admitID_Con.Close() 'code ends for adding ID in admitID table
    End Sub

#End Region

#Region "Function for saving blank record , remove bed , adding ID"
    Sub save_record_Solution()
        Dim save_Comstr As String

        If Rtxt_Remark.Text = "" Then
            Rtxt_Remark.Text = "No remark"

            save_Comstr = "insert into admit(Regno,Bedcategory,Bedno,Charge,Name,Age,Address,Phone,Admitdatetime,Disease,Bloodgroup,Doctor,Remarks)values" & _
                   ("(" + "'" + txtRegno.Text + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," _
                   + "'" + CB_Bedno.Text + "'" + "," + "'" + txtBedCharge.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," _
                   + "'" + txtAge.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," _
                   + "'" + txtphone.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," _
                   + "'" + txtDisease.Text + "'" + "," + "'" + CB_Bloodgp.Text + "'" + "," _
                   + "'" + cbDoctor.Text + "'" + "," + "'" + Rtxt_Remark.Text + "'" + ")")
        End If


        Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim save_Comstr As String
        Dim save_Con As OleDbConnection
        Dim save_Com As OleDbCommand
        Dim save_DA As OleDbDataAdapter
        Dim save_DS As DataSet
        Dim save_DT As DataTable


        save_Con = New OleDbConnection(save_Constr)
        save_Com = New OleDbCommand(save_Comstr, save_Con)

        save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
        Dim save_builder As New OleDbCommandBuilder(save_DA)

        save_DS = New DataSet()

        Try
            save_DA.Fill(save_DS, "hm")
            MsgBox("Patient record inserted successfully.", MsgBoxStyle.Information, "Register Patient")

            CB_BedCategory.Enabled = False
            CB_Bedno.Enabled = False

            Dim obj As New frmMain()
            obj.setFlag = 0
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        save_Con.Close()

        '''''''''''''''''''''''''''''
        'code for removing the bed no which is selected above
        'from the general ward.
        If CB_BedCategory.Text = "General" Then
            Dim remove_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            Dim remove_Comstr As String
            Dim remove_Con As OleDbConnection
            Dim remove_Com As OleDbCommand

            Dim ID As String

            ID = CB_Bedno.Text
            remove_Comstr = "Delete from Generalward WHERE Bedno=" & ID

            remove_Con = New OleDbConnection(remove_Constr)
            remove_Com = New OleDbCommand(remove_Comstr, remove_Con)

            remove_Con.Open()
            Try
                remove_Com.ExecuteNonQuery()
                'MsgBox("Bed deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
            Catch ex As OleDbException
                'MsgBox(ex.Message)
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            remove_Con.Close()

        ElseIf CB_BedCategory.Text = "Special" Then
            Dim remove_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            Dim remove_Comstr As String
            Dim remove_Con As OleDbConnection
            Dim remove_Com As OleDbCommand

            Dim ID As String

            ID = CB_Bedno.Text
            remove_Comstr = "Delete from Specialward WHERE Bedno=" & ID

            remove_Con = New OleDbConnection(remove_Constr)
            remove_Com = New OleDbCommand(remove_Comstr, remove_Con)

            remove_Con.Open()
            Try
                remove_Com.ExecuteNonQuery()
                'MsgBox("Bed deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
            Catch ex As OleDbException
                'MsgBox(ex.Message)
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            remove_Con.Close()

        End If

        'code for adding ID in admitID table
        Dim save_admitID_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_admitID_Comstr As String
        Dim save_admitID_Con As OleDbConnection
        Dim save_admitID_Com As OleDbCommand
        Dim save_admitID_DA As OleDbDataAdapter
        Dim save_admitID_DS As DataSet
        Dim save_admitID_DT As DataTable

        save_admitID_Comstr = "insert into admitID(aID)values" & _
        ("(" + "'" + txtRegno.Text + "'" + ")")

        save_admitID_Con = New OleDbConnection(save_admitID_Constr)
        save_admitID_Com = New OleDbCommand(save_admitID_Comstr, save_admitID_Con)

        save_admitID_DA = New OleDbDataAdapter(save_admitID_Comstr, save_admitID_Constr)
        Dim save_admitID_builder As New OleDbCommandBuilder(save_admitID_DA)

        save_admitID_DS = New DataSet()

        Try
            save_admitID_DA.Fill(save_admitID_DS, "hm")
            'MsgBox("ID inserted in admitID table successfully.", MsgBoxStyle.Information, "Register Patient")
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        save_admitID_Con.Close() 'code ends for adding ID in admitID table
    End Sub
#End Region

#Region "Function for checking blank values"
    'this function check for blank values in form
    Sub Check_Textbox()

        Dim age As Integer
        age = txtAge.Text

        If age > 100 Then
            MessageBox.Show("Age should be within 120.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim r As DialogResult

        If CB_BedCategory.Text = "" Or CB_Bedno.Text = "" _
            Or txtBedCharge.Text = "" Or txtPatiname.Text = "" _
            Or txtAge.Text = "" _
            Or RtxtAddress.Text = "" _
            Or txtphone.Text = "" _
            Or txtDisease.Text = "" _
            Or CB_Bloodgp.Text = "" _
            Or cbDoctor.Text = "" Then
            'Or Rtxt_Remark.Text = "" Then
            MessageBox.Show("You must fill all the fields.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        ElseIf (CB_BedCategory.Text <> "" Or CB_Bedno.Text <> "" _
                   Or txtBedCharge.Text <> "" Or txtPatiname.Text <> "" _
                   Or txtAge.Text <> "" _
                   Or RtxtAddress.Text <> "" _
                   Or txtphone.Text <> "" _
                   Or txtDisease.Text <> "" _
                   Or CB_Bloodgp.Text <> "" _
                   Or cbDoctor.Text <> "") And Rtxt_Remark.Text = "" Then
            'Or Rtxt_Remark.Text = "" Then
            'MessageBox.Show("You must fill all the fields.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            r = MessageBox.Show("Remark field is empty. Do you want to save it?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        Else
            save_record()
            fill_listview()

            If LV_PatiView.Items.Count = 0 Then
                btnRegisterPatient.Enabled = True
                btnEditRegistration.Enabled = False
                'btnDeleteRegistration.Enabled = False
                btnSaveRegistration.Enabled = False
                btnCancelRegistration.Enabled = False
                btnDischarge.Enabled = False
                'btnPrint.Enabled = False
                btnCustomprint.Enabled = False
                btnSearch.Enabled = False
                btnExcel.Enabled = False
            Else
                btnRegisterPatient.Enabled = True
                btnEditRegistration.Enabled = True
                'btnDeleteRegistration.Enabled = True
                btnSaveRegistration.Enabled = False
                btnCancelRegistration.Enabled = False
                btnDischarge.Enabled = True
                'btnPrint.Enabled = True
                btnCustomprint.Enabled = True
                btnSearch.Enabled = True
                btnExcel.Enabled = True
            End If
        End If

        If r = DialogResult.Yes Then
            save_record_Solution()
            fill_listview()

            If LV_PatiView.Items.Count = 0 Then
                btnRegisterPatient.Enabled = True
                btnEditRegistration.Enabled = False
                'btnDeleteRegistration.Enabled = False
                btnSaveRegistration.Enabled = False
                btnCancelRegistration.Enabled = False
                btnDischarge.Enabled = False
                'btnPrint.Enabled = False
                btnCustomprint.Enabled = False
                btnSearch.Enabled = False
                btnExcel.Enabled = False
            Else
                btnRegisterPatient.Enabled = True
                btnEditRegistration.Enabled = True
                'btnDeleteRegistration.Enabled = True
                btnSaveRegistration.Enabled = False
                btnCancelRegistration.Enabled = False
                btnDischarge.Enabled = True
                'btnPrint.Enabled = True
                btnCustomprint.Enabled = True
                btnSearch.Enabled = True
                btnExcel.Enabled = True
            End If
        End If

    End Sub
#End Region

#Region "All textbox Keypress event"
    Private Sub txtBedCharge_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBedCharge.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtPatiname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPatiname.KeyPress
        If Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtAge_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAge.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub

    'Private Sub RtxtAddress_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RtxtAddress.KeyPress
    '    If Char.IsLetterOrDigit(e.KeyChar) = False Then
    '        If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
    '            e.Handled = False
    '        Else
    '            e.Handled = True
    '        End If
    '    End If
    'End Sub

    Private Sub txtphone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtphone.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtDisease_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDisease.KeyPress
        If Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub

    'Private Sub Rtxt_Remark_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Rtxt_Remark.KeyPress
    '    If Char.IsLetterOrDigit(e.KeyChar) = False Then
    '        If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
    '            e.Handled = False
    '        Else
    '            e.Handled = True
    '        End If
    '    End If
    'End Sub

#End Region

#Region "CB_BedCategory event i.e. bed type"

    Private Sub CB_BedCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_BedCategory.SelectedIndexChanged
        If CB_BedCategory.Text = "General" Then
            CB_Bedno.Enabled = True
            CB_Bedno.Items.Clear()

            Dim row1 As String
            Dim insertCB_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
            Dim insertCB_cmdstr As String = "select * from Generalward"
            Dim insertCB_con As New OleDbConnection(insertCB_constr)
            Dim insertCB_com As New OleDbCommand(insertCB_cmdstr, insertCB_con)
            insertCB_con.Open()
            Dim insertCB_DR As OleDbDataReader = insertCB_com.ExecuteReader
            While insertCB_DR.Read
                row1 = insertCB_DR(0)
                CB_Bedno.Items.Add(row1)
            End While
            insertCB_con.Close()

        ElseIf CB_BedCategory.Text = "Special" Then
            CB_Bedno.Enabled = True
            CB_Bedno.Items.Clear()

            Dim row1 As String
            Dim insertCB_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
            Dim insertCB_cmdstr As String = "select * from Specialward"
            Dim insertCB_con As New OleDbConnection(insertCB_constr)
            Dim insertCB_com As New OleDbCommand(insertCB_cmdstr, insertCB_con)
            insertCB_con.Open()
            Dim insertCB_DR As OleDbDataReader = insertCB_com.ExecuteReader
            While insertCB_DR.Read
                row1 = insertCB_DR(0)
                CB_Bedno.Items.Add(row1)
            End While
            insertCB_con.Close()
        End If
    End Sub
#End Region

#Region "ALL BUTTON EVENT"

    Private Sub btnRegisterPatient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegisterPatient.Click

        ''''''''''GENERAL WARD ''''''''''''''''''''''''''''''''''''''''''''''''
        Dim row As String
        Dim insertCB_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
        Dim insertCB_cmdstr As String = "select * from Generalward"
        Dim insertCB_con As New OleDbConnection(insertCB_constr)
        Dim insertCB_com As New OleDbCommand(insertCB_cmdstr, insertCB_con)
        insertCB_con.Open()
        Dim insertCB_DR As OleDbDataReader = insertCB_com.ExecuteReader
        While insertCB_DR.Read
            row = insertCB_DR(0)
        End While
        insertCB_con.Close()
        '''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''SPECIAL WARD''''''''''''''''''''''''''''''''''''''
        Dim row1 As String
        Dim insertCB_constr1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
        Dim insertCB_cmdstr1 As String = "select * from Specialward"
        Dim insertCB_con1 As New OleDbConnection(insertCB_constr1)
        Dim insertCB_com1 As New OleDbCommand(insertCB_cmdstr1, insertCB_con1)
        insertCB_con1.Open()
        Dim insertCB_DR1 As OleDbDataReader = insertCB_com1.ExecuteReader
        While insertCB_DR1.Read
            row1 = insertCB_DR1(0)
        End While
        insertCB_con1.Close()
        ''''''''''''''''''''''''''''''''''

        If row = Nothing And row1 = Nothing Then
            MessageBox.Show("Beds are not available.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'CB_BedCategory.Items.Clear()
            'btnRegisterPatient.Enabled = False
            'btnEditRegistration.Enabled = False
            ''btnDeleteRegistration.Enabled = False
            'btnSaveRegistration.Enabled = False
            'btnCancelRegistration.Enabled = True
            'btnDischarge.Enabled = False
            ''btnPrint.Enabled = False
            'btnCustomprint.Enabled = False
            'btnSearch.Enabled = False
            'btnExcel.Enabled = False

            'CB_BedCategory.Enabled = False
            'CB_Bedno.Enabled = False


            ''code when listview is empty
            'txtBedCharge.ReadOnly = True
            'txtPatiname.ReadOnly = True
            'txtAge.ReadOnly = True
            'RtxtAddress.ReadOnly = True
            'txtphone.ReadOnly = True
            'DTP_date.Enabled = False
            'txtDisease.ReadOnly = True
            'CB_Bloodgp.Enabled = False
            'cbDoctor.Enabled = False
            'Rtxt_Remark.ReadOnly = True
            Exit Sub
        End If

        If row = Nothing Then
            CB_BedCategory.Items.Clear()
            CB_BedCategory.Items.Add("Special")
        ElseIf row1 = Nothing Then
            CB_BedCategory.Items.Clear()
            CB_BedCategory.Items.Add("General")
        ElseIf ((Not row = Nothing) And (Not row1 = Nothing)) Then
            CB_BedCategory.Items.Clear()
            CB_BedCategory.Items.Add("General")
            CB_BedCategory.Items.Add("Special")
        End If

        ''''''''''''CODE''''''''''''''''''

        Dim r0, r1, r2, r3 As String

        Dim insert_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
        'Dim insert_cmdstr As String = "select * from admit" 'admitpatients"
        Dim insert_cmdstr As String = "select * from admitID"
        Dim insert_con As New OleDbConnection(insert_constr)
        Dim insert_com As New OleDbCommand(insert_cmdstr, insert_con)
        insert_con.Open()
        Dim insert_DR As OleDbDataReader = insert_com.ExecuteReader
        While insert_DR.Read
            r0 = insert_DR(0)
        End While
        insert_con.Close()

        CB_BedCategory.Text = ""
        CB_Bedno.Text = ""
        txtBedCharge.Text = ""
        txtPatiname.Text = ""
        txtAge.Text = ""
        RtxtAddress.Text = ""
        txtphone.Text = ""
        DTP_date.Value = System.DateTime.Today
        'DTP_time = System.DateTime.Now
        txtDisease.Text = ""
        CB_Bloodgp.Text = ""
        cbDoctor.Text = ""
        Rtxt_Remark.Text = ""

        txtRegno.Text = r0 + 1

        Dim obj As New frmMain()
        obj.setFlag = 1

        btnRegisterPatient.Enabled = False
        btnEditRegistration.Enabled = False
        'btnDeleteRegistration.Enabled = False
        btnSaveRegistration.Enabled = True
        btnCancelRegistration.Enabled = True
        btnDischarge.Enabled = False

        CB_BedCategory.Enabled = True

        'code when listview is empty
        txtBedCharge.ReadOnly = False
        txtPatiname.ReadOnly = False
        txtAge.ReadOnly = False
        RtxtAddress.ReadOnly = False
        txtphone.ReadOnly = False
        DTP_date.Enabled = True
        txtDisease.ReadOnly = False
        CB_Bloodgp.Enabled = True
        cbDoctor.Enabled = True
        Rtxt_Remark.ReadOnly = False


    End Sub

    Private Sub btnSaveRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveRegistration.Click
        Check_Textbox()
    End Sub

    Private Sub btnEditRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditRegistration.Click

        Dim update_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim update_Comstr As String
        Dim update_Con As OleDbConnection
        Dim update_Com As OleDbCommand

        Dim ID As String

        ID = txtRegno.Text

        update_Comstr = "UPDATE admit SET Name=" & "'" & txtPatiname.Text & "'" & "," & _
        "Bedcategory=" & "'" & CB_BedCategory.Text & "'" & "," & _
        "Bedno=" & "'" & CB_Bedno.Text & "'" & "," & _
        "Charge=" & "'" & txtBedCharge.Text & "'" & "," & _
        "Age=" & "'" & txtAge.Text & "'" & "," & _
        "Address=" & "'" & RtxtAddress.Text & "'" & "," & _
        "Phone=" & "'" & txtphone.Text & "'" & "," & _
        "Admitdatetime=" & "'" & DTP_date.Value & "'" & "," & _
        "Disease=" & "'" & txtDisease.Text & "'" & "," & _
        "Bloodgroup=" & "'" & CB_Bloodgp.Text & "'" & "," & _
        "Doctor=" & "'" & cbDoctor.Text & "'" & "," & _
         "Remarks=" & "'" & Rtxt_Remark.Text & "'" & " WHERE Regno=" & ID & " ; "

        update_Con = New OleDbConnection(update_Constr)
        update_Com = New OleDbCommand(update_Comstr, update_Con)

        update_Con.Open()
        Try
            update_Com.ExecuteNonQuery()
            MsgBox("Patient record updated Successfully.", MsgBoxStyle.Information, "Edit Registration")
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
        update_Con.Close()
        fill_listview()

        If LV_PatiView.Items.Count = 0 Then
            btnRegisterPatient.Enabled = True
            btnEditRegistration.Enabled = False
            'btnDeleteRegistration.Enabled = False
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False
            btnDischarge.Enabled = False
        Else
            btnRegisterPatient.Enabled = True
            btnEditRegistration.Enabled = True
            'btnDeleteRegistration.Enabled = True
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False
            btnDischarge.Enabled = True
        End If
    End Sub

    'DELETE BUTTON CODE
    'Private Sub btnDeleteRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteRegistration.Click
    '    delete_Record()
    '    'btnRegisterPatient.Enabled = True
    '    'btnEditRegistration.Enabled = True
    '    'btnDeleteRegistration.Enabled = True
    '    'btnSaveRegistration.Enabled = False
    '    'btnCancelRegistration.Enabled = False

    '    If LV_PatiView.Items.Count = 0 Then
    '        btnRegisterPatient.Enabled = True
    '        btnEditRegistration.Enabled = False
    '        btnDeleteRegistration.Enabled = False
    '        btnSaveRegistration.Enabled = False
    '        btnCancelRegistration.Enabled = False
    '        btnDischarge.Enabled = False
    '        'btnPrint.Enabled = False
    '        btnCustomprint.Enabled = False
    '        btnSearch.Enabled = False
    '        btnExcel.Enabled = False
    '    Else
    '        btnRegisterPatient.Enabled = True
    '        btnEditRegistration.Enabled = True
    '        btnDeleteRegistration.Enabled = True
    '        btnSaveRegistration.Enabled = False
    '        btnCancelRegistration.Enabled = False
    '        btnDischarge.Enabled = True
    '    End If
    'End Sub

    Private Sub btnCancelRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelRegistration.Click

        'btnRegisterPatient.Enabled = True
        'btnEditRegistration.Enabled = True
        'btnDeleteRegistration.Enabled = True
        'btnSaveRegistration.Enabled = False
        'btnCancelRegistration.Enabled = False

        'If LV_PatiView.Items.Count = 0 Then
        '    btnDischarge.Enabled = False
        'Else
        '    btnDischarge.Enabled = True
        'End If

        txtRegno.Text = ""
        CB_BedCategory.Text = ""
        CB_Bedno.Text = ""
        txtBedCharge.Text = ""
        txtPatiname.Text = ""
        txtAge.Text = ""
        RtxtAddress.Text = ""
        txtphone.Text = ""
        'DTP_date.Value = 
        'DTP_time.Value = ""
        txtDisease.Text = ""
        CB_Bloodgp.Text = ""
        cbDoctor.Text = ""
        Rtxt_Remark.Text = ""

        Fill_Textbox_Admit()

        If LV_PatiView.Items.Count = 0 Then
            btnRegisterPatient.Enabled = True
            btnEditRegistration.Enabled = False
            'btnDeleteRegistration.Enabled = False
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False
            btnDischarge.Enabled = False

            'code when listview is empty
            txtBedCharge.ReadOnly = True
            txtPatiname.ReadOnly = True
            txtAge.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtphone.ReadOnly = True
            DTP_date.Enabled = False
            txtDisease.ReadOnly = True
            CB_Bloodgp.Enabled = False
            cbDoctor.Enabled = False
            Rtxt_Remark.ReadOnly = True
        Else
            btnRegisterPatient.Enabled = True
            btnEditRegistration.Enabled = True
            'btnDeleteRegistration.Enabled = True
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False
            btnDischarge.Enabled = True

            'code when listview is empty
            txtBedCharge.ReadOnly = False
            txtPatiname.ReadOnly = False
            txtAge.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtphone.ReadOnly = False
            DTP_date.Enabled = True
            txtDisease.ReadOnly = False
            CB_Bloodgp.Enabled = True
            cbDoctor.Enabled = True
            Rtxt_Remark.ReadOnly = False
        End If

        CB_BedCategory.Enabled = False
        CB_Bedno.Enabled = False

        Dim obj As New frmMain()
        obj.setFlag = 0
    End Sub

    Private Sub btnDischarge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDischarge.Click
        ''code for saving record into discharge table
        'Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim save_Comstr As String
        'Dim save_Con As OleDbConnection
        'Dim save_Com As OleDbCommand
        'Dim save_DA As OleDbDataAdapter
        'Dim save_DS As DataSet
        'Dim save_DT As DataTable

        Dim currentdate As Date
        currentdate = currentdate.Now

        'save_Comstr = "insert into discharge(Regno,Bedcategory,Bedno,Charge,Name,Age,Address,Phone,Admitdatetime,Disease,Bloodgroup,Doctor,Remarks,Dischargedate)values" & _
        '("(" + "'" + txtRegno.Text + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + "'" + CB_Bedno.Text + "'" + "," + "'" + txtBedCharge.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" _
        '+ txtAge.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + txtphone.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + txtDisease.Text + "'" _
        '+ "," + "'" + CB_Bloodgp.Text + "'" + "," + "'" + cbDoctor.Text + "'" + "," + "'" + Rtxt_Remark.Text + "'" + "," + "'" + currentdate + "'" + ")")

        'save_Con = New OleDbConnection(save_Constr)
        'save_Com = New OleDbCommand(save_Comstr, save_Con)

        'save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
        'Dim save_builder As New OleDbCommandBuilder(save_DA)

        'save_DS = New DataSet()

        'Try
        '    save_DA.Fill(save_DS, "hm")
        '    MsgBox("Patient discharged successfully.", MsgBoxStyle.Information, "Register Patient")
        '    'save_Con.Close()

        '    'call this function for deleting the patient from admit list after discharging
        '    delete_Record()
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try 'code end for saving record into discharge table


        ''''''''''''''''''''''''''
        'code for saving data into bill table
        Dim savebill_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim savebill_Comstr As String
        Dim savebill_Con As OleDbConnection
        Dim savebill_Com As OleDbCommand
        Dim savebill_DA As OleDbDataAdapter
        Dim savebill_DS As DataSet
        Dim savebill_DT As DataTable

        Dim bill_currentdate As Date
        bill_currentdate = currentdate.Today


        Dim totBedcharge As Integer
        Dim totDay As Integer
        Dim bedchargeperday As Integer

        totDay = bill_currentdate.Today.Day - DTP_date.Value.Day
        bedchargeperday = Val(txtBedCharge.Text)
        totBedcharge = totDay * bedchargeperday

        Dim txtTotBedcharge As New TextBox()
        Dim txtOperationother As New TextBox()
        Dim txtTotal As New TextBox()
        txtTotBedcharge.Text = totBedcharge
        txtOperationother.Text = "0"
        txtTotal.Text = "0"

        'Dim Total_Bill As Integer
        'Total_Bill = totBedcharge + Othercharge
        'Try
        '    savebill_Comstr = "insert into Bill(Regno,Name,Admitdate,Bedtype,Dischargedate,Chargeperday,Bedcharge,Operationother,Total)values" & _
        '("(" + "'" + txtRegno.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + "'" + currentdate + "'" + "," + _
        '"'" + txtBedCharge.Text + "'" + "," + "'" + "100" + "'" + "," + "'" + "" + "'" + "," + "'" + "" + "'" + ")")
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        savebill_Comstr = "insert into Bill(Regno,Name,Admitdate,Dischargedate,Bedtype,Chargeperday,Bedcharge,Operationother,Total)values" & _
        ("(" + "'" + txtRegno.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + currentdate + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + _
        "'" + txtBedCharge.Text + "'" + "," + "'" + txtTotBedcharge.Text + "'" + "," + "'" + txtOperationother.Text + "'" + "," + "'" + txtTotal.Text + "'" + ")")

        savebill_Con = New OleDbConnection(savebill_Constr)
        savebill_Com = New OleDbCommand(savebill_Comstr, savebill_Con)

        savebill_DA = New OleDbDataAdapter(savebill_Comstr, savebill_Constr)
        Dim savebill_builder As New OleDbCommandBuilder(savebill_DA)

        savebill_DS = New DataSet()

        Try
            savebill_DA.Fill(savebill_DS, "hm")
            MsgBox("Patient bill amount filled successfully in bill record.", MsgBoxStyle.Information, "Register Patient")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        savebill_Con.Close() 'code end for saving data into bill table


        'code for refilling the bed occupied by the discharged patient
        Dim saveBed_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim saveBed_Comstr As String
        Dim saveBed_Con As OleDbConnection
        Dim saveBed_Com As OleDbCommand
        Dim saveBed_DA As OleDbDataAdapter
        Dim saveBed_DS As DataSet
        Dim saveBed_DT As DataTable

        Dim saveBed_currentdate As Date

        'check the bed type

        If CB_BedCategory.Text = "General" Then

            saveBed_Comstr = "insert into Generalward(Bedno)values" & _
                   ("(" + "'" + CB_Bedno.Text + "'" + ")")

        ElseIf CB_BedCategory.Text = "Special" Then

            saveBed_Comstr = "insert into Specialward(Bedno)values" & _
                   ("(" + "'" + CB_Bedno.Text + "'" + ")")

        Else
            Exit Sub

        End If


        saveBed_Con = New OleDbConnection(saveBed_Constr)
        saveBed_Com = New OleDbCommand(saveBed_Comstr, saveBed_Con)

        saveBed_DA = New OleDbDataAdapter(saveBed_Comstr, saveBed_Constr)
        Dim saveBed_builder As New OleDbCommandBuilder(saveBed_DA)

        saveBed_DS = New DataSet()

        Try
            saveBed_DA.Fill(saveBed_DS, "hm")
            MsgBox("Bed refilled in bed record.", MsgBoxStyle.Information, "Register Patient")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        saveBed_Con.Close() 'code ends for refilling the bed occupied by the discharged patient

        'code for saving record into discharge table
        Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_Comstr As String
        Dim save_Con As OleDbConnection
        Dim save_Com As OleDbCommand
        Dim save_DA As OleDbDataAdapter
        Dim save_DS As DataSet
        Dim save_DT As DataTable

        'Dim currentdate As Date
        'currentdate = currentdate.Now

        save_Comstr = "insert into discharge(Regno,Bedcategory,Bedno,Charge,Name,Age,Address,Phone,Admitdatetime,Disease,Bloodgroup,Doctor,Remarks,Dischargedate)values" & _
        ("(" + "'" + txtRegno.Text + "'" + "," + "'" + CB_BedCategory.Text + "'" + "," + "'" + CB_Bedno.Text + "'" + "," + "'" + txtBedCharge.Text + "'" + "," + "'" + txtPatiname.Text + "'" + "," + "'" _
        + txtAge.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + txtphone.Text + "'" + "," + "'" + DTP_date.Value + "'" + "," + "'" + txtDisease.Text + "'" _
        + "," + "'" + CB_Bloodgp.Text + "'" + "," + "'" + cbDoctor.Text + "'" + "," + "'" + Rtxt_Remark.Text + "'" + "," + "'" + currentdate + "'" + ")")

        save_Con = New OleDbConnection(save_Constr)
        save_Com = New OleDbCommand(save_Comstr, save_Con)

        save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
        Dim save_builder As New OleDbCommandBuilder(save_DA)

        save_DS = New DataSet()

        Try
            save_DA.Fill(save_DS, "hm")
            MsgBox("Patient discharged successfully.", MsgBoxStyle.Information, "Register Patient")
            'save_Con.Close()

            'call this function for deleting the patient from admit list after discharging
            delete_Record()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try 'code end for saving record into discharge table

        'btnRegisterPatient.Enabled = True
        'btnEditRegistration.Enabled = True
        'btnDeleteRegistration.Enabled = True
        'btnSaveRegistration.Enabled = False
        'btnCancelRegistration.Enabled = False
        'btnDischarge.Enabled = False

        If LV_PatiView.Items.Count = 0 Then
            btnRegisterPatient.Enabled = True
            btnEditRegistration.Enabled = False
            'btnDeleteRegistration.Enabled = False
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False
            btnDischarge.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False

            'code when listview is empty
            txtBedCharge.ReadOnly = True
            txtPatiname.ReadOnly = True
            txtAge.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtphone.ReadOnly = True
            DTP_date.Enabled = False
            txtDisease.ReadOnly = True
            CB_Bloodgp.Enabled = False
            cbDoctor.Enabled = False
            Rtxt_Remark.ReadOnly = True
        Else
            btnRegisterPatient.Enabled = True
            btnEditRegistration.Enabled = True
            'btnDeleteRegistration.Enabled = True
            btnSaveRegistration.Enabled = False
            btnCancelRegistration.Enabled = False
            btnDischarge.Enabled = True

            'code when listview is empty
            txtBedCharge.ReadOnly = False
            txtPatiname.ReadOnly = False
            txtAge.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtphone.ReadOnly = False
            DTP_date.Enabled = True
            txtDisease.ReadOnly = False
            CB_Bloodgp.Enabled = True
            cbDoctor.Enabled = True
            Rtxt_Remark.ReadOnly = False
        End If
    End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        Me.Cursor = Cursors.WaitCursor
        Dim objclsConvertToExcel As New clsConvertToExcel()
        objclsConvertToExcel.Load_Excel_Details("c:\hm.mdb", "admit")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim sfrm = New frmSearch(Me, "c:\hm.mdb", "[" + "admit" + "]", "Select * From [" + "admit" + " ]")
        sfrm.Show()
    End Sub



    'Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

    '    'Dim objfrmPrint = New frmPrint(Me, "c:\hm.mdb", "admit")
    '    'objfrmPrint.show()
    '    Dim objfrmAdmit_CurrentPrint As New frmAdmit_CurrentPrint()

    '    objfrmAdmit_CurrentPrint.getAdmit_Print_Values(txtRegno.Text, CB_BedCategory.Text, _
    '    CB_Bedno.Text, txtBedCharge.Text, txtPatiname.Text, txtAge.Text, RtxtAddress.Text, _
    '    txtphone.Text, DTP_date.Value, txtDisease.Text, CB_Bloodgp.Text, cbDoctor.Text, _
    '    Rtxt_Remark.Text)

    '    objfrmAdmit_CurrentPrint.Show()
    'End Sub

    Private Sub btnCustomprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustomprint.Click
        Dim objfrmcheckbox = New frmCheckBox(Me, "c:\hm.mdb", "admit")
        objfrmcheckbox.show()
    End Sub


#End Region

End Class
