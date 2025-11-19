Imports System.Data
Imports System.Data.OleDb

Public Class frmDoctor
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents DTP_DOJ As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GB_Doctor As System.Windows.Forms.GroupBox
    Friend WithEvents GB_DoctorView As System.Windows.Forms.GroupBox
    Friend WithEvents LV_DoctorView As System.Windows.Forms.ListView
    Friend WithEvents btnCustomprint As System.Windows.Forms.Button
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSaveDoctor As System.Windows.Forms.Button
    Friend WithEvents btnDeleteDoctor As System.Windows.Forms.Button
    Friend WithEvents btnUpdateDoctor As System.Windows.Forms.Button
    Friend WithEvents btnAddDoctor As System.Windows.Forms.Button
    Friend WithEvents RtxtAddress As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GB_Doctor = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSaveDoctor = New System.Windows.Forms.Button()
        Me.btnDeleteDoctor = New System.Windows.Forms.Button()
        Me.btnUpdateDoctor = New System.Windows.Forms.Button()
        Me.btnAddDoctor = New System.Windows.Forms.Button()
        Me.DTP_DOJ = New System.Windows.Forms.DateTimePicker()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RtxtAddress = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.btnCustomprint = New System.Windows.Forms.Button()
        Me.GB_DoctorView = New System.Windows.Forms.GroupBox()
        Me.LV_DoctorView = New System.Windows.Forms.ListView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GB_Doctor.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GB_DoctorView.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GB_Doctor
        '
        Me.GB_Doctor.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox3, Me.DTP_DOJ, Me.txtPhone, Me.txtName, Me.txtID, Me.Label5, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.RtxtAddress})
        Me.GB_Doctor.Location = New System.Drawing.Point(16, 16)
        Me.GB_Doctor.Name = "GB_Doctor"
        Me.GB_Doctor.Size = New System.Drawing.Size(504, 296)
        Me.GB_Doctor.TabIndex = 0
        Me.GB_Doctor.TabStop = False
        Me.GB_Doctor.Text = "Doctor Detail"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCancel, Me.btnSaveDoctor, Me.btnDeleteDoctor, Me.btnUpdateDoctor, Me.btnAddDoctor})
        Me.GroupBox3.Location = New System.Drawing.Point(328, 48)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(136, 208)
        Me.GroupBox3.TabIndex = 31
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Transaction"
        '
        'btnCancel
        '
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancel.Location = New System.Drawing.Point(24, 160)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 23)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel"
        '
        'btnSaveDoctor
        '
        Me.btnSaveDoctor.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSaveDoctor.Location = New System.Drawing.Point(24, 128)
        Me.btnSaveDoctor.Name = "btnSaveDoctor"
        Me.btnSaveDoctor.Size = New System.Drawing.Size(88, 23)
        Me.btnSaveDoctor.TabIndex = 8
        Me.btnSaveDoctor.Text = "Save Doctor"
        '
        'btnDeleteDoctor
        '
        Me.btnDeleteDoctor.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDeleteDoctor.Location = New System.Drawing.Point(24, 96)
        Me.btnDeleteDoctor.Name = "btnDeleteDoctor"
        Me.btnDeleteDoctor.Size = New System.Drawing.Size(88, 23)
        Me.btnDeleteDoctor.TabIndex = 7
        Me.btnDeleteDoctor.Text = "Delete Doctor"
        '
        'btnUpdateDoctor
        '
        Me.btnUpdateDoctor.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnUpdateDoctor.Location = New System.Drawing.Point(24, 64)
        Me.btnUpdateDoctor.Name = "btnUpdateDoctor"
        Me.btnUpdateDoctor.Size = New System.Drawing.Size(88, 23)
        Me.btnUpdateDoctor.TabIndex = 6
        Me.btnUpdateDoctor.Text = "Edit Doctor"
        '
        'btnAddDoctor
        '
        Me.btnAddDoctor.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddDoctor.Location = New System.Drawing.Point(24, 29)
        Me.btnAddDoctor.Name = "btnAddDoctor"
        Me.btnAddDoctor.Size = New System.Drawing.Size(88, 23)
        Me.btnAddDoctor.TabIndex = 5
        Me.btnAddDoctor.Text = "Add Doctor"
        '
        'DTP_DOJ
        '
        Me.DTP_DOJ.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DTP_DOJ.Location = New System.Drawing.Point(168, 240)
        Me.DTP_DOJ.Name = "DTP_DOJ"
        Me.DTP_DOJ.Size = New System.Drawing.Size(104, 20)
        Me.DTP_DOJ.TabIndex = 4
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(168, 200)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.TabIndex = 3
        Me.txtPhone.Text = ""
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(168, 88)
        Me.txtName.Name = "txtName"
        Me.txtName.TabIndex = 1
        Me.txtName.Text = ""
        '
        'txtID
        '
        Me.txtID.Enabled = False
        Me.txtID.Location = New System.Drawing.Point(168, 48)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.TabIndex = 20
        Me.txtID.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(48, 240)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Date of joining"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(48, 200)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Phone"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(48, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Address"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(48, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Name"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "ID"
        '
        'RtxtAddress
        '
        Me.RtxtAddress.Location = New System.Drawing.Point(168, 120)
        Me.RtxtAddress.Multiline = True
        Me.RtxtAddress.Name = "RtxtAddress"
        Me.RtxtAddress.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.RtxtAddress.Size = New System.Drawing.Size(100, 56)
        Me.RtxtAddress.TabIndex = 2
        Me.RtxtAddress.Text = ""
        '
        'btnSearch
        '
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSearch.Location = New System.Drawing.Point(24, 32)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.TabIndex = 11
        Me.btnSearch.Text = "Search"
        '
        'btnExcel
        '
        Me.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExcel.Location = New System.Drawing.Point(24, 72)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.TabIndex = 12
        Me.btnExcel.Text = "Excel"
        '
        'btnCustomprint
        '
        Me.btnCustomprint.BackColor = System.Drawing.Color.LightGray
        Me.btnCustomprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCustomprint.Location = New System.Drawing.Point(24, 48)
        Me.btnCustomprint.Name = "btnCustomprint"
        Me.btnCustomprint.Size = New System.Drawing.Size(72, 24)
        Me.btnCustomprint.TabIndex = 10
        Me.btnCustomprint.Text = "List"
        '
        'GB_DoctorView
        '
        Me.GB_DoctorView.Controls.AddRange(New System.Windows.Forms.Control() {Me.LV_DoctorView})
        Me.GB_DoctorView.Location = New System.Drawing.Point(24, 336)
        Me.GB_DoctorView.Name = "GB_DoctorView"
        Me.GB_DoctorView.Size = New System.Drawing.Size(648, 280)
        Me.GB_DoctorView.TabIndex = 1
        Me.GB_DoctorView.TabStop = False
        Me.GB_DoctorView.Text = "Doctor View"
        '
        'LV_DoctorView
        '
        Me.LV_DoctorView.FullRowSelect = True
        Me.LV_DoctorView.GridLines = True
        Me.LV_DoctorView.Location = New System.Drawing.Point(24, 32)
        Me.LV_DoctorView.Name = "LV_DoctorView"
        Me.LV_DoctorView.Size = New System.Drawing.Size(600, 224)
        Me.LV_DoctorView.TabIndex = 0
        Me.LV_DoctorView.View = System.Windows.Forms.View.Details
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCustomprint})
        Me.GroupBox1.Location = New System.Drawing.Point(544, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(120, 112)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSearch, Me.btnExcel})
        Me.GroupBox2.Location = New System.Drawing.Point(544, 200)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(120, 112)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Seaech/Export"
        '
        'frmDoctor
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(696, 632)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.GroupBox1, Me.GB_DoctorView, Me.GB_Doctor})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmDoctor"
        Me.Text = "frmDoctor"
        Me.GB_Doctor.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GB_DoctorView.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Function for binding textbox with record"
    'code for binding text box with 1st record of table
    Sub Fill_Textbox_Doctors()
        Dim r0, r1, r2, r3, r4 As String

        Dim Fill_Textbox_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim Fill_Textbox_cmdstr As String = "select Regno,Name,Age,Address,Phone,Checkupdate,Checkuptime,Disease,Bloodgroup,Doctor,Remarks from clinicpatients where Regno=1"
        Dim Fill_Textbox_cmdstr As String = "select * from Doctors"
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

            Exit While
        End While
        Fill_Textbox_con.Close()

        txtID.Text = r0
        txtName.Text = r1
        RtxtAddress.Text = r2
        txtPhone.Text = r3
        If r4 = Nothing Then
            DTP_DOJ.Text = ""
        Else
            DTP_DOJ.Value = r4
        End If

    End Sub
#End Region

#Region "Function for filling Listview"
    'fill listview
    Sub fill_listview()
        ' Clear list view column headers and items
        LV_DoctorView.Columns.Clear()
        LV_DoctorView.Items.Clear()

        'If LV.Items.Count > 0 Then
        'LV.Items.Clear()
        'End If

        ' Get SQL Query from textbox
        Dim fillLV_comstr As String = "select * from Doctors"
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
                LV_DoctorView.Columns.Add(header)
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
                LV_DoctorView.Items.Add(NewItem)
            End While

            ' Close data reader
            NewReader.Close()

        Catch ex As OleDbException
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "oledb Error"
            LV_DoctorView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem(ex.Message)
            LV_DoctorView.Items.Add(ErrorItem)

        Catch ex As Exception
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "Error"
            LV_DoctorView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem("An error has occurred")
            LV_DoctorView.Items.Add(ErrorItem)

        Finally

            fillLV_Con.Close()

        End Try
    End Sub
#End Region

#Region "Listview event"
    Private Sub LV_PatiView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LV_DoctorView.SelectedIndexChanged
        Dim SelectedItems As ListView.SelectedListViewItemCollection = _
      CType(sender, ListView).SelectedItems
        If (SelectedItems.Count > 0) Then
            'MessageBox.Show(SelectedItems(0).SubItems(0).Text & _
            '    " " & SelectedItems(0).SubItems(1).Text & _
            '    " " & SelectedItems(0).SubItems(2).Text & _
            '    " " & SelectedItems(0).SubItems(3).Text)

            txtID.Text = SelectedItems(0).SubItems(0).Text
            txtName.Text = SelectedItems(0).SubItems(1).Text
            RtxtAddress.Text = SelectedItems(0).SubItems(2).Text
            txtPhone.Text = SelectedItems(0).SubItems(3).Text
            DTP_DOJ.Value = SelectedItems(0).SubItems(4).Text
        End If
    End Sub
#End Region

#Region "Function for saving record"
    Sub save_file()
        Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_Comstr As String
        Dim save_Con As OleDbConnection
        Dim save_Com As OleDbCommand
        Dim save_DA As OleDbDataAdapter
        Dim save_DS As DataSet
        Dim save_DT As DataTable
        'Dim txt As String
        'txt = CType(RtxtAddress.Text, String)
        If txtName.Text = "" Then
            txtName.Text = Nothing
        End If

        If RtxtAddress.Text = "" Then
            RtxtAddress.Text = Nothing
        End If
        If txtPhone.Text = "" Then
            txtPhone.Text = Nothing
        End If

        'txtRegno.Text = ""
        'DTP_date.Value = 
        'DTP_time.Value = ""

        save_Comstr = "insert into Doctors(ID,Name,Address,Phone,JoiningDate)values" & _
        ("(" + "'" + txtID.Text + "'" + "," + "'" + txtName.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + txtPhone.Text + "'" + "," + "'" + DTP_DOJ.Value + "'" + ")")

        save_Con = New OleDbConnection(save_Constr)
        save_Com = New OleDbCommand(save_Comstr, save_Con)

        save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
        Dim save_builder As New OleDbCommandBuilder(save_DA)

        save_DS = New DataSet()

        Try
            save_DA.Fill(save_DS, "hm")
            MsgBox("Doctor's record inserted successfully.", MsgBoxStyle.Information, "Register Patient")
            Dim obj As New frmMain()
            obj.setFlag = 0
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        save_Con.Close()

        'code for adding ID in doctorID table
        Dim save_doctorID_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_doctorID_Comstr As String
        Dim save_doctorID_Con As OleDbConnection
        Dim save_doctorID_Com As OleDbCommand
        Dim save_doctorID_DA As OleDbDataAdapter
        Dim save_doctorID_DS As DataSet
        Dim save_doctorID_DT As DataTable

        save_doctorID_Comstr = "insert into doctorID(dID)values" & _
        ("(" + "'" + txtID.Text + "'" + ")")

        save_doctorID_Con = New OleDbConnection(save_doctorID_Constr)
        save_doctorID_Com = New OleDbCommand(save_doctorID_Comstr, save_doctorID_Con)

        save_doctorID_DA = New OleDbDataAdapter(save_doctorID_Comstr, save_doctorID_Constr)
        Dim save_doctorID_builder As New OleDbCommandBuilder(save_doctorID_DA)

        save_doctorID_DS = New DataSet()

        Try
            save_doctorID_DA.Fill(save_doctorID_DS, "hm")
            'MsgBox("ID inserted in admitID table successfully.", MsgBoxStyle.Information, "Register Patient")
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        save_doctorID_Con.Close() 'code ends for adding ID in doctorID table
    End Sub
#End Region

#Region "Function for checking blank values in textbox"
    'this function check for blank values in form
    Sub Check_Textbox()
        Dim r As DialogResult
        If txtName.Text = "" _
           Or RtxtAddress.Text = "" _
           Or txtPhone.Text = "" Then
            r = MessageBox.Show("You must fill al the fields.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            save_file()
            fill_listview()

            If LV_DoctorView.Items.Count = 0 Then
                btnAddDoctor.Enabled = True
                btnUpdateDoctor.Enabled = False
                btnDeleteDoctor.Enabled = False
                btnSaveDoctor.Enabled = False
                btnCancel.Enabled = False
                'btnPrint.Enabled = False
                btnCustomprint.Enabled = False
                btnSearch.Enabled = False
                btnExcel.Enabled = False
            Else
                btnAddDoctor.Enabled = True
                btnUpdateDoctor.Enabled = True
                btnDeleteDoctor.Enabled = True
                btnSaveDoctor.Enabled = False
                btnCancel.Enabled = False
                'btnPrint.Enabled = True
                btnCustomprint.Enabled = True
                btnSearch.Enabled = True
                btnExcel.Enabled = True
            End If
        End If

        'If r = DialogResult.Yes Then
        '    save_file()
        '    fill_listview()
        'End If
    End Sub
#End Region


#Region "All keypress events"

    Private Sub txtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress
        If Char.IsLetter(e.KeyChar) = False Then
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

    Private Sub txtphone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPhone.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub
#End Region


#Region "Load event of frmDoctor"
    Private Sub frmDoctor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Fill_Textbox_Doctors()
        fill_listview()

        If LV_DoctorView.Items.Count = 0 Then
            'MessageBox.Show("There is no record to show.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnAddDoctor.Enabled = True
            btnUpdateDoctor.Enabled = False
            btnDeleteDoctor.Enabled = False
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtPhone.ReadOnly = True
            DTP_DOJ.Enabled = False
        Else
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtPhone.ReadOnly = False
            DTP_DOJ.Enabled = True

        End If
    End Sub
#End Region


#Region "ALL TRANSACTION BUTTON EVENT"

    Private Sub btnAddDoctor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddDoctor.Click

        Dim r0, r1, r2, r3 As String

        Dim insert_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
        'Dim insert_cmdstr As String = "select * from Doctors"
        Dim insert_cmdstr As String = "select * from doctorID"
        Dim insert_con As New OleDbConnection(insert_constr)
        Dim insert_com As New OleDbCommand(insert_cmdstr, insert_con)
        insert_con.Open()
        Dim insert_DR As OleDbDataReader = insert_com.ExecuteReader
        While insert_DR.Read
            r0 = insert_DR(0)
        End While
        insert_con.Close()

        txtName.Text = ""
        RtxtAddress.Text = ""
        txtPhone.Text = ""
        DTP_DOJ.Value = System.DateTime.Today

        txtID.Text = r0 + 1

        Dim obj As New frmMain()
        obj.setFlag = 1

        btnAddDoctor.Enabled = False
        btnUpdateDoctor.Enabled = False
        btnDeleteDoctor.Enabled = False
        btnSaveDoctor.Enabled = True
        btnCancel.Enabled = True

        'code when listview is empty
        txtName.ReadOnly = False
        RtxtAddress.ReadOnly = False
        txtPhone.ReadOnly = False
        DTP_DOJ.Enabled = True
    End Sub

    Private Sub btnUpdateDoctor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateDoctor.Click
        Dim update_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim update_Comstr As String
        Dim update_Con As OleDbConnection
        Dim update_Com As OleDbCommand

        Dim ID As String

        ID = txtID.Text

        update_Comstr = "UPDATE Doctors SET Name=" & "'" & txtName.Text & "'" & "," & _
        "Address=" & "'" & RtxtAddress.Text & "'" & "," & _
        "Phone=" & "'" & txtPhone.Text & "'" & "," & _
        "JoiningDate=" & "'" & DTP_DOJ.Value & "'" & " WHERE ID=" & ID & " ; "

        update_Con = New OleDbConnection(update_Constr)
        update_Com = New OleDbCommand(update_Comstr, update_Con)

        update_Con.Open()
        Try
            update_Com.ExecuteNonQuery()
            MsgBox("Doctor's record updated Successfully.", MsgBoxStyle.Information, "Edit Registration")
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
        update_Con.Close()
        fill_listview()

        If LV_DoctorView.Items.Count = 0 Then
            btnAddDoctor.Enabled = True
            btnUpdateDoctor.Enabled = False
            btnDeleteDoctor.Enabled = False
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False
        Else
            btnAddDoctor.Enabled = True
            btnUpdateDoctor.Enabled = True
            btnDeleteDoctor.Enabled = True
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False
        End If
    End Sub

    Private Sub btnDeleteDoctor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteDoctor.Click
        Dim delete_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim delete_Comstr As String
        Dim delete_Con As OleDbConnection
        Dim delete_Com As OleDbCommand

        Dim ID As String

        ID = txtID.Text
        delete_Comstr = "Delete from Doctors WHERE ID=" & ID

        delete_Con = New OleDbConnection(delete_Constr)
        delete_Com = New OleDbCommand(delete_Comstr, delete_Con)

        delete_Con.Open()
        Try
            delete_Com.ExecuteNonQuery()
            MsgBox("Doctor's record deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
        delete_Con.Close()

        Fill_Textbox_Doctors()
        fill_listview()

        If LV_DoctorView.Items.Count = 0 Then
            btnAddDoctor.Enabled = True
            btnUpdateDoctor.Enabled = False
            btnDeleteDoctor.Enabled = False
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtPhone.ReadOnly = True
            DTP_DOJ.Enabled = False
        Else
            btnAddDoctor.Enabled = True
            btnUpdateDoctor.Enabled = True
            btnDeleteDoctor.Enabled = True
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtPhone.ReadOnly = False
            DTP_DOJ.Enabled = True
        End If
    End Sub

    Private Sub btnSaveDoctor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveDoctor.Click
        Check_Textbox()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        txtID.Text = ""
        txtName.Text = ""
        RtxtAddress.Text = ""
        txtPhone.Text = ""
        'DTP_date.Value = 

        Fill_Textbox_Doctors()

        If LV_DoctorView.Items.Count = 0 Then
            btnAddDoctor.Enabled = True
            btnUpdateDoctor.Enabled = False
            btnDeleteDoctor.Enabled = False
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtPhone.ReadOnly = True
            DTP_DOJ.Enabled = False
        Else
            btnAddDoctor.Enabled = True
            btnUpdateDoctor.Enabled = True
            btnDeleteDoctor.Enabled = True
            btnSaveDoctor.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtPhone.ReadOnly = False
            DTP_DOJ.Enabled = True
        End If

        Dim obj As New frmMain()
        obj.setFlag = 0
    End Sub

    Private Sub btnCustomprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustomprint.Click
        Dim objfrmcheckbox = New frmCheckBox(Me, "c:\hm.mdb", "doctors")
        objfrmcheckbox.show()
    End Sub

    'Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
    '    'Dim objfrmPrint = New frmPrint(Me, "c:\hm.mdb", "doctors")
    '    'objfrmPrint.show()
    '    Dim objfrmDoctor_CurrentPrint As New frmDoctor_CurrentPrint()
    '    objfrmDoctor_CurrentPrint.getDoctor_Print_Values(txtID.Text, txtName.Text, _
    '    RtxtAddress.Text, txtPhone.Text, DTP_DOJ.Value)

    '    objfrmDoctor_CurrentPrint.Show()
    'End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        Me.Cursor = Cursors.WaitCursor
        Dim objclsConvertToExcel As New clsConvertToExcel()
        objclsConvertToExcel.Load_Excel_Details("c:\hm.mdb", "doctors")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim sfrm = New frmSearch(Me, "c:\hm.mdb", "[" + "doctors" + "]", "Select * From [" + "doctors" + " ]")
        sfrm.Show()
    End Sub

#End Region

End Class
