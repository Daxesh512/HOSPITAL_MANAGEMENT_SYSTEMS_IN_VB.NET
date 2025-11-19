Imports System.Data
Imports System.Data.OleDb

Public Class frmSisterWardboy
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
    Friend WithEvents LV_SWView As System.Windows.Forms.ListView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CB_Category As System.Windows.Forms.ComboBox
    Friend WithEvents btnCustomprint As System.Windows.Forms.Button
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSaveS_W As System.Windows.Forms.Button
    Friend WithEvents btnDeleteS_W As System.Windows.Forms.Button
    Friend WithEvents btnUpdateS_W As System.Windows.Forms.Button
    Friend WithEvents btnAddS_W As System.Windows.Forms.Button
    Friend WithEvents RtxtAddress As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GB_Doctor = New System.Windows.Forms.GroupBox()
        Me.RtxtAddress = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSaveS_W = New System.Windows.Forms.Button()
        Me.btnDeleteS_W = New System.Windows.Forms.Button()
        Me.btnUpdateS_W = New System.Windows.Forms.Button()
        Me.btnAddS_W = New System.Windows.Forms.Button()
        Me.CB_Category = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DTP_DOJ = New System.Windows.Forms.DateTimePicker()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.btnCustomprint = New System.Windows.Forms.Button()
        Me.GB_DoctorView = New System.Windows.Forms.GroupBox()
        Me.LV_SWView = New System.Windows.Forms.ListView()
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
        Me.GB_Doctor.Controls.AddRange(New System.Windows.Forms.Control() {Me.RtxtAddress, Me.GroupBox3, Me.CB_Category, Me.Label6, Me.DTP_DOJ, Me.txtPhone, Me.txtName, Me.txtID, Me.Label5, Me.Label4, Me.Label3, Me.Label2, Me.Label1})
        Me.GB_Doctor.Location = New System.Drawing.Point(16, 16)
        Me.GB_Doctor.Name = "GB_Doctor"
        Me.GB_Doctor.Size = New System.Drawing.Size(496, 328)
        Me.GB_Doctor.TabIndex = 0
        Me.GB_Doctor.TabStop = False
        Me.GB_Doctor.Text = "Sister/Wardboy Detail"
        '
        'RtxtAddress
        '
        Me.RtxtAddress.Location = New System.Drawing.Point(168, 120)
        Me.RtxtAddress.Multiline = True
        Me.RtxtAddress.Name = "RtxtAddress"
        Me.RtxtAddress.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.RtxtAddress.Size = New System.Drawing.Size(100, 64)
        Me.RtxtAddress.TabIndex = 2
        Me.RtxtAddress.Text = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCancel, Me.btnSaveS_W, Me.btnDeleteS_W, Me.btnUpdateS_W, Me.btnAddS_W})
        Me.GroupBox3.Location = New System.Drawing.Point(320, 48)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(144, 248)
        Me.GroupBox3.TabIndex = 33
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Transaction"
        '
        'btnCancel
        '
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancel.Location = New System.Drawing.Point(32, 192)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        '
        'btnSaveS_W
        '
        Me.btnSaveS_W.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSaveS_W.Location = New System.Drawing.Point(32, 152)
        Me.btnSaveS_W.Name = "btnSaveS_W"
        Me.btnSaveS_W.Size = New System.Drawing.Size(88, 23)
        Me.btnSaveS_W.TabIndex = 9
        Me.btnSaveS_W.Text = "Save S/W"
        '
        'btnDeleteS_W
        '
        Me.btnDeleteS_W.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDeleteS_W.Location = New System.Drawing.Point(32, 112)
        Me.btnDeleteS_W.Name = "btnDeleteS_W"
        Me.btnDeleteS_W.Size = New System.Drawing.Size(88, 23)
        Me.btnDeleteS_W.TabIndex = 8
        Me.btnDeleteS_W.Text = "Delete S/W"
        '
        'btnUpdateS_W
        '
        Me.btnUpdateS_W.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnUpdateS_W.Location = New System.Drawing.Point(32, 72)
        Me.btnUpdateS_W.Name = "btnUpdateS_W"
        Me.btnUpdateS_W.Size = New System.Drawing.Size(88, 23)
        Me.btnUpdateS_W.TabIndex = 7
        Me.btnUpdateS_W.Text = "Edit S/W"
        '
        'btnAddS_W
        '
        Me.btnAddS_W.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddS_W.Location = New System.Drawing.Point(32, 32)
        Me.btnAddS_W.Name = "btnAddS_W"
        Me.btnAddS_W.Size = New System.Drawing.Size(88, 23)
        Me.btnAddS_W.TabIndex = 6
        Me.btnAddS_W.Text = "Add S/W"
        '
        'CB_Category
        '
        Me.CB_Category.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_Category.Items.AddRange(New Object() {"Sister", "Wardboy"})
        Me.CB_Category.Location = New System.Drawing.Point(168, 200)
        Me.CB_Category.Name = "CB_Category"
        Me.CB_Category.Size = New System.Drawing.Size(104, 21)
        Me.CB_Category.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label6.Location = New System.Drawing.Point(48, 200)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Category"
        '
        'DTP_DOJ
        '
        Me.DTP_DOJ.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DTP_DOJ.Location = New System.Drawing.Point(168, 280)
        Me.DTP_DOJ.Name = "DTP_DOJ"
        Me.DTP_DOJ.Size = New System.Drawing.Size(104, 20)
        Me.DTP_DOJ.TabIndex = 5
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(168, 240)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.TabIndex = 4
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
        Me.Label5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label5.Location = New System.Drawing.Point(48, 280)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Date of joining"
        '
        'Label4
        '
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label4.Location = New System.Drawing.Point(48, 240)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Phone"
        '
        'Label3
        '
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label3.Location = New System.Drawing.Point(48, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Address"
        '
        'Label2
        '
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label2.Location = New System.Drawing.Point(48, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Name"
        '
        'Label1
        '
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Location = New System.Drawing.Point(48, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "ID"
        '
        'btnSearch
        '
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSearch.Location = New System.Drawing.Point(24, 32)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(72, 23)
        Me.btnSearch.TabIndex = 12
        Me.btnSearch.Text = "Search"
        '
        'btnExcel
        '
        Me.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExcel.Location = New System.Drawing.Point(24, 72)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(72, 23)
        Me.btnExcel.TabIndex = 13
        Me.btnExcel.Text = "Excel"
        '
        'btnCustomprint
        '
        Me.btnCustomprint.BackColor = System.Drawing.Color.LightGray
        Me.btnCustomprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCustomprint.Location = New System.Drawing.Point(24, 48)
        Me.btnCustomprint.Name = "btnCustomprint"
        Me.btnCustomprint.Size = New System.Drawing.Size(72, 24)
        Me.btnCustomprint.TabIndex = 11
        Me.btnCustomprint.Text = "List"
        '
        'GB_DoctorView
        '
        Me.GB_DoctorView.Controls.AddRange(New System.Windows.Forms.Control() {Me.LV_SWView})
        Me.GB_DoctorView.Location = New System.Drawing.Point(16, 376)
        Me.GB_DoctorView.Name = "GB_DoctorView"
        Me.GB_DoctorView.Size = New System.Drawing.Size(656, 248)
        Me.GB_DoctorView.TabIndex = 1
        Me.GB_DoctorView.TabStop = False
        Me.GB_DoctorView.Text = "Sister/Wardboy View"
        '
        'LV_SWView
        '
        Me.LV_SWView.FullRowSelect = True
        Me.LV_SWView.GridLines = True
        Me.LV_SWView.Location = New System.Drawing.Point(24, 24)
        Me.LV_SWView.Name = "LV_SWView"
        Me.LV_SWView.Size = New System.Drawing.Size(608, 200)
        Me.LV_SWView.TabIndex = 0
        Me.LV_SWView.View = System.Windows.Forms.View.Details
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCustomprint})
        Me.GroupBox1.Location = New System.Drawing.Point(544, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(120, 120)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSearch, Me.btnExcel})
        Me.GroupBox2.Location = New System.Drawing.Point(544, 224)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(120, 120)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Search/Export"
        '
        'frmSisterWardboy
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(696, 632)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.GroupBox1, Me.GB_DoctorView, Me.GB_Doctor})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmSisterWardboy"
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
    Sub Fill_Textbox_SisWard()
        Dim r0, r1, r2, r3, r4, r5 As String

        Dim Fill_Textbox_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim Fill_Textbox_cmdstr As String = "select Regno,Name,Age,Address,Phone,Checkupdate,Checkuptime,Disease,Bloodgroup,Doctor,Remarks from clinicpatients where Regno=1"
        Dim Fill_Textbox_cmdstr As String = "select * from sisward"
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

            Exit While
        End While
        Fill_Textbox_con.Close()

        txtID.Text = r0
        txtName.Text = r1
        RtxtAddress.Text = r2
        CB_Category.Text = r3
        txtPhone.Text = r4
        If r5 = Nothing Then
            DTP_DOJ.Text = ""
        Else
            DTP_DOJ.Value = r5
        End If

    End Sub
#End Region

#Region "Function for filling Listview"
    'fill listview
    Sub fill_listview()
        ' Clear list view column headers and items
        LV_SWView.Columns.Clear()
        LV_SWView.Items.Clear()

        'If LV.Items.Count > 0 Then
        'LV.Items.Clear()
        'End If

        ' Get SQL Query from textbox
        Dim fillLV_comstr As String = "select * from sisward"
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
                LV_SWView.Columns.Add(header)
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
                LV_SWView.Items.Add(NewItem)
            End While

            ' Close data reader
            NewReader.Close()

        Catch ex As OleDbException
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "oledb Error"
            LV_SWView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem(ex.Message)
            LV_SWView.Items.Add(ErrorItem)

        Catch ex As Exception
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "Error"
            LV_SWView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem("An error has occurred")
            LV_SWView.Items.Add(ErrorItem)

        Finally

            fillLV_Con.Close()

        End Try
    End Sub
#End Region

#Region "Listview event"
    Private Sub LV_PatiView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LV_SWView.SelectedIndexChanged
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
            CB_Category.Text = SelectedItems(0).SubItems(3).Text
            txtPhone.Text = SelectedItems(0).SubItems(4).Text
            DTP_DOJ.Value = SelectedItems(0).SubItems(5).Text
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

        save_Comstr = "insert into sisward(ID,Name,Address,Category,Phone,JoiningDate)values" & _
        ("(" + "'" + txtID.Text + "'" + "," + "'" + txtName.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + CB_Category.Text + "'" + "," + "'" + txtPhone.Text + "'" + "," + "'" + DTP_DOJ.Value + "'" + ")")

        save_Con = New OleDbConnection(save_Constr)
        save_Com = New OleDbCommand(save_Comstr, save_Con)

        save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
        Dim save_builder As New OleDbCommandBuilder(save_DA)

        save_DS = New DataSet()

        Try
            save_DA.Fill(save_DS, "hm")
            MsgBox("Record inserted successfully.", MsgBoxStyle.Information, "Register Patient")

            Dim obj As New frmMain()
            obj.setFlag = 0
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        save_Con.Close()

        'code for adding ID in siswardID table
        Dim save_siswardID_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_siswardID_Comstr As String
        Dim save_siswardID_Con As OleDbConnection
        Dim save_siswardID_Com As OleDbCommand
        Dim save_siswardID_DA As OleDbDataAdapter
        Dim save_siswardID_DS As DataSet
        Dim save_siswardID_DT As DataTable

        save_siswardID_Comstr = "insert into siswardID(swID)values" & _
        ("(" + "'" + txtID.Text + "'" + ")")

        save_siswardID_Con = New OleDbConnection(save_siswardID_Constr)
        save_siswardID_Com = New OleDbCommand(save_siswardID_Comstr, save_siswardID_Con)

        save_siswardID_DA = New OleDbDataAdapter(save_siswardID_Comstr, save_siswardID_Constr)
        Dim save_siswardID_builder As New OleDbCommandBuilder(save_siswardID_DA)

        save_siswardID_DS = New DataSet()

        Try
            save_siswardID_DA.Fill(save_siswardID_DS, "hm")
            'MsgBox("ID inserted in siswardID table successfully.", MsgBoxStyle.Information, "Register Patient")
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        save_siswardID_Con.Close() 'code ends for adding ID in siswardID table
    End Sub
#End Region

#Region "Function for saving blank record"

    Sub save_file_Solution()

        Dim save_Comstr As String

        If txtPhone.Text = "" Then
            txtPhone.Text = "0"

            save_Comstr = "insert into sisward(ID,Name,Address,Category,Phone,JoiningDate)values" & _
               ("(" + "'" + txtID.Text + "'" + "," + "'" + txtName.Text + "'" + "," + "'" + RtxtAddress.Text + "'" + "," + "'" + CB_Category.Text + "'" + "," + "'" + txtPhone.Text + "'" + "," + "'" + DTP_DOJ.Value + "'" + ")")
        End If


        Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
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
            MsgBox("Record inserted successfully.", MsgBoxStyle.Information, "Register Patient")

            Dim obj As New frmMain()
            obj.setFlag = 0
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        save_Con.Close()

        'code for adding ID in siswardID table
        Dim save_siswardID_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim save_siswardID_Comstr As String
        Dim save_siswardID_Con As OleDbConnection
        Dim save_siswardID_Com As OleDbCommand
        Dim save_siswardID_DA As OleDbDataAdapter
        Dim save_siswardID_DS As DataSet
        Dim save_siswardID_DT As DataTable

        save_siswardID_Comstr = "insert into siswardID(swID)values" & _
        ("(" + "'" + txtID.Text + "'" + ")")

        save_siswardID_Con = New OleDbConnection(save_siswardID_Constr)
        save_siswardID_Com = New OleDbCommand(save_siswardID_Comstr, save_siswardID_Con)

        save_siswardID_DA = New OleDbDataAdapter(save_siswardID_Comstr, save_siswardID_Constr)
        Dim save_siswardID_builder As New OleDbCommandBuilder(save_siswardID_DA)

        save_siswardID_DS = New DataSet()

        Try
            save_siswardID_DA.Fill(save_siswardID_DS, "hm")
            'MsgBox("ID inserted in siswardID table successfully.", MsgBoxStyle.Information, "Register Patient")
        Catch ex As Exception
            'MsgBox(ex.Message)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        save_siswardID_Con.Close() 'code ends for adding ID in siswardID table
    End Sub

#End Region

#Region "Function for checking blank values in textbox"
    'this function check for blank values in form
    Sub Check_Textbox()
        Dim r As DialogResult
        If txtName.Text = "" _
           Or RtxtAddress.Text = "" _
           Or CB_Category.Text = "" Then
            'Or txtPhone.Text = "" Then
            MessageBox.Show("You must fill all the fields.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        ElseIf (txtName.Text <> "" _
           Or RtxtAddress.Text <> "" _
           Or CB_Category.Text <> "") _
            And txtPhone.Text = "" Then
            r = MessageBox.Show("Some fields are empty. Do you want Save it?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        Else
            save_file()
            fill_listview()

            If LV_SWView.Items.Count = 0 Then
                btnAddS_W.Enabled = True
                btnUpdateS_W.Enabled = False
                btnDeleteS_W.Enabled = False
                btnSaveS_W.Enabled = False
                btnCancel.Enabled = False
                'btnPrint.Enabled = False
                btnCustomprint.Enabled = False
                btnSearch.Enabled = False
                btnExcel.Enabled = False
            Else
                btnAddS_W.Enabled = True
                btnUpdateS_W.Enabled = True
                btnDeleteS_W.Enabled = True
                btnSaveS_W.Enabled = False
                btnCancel.Enabled = False
                'btnPrint.Enabled = True
                btnCustomprint.Enabled = True
                btnSearch.Enabled = True
                btnExcel.Enabled = True
            End If
        End If

        If r = DialogResult.Yes Then
            save_file_Solution()
            fill_listview()

            If LV_SWView.Items.Count = 0 Then
                btnAddS_W.Enabled = True
                btnUpdateS_W.Enabled = False
                btnDeleteS_W.Enabled = False
                btnSaveS_W.Enabled = False
                btnCancel.Enabled = False
                'btnPrint.Enabled = False
                btnCustomprint.Enabled = False
                btnSearch.Enabled = False
                btnExcel.Enabled = False
            Else
                btnAddS_W.Enabled = True
                btnUpdateS_W.Enabled = True
                btnDeleteS_W.Enabled = True
                btnSaveS_W.Enabled = False
                btnCancel.Enabled = False
                ' btnPrint.Enabled = True
                btnCustomprint.Enabled = True
                btnSearch.Enabled = True
                btnExcel.Enabled = True
            End If
        End If
    End Sub

#End Region

#Region "All Keypress event"

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

#Region "Load event of frmSisterWardboy "
    Private Sub frmDoctor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Fill_Textbox_SisWard()
        fill_listview()

        'MessageBox.Show("There is no record to show.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        If LV_SWView.Items.Count = 0 Then
            btnAddS_W.Enabled = True
            btnUpdateS_W.Enabled = False
            btnDeleteS_W.Enabled = False
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtPhone.ReadOnly = True
            CB_Category.Enabled = False
            DTP_DOJ.Enabled = False
        Else
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtPhone.ReadOnly = False
            CB_Category.Enabled = True
            DTP_DOJ.Enabled = True
        End If
    End Sub
#End Region

#Region "ALL TRANSACTION BUTTON EVENT"

    Private Sub btnAddS_W_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddS_W.Click
        Dim r0, r1, r2, r3 As String

        Dim insert_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim cmdstr As String = "select ID,name,address,bloodgroup,phone,DOA from inpatients where ID=2"
        Dim insert_cmdstr As String = "select * from siswardID"
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
        CB_Category.Text = ""
        txtPhone.Text = ""
        DTP_DOJ.Value = System.DateTime.Today

        txtID.Text = r0 + 1
        'save_flag = False
        'Dim o As New frmMain()
        'o.increase_counter_routine()

        Dim obj As New frmMain()
        obj.setFlag = 1

        btnAddS_W.Enabled = False
        btnUpdateS_W.Enabled = False
        btnDeleteS_W.Enabled = False
        btnSaveS_W.Enabled = True
        btnCancel.Enabled = True

        'code when listview is empty
        txtName.ReadOnly = False
        RtxtAddress.ReadOnly = False
        txtPhone.ReadOnly = False
        CB_Category.Enabled = True
        DTP_DOJ.Enabled = True
    End Sub

    Private Sub btnUpdateS_W_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdateS_W.Click
        Dim update_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim update_Comstr As String
        Dim update_Con As OleDbConnection
        Dim update_Com As OleDbCommand

        Dim ID As String

        ID = txtID.Text

        update_Comstr = "UPDATE sisward SET Name=" & "'" & txtName.Text & "'" & "," & _
        "Address=" & "'" & RtxtAddress.Text & "'" & "," & _
        "Category=" & "'" & CB_Category.Text & "'" & "," & _
        "Phone=" & "'" & txtPhone.Text & "'" & "," & _
        "JoiningDate=" & "'" & DTP_DOJ.Value & "'" & " WHERE ID=" & ID & " ; "

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

        If LV_SWView.Items.Count = 0 Then
            btnAddS_W.Enabled = True
            btnUpdateS_W.Enabled = False
            btnDeleteS_W.Enabled = False
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False
        Else
            btnAddS_W.Enabled = True
            btnUpdateS_W.Enabled = True
            btnDeleteS_W.Enabled = True
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False
        End If
    End Sub

    Private Sub btnDeleteS_W_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteS_W.Click
        Dim delete_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim delete_Comstr As String
        Dim delete_Con As OleDbConnection
        Dim delete_Com As OleDbCommand

        Dim ID As String

        ID = txtID.Text
        delete_Comstr = "Delete from sisward WHERE ID=" & ID

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

        Fill_Textbox_SisWard()
        fill_listview()

        If LV_SWView.Items.Count = 0 Then
            btnAddS_W.Enabled = True
            btnUpdateS_W.Enabled = False
            btnDeleteS_W.Enabled = False
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtPhone.ReadOnly = True
            CB_Category.Enabled = False
            DTP_DOJ.Enabled = False
        Else
            btnAddS_W.Enabled = True
            btnUpdateS_W.Enabled = True
            btnDeleteS_W.Enabled = True
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtPhone.ReadOnly = False
            CB_Category.Enabled = True
            DTP_DOJ.Enabled = True
        End If
    End Sub

    Private Sub btnSaveS_W_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveS_W.Click
        Check_Textbox()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        txtID.Text = ""
        txtName.Text = ""
        RtxtAddress.Text = ""
        CB_Category.Text = ""
        txtPhone.Text = ""
        'DTP_date.Value = 

        Fill_Textbox_SisWard()

        If LV_SWView.Items.Count = 0 Then
            btnAddS_W.Enabled = True
            btnUpdateS_W.Enabled = False
            btnDeleteS_W.Enabled = False
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = True
            RtxtAddress.ReadOnly = True
            txtPhone.ReadOnly = True
            CB_Category.Enabled = False
            DTP_DOJ.Enabled = False
        Else
            btnAddS_W.Enabled = True
            btnUpdateS_W.Enabled = True
            btnDeleteS_W.Enabled = True
            btnSaveS_W.Enabled = False
            btnCancel.Enabled = False

            'code when listview is empty
            txtName.ReadOnly = False
            RtxtAddress.ReadOnly = False
            txtPhone.ReadOnly = False
            CB_Category.Enabled = True
            DTP_DOJ.Enabled = True
        End If

        Dim obj As New frmMain()
        obj.setFlag = 0
    End Sub


    Private Sub btnCustomprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustomprint.Click
        Dim objfrmcheckbox = New frmCheckBox(Me, "c:\hm.mdb", "sisward")
        objfrmcheckbox.show()
    End Sub

    'Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
    '    'Dim objfrmPrint = New frmPrint(Me, "c:\hm.mdb", "sisward")
    '    'objfrmPrint.show()
    '    Dim objfrmSisward_CurrentPrint As New frmSisward_CurrentPrint()

    '    objfrmSisward_CurrentPrint.getSisward_Print_Values(txtID.Text, txtName.Text, _
    '    RtxtAddress.Text, CB_Category.Text, txtPhone.Text, DTP_DOJ.Value)

    '    objfrmSisward_CurrentPrint.Show()
    'End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        Me.Cursor = Cursors.WaitCursor
        Dim objclsConvertToExcel As New clsConvertToExcel()
        objclsConvertToExcel.Load_Excel_Details("c:\hm.mdb", "sisward")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim sfrm = New frmSearch(Me, "c:\hm.mdb", "[" + "sisward" + "]", "Select * From [" + "sisward" + " ]")
        sfrm.Show()
    End Sub
#End Region

End Class
