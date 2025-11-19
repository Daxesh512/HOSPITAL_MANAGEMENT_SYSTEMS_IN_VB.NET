Imports System.Data.OleDb

Public Class frmDischarge
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtBedCharge As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
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
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents LV_Discharge As System.Windows.Forms.ListView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnDeleteRegistration As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCustomprint As System.Windows.Forms.Button
    Friend WithEvents Rtxt_Remark As System.Windows.Forms.TextBox
    Friend WithEvents RtxtAddress As System.Windows.Forms.TextBox
    Friend WithEvents CB_BedCategory As System.Windows.Forms.TextBox
    Friend WithEvents CB_Bedno As System.Windows.Forms.TextBox
    Friend WithEvents DTP_date As System.Windows.Forms.TextBox
    Friend WithEvents CB_Bloodgp As System.Windows.Forms.TextBox
    Friend WithEvents cbDoctor As System.Windows.Forms.TextBox
    Friend WithEvents DTP_DOD As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.DTP_DOD = New System.Windows.Forms.TextBox()
        Me.cbDoctor = New System.Windows.Forms.TextBox()
        Me.CB_Bloodgp = New System.Windows.Forms.TextBox()
        Me.DTP_date = New System.Windows.Forms.TextBox()
        Me.CB_Bedno = New System.Windows.Forms.TextBox()
        Me.CB_BedCategory = New System.Windows.Forms.TextBox()
        Me.RtxtAddress = New System.Windows.Forms.TextBox()
        Me.Rtxt_Remark = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnDeleteRegistration = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtBedCharge = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
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
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.LV_Discharge = New System.Windows.Forms.ListView()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.btnCustomprint = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.LightGray
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.DTP_DOD, Me.cbDoctor, Me.CB_Bloodgp, Me.DTP_date, Me.CB_Bedno, Me.CB_BedCategory, Me.RtxtAddress, Me.Rtxt_Remark, Me.GroupBox3, Me.Label14, Me.Label13, Me.txtBedCharge, Me.Label10, Me.Label9, Me.Label8, Me.Label7, Me.Label6, Me.Label5, Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.txtDisease, Me.txtphone, Me.txtAge, Me.txtPatiname, Me.txtRegno, Me.Label11, Me.Label12})
        Me.GroupBox1.Location = New System.Drawing.Point(16, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(512, 504)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Patient Detail"
        '
        'DTP_DOD
        '
        Me.DTP_DOD.Location = New System.Drawing.Point(160, 472)
        Me.DTP_DOD.Name = "DTP_DOD"
        Me.DTP_DOD.ReadOnly = True
        Me.DTP_DOD.TabIndex = 159
        Me.DTP_DOD.Text = ""
        '
        'cbDoctor
        '
        Me.cbDoctor.Location = New System.Drawing.Point(160, 392)
        Me.cbDoctor.Name = "cbDoctor"
        Me.cbDoctor.ReadOnly = True
        Me.cbDoctor.TabIndex = 158
        Me.cbDoctor.Text = ""
        '
        'CB_Bloodgp
        '
        Me.CB_Bloodgp.Location = New System.Drawing.Point(160, 360)
        Me.CB_Bloodgp.Name = "CB_Bloodgp"
        Me.CB_Bloodgp.ReadOnly = True
        Me.CB_Bloodgp.TabIndex = 157
        Me.CB_Bloodgp.Text = ""
        '
        'DTP_date
        '
        Me.DTP_date.Location = New System.Drawing.Point(160, 296)
        Me.DTP_date.Name = "DTP_date"
        Me.DTP_date.ReadOnly = True
        Me.DTP_date.TabIndex = 156
        Me.DTP_date.Text = ""
        '
        'CB_Bedno
        '
        Me.CB_Bedno.Location = New System.Drawing.Point(160, 88)
        Me.CB_Bedno.Name = "CB_Bedno"
        Me.CB_Bedno.ReadOnly = True
        Me.CB_Bedno.TabIndex = 155
        Me.CB_Bedno.Text = ""
        '
        'CB_BedCategory
        '
        Me.CB_BedCategory.Location = New System.Drawing.Point(160, 56)
        Me.CB_BedCategory.Name = "CB_BedCategory"
        Me.CB_BedCategory.ReadOnly = True
        Me.CB_BedCategory.Size = New System.Drawing.Size(104, 20)
        Me.CB_BedCategory.TabIndex = 154
        Me.CB_BedCategory.Text = ""
        '
        'RtxtAddress
        '
        Me.RtxtAddress.Location = New System.Drawing.Point(160, 216)
        Me.RtxtAddress.Multiline = True
        Me.RtxtAddress.Name = "RtxtAddress"
        Me.RtxtAddress.ReadOnly = True
        Me.RtxtAddress.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.RtxtAddress.Size = New System.Drawing.Size(104, 40)
        Me.RtxtAddress.TabIndex = 7
        Me.RtxtAddress.Text = ""
        '
        'Rtxt_Remark
        '
        Me.Rtxt_Remark.Location = New System.Drawing.Point(160, 424)
        Me.Rtxt_Remark.Multiline = True
        Me.Rtxt_Remark.Name = "Rtxt_Remark"
        Me.Rtxt_Remark.ReadOnly = True
        Me.Rtxt_Remark.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.Rtxt_Remark.Size = New System.Drawing.Size(100, 40)
        Me.Rtxt_Remark.TabIndex = 13
        Me.Rtxt_Remark.Text = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnDeleteRegistration})
        Me.GroupBox3.Location = New System.Drawing.Point(320, 32)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(176, 96)
        Me.GroupBox3.TabIndex = 153
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Transaction"
        '
        'btnDeleteRegistration
        '
        Me.btnDeleteRegistration.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDeleteRegistration.Location = New System.Drawing.Point(32, 40)
        Me.btnDeleteRegistration.Name = "btnDeleteRegistration"
        Me.btnDeleteRegistration.Size = New System.Drawing.Size(112, 32)
        Me.btnDeleteRegistration.TabIndex = 15
        Me.btnDeleteRegistration.Text = "Delete"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(40, 472)
        Me.Label14.Name = "Label14"
        Me.Label14.TabIndex = 151
        Me.Label14.Text = "Discharge Date"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(40, 120)
        Me.Label13.Name = "Label13"
        Me.Label13.TabIndex = 150
        Me.Label13.Text = "Bed Charge/Day"
        '
        'txtBedCharge
        '
        Me.txtBedCharge.Location = New System.Drawing.Point(160, 120)
        Me.txtBedCharge.Name = "txtBedCharge"
        Me.txtBedCharge.ReadOnly = True
        Me.txtBedCharge.Size = New System.Drawing.Size(104, 20)
        Me.txtBedCharge.TabIndex = 4
        Me.txtBedCharge.Text = ""
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(40, 432)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 144
        Me.Label10.Text = "Remarks"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(40, 360)
        Me.Label9.Name = "Label9"
        Me.Label9.TabIndex = 143
        Me.Label9.Text = "Blood Group"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(40, 392)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 140
        Me.Label8.Text = "Doctor"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(40, 328)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 139
        Me.Label7.Text = "Disease"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(40, 296)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 138
        Me.Label6.Text = "Date"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(40, 264)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 137
        Me.Label5.Text = "Phone No"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(40, 216)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 136
        Me.Label4.Text = "Address"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(40, 184)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 135
        Me.Label3.Text = "Age"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(40, 152)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 134
        Me.Label2.Text = "Patient Name"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(40, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 133
        Me.Label1.Text = "Registration No"
        '
        'txtDisease
        '
        Me.txtDisease.Location = New System.Drawing.Point(160, 328)
        Me.txtDisease.Name = "txtDisease"
        Me.txtDisease.ReadOnly = True
        Me.txtDisease.Size = New System.Drawing.Size(104, 20)
        Me.txtDisease.TabIndex = 10
        Me.txtDisease.Text = ""
        '
        'txtphone
        '
        Me.txtphone.Location = New System.Drawing.Point(160, 264)
        Me.txtphone.Name = "txtphone"
        Me.txtphone.ReadOnly = True
        Me.txtphone.Size = New System.Drawing.Size(104, 20)
        Me.txtphone.TabIndex = 8
        Me.txtphone.Text = ""
        '
        'txtAge
        '
        Me.txtAge.Location = New System.Drawing.Point(160, 184)
        Me.txtAge.Name = "txtAge"
        Me.txtAge.ReadOnly = True
        Me.txtAge.Size = New System.Drawing.Size(104, 20)
        Me.txtAge.TabIndex = 6
        Me.txtAge.Text = ""
        '
        'txtPatiname
        '
        Me.txtPatiname.Location = New System.Drawing.Point(160, 152)
        Me.txtPatiname.Name = "txtPatiname"
        Me.txtPatiname.ReadOnly = True
        Me.txtPatiname.Size = New System.Drawing.Size(104, 20)
        Me.txtPatiname.TabIndex = 5
        Me.txtPatiname.Text = ""
        '
        'txtRegno
        '
        Me.txtRegno.Location = New System.Drawing.Point(160, 24)
        Me.txtRegno.Name = "txtRegno"
        Me.txtRegno.ReadOnly = True
        Me.txtRegno.Size = New System.Drawing.Size(104, 20)
        Me.txtRegno.TabIndex = 1
        Me.txtRegno.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(40, 56)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 23)
        Me.Label11.TabIndex = 125
        Me.Label11.Text = "Bed Type"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(40, 88)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 24)
        Me.Label12.TabIndex = 147
        Me.Label12.Text = "Bed No"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.LV_Discharge})
        Me.GroupBox2.Location = New System.Drawing.Point(16, 520)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(560, 104)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Patient List View"
        '
        'LV_Discharge
        '
        Me.LV_Discharge.FullRowSelect = True
        Me.LV_Discharge.GridLines = True
        Me.LV_Discharge.Location = New System.Drawing.Point(24, 16)
        Me.LV_Discharge.Name = "LV_Discharge"
        Me.LV_Discharge.Size = New System.Drawing.Size(512, 80)
        Me.LV_Discharge.TabIndex = 0
        Me.LV_Discharge.View = System.Windows.Forms.View.Details
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSearch, Me.btnExcel})
        Me.GroupBox4.Location = New System.Drawing.Point(544, 232)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(136, 128)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Search/Export"
        '
        'btnSearch
        '
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSearch.Location = New System.Drawing.Point(32, 32)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 32)
        Me.btnSearch.TabIndex = 17
        Me.btnSearch.Text = "Search"
        '
        'btnExcel
        '
        Me.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExcel.Location = New System.Drawing.Point(32, 80)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(75, 32)
        Me.btnExcel.TabIndex = 18
        Me.btnExcel.Text = "Excel"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCustomprint})
        Me.GroupBox5.Location = New System.Drawing.Point(544, 80)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(136, 104)
        Me.GroupBox5.TabIndex = 9
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Print"
        '
        'btnCustomprint
        '
        Me.btnCustomprint.BackColor = System.Drawing.Color.LightGray
        Me.btnCustomprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCustomprint.Location = New System.Drawing.Point(32, 40)
        Me.btnCustomprint.Name = "btnCustomprint"
        Me.btnCustomprint.Size = New System.Drawing.Size(72, 32)
        Me.btnCustomprint.TabIndex = 16
        Me.btnCustomprint.Text = "List"
        '
        'frmDischarge
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(696, 632)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.GroupBox1, Me.GroupBox5, Me.GroupBox4})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmDischarge"
        Me.Text = "frmDischarge"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Function for binding textbox with record"
    'code for binding text box with 1st record of table
    Sub Fill_Textbox_Discharge()
        Dim r0, r1, r2, r3, r4, r5, r6, r7, r8, r9, r10, r11, r12, r13 As String

        Dim Fill_Textbox_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim Fill_Textbox_cmdstr As String = "select Regno,Name,Age,Address,Phone,Checkupdate,Checkuptime,Disease,Bloodgroup,Doctor,Remarks from clinicpatients where Regno=1"
        Dim Fill_Textbox_cmdstr As String = "select * from discharge"
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
            r13 = insert_DR(13)
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
        DTP_date.Text = r8
        txtDisease.Text = r9
        CB_Bloodgp.Text = r10
        cbDoctor.Text = r11
        Rtxt_Remark.Text = r12
        DTP_DOD.Text = r13
    End Sub
#End Region

#Region "Fuction for filling Listview"

    'fill listview
    Sub fill_listview()
        ' Clear list view column headers and items
        LV_Discharge.Columns.Clear()
        LV_Discharge.Items.Clear()

        'If LV.Items.Count > 0 Then
        'LV.Items.Clear()
        'End If

        ' Get SQL Query from textbox
        Dim fillLV_comstr As String = "select * from discharge"
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
                LV_Discharge.Columns.Add(header)
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
                LV_Discharge.Items.Add(NewItem)
            End While

            ' Close data reader
            NewReader.Close()

        Catch ex As OleDbException
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "oledb Error"
            LV_Discharge.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem(ex.Message)
            LV_Discharge.Items.Add(ErrorItem)

        Catch ex As Exception
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "Error"
            LV_Discharge.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem("An error has occurred")
            LV_Discharge.Items.Add(ErrorItem)

        Finally
            fillLV_Con.Close()
        End Try
    End Sub
#End Region

#Region "Transaction Button Event"
    Private Sub btnDeleteRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteRegistration.Click
        Dim delete_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim delete_Comstr As String
        Dim delete_Con As OleDbConnection
        Dim delete_Com As OleDbCommand

        Dim ID As String

        ID = txtRegno.Text
        delete_Comstr = "Delete from discharge WHERE Regno=" & ID

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

        Fill_Textbox_Discharge()
        fill_listview()

        If LV_Discharge.Items.Count = 0 Then
            btnDeleteRegistration.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False
        Else
            btnDeleteRegistration.Enabled = True
        End If

    End Sub
#End Region

    Private Sub frmDischarge_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Fill_Textbox_Discharge()
        fill_listview()
        If LV_Discharge.Items.Count = 0 Then
            'MessageBox.Show("There is no record to show.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnDeleteRegistration.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False

        End If
    End Sub

    'Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
    '    Dim objfrmDischarge_CurrentPrint As New frmDischarge_CurrentPrint()

    '    objfrmDischarge_CurrentPrint.getDischarge_Print_Values(txtRegno.Text, CB_BedCategory.Text, _
    '    CB_Bedno.Text, txtBedCharge.Text, txtPatiname.Text, txtAge.Text, RtxtAddress.Text, _
    '    txtphone.Text, DTP_date.Value, txtDisease.Text, CB_Bloodgp.Text, cbDoctor.Text, _
    '    Rtxt_Remark.Text, DTP_DOD.Value)

    '    objfrmDischarge_CurrentPrint.Show()
    'End Sub

    Private Sub btnCustomprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustomprint.Click
        Dim objfrmcheckbox = New frmCheckBox(Me, "c:\hm.mdb", "discharge")
        objfrmcheckbox.show()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim sfrm = New frmSearch(Me, "c:\hm.mdb", "[" + "discharge" + "]", "Select * From [" + "discharge" + " ]")
        sfrm.Show()
    End Sub

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        Me.Cursor = Cursors.WaitCursor
        Dim objclsConvertToExcel As New clsConvertToExcel()
        objclsConvertToExcel.Load_Excel_Details("c:\hm.mdb", "discharge")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub LV_Discharge_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LV_Discharge.SelectedIndexChanged
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
            DTP_date.Text = SelectedItems(0).SubItems(8).Text
            txtDisease.Text = SelectedItems(0).SubItems(9).Text
            CB_Bloodgp.Text = SelectedItems(0).SubItems(10).Text
            cbDoctor.Text = SelectedItems(0).SubItems(11).Text
            Rtxt_Remark.Text = SelectedItems(0).SubItems(12).Text
            DTP_DOD.Text = SelectedItems(0).SubItems(13).Text

        End If
    End Sub
End Class
