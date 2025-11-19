Imports System.Data.OleDb

Public Class frmBill
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
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtBedCharge As System.Windows.Forms.TextBox
    Friend WithEvents btnEditRegistration As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPatiname As System.Windows.Forms.TextBox
    Friend WithEvents txtRegno As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents GrpBox_Patibill As System.Windows.Forms.GroupBox
    Friend WithEvents GB_PatientBillView As System.Windows.Forms.GroupBox
    Friend WithEvents LV_PatiBillView As System.Windows.Forms.ListView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtTotBedcharge As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtOthercharge As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalbill As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtBedCategory As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCustomprint As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents btnExcel As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents DTP_date As System.Windows.Forms.TextBox
    Friend WithEvents DTP_DOD As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GrpBox_Patibill = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnEditRegistration = New System.Windows.Forms.Button()
        Me.txtBedCategory = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtTotalbill = New System.Windows.Forms.TextBox()
        Me.txtOthercharge = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtTotBedcharge = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtBedCharge = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtPatiname = New System.Windows.Forms.TextBox()
        Me.txtRegno = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.DTP_date = New System.Windows.Forms.TextBox()
        Me.DTP_DOD = New System.Windows.Forms.TextBox()
        Me.GB_PatientBillView = New System.Windows.Forms.GroupBox()
        Me.LV_PatiBillView = New System.Windows.Forms.ListView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnCustomprint = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.btnExcel = New System.Windows.Forms.Button()
        Me.GrpBox_Patibill.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GB_PatientBillView.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GrpBox_Patibill
        '
        Me.GrpBox_Patibill.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox3, Me.txtBedCategory, Me.Label7, Me.txtTotalbill, Me.txtOthercharge, Me.Label5, Me.txtTotBedcharge, Me.Label4, Me.Label3, Me.Label13, Me.txtBedCharge, Me.Label6, Me.Label2, Me.Label1, Me.txtPatiname, Me.txtRegno, Me.Label11, Me.DTP_date, Me.DTP_DOD})
        Me.GrpBox_Patibill.Location = New System.Drawing.Point(32, 8)
        Me.GrpBox_Patibill.Name = "GrpBox_Patibill"
        Me.GrpBox_Patibill.Size = New System.Drawing.Size(456, 368)
        Me.GrpBox_Patibill.TabIndex = 6
        Me.GrpBox_Patibill.TabStop = False
        Me.GrpBox_Patibill.Text = "Patient Bill Detail"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnDelete, Me.btnEditRegistration})
        Me.GroupBox3.Location = New System.Drawing.Point(296, 32)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(136, 152)
        Me.GroupBox3.TabIndex = 78
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Transaction"
        '
        'btnDelete
        '
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDelete.Location = New System.Drawing.Point(24, 104)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(88, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "Delete"
        '
        'btnEditRegistration
        '
        Me.btnEditRegistration.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnEditRegistration.Location = New System.Drawing.Point(24, 40)
        Me.btnEditRegistration.Name = "btnEditRegistration"
        Me.btnEditRegistration.Size = New System.Drawing.Size(88, 32)
        Me.btnEditRegistration.TabIndex = 2
        Me.btnEditRegistration.Text = "Save"
        '
        'txtBedCategory
        '
        Me.txtBedCategory.Location = New System.Drawing.Point(144, 160)
        Me.txtBedCategory.Name = "txtBedCategory"
        Me.txtBedCategory.ReadOnly = True
        Me.txtBedCategory.TabIndex = 5
        Me.txtBedCategory.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 288)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 76
        Me.Label7.Text = "Total Bill"
        '
        'txtTotalbill
        '
        Me.txtTotalbill.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalbill.Location = New System.Drawing.Point(144, 288)
        Me.txtTotalbill.Name = "txtTotalbill"
        Me.txtTotalbill.ReadOnly = True
        Me.txtTotalbill.TabIndex = 8
        Me.txtTotalbill.Text = ""
        '
        'txtOthercharge
        '
        Me.txtOthercharge.Location = New System.Drawing.Point(144, 256)
        Me.txtOthercharge.Name = "txtOthercharge"
        Me.txtOthercharge.TabIndex = 1
        Me.txtOthercharge.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 256)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 73
        Me.Label5.Text = "Other Charge"
        '
        'txtTotBedcharge
        '
        Me.txtTotBedcharge.Location = New System.Drawing.Point(144, 224)
        Me.txtTotBedcharge.Name = "txtTotBedcharge"
        Me.txtTotBedcharge.ReadOnly = True
        Me.txtTotBedcharge.TabIndex = 7
        Me.txtTotBedcharge.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 224)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 71
        Me.Label4.Text = "Total Bed Charge"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 69
        Me.Label3.Text = "Discharge Date"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(24, 192)
        Me.Label13.Name = "Label13"
        Me.Label13.TabIndex = 68
        Me.Label13.Text = "Bed Charge/Day"
        '
        'txtBedCharge
        '
        Me.txtBedCharge.Location = New System.Drawing.Point(144, 192)
        Me.txtBedCharge.Name = "txtBedCharge"
        Me.txtBedCharge.ReadOnly = True
        Me.txtBedCharge.Size = New System.Drawing.Size(104, 20)
        Me.txtBedCharge.TabIndex = 6
        Me.txtBedCharge.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 96)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Admit Date"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 64)
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
        'txtPatiname
        '
        Me.txtPatiname.Location = New System.Drawing.Point(144, 64)
        Me.txtPatiname.Name = "txtPatiname"
        Me.txtPatiname.ReadOnly = True
        Me.txtPatiname.Size = New System.Drawing.Size(104, 20)
        Me.txtPatiname.TabIndex = 2
        Me.txtPatiname.Text = ""
        '
        'txtRegno
        '
        Me.txtRegno.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegno.Location = New System.Drawing.Point(144, 32)
        Me.txtRegno.Name = "txtRegno"
        Me.txtRegno.ReadOnly = True
        Me.txtRegno.Size = New System.Drawing.Size(104, 22)
        Me.txtRegno.TabIndex = 1
        Me.txtRegno.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(24, 160)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 23)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Bed Type"
        '
        'DTP_date
        '
        Me.DTP_date.Location = New System.Drawing.Point(144, 96)
        Me.DTP_date.Name = "DTP_date"
        Me.DTP_date.ReadOnly = True
        Me.DTP_date.TabIndex = 4
        Me.DTP_date.Text = ""
        '
        'DTP_DOD
        '
        Me.DTP_DOD.Location = New System.Drawing.Point(144, 128)
        Me.DTP_DOD.Name = "DTP_DOD"
        Me.DTP_DOD.ReadOnly = True
        Me.DTP_DOD.TabIndex = 4
        Me.DTP_DOD.Text = ""
        '
        'GB_PatientBillView
        '
        Me.GB_PatientBillView.Controls.AddRange(New System.Windows.Forms.Control() {Me.LV_PatiBillView})
        Me.GB_PatientBillView.Location = New System.Drawing.Point(32, 392)
        Me.GB_PatientBillView.Name = "GB_PatientBillView"
        Me.GB_PatientBillView.Size = New System.Drawing.Size(648, 224)
        Me.GB_PatientBillView.TabIndex = 7
        Me.GB_PatientBillView.TabStop = False
        Me.GB_PatientBillView.Text = "Patient Bill View"
        '
        'LV_PatiBillView
        '
        Me.LV_PatiBillView.FullRowSelect = True
        Me.LV_PatiBillView.GridLines = True
        Me.LV_PatiBillView.Location = New System.Drawing.Point(16, 24)
        Me.LV_PatiBillView.Name = "LV_PatiBillView"
        Me.LV_PatiBillView.Size = New System.Drawing.Size(608, 184)
        Me.LV_PatiBillView.TabIndex = 0
        Me.LV_PatiBillView.View = System.Windows.Forms.View.Details
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCustomprint})
        Me.GroupBox1.Location = New System.Drawing.Point(528, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(120, 112)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Print"
        '
        'btnCustomprint
        '
        Me.btnCustomprint.BackColor = System.Drawing.Color.LightGray
        Me.btnCustomprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCustomprint.Location = New System.Drawing.Point(24, 40)
        Me.btnCustomprint.Name = "btnCustomprint"
        Me.btnCustomprint.Size = New System.Drawing.Size(72, 32)
        Me.btnCustomprint.TabIndex = 4
        Me.btnCustomprint.Text = "List"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSearch, Me.btnExcel})
        Me.GroupBox2.Location = New System.Drawing.Point(528, 240)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(120, 136)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Search/Export"
        '
        'btnSearch
        '
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSearch.Location = New System.Drawing.Point(24, 32)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(72, 32)
        Me.btnSearch.TabIndex = 5
        Me.btnSearch.Text = "Search"
        '
        'btnExcel
        '
        Me.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExcel.Location = New System.Drawing.Point(24, 80)
        Me.btnExcel.Name = "btnExcel"
        Me.btnExcel.Size = New System.Drawing.Size(72, 32)
        Me.btnExcel.TabIndex = 6
        Me.btnExcel.Text = "Excel"
        '
        'frmBill
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(696, 632)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox2, Me.GroupBox1, Me.GrpBox_Patibill, Me.GB_PatientBillView})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmBill"
        Me.Text = "frmBill"
        Me.GrpBox_Patibill.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GB_PatientBillView.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Function for binding textbox with record"

    'code for binding text box with 1st record of table
    Sub Fill_Textbox_Bill()
        Dim r0, r1, r2, r3, r4, r5, r6, r7, r8 As String

        Dim Fill_Textbox_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        'Dim Fill_Textbox_cmdstr As String = "select Regno,Name,Age,Address,Phone,Checkupdate,Checkuptime,Disease,Bloodgroup,Doctor,Remarks from clinicpatients where Regno=1"
        Dim Fill_Textbox_cmdstr As String = "select * from Bill"
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

            Exit While
        End While
        Fill_Textbox_con.Close()

        txtRegno.Text = r0
        txtPatiname.Text = r1
        DTP_date.Text = r2
        DTP_DOD.Text = r3
        txtBedCategory.Text = r4
        txtBedCharge.Text = r5
        txtTotBedcharge.Text = r6
        txtOthercharge.Text = r7
        txtTotalbill.Text = r8
    End Sub

#End Region

#Region "Save Button event"
    'update record
    Private Sub btnEditRegistration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditRegistration.Click
        Dim update_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim update_Comstr As String
        Dim update_Con As OleDbConnection
        Dim update_Com As OleDbCommand

        Dim ID As String

        ID = txtRegno.Text


        update_Comstr = "UPDATE Bill SET Operationother=" & "'" & txtOthercharge.Text & "'" & "," & _
        "Total=" & "'" & txtTotalbill.Text & "'" & " WHERE Regno=" & ID & " ; "

        update_Con = New OleDbConnection(update_Constr)
        update_Com = New OleDbCommand(update_Comstr, update_Con)

        update_Con.Open()
        Try
            update_Com.ExecuteNonQuery()
            MsgBox("Patient bill record saved Successfully.", MsgBoxStyle.Information, "Edit Registration")

            Dim obj As New frmMain()
            obj.setFlag = 0
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
        update_Con.Close()
        fill_listview()
    End Sub
#End Region

#Region "Function for filling Listview"
    'fill listview
    Sub fill_listview()
        ' Clear list view column headers and items
        LV_PatiBillView.Columns.Clear()
        LV_PatiBillView.Items.Clear()

        'If LV.Items.Count > 0 Then
        'LV.Items.Clear()
        'End If

        ' Get SQL Query from textbox
        Dim fillLV_comstr As String = "select * from Bill"
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
                LV_PatiBillView.Columns.Add(header)
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
                LV_PatiBillView.Items.Add(NewItem)
            End While

            ' Close data reader
            NewReader.Close()

        Catch ex As OleDbException
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "oledb Error"
            LV_PatiBillView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem(ex.Message)
            LV_PatiBillView.Items.Add(ErrorItem)

        Catch ex As Exception
            ' Create and error column header
            Dim ErrorHeader As New ColumnHeader()
            ErrorHeader.Text = "Error"
            LV_PatiBillView.Columns.Add(ErrorHeader)

            ' Add Error List Item
            Dim ErrorItem As New ListViewItem("An error has occurred")
            LV_PatiBillView.Items.Add(ErrorItem)

        Finally

            fillLV_Con.Close()

        End Try
    End Sub

#End Region

#Region "Listview Event"

    Private Sub LV_PatiBillView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LV_PatiBillView.SelectedIndexChanged
        Dim SelectedItems As ListView.SelectedListViewItemCollection = _
      CType(sender, ListView).SelectedItems
        If (SelectedItems.Count > 0) Then
            txtRegno.Text = SelectedItems(0).SubItems(0).Text
            txtPatiname.Text = SelectedItems(0).SubItems(1).Text
            DTP_date.Text = SelectedItems(0).SubItems(2).Text
            DTP_DOD.Text = SelectedItems(0).SubItems(3).Text
            txtBedCategory.Text = SelectedItems(0).SubItems(4).Text
            txtBedCharge.Text = SelectedItems(0).SubItems(5).Text
            txtTotBedcharge.Text = SelectedItems(0).SubItems(6).Text
            txtOthercharge.Text = SelectedItems(0).SubItems(7).Text
            txtTotalbill.Text = SelectedItems(0).SubItems(8).Text
        End If
    End Sub
#End Region

#Region "Keypress event"

    Private Sub txtOthercharge_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOthercharge.KeyPress
        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = False
            Else
                e.Handled = True
            End If
        Else
            Dim obj As New frmMain()
            obj.setFlag = 1
        End If
    End Sub
#End Region

#Region "Load event of frmBill"
    Private Sub frmBill_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Fill_Textbox_Bill()
        fill_listview()

        If LV_PatiBillView.Items.Count = 0 Then
            btnEditRegistration.Enabled = False
            btnDelete.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False
            txtOthercharge.ReadOnly = True
        Else
            btnEditRegistration.Enabled = True
            btnDelete.Enabled = True
            txtOthercharge.ReadOnly = False
        End If
    End Sub
#End Region

#Region "Textchanged event"

    Private Sub txtOthercharge_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOthercharge.TextChanged
        txtTotalbill.Text = Val(txtTotBedcharge.Text) + Val(txtOthercharge.Text)

        'Dim obj As New frmMain()
        'obj.setFlag = 1
    End Sub
#End Region

   

    Private Sub btnExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        Me.Cursor = Cursors.WaitCursor
        Dim objclsConvertToExcel As New clsConvertToExcel()
        objclsConvertToExcel.Load_Excel_Details("c:\hm.mdb", "bill")
        Me.Cursor = Cursors.Default
    End Sub
 
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        If LV_PatiBillView.Items.Count = 0 Then
            MessageBox.Show("There is no record available for searching.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim sfrm = New frmSearch(Me, "c:\hm.mdb", "[" + "bill" + "]", "Select * From [" + "bill" + " ]")
            sfrm.Show()
        End If
    End Sub

    'Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
    '    Dim objfrmBill_CurrentPrint As New frmBill_CurrentPrint()

    '    objfrmBill_CurrentPrint.getBill_Print_Values(txtRegno.Text, txtPatiname.Text, _
    '    DTP_date.Value, DTP_DOD.Value, txtBedCategory.Text, txtBedCharge.Text, _
    '    txtTotBedcharge.Text, txtOthercharge.Text, txtTotalbill.Text)

    '    objfrmBill_CurrentPrint.Show()
    'End Sub

    Private Sub btnCustomprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCustomprint.Click
        Dim objfrmcheckbox = New frmCheckBox(Me, "c:\hm.mdb", "bill")
        objfrmcheckbox.show()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim delete_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
        Dim delete_Comstr As String
        Dim delete_Con As OleDbConnection
        Dim delete_Com As OleDbCommand

        Dim ID As String

        ID = txtRegno.Text
        delete_Comstr = "Delete from bill WHERE Regno=" & ID

        delete_Con = New OleDbConnection(delete_Constr)
        delete_Com = New OleDbCommand(delete_Comstr, delete_Con)

        delete_Con.Open()
        Try
            delete_Com.ExecuteNonQuery()
            MsgBox("Patient's bill record deleted  successfully.", MsgBoxStyle.Information, "Delete Registration")
        Catch ex As OleDbException
            MsgBox(ex.Message)
        End Try
        delete_Con.Close()

        Fill_Textbox_Bill()
        fill_listview()

        If LV_PatiBillView.Items.Count = 0 Then
            btnEditRegistration.Enabled = False
            btnDelete.Enabled = False
            'btnPrint.Enabled = False
            btnCustomprint.Enabled = False
            btnSearch.Enabled = False
            btnExcel.Enabled = False
            txtOthercharge.ReadOnly = True
        Else
            btnEditRegistration.Enabled = True
            btnDelete.Enabled = True
            txtOthercharge.ReadOnly = False
        End If
    End Sub
End Class
