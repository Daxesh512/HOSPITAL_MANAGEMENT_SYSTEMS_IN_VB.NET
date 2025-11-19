Imports System.Data.OleDb

Public Class frmCheckBox
    Inherits System.Windows.Forms.Form

#Region "Variables"
    Private colType() As String 'This array holds all the columnTypes;        

    Dim DB_Filename As String = ""
    Dim DB_Tablename As String = ""
    Dim query, dummyquery As String
    Dim clbFlag As Boolean = True
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal motherFrm As Form, ByVal datasource As String, ByVal table_Name As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'frmcheck_DG.Visible = False
        Add_CBL(datasource, table_Name)

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
    Friend WithEvents clbforcolumn As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnSelectall As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents frmcheck_DG As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCheckBox))
        Me.clbforcolumn = New System.Windows.Forms.CheckedListBox()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.btnSelectall = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.frmcheck_DG = New System.Windows.Forms.DataGrid()
        Me.GroupBox1.SuspendLayout()
        CType(Me.frmcheck_DG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'clbforcolumn
        '
        Me.clbforcolumn.CheckOnClick = True
        Me.clbforcolumn.Location = New System.Drawing.Point(16, 16)
        Me.clbforcolumn.Name = "clbforcolumn"
        Me.clbforcolumn.Size = New System.Drawing.Size(160, 214)
        Me.clbforcolumn.TabIndex = 0
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnPrint.Location = New System.Drawing.Point(328, 104)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(104, 32)
        Me.btnPrint.TabIndex = 3
        Me.btnPrint.Text = "Print"
        '
        'btnSelectall
        '
        Me.btnSelectall.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSelectall.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnSelectall.Location = New System.Drawing.Point(328, 56)
        Me.btnSelectall.Name = "btnSelectall"
        Me.btnSelectall.Size = New System.Drawing.Size(104, 32)
        Me.btnSelectall.TabIndex = 4
        Me.btnSelectall.Text = "Select All"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.frmcheck_DG})
        Me.GroupBox1.Location = New System.Drawing.Point(8, 256)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(608, 264)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "View"
        '
        'frmcheck_DG
        '
        Me.frmcheck_DG.DataMember = ""
        Me.frmcheck_DG.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.frmcheck_DG.Location = New System.Drawing.Point(20, 32)
        Me.frmcheck_DG.Name = "frmcheck_DG"
        Me.frmcheck_DG.Size = New System.Drawing.Size(568, 216)
        Me.frmcheck_DG.TabIndex = 3
        '
        'frmCheckBox
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(624, 534)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox1, Me.btnSelectall, Me.btnPrint, Me.clbforcolumn})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCheckBox"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmCheckBox"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.frmcheck_DG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "code for adding CBL"

    Sub Add_CBL(ByVal datasource As String, ByVal table_Name As String)
        DB_Filename = datasource
        DB_Tablename = table_Name

        Dim colType() As String

        ' Create a Connection object

        Dim da As OleDbDataAdapter
        Dim ds As New DataSet()
        'Dim dt As DataTable

        Dim dbConn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DB_Filename)

        ' Create the data adapter object pointing to the authors table
        da = New OleDbDataAdapter( _
              "SELECT * FROM " + DB_Tablename, dbConn)

        ' Fill the DataSet
        ds = New DataSet("acc")

        Try
            da.Fill(ds)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



        ' Attempt to fill the dataset through the accessDataAdapter
        da.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", DB_Tablename)})
        'populate the DataSet with existing constraints information from a data source
        da.FillSchema(ds, SchemaType.Source, DB_Tablename)

        ' Fill the dataser
        da.Fill(ds)

        Dim dt As DataTable
        dt = ds.Tables(DB_Tablename)

        ' Display the records in a DataGrid component
        'DataGrid1.DataSource = ds.Tables(0)

        ' Dynamic Controls Postions
        Dim controlTop As Integer = 10
        Dim controlLeft As Integer = 10

        'Get all the System.DataTypes of all the
        'columns in the table and assign them to the
        'array colType

        colType = New String(dt.Columns.Count) {}



        'Here AutoMenu is created which would allow the user
        'to insert automatic incrementation of numbers (+1 on the
        'last cell) or insert today's date for datetime type
        'columns
        'AutoMenu = New ContextMenu(accessDataTable.Columns.Count) {}

        'Create dynamically all the textboxes and labels
        'which will hold and link information to the database
        'making it easier to input data

        'Dim objcheckbox As New frmCheckBox()
        'Dim ii As Integer
        'ii = dt.Columns.Count

        Dim i As Integer
        For i = 0 To dt.Columns.Count - 1
            colType(i) = dt.Columns(i).DataType.ToString()

            'Create the control (CheckBox)
            Dim CheckBoxControl As CheckBox = CType(clsCreateControls.MakeControl("CheckBox", 30, 100, controlLeft, controlTop + 3, dt.Columns(i).Caption, "cCheckBox" + i.ToString()), CheckBox)
            clbforcolumn.Items.Add(CheckBoxControl.Text)

            ''Finally add the controls to the form                 
            'Me.Controls.Add(CheckBoxControl)

            ''Here we arrange the controls on the form
            'If i = dt.Columns.Count - 1 Then
            '    Exit For
            'End If
            'If controlLeft = 10 Then
            '    controlLeft = 220
            'Else
            '    If controlLeft = 220 And dt.Columns.Count > 6 Then
            '        controlLeft = 430
            '        Me.Width = 680
            '    Else
            '        controlTop += CheckBoxControl.Height + 10
            '        controlLeft = 10
            '    End If
            'End If
        Next i
    End Sub

#End Region

#Region "Code for filling datagrid"

    Sub Fill_DG()
        'DB_Filename = datasource
        'DB_Tablename = table_Name

        Dim colType() As String

        query = ""
        Dim counter As Integer = 1


        If (clbforcolumn.SelectedItem <> "") Then
            Dim Entry As Object
            For Each Entry In clbforcolumn.CheckedItems
                counter = counter + 1
                'MessageBox.Show(Entry.ToString())

                'MsgBox(CheckedListBox1.CheckedItems.Count, MsgBoxStyle.Information)
                If counter > clbforcolumn.CheckedItems.Count Then
                    query &= Entry.ToString()
                Else
                    query &= Entry.ToString() + ","
                End If
            Next
        Else
            MsgBox("You must select at least one field.", MsgBoxStyle.Information, "Error Information")
            Exit Sub
        End If
        query = "select " + query + " from "
        dummyquery = "select " + dummyquery + " from "

        ' Create a Connection object

        Dim da As OleDbDataAdapter
        Dim ds As New DataSet()
        'Dim dt As DataTable

        Dim dbConn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DB_Filename)

        ' Create the data adapter object pointing to the authors table
        'If clbFlag = False Then
        '    da = New OleDbDataAdapter( _
        '      dummyquery + DB_Tablename, dbConn)
        '    clbFlag = True
        'Else
        da = New OleDbDataAdapter( _
                          query + DB_Tablename, dbConn)
        'End If


        ' Fill the DataSet
        ds = New DataSet("acc")

        Try
            da.Fill(ds)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' Attempt to fill the dataset through the accessDataAdapter
        da.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", DB_Tablename)})
        'populate the DataSet with existing constraints information from a data source
        da.FillSchema(ds, SchemaType.Source, DB_Tablename)

        ' Fill the dataser
        da.Fill(ds)

        Dim dt As DataTable
        dt = ds.Tables(DB_Tablename)

        ' Display the records in a DataGrid component
        frmcheck_DG.DataSource = ds.Tables(0)
    End Sub

#End Region

#Region "All button events"

    'Private Sub btnShowDG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowDG.Click
    '    If clbforcolumn.CheckedItems.Count = 0 Then
    '        Dim count_cbl As Integer
    '        Dim clbString As Object

    '        clbString = clbforcolumn.SelectedItem
    '        count_cbl = clbforcolumn.SelectedIndex
    '        MsgBox("You must stay with at least one field.", MsgBoxStyle.Information, "Error Information")

    '        clbforcolumn.SetItemChecked(count_cbl, True)
    '        'clbFlag = False
    '        'dummyquery = clbString
    '        Fill_DG()
    '    Else
    '        frmcheck_DG.Visible = True
    '        Fill_DG()
    '    End If
    'End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        'Create a new instance of the frmSelectedPrint form
        'Specifying the datasource and table to view
        Dim objfrmSelectedPrint As New frmSelectedPrint(Me, DB_Filename, DB_Tablename, query)

        objfrmSelectedPrint.Show()
    End Sub

    Private Sub btnSelectall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectall.Click
        Dim counter As Integer

        For counter = 0 To clbforcolumn.Items.Count - 1
            clbforcolumn.SetItemChecked(counter, True)
        Next
        clbforcolumn.SetSelected(counter - 1, True)

    End Sub

#End Region

#Region "Code for clbforcolumn event"

    Private Sub clbforcolumn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clbforcolumn.SelectedIndexChanged
        If clbforcolumn.CheckedItems.Count = 0 Then
            Dim count_cbl As Integer
            Dim clbString As Object

            clbString = clbforcolumn.SelectedItem
            count_cbl = clbforcolumn.SelectedIndex
            MsgBox("You must stay with at least one field.", MsgBoxStyle.Information, "Error Information")

            clbforcolumn.SetItemChecked(count_cbl, True)
            'clbFlag = False
            'dummyquery = clbString
            Fill_DG()
            btnPrint.Enabled = True
        Else
            'frmcheck_DG.Visible = True
            Fill_DG()
            btnPrint.Enabled = True
        End If
    End Sub

#End Region

End Class

