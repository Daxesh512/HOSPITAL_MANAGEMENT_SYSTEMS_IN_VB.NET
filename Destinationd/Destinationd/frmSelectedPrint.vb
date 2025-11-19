Imports System.Data.OleDb


Public Class frmSelectedPrint
    Inherits System.Windows.Forms.Form

#Region "Private members"
    Private GridPrinter As clsDataGridPrinter

#End Region

#Region "variables"

    Dim DB_Filename As String = ""
    Dim DB_Tablename As String = ""
    Dim colType() As String  'This array holds all the columnTypes;
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal motherFrm As Form, ByVal datasource As String, ByVal table_Name As String, ByVal getquery As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Call PopulateColourList(Me.ComboBox_ColourBodyline)
        Call PopulateColourList(Me.ComboBox_ColourFooterLine)
        Call PopulateColourList(Me.ComboBox_ColourHeaderLine)

        Call PopulateBrushList(Me.ComboBox_EvenBrush)
        Call PopulateBrushList(Me.ComboBox_FooterBrush)
        Call PopulateBrushList(Me.ComboBox_HeaderBrush)
        Call PopulateBrushList(Me.ComboBox_OddRowBrush)
        Call PopulateBrushList(Me.ComboBox_ColumnHeaderBrush)

        '\\ Populate teh data grids with some bumpf
        'Dim MyTable As New DataTable()
        'MyTable.Columns.Add(New DataColumn("Team", GetType(String)))
        'MyTable.Columns.Add(New DataColumn("Played", GetType(Integer)))
        'MyTable.Columns.Add(New DataColumn("Goals For", GetType(Integer)))
        'MyTable.Columns.Add(New DataColumn("Goals Against", GetType(Integer)))
        'MyTable.Columns.Add(New DataColumn("Points", GetType(Integer)))

        'Me.DataGrid1.DataSource = MyTable

        DB_Filename = datasource
        DB_Tablename = table_Name


        ' Create a Connection object

        Dim accessDataTable As DataTable
        Dim accessDataSet As DataSet 'the main DataSet
        Dim accessDataAdapter As OleDbDataAdapter 'the adapter to be used in conjunction with

        Dim da As OleDbDataAdapter
        Dim ds As DataSet

        Dim dbConn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DB_Filename)

        ' Create the data adapter object pointing to the authors table
        da = New OleDbDataAdapter( _
              getquery + DB_Tablename, dbConn)

        ' Fill the DataSet
        ds = New DataSet("acc")

        Try
            da.Fill(ds)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' Display the records in a DataGrid component
        Try
            DataGrid1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MsgBox("You must select some fields in checkbox.")
        End Try

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
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents MainMenu_App As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem_File As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem_File_PageSetup As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem_File_print As System.Windows.Forms.MenuItem
    Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_HeaderHeightPercentage As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown_FooterHeightPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_InterSectionSpacingPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_ColourHeaderLine As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_ColourFooterLine As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_ColourBodyline As System.Windows.Forms.ComboBox
    Friend WithEvents PrintPreviewDialog1 As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox_FooterBrush As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_HeaderBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_EvenBrush As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_OddRowBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_ColumnHeaderBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown_PagesAcross As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label13 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSelectedPrint))
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.MainMenu_App = New System.Windows.Forms.MainMenu()
        Me.MenuItem_File = New System.Windows.Forms.MenuItem()
        Me.MenuItem_File_PageSetup = New System.Windows.Forms.MenuItem()
        Me.MenuItem_File_print = New System.Windows.Forms.MenuItem()
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown_InterSectionSpacingPercent = New System.Windows.Forms.NumericUpDown()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.NumericUpDown_FooterHeightPercent = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.NumericUpDown_HeaderHeightPercentage = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.ComboBox_ColourBodyline = New System.Windows.Forms.ComboBox()
        Me.ComboBox_ColourFooterLine = New System.Windows.Forms.ComboBox()
        Me.ComboBox_ColourHeaderLine = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ComboBox_ColumnHeaderBrush = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ComboBox_EvenBrush = New System.Windows.Forms.ComboBox()
        Me.ComboBox_OddRowBrush = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ComboBox_FooterBrush = New System.Windows.Forms.ComboBox()
        Me.ComboBox_HeaderBrush = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.NumericUpDown_PagesAcross = New System.Windows.Forms.NumericUpDown()
        Me.Label13 = New System.Windows.Forms.Label()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown_InterSectionSpacingPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_FooterHeightPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_HeaderHeightPercentage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.NumericUpDown_PagesAcross, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.LightGray
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Dock = System.Windows.Forms.DockStyle.Top
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.Size = New System.Drawing.Size(552, 128)
        Me.DataGrid1.TabIndex = 0
        '
        'MainMenu_App
        '
        Me.MainMenu_App.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem_File})
        '
        'MenuItem_File
        '
        Me.MenuItem_File.Index = 0
        Me.MenuItem_File.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem_File_PageSetup, Me.MenuItem_File_print})
        Me.MenuItem_File.Text = "&File"
        '
        'MenuItem_File_PageSetup
        '
        Me.MenuItem_File_PageSetup.Index = 0
        Me.MenuItem_File_PageSetup.Text = "Page &Setup"
        '
        'MenuItem_File_print
        '
        Me.MenuItem_File_print.Index = 1
        Me.MenuItem_File_print.Text = "&Print"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(88, 136)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(456, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "Page Heading"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Page Heading"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.NumericUpDown_InterSectionSpacingPercent, Me.Label4, Me.NumericUpDown_FooterHeightPercent, Me.Label3, Me.NumericUpDown_HeaderHeightPercentage, Me.Label2})
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 160)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(536, 56)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Section Heights"
        '
        'NumericUpDown_InterSectionSpacingPercent
        '
        Me.NumericUpDown_InterSectionSpacingPercent.Location = New System.Drawing.Point(384, 24)
        Me.NumericUpDown_InterSectionSpacingPercent.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NumericUpDown_InterSectionSpacingPercent.Name = "NumericUpDown_InterSectionSpacingPercent"
        Me.NumericUpDown_InterSectionSpacingPercent.Size = New System.Drawing.Size(40, 20)
        Me.NumericUpDown_InterSectionSpacingPercent.TabIndex = 5
        Me.NumericUpDown_InterSectionSpacingPercent.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(248, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Inter-section spacing"
        '
        'NumericUpDown_FooterHeightPercent
        '
        Me.NumericUpDown_FooterHeightPercent.Location = New System.Drawing.Point(176, 24)
        Me.NumericUpDown_FooterHeightPercent.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.NumericUpDown_FooterHeightPercent.Name = "NumericUpDown_FooterHeightPercent"
        Me.NumericUpDown_FooterHeightPercent.Size = New System.Drawing.Size(40, 20)
        Me.NumericUpDown_FooterHeightPercent.TabIndex = 3
        Me.NumericUpDown_FooterHeightPercent.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(120, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Footer"
        '
        'NumericUpDown_HeaderHeightPercentage
        '
        Me.NumericUpDown_HeaderHeightPercentage.Location = New System.Drawing.Point(64, 24)
        Me.NumericUpDown_HeaderHeightPercentage.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.NumericUpDown_HeaderHeightPercentage.Name = "NumericUpDown_HeaderHeightPercentage"
        Me.NumericUpDown_HeaderHeightPercentage.Size = New System.Drawing.Size(40, 20)
        Me.NumericUpDown_HeaderHeightPercentage.TabIndex = 1
        Me.NumericUpDown_HeaderHeightPercentage.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Header"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.ComboBox_ColourBodyline, Me.ComboBox_ColourFooterLine, Me.ComboBox_ColourHeaderLine, Me.Label5, Me.Label6, Me.Label7})
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(8, 224)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(536, 56)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Grid line colours"
        '
        'ComboBox_ColourBodyline
        '
        Me.ComboBox_ColourBodyline.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourBodyline.Location = New System.Drawing.Point(400, 24)
        Me.ComboBox_ColourBodyline.Name = "ComboBox_ColourBodyline"
        Me.ComboBox_ColourBodyline.Size = New System.Drawing.Size(128, 21)
        Me.ComboBox_ColourBodyline.TabIndex = 10
        '
        'ComboBox_ColourFooterLine
        '
        Me.ComboBox_ColourFooterLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourFooterLine.Location = New System.Drawing.Point(216, 24)
        Me.ComboBox_ColourFooterLine.Name = "ComboBox_ColourFooterLine"
        Me.ComboBox_ColourFooterLine.Size = New System.Drawing.Size(104, 21)
        Me.ComboBox_ColourFooterLine.TabIndex = 9
        '
        'ComboBox_ColourHeaderLine
        '
        Me.ComboBox_ColourHeaderLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourHeaderLine.Location = New System.Drawing.Point(48, 24)
        Me.ComboBox_ColourHeaderLine.Name = "ComboBox_ColourHeaderLine"
        Me.ComboBox_ColourHeaderLine.Size = New System.Drawing.Size(112, 21)
        Me.ComboBox_ColourHeaderLine.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(336, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Body"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(168, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 16)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Footer"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Header"
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Location = New System.Drawing.Point(279, 17)
        Me.PrintPreviewDialog1.MaximumSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Opacity = 1
        Me.PrintPreviewDialog1.TransparencyKey = System.Drawing.Color.Empty
        Me.PrintPreviewDialog1.Visible = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.ComboBox_ColumnHeaderBrush, Me.Label12, Me.ComboBox_EvenBrush, Me.ComboBox_OddRowBrush, Me.Label8, Me.Label11, Me.ComboBox_FooterBrush, Me.ComboBox_HeaderBrush, Me.Label9, Me.Label10})
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(8, 288)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(536, 56)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Grid background  colours"
        '
        'ComboBox_ColumnHeaderBrush
        '
        Me.ComboBox_ColumnHeaderBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_ColumnHeaderBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColumnHeaderBrush.Location = New System.Drawing.Point(488, 24)
        Me.ComboBox_ColumnHeaderBrush.Name = "ComboBox_ColumnHeaderBrush"
        Me.ComboBox_ColumnHeaderBrush.Size = New System.Drawing.Size(40, 21)
        Me.ComboBox_ColumnHeaderBrush.TabIndex = 15
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(424, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(60, 16)
        Me.Label12.TabIndex = 14
        Me.Label12.Text = "Columns"
        '
        'ComboBox_EvenBrush
        '
        Me.ComboBox_EvenBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_EvenBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_EvenBrush.Location = New System.Drawing.Point(368, 24)
        Me.ComboBox_EvenBrush.Name = "ComboBox_EvenBrush"
        Me.ComboBox_EvenBrush.Size = New System.Drawing.Size(40, 21)
        Me.ComboBox_EvenBrush.TabIndex = 13
        '
        'ComboBox_OddRowBrush
        '
        Me.ComboBox_OddRowBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_OddRowBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_OddRowBrush.Location = New System.Drawing.Point(256, 24)
        Me.ComboBox_OddRowBrush.Name = "ComboBox_OddRowBrush"
        Me.ComboBox_OddRowBrush.Size = New System.Drawing.Size(40, 21)
        Me.ComboBox_OddRowBrush.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(304, 24)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 16)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Even rows"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(200, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 16)
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "Odd rows"
        '
        'ComboBox_FooterBrush
        '
        Me.ComboBox_FooterBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_FooterBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_FooterBrush.Location = New System.Drawing.Point(152, 24)
        Me.ComboBox_FooterBrush.Name = "ComboBox_FooterBrush"
        Me.ComboBox_FooterBrush.Size = New System.Drawing.Size(40, 21)
        Me.ComboBox_FooterBrush.TabIndex = 9
        '
        'ComboBox_HeaderBrush
        '
        Me.ComboBox_HeaderBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_HeaderBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_HeaderBrush.Location = New System.Drawing.Point(64, 24)
        Me.ComboBox_HeaderBrush.Name = "ComboBox_HeaderBrush"
        Me.ComboBox_HeaderBrush.Size = New System.Drawing.Size(40, 21)
        Me.ComboBox_HeaderBrush.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(112, 24)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 16)
        Me.Label9.TabIndex = 6
        Me.Label9.Text = "Footer"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 24)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 16)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Header"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.AddRange(New System.Windows.Forms.Control() {Me.NumericUpDown_PagesAcross, Me.Label13})
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox4.Location = New System.Drawing.Point(8, 352)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(536, 40)
        Me.GroupBox4.TabIndex = 6
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Page layout"
        '
        'NumericUpDown_PagesAcross
        '
        Me.NumericUpDown_PagesAcross.Location = New System.Drawing.Point(368, 10)
        Me.NumericUpDown_PagesAcross.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.NumericUpDown_PagesAcross.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown_PagesAcross.Name = "NumericUpDown_PagesAcross"
        Me.NumericUpDown_PagesAcross.Size = New System.Drawing.Size(48, 20)
        Me.NumericUpDown_PagesAcross.TabIndex = 3
        Me.NumericUpDown_PagesAcross.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(24, 16)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(328, 16)
        Me.Label13.TabIndex = 2
        Me.Label13.Text = "Minimum number of pages across to split the columns over"
        '
        'frmSelectedPrint
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(552, 393)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox4, Me.GroupBox3, Me.GroupBox2, Me.GroupBox1, Me.Label1, Me.TextBox1, Me.DataGrid1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu_App
        Me.MinimizeBox = False
        Me.Name = "frmSelectedPrint"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Record Print"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.NumericUpDown_InterSectionSpacingPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_FooterHeightPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_HeaderHeightPercentage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.NumericUpDown_PagesAcross, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Menu handlers"

    Private Sub MenuItem_File_PageSetup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem_File_PageSetup.Click

        If GridPrinter Is Nothing Then
            GridPrinter = New clsDataGridPrinter(Me.DataGrid1)
        End If

        With Me.PageSetupDialog1
            .Document = GridPrinter.PrintDocument
            .ShowDialog()
        End With

    End Sub

    Private Sub MenuItem_File_print_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem_File_print.Click
        If GridPrinter Is Nothing Then
            GridPrinter = New clsDataGridPrinter(Me.DataGrid1)
        End If

        With GridPrinter
            .HeaderText = Me.TextBox1.Text
            .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
            .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
            .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
            .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
            .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
            .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
            .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
            .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
            .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
            .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
            .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
            .PagesAcross = CInt(Me.NumericUpDown_PagesAcross.Value)
        End With

        With Me.PrintPreviewDialog1
            .Document = GridPrinter.PrintDocument
            If .ShowDialog = DialogResult.OK Then
                GridPrinter.Print()
            End If
        End With
    End Sub

#End Region

#Region "Private methods"
    Private Sub PopulateColourList(ByVal combo As ComboBox)

        combo.Items.Clear()
        combo.Items.Add(System.Drawing.Color.AliceBlue)
        combo.Items.Add(System.Drawing.Color.Aqua)
        combo.Items.Add(System.Drawing.Color.Azure)
        combo.Items.Add(System.Drawing.Color.Beige)
        combo.Items.Add(System.Drawing.Color.Black)
        combo.Items.Add(System.Drawing.Color.Blue)
        combo.Items.Add(System.Drawing.Color.Green)
        combo.Items.Add(System.Drawing.Color.Red)
        combo.SelectedIndex = 0
    End Sub

    Private Sub PopulateBrushList(ByVal Combo As ComboBox)
        Combo.Items.Clear()
        Combo.Items.Add(System.Drawing.Brushes.White)
        Combo.Items.Add(System.Drawing.Brushes.Beige)
        Combo.Items.Add(System.Drawing.Brushes.Bisque)
        Combo.Items.Add(System.Drawing.Brushes.BlanchedAlmond)
        Combo.Items.Add(System.Drawing.Brushes.Blue)
        Combo.Items.Add(System.Drawing.Brushes.BlueViolet)
        Combo.Items.Add(System.Drawing.Brushes.Brown)
        Combo.Items.Add(System.Drawing.Brushes.BurlyWood)
        Combo.Items.Add(System.Drawing.Brushes.CadetBlue)
        Combo.Items.Add(System.Drawing.Brushes.Chartreuse)
        Combo.Items.Add(System.Drawing.Brushes.Chocolate)
        Combo.Items.Add(System.Drawing.Brushes.Coral)
        Combo.Items.Add(System.Drawing.Brushes.CornflowerBlue)
        Combo.Items.Add(System.Drawing.Brushes.Cornsilk)
        Combo.Items.Add(System.Drawing.Brushes.Crimson)
        Combo.Items.Add(System.Drawing.Brushes.Cyan)
        Combo.SelectedIndex = 0
    End Sub
#End Region

#Region "All ComboBox code"

    Private Sub ComboBox_EvenBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_EvenBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_EvenBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub


    Private Sub ComboBox_FooterBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_FooterBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_FooterBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub


    Private Sub ComboBox_OddRowBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_OddRowBrush.DrawItem

        e.Graphics.FillRectangle(CType(ComboBox_OddRowBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)

    End Sub


    Private Sub ComboBox_HeaderBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_HeaderBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_HeaderBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub


    Private Sub ComboBox_ColumnHeaderBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_ColumnHeaderBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_ColumnHeaderBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
#End Region

End Class

