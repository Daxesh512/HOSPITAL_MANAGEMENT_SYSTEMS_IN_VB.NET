Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.OleDb

Public Class frmSearch
    Inherits System.Windows.Forms.Form

#Region "Private Declarations"
    Private activateMouse As Boolean = False 'a boolian to allow mouse click on the DGrid
    Private elements As Collection 'the filter menues' elements are stored here
    Private checkedMenu() As String 'the element you are looking for in the coulmn
    Private accessDataSet As New DataSet()
    Private accessConnection As New OleDbConnection()

    Private tableName As String = "" 'DataBase table name
    Private doUpdate As Boolean = True 'Update the filter menues or not 
    Private columnHit As Integer 'which column in the table is hit
    Private WithEvents dGrid As System.Windows.Forms.DataGrid
    Private FilterMenu() As System.Windows.Forms.ContextMenu 'the filter menues for all the columns
    Private cBoxParamets As System.Windows.Forms.ComboBox 'the combobox which holds the column names
    'to chose from in the search
    Private WithEvents searchTxt As System.Windows.Forms.TextBox 'the search element is put here
    'for the text based search
    'find row in data according to text based search
    Private WithEvents btnRestore As System.Windows.Forms.Button 'remove all filters button
    Private MFRM As Form
#End Region

#Region "variables"
    Dim DB_FileName As String
    Dim DB_TableName As String
    Dim DB_Query As String

    Private components As System.ComponentModel.Container = Nothing
#End Region

#Region "New and Dispose methods"
    'The search class constructor
    Public Sub New(ByVal motherFrm As Form, ByVal datasource As String, ByVal table_Name As String, ByVal SelectString As String)
        InitializeComponent()

        DB_FileName = datasource
        DB_TableName = table_Name

        'This refers to the parent form
        MFRM = motherFrm
        'the tablename to do the search on
        tableName = table_Name

        'Initializing the connection here to the source mdb file
        CType(Me.accessDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        accessConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + datasource '
        accessConnection.Open()
        loadData(SelectString)
        'lblSelectString.Text = "Select Command = " + SelectString
        activateMouse = True
        accessConnection.Close()

        Dim myDataTable As DataTable = accessDataSet.Tables(tableName)

        'Find all columns and put them in the combobox
        'cBoxParamets
        Dim i As Integer
        For i = 0 To myDataTable.Columns.Count - 1
            cBoxParamets.Items.Add(myDataTable.Columns(i).Caption)
            If i = 0 Then
                cBoxParamets.Text = myDataTable.Columns(i).Caption
            End If
        Next i
        'dGrid.Height = Me.Height - 135
    End Sub 'New

    Protected Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub 'Dispose
#End Region

#Region "Windows Form Designer generated code"
    '/ <summary>
    '/ Required method for Designer support - do not modify
    '/ the contents of this method with the code editor.
    '/ </summary>
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSearch))
        Me.dGrid = New System.Windows.Forms.DataGrid()
        Me.cBoxParamets = New System.Windows.Forms.ComboBox()
        Me.searchTxt = New System.Windows.Forms.TextBox()
        Me.btnRestore = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnPrint = New System.Windows.Forms.Button()
        CType(Me.dGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dGrid
        '
        Me.dGrid.AccessibleRole = System.Windows.Forms.AccessibleRole.ColumnHeader
        Me.dGrid.AllowSorting = False
        Me.dGrid.BackgroundColor = System.Drawing.Color.LightGray
        Me.dGrid.DataMember = ""
        Me.dGrid.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dGrid.Location = New System.Drawing.Point(0, 174)
        Me.dGrid.Name = "dGrid"
        Me.dGrid.ReadOnly = True
        Me.dGrid.Size = New System.Drawing.Size(592, 232)
        Me.dGrid.TabIndex = 1
        '
        'cBoxParamets
        '
        Me.cBoxParamets.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cBoxParamets.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.cBoxParamets.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cBoxParamets.Location = New System.Drawing.Point(112, 40)
        Me.cBoxParamets.Name = "cBoxParamets"
        Me.cBoxParamets.Size = New System.Drawing.Size(120, 21)
        Me.cBoxParamets.TabIndex = 2
        '
        'searchTxt
        '
        Me.searchTxt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.searchTxt.Location = New System.Drawing.Point(112, 80)
        Me.searchTxt.Name = "searchTxt"
        Me.searchTxt.Size = New System.Drawing.Size(120, 20)
        Me.searchTxt.TabIndex = 3
        Me.searchTxt.Text = ""
        '
        'btnRestore
        '
        Me.btnRestore.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnRestore.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnRestore.Location = New System.Drawing.Point(336, 40)
        Me.btnRestore.Name = "btnRestore"
        Me.btnRestore.Size = New System.Drawing.Size(80, 32)
        Me.btnRestore.TabIndex = 5
        Me.btnRestore.Text = "Restore"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 24)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Select Category"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 24)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Enter input"
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPrint.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnPrint.Location = New System.Drawing.Point(336, 88)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(80, 32)
        Me.btnPrint.TabIndex = 8
        Me.btnPrint.Text = "Print"
        '
        'frmSearch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(592, 406)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnPrint, Me.Label2, Me.Label1, Me.btnRestore, Me.searchTxt, Me.cBoxParamets, Me.dGrid})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(352, 360)
        Me.Name = "frmSearch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Search Record"
        CType(Me.dGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub 'InitializeComponent

#End Region

#Region "load Data"
    'This routine loads data from a mdb file
    'into the DGrid
    Private Sub loadData(ByVal SelectString As String)
        accessDataSet.RejectChanges()
        accessDataSet.Clear()

        Dim accessSelectCommand As New OleDbCommand()
        Dim accessInsertCommand As New OleDbCommand()
        Dim accessDataAdapter As New OleDbDataAdapter()

        accessSelectCommand.CommandText = SelectString
        accessSelectCommand.Connection = accessConnection
        accessDataAdapter.SelectCommand = accessSelectCommand

        'DB_Query = accessSelectCommand.CommandText

        ' Attempt to fill the dataset through the OleDbDataAdapter1.
        accessDataAdapter.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", tableName)})
        accessDataAdapter.Fill(accessDataSet)

        dGrid.SetDataBinding(accessDataSet, tableName)

        Dim col As Integer = accessDataSet.Tables(tableName).Columns.Count
        Dim row As Integer = accessDataSet.Tables(tableName).Rows.Count

        If doUpdate = True Then
            checkedMenu = New [String](col - 1) {}
        End If

        elements = New Collection()

        'elements = New Object()(col) {}
        FilterMenu = New ContextMenu(col - 1) {}

        Dim i As Integer
        For i = 0 To col - 1
            elements.Add(New Object(row - 1) {})
            If doUpdate = True Then
                checkedMenu(i) = "None"
            End If
        Next i

        For i = 0 To col - 1
            Dim j As Integer
            For j = 0 To row - 1
                If dGrid(j, i).GetType().Name = "Int32" Or dGrid(j, i).GetType().Name = "DateTime" Or dGrid(j, i).GetType().Name = "Decimal" Then
                    elements.Item(i + 1)(j) = dGrid(j, i)
                Else
                    elements.Item(i + 1)(j) = dGrid(j, i).ToString()
                End If
            Next j
        Next i

        For i = 0 To col - 1
            Try
                Array.Sort(elements.Item(i + 1))
            Catch
            End Try
            'FilterMenu(i) = New ContextMenu()
            'Console.WriteLine(elements(i + 1)(0))
            'make_menues(elements(i + 1), FilterMenu(i))
        Next i
    End Sub 'loadData

#End Region

#Region "Make the Select Command"
    'This routine creates the seacrh command to be used
    'as a select command based on the options specified
    'by the user through text based search or menu based
    'filter
    Private Function MakeSelectString(ByVal MenuChecked() As String) As String
        Dim accessDataTable As DataTable = accessDataSet.Tables(tableName)
        Dim STselect As String = "Select * From " + tableName + " Where "
        Dim there_is_Change As Boolean = False
        Dim i As Integer
        For i = 0 To MenuChecked.Length - 1

            Dim colType As String = accessDataTable.Columns(i).DataType.ToString()


            If MenuChecked(i) <> "None" Then
                If i <> 0 And there_is_Change = True Then

                    If colType = "System.String" Then
                        STselect += " And [" + accessDataSet.Tables(tableName).Columns(i).Caption + " ] = '" + MenuChecked(i) + "'"
                    Else
                        If colType = "System.DateTime" Then
                            STselect += " And [" + accessDataSet.Tables(tableName).Columns(i).Caption + " ] = #" + MenuChecked(i) + "#"
                        Else
                            STselect += " And [" + accessDataSet.Tables(tableName).Columns(i).Caption + " ] = " + MenuChecked(i)
                        End If
                    End If
                Else
                    If colType = "System.String" Then
                        STselect += " [" + accessDataSet.Tables(tableName).Columns(i).Caption + " ] = '" + MenuChecked(i) + "'"
                    Else
                        If colType = "System.DateTime" Then
                            STselect += " [" + accessDataSet.Tables(tableName).Columns(i).Caption + " ] = #" + MenuChecked(i) + "#"
                        Else
                            STselect += " [" + accessDataSet.Tables(tableName).Columns(i).Caption + " ] = " + MenuChecked(i)
                        End If
                    End If
                End If
                there_is_Change = True
            End If
        Next i
        If there_is_Change = False Then
            STselect = "Select * From " + tableName
        End If
        'lblSelectString.Text = "Select Command = " + STselect
        Return STselect
    End Function 'MakeSelectString

#End Region

#Region "find_the_data Method"
    'Button find is clicked
    'Private Sub Findbtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Findbtn.Click
    '    find_the_data()
    'End Sub 'Findbtn_Click


    'based on the element required in the search string
    Private Sub find_the_data()
        Dim index As Integer = 0
        If searchTxt.Text = "" Then
            Return
        End If
        Dim accessDataTable As DataTable = accessDataSet.Tables(tableName)
        Dim CommandText As String = "SELECT * FROM " + tableName + " Where [" + cBoxParamets.Text + "] = "
        Try
            Dim i As Integer
            For i = 0 To (accessDataSet.Tables(tableName).Columns.Count) - 1
                If cBoxParamets.Text = cBoxParamets.Items(i).ToString() Then
                    index = i
                End If
                checkedMenu(i) = "None"
            Next i

            If accessDataTable.Columns(index).DataType.ToString() = "System.Byte[]" Then
                MessageBox.Show("This DataType Cannot Be Filtered", "Unable To Do Filter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            If accessDataTable.Columns(index).DataType.ToString() = "System.String" Then
                CommandText += " '" + searchTxt.Text + "'"
            Else
                If accessDataTable.Columns(index).DataType.ToString() = "System.DateTime" Then
                    CommandText += " #" + searchTxt.Text + "#"
                Else
                    CommandText += searchTxt.Text
                End If
            End If
            loadData(CommandText)
            'lblSelectString.Text = "Select Command = " + CommandText
        Catch
        End Try
    End Sub 'find_the_data

#End Region

#Region "seachTxt events"
    'The user clicked enter instead of buttonFind
    'should give same affect
    Private Sub searchTxt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles searchTxt.KeyDown
        If e.KeyCode = Keys.Enter Then
            find_the_data()
        End If
    End Sub 'searchTxt_KeyDown

    'Dynamic Search with text change
    Private Sub searchTxt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles searchTxt.TextChanged
        Dim index As Integer = 0
        If searchTxt.Text = "" Or dGrid.CurrentRowIndex = -1 Then
            btnPrint.Enabled = False
        Else
            btnPrint.Enabled = True
        End If

        If dGrid.CurrentRowIndex = -1 Then
            btnPrint.Enabled = False
        End If
        If searchTxt.Text = "" Then
            'lblSelectString.Text = "Select Command = Select * From " + tableName
            loadData(("Select * From " + tableName))
            Return
        End If
        Dim accessDataTable As DataTable = accessDataSet.Tables(tableName)
        Dim CommandText As String = "SELECT * FROM " + tableName + " Where [" + cBoxParamets.Text + "] Like "
        DB_Query = "SELECT * FROM " + tableName + " Where [" + cBoxParamets.Text + "]="
        Try
            Dim i As Integer
            For i = 0 To (accessDataSet.Tables(tableName).Columns.Count) - 1
                If cBoxParamets.Text = cBoxParamets.Items(i).ToString() Then
                    index = i
                End If
                checkedMenu(i) = "None"
            Next i

            If accessDataTable.Columns(index).DataType.ToString() = "System.Byte[]" Then
                MessageBox.Show("This DataType Cannot Be Filtered", "Unable To Do Filter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            CommandText += " '" + searchTxt.Text + "%'"
            DB_Query = CommandText '+= searchTxt.Text
            loadData(CommandText)
            'lblSelectString.Text = "Select Command = " + CommandText
        Catch
        End Try

        If searchTxt.Text <> "" And dGrid.CurrentRowIndex = -1 Then
            btnPrint.Enabled = False
        End If
    End Sub 'searchTxt_TextChanged

#End Region

#Region "search Form events"
    'From Resize routine
    'Private Sub frmSearch_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    '    dGrid.Height = Me.Height - 135
    'End Sub 'frmSearch_Resize


    'Form Closing Routine
    Private Sub frmSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        MFRM.Enabled = True
    End Sub 'frmSearch_Closing
#End Region

#Region "Button event"
    'This routine removes all the filters
    'and displays all the data
    Private Sub btnRestore_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRestore.Click
        Dim i As Integer
        For i = 0 To (accessDataSet.Tables(tableName).Columns.Count) - 1
            checkedMenu(i) = "None"
        Next i
        loadData(("Select * From " + tableName))
        'lblSelectString.Text = "Select Command = Select * From " + tableName
    End Sub 'btnRestore_Click

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'DB_Query = "select * from " & DB_TableName & " where ID=1"
        Dim objfrmPrint As New frmPrint(Me, DB_FileName, DB_TableName, DB_Query)
        objfrmPrint.Show()
    End Sub
#End Region

End Class