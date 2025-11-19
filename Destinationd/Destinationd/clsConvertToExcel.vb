Imports System.Data.OleDb
Imports System.IO

Public Class clsConvertToExcel


#Region "Function for converting table into excel file"

    Sub Load_Excel_Details(ByVal DB_Name As String, ByVal DB_Table As String)

        Dim saveDlg As New SaveFileDialog()

        ''''''''''''''
        saveDlg.Filter = "All Microsoft Office Excel Files | *.xls"
        If saveDlg.ShowDialog = DialogResult.OK Then



            'Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=c:\hm.mdb")
            Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=" + DB_Name)
            Dim ComDset As New DataSet()
            Dim ComDset1 As New DataSet()

            Dim str, filename As String
            Dim col, row As Integer

            'str = "SELECT * from clinicpatients"
            str = "SELECT * from " + DB_Table

            Dim adp As New OleDbDataAdapter(str, conn)
            ComDset.Reset()
            adp.Fill(ComDset, "TTbl")

            If ComDset.Tables.Count < 0 Or ComDset.Tables(0).Rows.Count <= 0 Then
                Exit Sub
            End If

            Dim Excel As Object = CreateObject("Excel.Application")

            If Excel Is Nothing Then
                MessageBox.Show("It appears that Excel is not installed on this machine. This operation requires MS Excel to be installed on this machine.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            With Excel
                .SheetsInNewWorkbook = 1
                .Workbooks.Add()
                .Worksheets(1).Select()

                Dim i As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    .cells(1, i).value = ComDset.Tables(0).Columns(col).ColumnName
                    .cells(1, i).EntireRow.Font.Bold = True
                    i += 1
                Next

                i = 2
                Dim k As Integer = 1
                For col = 0 To ComDset.Tables(0).Columns.Count - 1
                    i = 2
                    For row = 0 To ComDset.Tables(0).Rows.Count - 1
                        .Cells(i, k).Value = ComDset.Tables(0).Rows(row).ItemArray(col)
                        i += 1
                    Next
                    k += 1
                Next

                filename = saveDlg.FileName '"c:\File_Exported" & Format(Now(), "dd-MM-yyyy_hh-mm-ss") & ".xls"
                .ActiveCell.Worksheet.SaveAs(filename)
            End With

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)

            Excel = Nothing
            'MsgBox("Data's are exported to Excel Succesfully", MsgBoxStyle.Information)
            'MessageBox.Show("Data's are exported to Excel Succesfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' The excel is created and opened for insert value. We most close this excel using this system
            Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")

            Dim ii As Process
            For Each ii In pro
                ii.Kill()
            Next
            MessageBox.Show("Data's are exported to Excel Succesfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

#End Region

End Class



