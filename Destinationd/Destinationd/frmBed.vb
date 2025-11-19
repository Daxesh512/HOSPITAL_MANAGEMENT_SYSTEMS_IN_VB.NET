Imports System.Data.OleDb

Public Class frmBed
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents txtGeneralWard As System.Windows.Forms.TextBox
    Friend WithEvents txtSpecialWard As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.txtGeneralWard = New System.Windows.Forms.TextBox()
        Me.txtSpecialWard = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOK, Me.txtGeneralWard, Me.txtSpecialWard, Me.Label2, Me.Label1})
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(352, 206)
        Me.Panel1.TabIndex = 0
        '
        'btnOK
        '
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnOK.Location = New System.Drawing.Point(134, 140)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(72, 32)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        '
        'txtGeneralWard
        '
        Me.txtGeneralWard.Location = New System.Drawing.Point(214, 76)
        Me.txtGeneralWard.Name = "txtGeneralWard"
        Me.txtGeneralWard.TabIndex = 2
        Me.txtGeneralWard.Text = ""
        '
        'txtSpecialWard
        '
        Me.txtSpecialWard.Location = New System.Drawing.Point(214, 28)
        Me.txtSpecialWard.Name = "txtSpecialWard"
        Me.txtSpecialWard.TabIndex = 1
        Me.txtSpecialWard.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(38, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(160, 24)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Enter beds  for General ward "
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(38, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 32)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Enter beds  for Special ward "
        '
        'frmBed
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(352, 206)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBed"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Bed Information"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "All button code"

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If txtSpecialWard.Text = "" Or txtGeneralWard.Text = "" Then
            MessageBox.Show("You should fill both fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Else
            Dim save_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            Dim save_Comstr As String
            Dim save_Con As OleDbConnection
            Dim save_Com As OleDbCommand
            Dim save_DA As OleDbDataAdapter
            Dim save_DS As DataSet
            Dim save_DT As DataTable

            save_Comstr = "insert into flag(setflag)values" & _
            ("(" + "'" + "1" + "'" + ")")

            save_Con = New OleDbConnection(save_Constr)
            save_Com = New OleDbCommand(save_Comstr, save_Con)

            save_DA = New OleDbDataAdapter(save_Comstr, save_Constr)
            Dim save_builder As New OleDbCommandBuilder(save_DA)

            save_DS = New DataSet()

            Try
                save_DA.Fill(save_DS, "hm")
                MsgBox("Flag set successfully.", MsgBoxStyle.Information, "Register Patient")
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
            save_Con.Close()

            'code for filling the General bed in Generalward table
            Dim saveGeneralBed_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            Dim saveGeneralBed_Comstr As String
            Dim saveGeneralBed_Con As OleDbConnection
            Dim saveGeneralBed_Com As OleDbCommand
            Dim saveGeneralBed_DA As OleDbDataAdapter
            Dim saveGeneralBed_DS As DataSet
            Dim saveGeneralBed_DT As DataTable

            Dim General_begining As Integer
            Dim General_ending As Integer = txtGeneralWard.Text

            For General_begining = 1 To General_ending
                Me.Cursor = Cursors.WaitCursor

                txtGeneralWard.Text = General_begining
                saveGeneralBed_Comstr = "insert into generalward(bedno)values" & _
                          ("(" + "'" + txtGeneralWard.Text + "'" + ")")

                saveGeneralBed_Con = New OleDbConnection(saveGeneralBed_Constr)
                saveGeneralBed_Com = New OleDbCommand(saveGeneralBed_Comstr, saveGeneralBed_Con)

                saveGeneralBed_DA = New OleDbDataAdapter(saveGeneralBed_Comstr, saveGeneralBed_Constr)
                Dim saveGeneralBed_builder As New OleDbCommandBuilder(saveGeneralBed_DA)

                saveGeneralBed_DS = New DataSet()

                Try
                    saveGeneralBed_DA.Fill(saveGeneralBed_DS, "hm")
                    'MsgBox("Bed capacity filled in setbed record.", MsgBoxStyle.Information, "Register Patient")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                Me.Cursor = Cursors.Arrow
            Next

            'code for filling the General bed in Specialward table
            Dim saveSpecialBed_Constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            Dim saveSpecialBed_Comstr As String
            Dim saveSpecialBed_Con As OleDbConnection
            Dim saveSpecialBed_Com As OleDbCommand
            Dim saveSpecialBed_DA As OleDbDataAdapter
            Dim saveSpecialBed_DS As DataSet
            Dim saveSpecialBed_DT As DataTable

            Dim Special_begining As Integer
            Dim Special_ending As Integer = txtSpecialWard.Text

            For Special_begining = 1 To Special_ending
                Me.Cursor = Cursors.WaitCursor

                txtSpecialWard.Text = Special_begining
                saveSpecialBed_Comstr = "insert into Specialward(bedno)values" & _
                          ("(" + "'" + txtSpecialWard.Text + "'" + ")")

                saveSpecialBed_Con = New OleDbConnection(saveSpecialBed_Constr)
                saveSpecialBed_Com = New OleDbCommand(saveSpecialBed_Comstr, saveSpecialBed_Con)

                saveSpecialBed_DA = New OleDbDataAdapter(saveSpecialBed_Comstr, saveSpecialBed_Constr)
                Dim saveSpecialBed_builder As New OleDbCommandBuilder(saveSpecialBed_DA)

                saveSpecialBed_DS = New DataSet()

                Try
                    saveSpecialBed_DA.Fill(saveSpecialBed_DS, "hm")
                    'MsgBox("Bed capacity filled in setbed record.", MsgBoxStyle.Information, "Register Patient")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                Me.Cursor = Cursors.Arrow
            Next
            Me.Hide()
            Dim objfrmMain As New frmMain()
            objfrmMain.Show()
        End If
    End Sub
#End Region

End Class
