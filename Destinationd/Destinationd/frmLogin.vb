Imports System.Data.OleDb

Public Class frmLogin
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txt_Username As System.Windows.Forms.TextBox
    Friend WithEvents txt_Password As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents SB_Login As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLogin))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txt_Username = New System.Windows.Forms.TextBox()
        Me.txt_Password = New System.Windows.Forms.TextBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SB_Login = New System.Windows.Forms.StatusBar()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Teal
        Me.Label1.Location = New System.Drawing.Point(184, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 32)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "User Name"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Bitmap)
        Me.PictureBox1.Location = New System.Drawing.Point(16, 48)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(144, 144)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Teal
        Me.Label2.Location = New System.Drawing.Point(184, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password"
        '
        'txt_Username
        '
        Me.txt_Username.Location = New System.Drawing.Point(320, 64)
        Me.txt_Username.Name = "txt_Username"
        Me.txt_Username.Size = New System.Drawing.Size(128, 20)
        Me.txt_Username.TabIndex = 1
        Me.txt_Username.Text = ""
        '
        'txt_Password
        '
        Me.txt_Password.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Password.Location = New System.Drawing.Point(320, 120)
        Me.txt_Password.Name = "txt_Password"
        Me.txt_Password.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txt_Password.Size = New System.Drawing.Size(128, 22)
        Me.txt_Password.TabIndex = 2
        Me.txt_Password.Text = ""
        '
        'btnOK
        '
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnOK.Location = New System.Drawing.Point(240, 192)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(72, 32)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancel.Location = New System.Drawing.Point(360, 192)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(72, 32)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'SB_Login
        '
        Me.SB_Login.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold)
        Me.SB_Login.Location = New System.Drawing.Point(0, 248)
        Me.SB_Login.Name = "SB_Login"
        Me.SB_Login.Size = New System.Drawing.Size(474, 32)
        Me.SB_Login.TabIndex = 7
        Me.SB_Login.Text = "Attempt : "
        '
        'frmLogin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(474, 280)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.SB_Login, Me.btnCancel, Me.btnOK, Me.txt_Password, Me.txt_Username, Me.Label2, Me.PictureBox1, Me.Label1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Variable"
    Public Shared counter As Integer = 3
    Dim str As String = "Attempt : "
#End Region

#Region "All button event"
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If txt_Username.Text = "hospital" And txt_Password.Text = "user" Then
            Me.Hide()

            Dim r0 As String

            Dim Flag_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
            'Dim Fill_Textbox_cmdstr As String = "select Regno,Name,Age,Address,Phone,Checkupdate,Checkuptime,Disease,Bloodgroup,Doctor,Remarks from clinicpatients where Regno=1"
            Dim Flag_cmdstr As String = "select * from flag"
            Dim Flag_con As New OleDbConnection(Flag_constr)
            Dim Flag_com As New OleDbCommand(Flag_cmdstr, Flag_con)
            Flag_con.Open()
            Dim Flag_DR As OleDbDataReader = Flag_com.ExecuteReader
            While Flag_DR.Read
                r0 = Flag_DR(0)
            End While
            Flag_con.Close()

            If r0 = 0 Then
                Dim objfrmBed As New frmBed()
                objfrmBed.Show()
            ElseIf r0 = 1 Then
                Dim objfrmMain As New frmMain()
                objfrmMain.Show()
            End If

        ElseIf txt_Username.Text = "" And txt_Password.Text = "" Then
            MessageBox.Show("Please fill the fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Else
            MessageBox.Show("Invalid login.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            counter = counter - 1
            SB_Login.Text = str & counter
            txt_Username.Text = ""
            txt_Password.Text = ""

            If counter = 0 Then
                btnOK.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
#End Region

#Region "frmLogin load event"
    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SB_Login.Text = str & counter
    End Sub
#End Region

#Region "txt_Password Events"
    Private Sub txt_Password_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_Password.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txt_Username.Text = "hospital" And txt_Password.Text = "user" Then
                Me.Hide()

                Dim r0 As String

                Dim Flag_constr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\hm.mdb"
                'Dim Fill_Textbox_cmdstr As String = "select Regno,Name,Age,Address,Phone,Checkupdate,Checkuptime,Disease,Bloodgroup,Doctor,Remarks from clinicpatients where Regno=1"
                Dim Flag_cmdstr As String = "select * from flag"
                Dim Flag_con As New OleDbConnection(Flag_constr)
                Dim Flag_com As New OleDbCommand(Flag_cmdstr, Flag_con)
                Flag_con.Open()
                Dim Flag_DR As OleDbDataReader = Flag_com.ExecuteReader
                While Flag_DR.Read
                    r0 = Flag_DR(0)
                End While
                Flag_con.Close()

                If r0 = 0 Then
                    Dim objfrmBed As New frmBed()
                    objfrmBed.Show()
                ElseIf r0 = 1 Then
                    Dim objfrmMain As New frmMain()
                    objfrmMain.Show()
                End If

            ElseIf txt_Username.Text = "" And txt_Password.Text = "" Then
                MessageBox.Show("Please fill the fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Else
                MessageBox.Show("Invalid login.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                counter = counter - 1
                SB_Login.Text = str & counter
                txt_Username.Text = ""
                txt_Password.Text = ""

                If counter = 0 Then
                    btnOK.Enabled = False
                End If
            End If
        End If
    End Sub
#End Region

End Class
