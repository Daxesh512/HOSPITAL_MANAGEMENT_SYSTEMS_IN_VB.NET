'24,32
'296, 384
'640, 478

Public Class frmHelp
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
    Friend WithEvents PB_Help As System.Windows.Forms.PictureBox
    Friend WithEvents txt_AddRecords As System.Windows.Forms.TextBox
    Friend WithEvents txt_print As System.Windows.Forms.TextBox
    Friend WithEvents txt_Search As System.Windows.Forms.TextBox
    Friend WithEvents txt_Export As System.Windows.Forms.TextBox
    Friend WithEvents txt_Email As System.Windows.Forms.TextBox
    Friend WithEvents txt_Calculator As System.Windows.Forms.TextBox
    Friend WithEvents txt_Notepad As System.Windows.Forms.TextBox
    Friend WithEvents GB_HelpIndex As System.Windows.Forms.GroupBox
    Friend WithEvents ll_Notepad As System.Windows.Forms.LinkLabel
    Friend WithEvents ll_Calculator As System.Windows.Forms.LinkLabel
    Friend WithEvents ll_Email As System.Windows.Forms.LinkLabel
    Friend WithEvents ll_Export As System.Windows.Forms.LinkLabel
    Friend WithEvents ll_Search As System.Windows.Forms.LinkLabel
    Friend WithEvents ll_Print As System.Windows.Forms.LinkLabel
    Friend WithEvents ll_AddRecords As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHelp))
        Me.PB_Help = New System.Windows.Forms.PictureBox()
        Me.txt_AddRecords = New System.Windows.Forms.TextBox()
        Me.txt_print = New System.Windows.Forms.TextBox()
        Me.txt_Search = New System.Windows.Forms.TextBox()
        Me.txt_Export = New System.Windows.Forms.TextBox()
        Me.txt_Email = New System.Windows.Forms.TextBox()
        Me.txt_Calculator = New System.Windows.Forms.TextBox()
        Me.txt_Notepad = New System.Windows.Forms.TextBox()
        Me.GB_HelpIndex = New System.Windows.Forms.GroupBox()
        Me.ll_Notepad = New System.Windows.Forms.LinkLabel()
        Me.ll_Calculator = New System.Windows.Forms.LinkLabel()
        Me.ll_Email = New System.Windows.Forms.LinkLabel()
        Me.ll_Export = New System.Windows.Forms.LinkLabel()
        Me.ll_Search = New System.Windows.Forms.LinkLabel()
        Me.ll_Print = New System.Windows.Forms.LinkLabel()
        Me.ll_AddRecords = New System.Windows.Forms.LinkLabel()
        Me.GB_HelpIndex.SuspendLayout()
        Me.SuspendLayout()
        '
        'PB_Help
        '
        Me.PB_Help.Image = CType(resources.GetObject("PB_Help.Image"), System.Drawing.Bitmap)
        Me.PB_Help.Location = New System.Drawing.Point(432, 40)
        Me.PB_Help.Name = "PB_Help"
        Me.PB_Help.Size = New System.Drawing.Size(136, 136)
        Me.PB_Help.TabIndex = 0
        Me.PB_Help.TabStop = False
        '
        'txt_AddRecords
        '
        Me.txt_AddRecords.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_AddRecords.Location = New System.Drawing.Point(24, 32)
        Me.txt_AddRecords.Multiline = True
        Me.txt_AddRecords.Name = "txt_AddRecords"
        Me.txt_AddRecords.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txt_AddRecords.Size = New System.Drawing.Size(296, 384)
        Me.txt_AddRecords.TabIndex = 9
        Me.txt_AddRecords.Text = "1. Clinical Checkup : In this record only      regular patient's record is added." & _
        "" & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "2. Admit Patient     : In this record only       admitted patient's record i" & _
        "s added." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "3. Doctor Record   : In this record only       doctor's record is ad" & _
        "ded." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "4. Nurse/Wardboy Record  : In this             record only nurse's or wa" & _
        "rdboy's             record is added." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "5. Discharge Record  : This record is   " & _
        "       automatically filled when you click on      Discharge Button in Admit Pat" & _
        "ient          form." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "6. Bill Record   : This record is                    auto" & _
        "matically filled when you click on      Discharge Button in Admit Patient       " & _
        "   form." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "* Note1 : All fields are neccessay to fill     in Clinical Checkup, " & _
        "Admit Patient          form except Remark field. All fields        are neccessay" & _
        " to fill in  Doctor, Nurse    or Wardboy form except Phone field       in Nurse " & _
        "or Wardboy form." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "* Note2 : Remark field will be                    considered" & _
        " as 'No remark' if you             leave this field as blank."
        '
        'txt_print
        '
        Me.txt_print.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_print.Location = New System.Drawing.Point(336, 456)
        Me.txt_print.Multiline = True
        Me.txt_print.Name = "txt_print"
        Me.txt_print.Size = New System.Drawing.Size(280, 32)
        Me.txt_print.TabIndex = 17
        Me.txt_print.Text = "1. List         : Prints the list of record." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "2. Current  : Print current or pa" & _
        "rticular        record through search form."
        '
        'txt_Search
        '
        Me.txt_Search.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Search.Location = New System.Drawing.Point(336, 408)
        Me.txt_Search.Multiline = True
        Me.txt_Search.Name = "txt_Search"
        Me.txt_Search.Size = New System.Drawing.Size(280, 40)
        Me.txt_Search.TabIndex = 18
        Me.txt_Search.Text = "By using this functionality you can find the records. There is no headche for rem" & _
        "embering patient name or other else. You can search record just by entering the " & _
        "patinet ID or name or admit date or address or age etc."
        '
        'txt_Export
        '
        Me.txt_Export.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Export.Location = New System.Drawing.Point(40, 488)
        Me.txt_Export.Multiline = True
        Me.txt_Export.Name = "txt_Export"
        Me.txt_Export.Size = New System.Drawing.Size(280, 40)
        Me.txt_Export.TabIndex = 19
        Me.txt_Export.Text = "By using this functionality you can export the list of patients , doctors , nurse" & _
        " or wardboy, discharge , bill  to excel."
        '
        'txt_Email
        '
        Me.txt_Email.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Email.Location = New System.Drawing.Point(40, 440)
        Me.txt_Email.Multiline = True
        Me.txt_Email.Name = "txt_Email"
        Me.txt_Email.Size = New System.Drawing.Size(280, 40)
        Me.txt_Email.TabIndex = 20
        Me.txt_Email.Text = "By using this functionality you can send email to anyone. It also provides the fu" & _
        "nctionality of attachment so that you can able to send the file attchament like " & _
        "list of patients , doctors etc." & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10) & "* Note : For sending Email through this form," & _
        " you have must install the IIS on your computer."
        '
        'txt_Calculator
        '
        Me.txt_Calculator.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Calculator.Location = New System.Drawing.Point(336, 496)
        Me.txt_Calculator.Multiline = True
        Me.txt_Calculator.Name = "txt_Calculator"
        Me.txt_Calculator.Size = New System.Drawing.Size(280, 24)
        Me.txt_Calculator.TabIndex = 21
        Me.txt_Calculator.Text = "It helps in calculations like calculating the the total amount of patient's bill " & _
        "etc."
        '
        'txt_Notepad
        '
        Me.txt_Notepad.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Notepad.Location = New System.Drawing.Point(336, 528)
        Me.txt_Notepad.Multiline = True
        Me.txt_Notepad.Name = "txt_Notepad"
        Me.txt_Notepad.Size = New System.Drawing.Size(280, 24)
        Me.txt_Notepad.TabIndex = 22
        Me.txt_Notepad.Text = "It helps to write notes."
        '
        'GB_HelpIndex
        '
        Me.GB_HelpIndex.Controls.AddRange(New System.Windows.Forms.Control() {Me.ll_Notepad, Me.ll_Calculator, Me.ll_Email, Me.ll_Export, Me.ll_Search, Me.ll_Print, Me.ll_AddRecords})
        Me.GB_HelpIndex.Location = New System.Drawing.Point(416, 192)
        Me.GB_HelpIndex.Name = "GB_HelpIndex"
        Me.GB_HelpIndex.Size = New System.Drawing.Size(168, 208)
        Me.GB_HelpIndex.TabIndex = 23
        Me.GB_HelpIndex.TabStop = False
        Me.GB_HelpIndex.Text = "Index"
        '
        'll_Notepad
        '
        Me.ll_Notepad.Location = New System.Drawing.Point(40, 176)
        Me.ll_Notepad.Name = "ll_Notepad"
        Me.ll_Notepad.Size = New System.Drawing.Size(88, 24)
        Me.ll_Notepad.TabIndex = 15
        Me.ll_Notepad.TabStop = True
        Me.ll_Notepad.Text = "Notepad"
        '
        'll_Calculator
        '
        Me.ll_Calculator.Location = New System.Drawing.Point(40, 152)
        Me.ll_Calculator.Name = "ll_Calculator"
        Me.ll_Calculator.Size = New System.Drawing.Size(88, 24)
        Me.ll_Calculator.TabIndex = 14
        Me.ll_Calculator.TabStop = True
        Me.ll_Calculator.Text = "Calculator"
        '
        'll_Email
        '
        Me.ll_Email.Location = New System.Drawing.Point(40, 128)
        Me.ll_Email.Name = "ll_Email"
        Me.ll_Email.Size = New System.Drawing.Size(88, 24)
        Me.ll_Email.TabIndex = 13
        Me.ll_Email.TabStop = True
        Me.ll_Email.Text = "Email"
        '
        'll_Export
        '
        Me.ll_Export.Location = New System.Drawing.Point(40, 104)
        Me.ll_Export.Name = "ll_Export"
        Me.ll_Export.Size = New System.Drawing.Size(88, 24)
        Me.ll_Export.TabIndex = 12
        Me.ll_Export.TabStop = True
        Me.ll_Export.Text = "Export"
        '
        'll_Search
        '
        Me.ll_Search.Location = New System.Drawing.Point(40, 80)
        Me.ll_Search.Name = "ll_Search"
        Me.ll_Search.Size = New System.Drawing.Size(88, 24)
        Me.ll_Search.TabIndex = 11
        Me.ll_Search.TabStop = True
        Me.ll_Search.Text = "Search"
        '
        'll_Print
        '
        Me.ll_Print.Location = New System.Drawing.Point(40, 56)
        Me.ll_Print.Name = "ll_Print"
        Me.ll_Print.Size = New System.Drawing.Size(88, 16)
        Me.ll_Print.TabIndex = 10
        Me.ll_Print.TabStop = True
        Me.ll_Print.Text = "Print"
        '
        'll_AddRecords
        '
        Me.ll_AddRecords.Location = New System.Drawing.Point(40, 32)
        Me.ll_AddRecords.Name = "ll_AddRecords"
        Me.ll_AddRecords.Size = New System.Drawing.Size(88, 16)
        Me.ll_AddRecords.TabIndex = 9
        Me.ll_AddRecords.TabStop = True
        Me.ll_AddRecords.Text = "Add Record"
        '
        'frmHelp
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(634, 446)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GB_HelpIndex, Me.txt_Notepad, Me.txt_Calculator, Me.txt_Email, Me.txt_Export, Me.txt_Search, Me.txt_print, Me.txt_AddRecords, Me.PB_Help})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmHelp"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Help"
        Me.GB_HelpIndex.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Form load event"

    Private Sub frmHelp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txt_AddRecords.Visible = True
        txt_print.Visible = False
        txt_Search.Visible = False
        txt_Export.Visible = False
        txt_Email.Visible = False
        txt_Calculator.Visible = False
        txt_Notepad.Visible = False
    End Sub

#End Region

#Region "All link label events"

    Private Sub ll_AddRecords_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ll_AddRecords.LinkClicked
        txt_AddRecords.Visible = True
        txt_print.Visible = False
        txt_Search.Visible = False
        txt_Export.Visible = False
        txt_Email.Visible = False
        txt_Calculator.Visible = False
        txt_Notepad.Visible = False
    End Sub

    Private Sub ll_Print_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ll_Print.LinkClicked
        txt_AddRecords.Visible = False
        txt_print.Visible = True
        txt_Search.Visible = False
        txt_Export.Visible = False
        txt_Email.Visible = False
        txt_Calculator.Visible = False
        txt_Notepad.Visible = False

        txt_print.ScrollBars = ScrollBars.Both
        txt_print.Location = New Point(24, 32)
        txt_print.Size = New Size(296, 384)
    End Sub

    Private Sub ll_Search_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ll_Search.LinkClicked
        txt_AddRecords.Visible = False
        txt_print.Visible = False
        txt_Search.Visible = True
        txt_Export.Visible = False
        txt_Email.Visible = False
        txt_Calculator.Visible = False
        txt_Notepad.Visible = False

        txt_Search.ScrollBars = ScrollBars.Both
        txt_Search.Location = New Point(24, 32)
        txt_Search.Size = New Size(296, 384)
    End Sub

    Private Sub ll_Export_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ll_Export.LinkClicked
        txt_AddRecords.Visible = False
        txt_print.Visible = False
        txt_Search.Visible = False
        txt_Export.Visible = True
        txt_Email.Visible = False
        txt_Calculator.Visible = False
        txt_Notepad.Visible = False

        txt_Export.ScrollBars = ScrollBars.Both
        txt_Export.Location = New Point(24, 32)
        txt_Export.Size = New Size(296, 384)
    End Sub

    Private Sub ll_Email_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ll_Email.LinkClicked
        txt_AddRecords.Visible = False
        txt_print.Visible = False
        txt_Search.Visible = False
        txt_Export.Visible = False
        txt_Email.Visible = True
        txt_Calculator.Visible = False
        txt_Notepad.Visible = False

        txt_Email.ScrollBars = ScrollBars.Both
        txt_Email.Location = New Point(24, 32)
        txt_Email.Size = New Size(296, 384)
    End Sub

    Private Sub ll_Calculator_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ll_Calculator.LinkClicked
        txt_AddRecords.Visible = False
        txt_print.Visible = False
        txt_Search.Visible = False
        txt_Export.Visible = False
        txt_Email.Visible = False
        txt_Calculator.Visible = True
        txt_Notepad.Visible = False

        txt_Calculator.ScrollBars = ScrollBars.Both
        txt_Calculator.Location = New Point(24, 32)
        txt_Calculator.Size = New Size(296, 384)
    End Sub

    Private Sub ll_Notepad_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles ll_Notepad.LinkClicked
        txt_AddRecords.Visible = False
        txt_print.Visible = False
        txt_Search.Visible = False
        txt_Export.Visible = False
        txt_Email.Visible = False
        txt_Calculator.Visible = False
        txt_Notepad.Visible = True

        txt_Notepad.ScrollBars = ScrollBars.Both
        txt_Notepad.Location = New Point(24, 32)
        txt_Notepad.Size = New Size(296, 384)
    End Sub

#End Region

#Region "All link label key event"

    Private Sub txt_AddRecords_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_AddRecords.KeyPress
        If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsNumber(e.KeyChar) Or Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = True
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_print_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_print.KeyPress
        If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsNumber(e.KeyChar) Or Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = True
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_Search_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Search.KeyPress
        If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsNumber(e.KeyChar) Or Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = True
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_Export_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Export.KeyPress
        If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsNumber(e.KeyChar) Or Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = True
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_Email_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Email.KeyPress
        If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsNumber(e.KeyChar) Or Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = True
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_Calculator_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Calculator.KeyPress
        If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsNumber(e.KeyChar) Or Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = True
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txt_Notepad_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_Notepad.KeyPress
        If Char.IsLetterOrDigit(e.KeyChar) Or Char.IsNumber(e.KeyChar) Or Char.IsLetter(e.KeyChar) = False Then
            If e.KeyChar = CChar(ChrW(Keys.Back)) Or e.KeyChar = CChar(ChrW(Keys.Space)) Then
                e.Handled = True
            Else
                e.Handled = True
            End If
        End If
    End Sub
#End Region

End Class
