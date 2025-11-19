Imports System.Web.Mail

Public Class frmSenEMail
    Inherits System.Windows.Forms.Form

#Region "Variables"
    Dim obj As System.Web.Mail.SmtpMail     ' Variable which will send the mail
    Dim Attachment As System.Web.Mail.MailAttachment    'Variable to store the attachments
    Dim Mailmsg As New System.Web.Mail.MailMessage()    'Variable to create the message to send

    Public Shared count_file As Integer
#End Region

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtFromDisplayName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtTo As System.Windows.Forms.TextBox
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lstAttachment As System.Windows.Forms.ListBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnSend As System.Windows.Forms.Button
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSenEMail))
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFromDisplayName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtTo = New System.Windows.Forms.TextBox()
        Me.OFD = New System.Windows.Forms.OpenFileDialog()
        Me.lstAttachment = New System.Windows.Forms.ListBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtSubject = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(144, 64)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(232, 20)
        Me.txtFrom.TabIndex = 3
        Me.txtFrom.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "From"
        '
        'txtFromDisplayName
        '
        Me.txtFromDisplayName.Location = New System.Drawing.Point(144, 96)
        Me.txtFromDisplayName.Name = "txtFromDisplayName"
        Me.txtFromDisplayName.Size = New System.Drawing.Size(232, 20)
        Me.txtFromDisplayName.TabIndex = 5
        Me.txtFromDisplayName.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 32)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "From Display name"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(32, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 32)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "From Display name"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(24, 128)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 16)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "To"
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(144, 128)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(232, 20)
        Me.txtTo.TabIndex = 7
        Me.txtTo.Text = ""
        '
        'OFD
        '
        Me.OFD.DefaultExt = "*.*"
        Me.OFD.InitialDirectory = "c:\"
        Me.OFD.Multiselect = True
        '
        'lstAttachment
        '
        Me.lstAttachment.Location = New System.Drawing.Point(144, 192)
        Me.lstAttachment.Name = "lstAttachment"
        Me.lstAttachment.Size = New System.Drawing.Size(232, 82)
        Me.lstAttachment.TabIndex = 10
        '
        'btnAdd
        '
        Me.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAdd.Location = New System.Drawing.Point(392, 192)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(120, 32)
        Me.btnAdd.TabIndex = 11
        Me.btnAdd.Text = "Add attachment"
        '
        'btnRemove
        '
        Me.btnRemove.Enabled = False
        Me.btnRemove.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnRemove.Location = New System.Drawing.Point(392, 240)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(120, 32)
        Me.btnRemove.TabIndex = 12
        Me.btnRemove.Text = "Remove attachment"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 192)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 32)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Attachments"
        '
        'txtSubject
        '
        Me.txtSubject.Location = New System.Drawing.Point(144, 288)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(232, 20)
        Me.txtSubject.TabIndex = 15
        Me.txtSubject.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(24, 288)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(104, 16)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "Subject"
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(144, 320)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtMessage.Size = New System.Drawing.Size(240, 144)
        Me.txtMessage.TabIndex = 17
        Me.txtMessage.Text = ""
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(24, 320)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(104, 16)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "Message"
        '
        'btnSend
        '
        Me.btnSend.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSend.Location = New System.Drawing.Point(400, 424)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(104, 32)
        Me.btnSend.TabIndex = 20
        Me.btnSend.Text = "Send"
        '
        'frmSenEMail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(560, 494)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnSend, Me.txtMessage, Me.Label10, Me.txtSubject, Me.Label9, Me.Label8, Me.btnRemove, Me.btnAdd, Me.lstAttachment, Me.Label7, Me.txtTo, Me.txtFromDisplayName, Me.Label3, Me.txtFrom, Me.Label2, Me.Label4})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSenEMail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Send email"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "All button event"
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        'Show open dialogue box to select the files to attach
        Dim Counter As Integer
        OFD.CheckFileExists = True
        OFD.Title = "Select file(s) to attach"
        OFD.ShowDialog()
        For Counter = 0 To UBound(OFD.FileNames)
            lstAttachment.Items.Add(OFD.FileNames(Counter))
            count_file = count_file + 1
            btnRemove.Enabled = True
            If count_file > 2 Then
                btnAdd.Enabled = False
            End If
        Next
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        'Remove the attachments
        If lstAttachment.SelectedIndex > -1 Then
            lstAttachment.Items.RemoveAt(lstAttachment.SelectedIndex)
            count_file = count_file - 1

            If count_file = 0 Then
                btnAdd.Enabled = True
                btnRemove.Enabled = False
            End If
        End If
    End Sub

    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend.Click

        Dim o As New clsCheckInternetConnection()
        Dim flag As Boolean
        flag = o.IsConnected()
        If flag = False Then
            MessageBox.Show("Internet connection is not available to your computer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If


        Dim Counter As Integer

        'Validate the data
        'If txtSMTPServer.Text = "" Then
        '    MsgBox("Enter the SMTP server info ...!!!", MsgBoxStyle.Information, "Send Email")
        '    Exit Sub
        'End If

        If txtFrom.Text = "" Then
            MsgBox("Enter the From email address.", MsgBoxStyle.Information, "Send Email")
            Exit Sub
        End If

        If txtTo.Text = "" Then
            MsgBox("Enter the Recipient email address.", MsgBoxStyle.Information, "Send Email")
            Exit Sub
        End If

        If txtSubject.Text = "" Then
            'MsgBox("Enter the Email subject.", MsgBoxStyle.Information, "Send Email")
            'Exit Sub
            'txtSubject.Text = "None"
        End If

        Dim bFromID, bToID As Boolean
        bFromID = ChkValidEmail(txtFrom.Text)
        If bFromID = False Then
            MessageBox.Show("Invalid Email ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        bToID = ChkValidEmail(txtTo.Text)
        If bToID = False Then
            MessageBox.Show("Invalid Email ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        'Set the properties

        'obj.SmtpServer = txtSMTPServer.Text
        'Multiple recepients can be specified using ; as the delimeter
        'nileshschaudhari1410@gmail.com
        Mailmsg.To = txtTo.Text
        Mailmsg.From = "\" & txtFromDisplayName.Text & "\ <" & txtFrom.Text & ">"

        'Specify the body format
        'If chkFormat.Checked = True Then
        'Mailmsg.BodyFormat = MailFormat.Html   'Send the mail in HTML Format
        'Else
        Mailmsg.BodyFormat = MailFormat.Text
        'End If

        'If you want you can add a reply to header 
        'Mailmsg.Headers.Add("Reply-To", "Manoj@geinetech.net")
        'custom headersare added like this
        'Mailmsg.Headers.Add("Manoj", "TestHeader")
        Mailmsg.Subject = txtSubject.Text
        For Counter = 0 To lstAttachment.Items.Count - 1
            Attachment = New MailAttachment(lstAttachment.Items(Counter))
            Mailmsg.Attachments.Add(Attachment)
        Next

        Mailmsg.Body = txtMessage.Text
        'obj.Send(Mailmsg)
        Try
            SmtpMail.Send(Mailmsg)
            MessageBox.Show("Email send successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)	
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Code for validating email ID"

    'Public Function ChkValidEmail(ByVal Value As String, Optional ByVal MaxLength As Integer = 255, Optional ByVal IsRequired As Boolean = True) As Boolean
    '    If Value Is Nothing OrElse Value.Length = 0 Then
    '        Return Not IsRequired
    '    ElseIf Value.Length > MaxLength Then
    '        Return False
    '    End If
    '    If Not System.Text.RegularExpressions.Regex.IsMatch(Value, "^[-A-Za-z0-9_@.]+$") Then
    '        MsgBox(" Invalid character in the e-mail address. Only(A-Z,a-z,_,0-9,@,.)is valid. ")
    '        Return False
    '    End If
    '    ' search the @ char       
    '    Dim i As Integer = Value.IndexOf("@"c)
    '    ' there must be at least three chars after the @      
    '    If i <= 0 Or i >= Value.Length - 3 Then
    '        MsgBox("INVALID E-Mail: '@' missing.")
    '        Exit Function
    '    End If
    '    ' ensure there is only one @ char      
    '    If Value.IndexOf("@"c, i + 1) >= 0 Then
    '        MsgBox(" INVALID E-Mail: '@'Cann't be allowed twice.")
    '        Exit Function
    '    End If
    '    ' check that the domain portion contains at least one dot      
    '    Dim j As Integer = Value.LastIndexOf("."c)
    '    ' it can't be before or immediately after the @ char    
    '    If j < 0 Or j <= i + 1 Then
    '        MsgBox(" INVALID E-Mail: '.' missing. ")
    '        Exit Function
    '    End If
    '    ' if we get here the address if validated      
    '    'MsgBox(" Congratulations Operator:- Entered E-Mail Address is valid. ")
    '    Return True
    'End Function

    Function ChkValidEmail(ByVal Value As String, Optional ByVal MaxLength As Integer = 255, Optional ByVal IsRequired As Boolean = True) As Boolean
        If Value Is Nothing OrElse Value.Length = 0 Then
            ' rule out the null string case        
            Return Not IsRequired
        ElseIf Value.Length > MaxLength Then
            ' rule out values that are longer than allowed        
            Return False
        End If
        ' search invalid chars    
        If Not System.Text.RegularExpressions.Regex.IsMatch(Value, _
        "^[-A-Za-z0-9_@.]+$") Then Return False
        ' search the @ char    
        Dim i As Integer = Value.IndexOf("@"c)
        ' there must be at least three chars after the @    
        If i <= 0 Or i >= Value.Length - 3 Then Return False
        ' ensure there is only one @ char    
        If Value.IndexOf("@"c, i + 1) >= 0 Then Return False
        ' check that the domain portion contains at least one dot   
        Dim j As Integer = Value.LastIndexOf("."c)
        ' it can't be before or immediately after the @ char   
        If j < 0 Or j <= i + 1 Then Return False
        ' if we get here the address if validated    
        Return True
    End Function

#End Region

End Class
