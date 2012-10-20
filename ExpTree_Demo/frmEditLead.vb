
Imports System
Imports ADODB
Imports System.Web.Mail
Imports System.Runtime.InteropServices


Imports System.IO
Imports System.IO.IsolatedStorage
Imports System.Data

Imports ExpTreeLib
Imports ExpTreeLib.CShItem
Imports ExpTreeLib.SystemImageListManager

Imports System.Threading
Public Class frmEditLead
    Inherits System.Windows.Forms.Form
    'avoid Globalization problem-- an empty timevalue
    Dim testTime As New DateTime(1, 1, 1, 0, 0, 0)

    Private LastSelectedCSI As CShItem

    Private Shared Event1 As New ManualResetEvent(True)

    Public clientno, cname, cstatus, amount 'client number and name
    Public leadno, desription, ddate, title
    Public onactivateform As Boolean = False
    Private journalpath As String

    ' Variable which will send the mail
    Dim obj As System.Web.Mail.SmtpMail
    'Variable to store the attachments 
    Dim Attachment As System.Web.Mail.MailAttachment
    'Variable to create the message to send
    Dim Mailmsg As New System.Web.Mail.MailMessage()
    Dim attachedfile As String

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Add any initialization after the InitializeComponent() call
        SystemImageListManager.SetListViewImageList(lv1, True, False)
        SystemImageListManager.SetListViewImageList(lv1, False, False)
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            editleads = False
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
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents tpgedit As System.Windows.Forms.TabPage
    Friend WithEvents tpgsendmail As System.Windows.Forms.TabPage
    Friend WithEvents pnlsendmail As System.Windows.Forms.Panel
    Friend WithEvents btnSendMail As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents lstAttachment As System.Windows.Forms.ListBox
    Friend WithEvents btnMaillAttach As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtMailBody As System.Windows.Forms.RichTextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents txtMailTo As System.Windows.Forms.ComboBox
    Friend WithEvents txtMailSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtMailCC As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents rtbjournal As System.Windows.Forms.RichTextBox
    Friend WithEvents btnsavejournal As System.Windows.Forms.Button
    Friend WithEvents pnlAddLead As System.Windows.Forms.Panel
    Friend WithEvents grpDetails As System.Windows.Forms.GroupBox
    Friend WithEvents dtpsniffed As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateSniffed As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents pnlviewjournal As System.Windows.Forms.Panel
    Friend WithEvents txttitle As System.Windows.Forms.TextBox
    Friend WithEvents lbltitle As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents tpgviewmail As System.Windows.Forms.TabPage
    Friend WithEvents btndeletelead As System.Windows.Forms.Button
    Friend WithEvents lblOpenjournal As System.Windows.Forms.LinkLabel
    Friend WithEvents btnconvert As System.Windows.Forms.Button
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents cmdCTest As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents WebService1 As System.Web.Services.WebService
    Friend WithEvents ExpTree1 As ExpTreeLib.ExpTree
    Friend WithEvents cb1 As System.Windows.Forms.ComboBox
    Friend WithEvents lv1 As System.Windows.Forms.ListView
    Friend WithEvents sbr1 As System.Windows.Forms.StatusBar
    Friend WithEvents txtdepartment As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboProspect As System.Windows.Forms.ComboBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents btndepartments As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEditLead))
        Me.btnExit = New System.Windows.Forms.Button
        Me.tpgedit = New System.Windows.Forms.TabPage
        Me.pnlAddLead = New System.Windows.Forms.Panel
        Me.btnconvert = New System.Windows.Forms.Button
        Me.lbltitle = New System.Windows.Forms.Label
        Me.txttitle = New System.Windows.Forms.TextBox
        Me.grpDetails = New System.Windows.Forms.GroupBox
        Me.btndepartments = New System.Windows.Forms.Button
        Me.cboProspect = New System.Windows.Forms.ComboBox
        Me.txtdepartment = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblOpenjournal = New System.Windows.Forms.LinkLabel
        Me.btndeletelead = New System.Windows.Forms.Button
        Me.lblAmount = New System.Windows.Forms.Label
        Me.txtAmount = New System.Windows.Forms.TextBox
        Me.dtpsniffed = New System.Windows.Forms.DateTimePicker
        Me.lblDateSniffed = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblCompany = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlviewjournal = New System.Windows.Forms.Panel
        Me.btnsavejournal = New System.Windows.Forms.Button
        Me.rtbjournal = New System.Windows.Forms.RichTextBox
        Me.tpgsendmail = New System.Windows.Forms.TabPage
        Me.pnlsendmail = New System.Windows.Forms.Panel
        Me.btnSendMail = New System.Windows.Forms.Button
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.lstAttachment = New System.Windows.Forms.ListBox
        Me.btnMaillAttach = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtMailBody = New System.Windows.Forms.RichTextBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.txtServer = New System.Windows.Forms.TextBox
        Me.lblServer = New System.Windows.Forms.Label
        Me.txtMailTo = New System.Windows.Forms.ComboBox
        Me.txtMailSubject = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtMailCC = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.tpgviewmail = New System.Windows.Forms.TabPage
        Me.sbr1 = New System.Windows.Forms.StatusBar
        Me.lv1 = New System.Windows.Forms.ListView
        Me.cb1 = New System.Windows.Forms.ComboBox
        Me.ExpTree1 = New ExpTreeLib.ExpTree
        Me.cmdRefresh = New System.Windows.Forms.Button
        Me.cmdCTest = New System.Windows.Forms.Button
        Me.cmdExit = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.tpgedit.SuspendLayout()
        Me.pnlAddLead.SuspendLayout()
        Me.grpDetails.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.pnlviewjournal.SuspendLayout()
        Me.tpgsendmail.SuspendLayout()
        Me.pnlsendmail.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.tpgviewmail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Location = New System.Drawing.Point(304, 516)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(120, 20)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "Close"
        '
        'tpgedit
        '
        Me.tpgedit.Controls.Add(Me.pnlAddLead)
        Me.tpgedit.Controls.Add(Me.pnlviewjournal)
        Me.tpgedit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tpgedit.Location = New System.Drawing.Point(4, 23)
        Me.tpgedit.Name = "tpgedit"
        Me.tpgedit.Size = New System.Drawing.Size(416, 485)
        Me.tpgedit.TabIndex = 2
        Me.tpgedit.Text = "Edit Lead"
        '
        'pnlAddLead
        '
        Me.pnlAddLead.Controls.Add(Me.btnconvert)
        Me.pnlAddLead.Controls.Add(Me.lbltitle)
        Me.pnlAddLead.Controls.Add(Me.txttitle)
        Me.pnlAddLead.Controls.Add(Me.grpDetails)
        Me.pnlAddLead.Controls.Add(Me.GroupBox1)
        Me.pnlAddLead.Controls.Add(Me.btnSave)
        Me.pnlAddLead.Location = New System.Drawing.Point(10, 9)
        Me.pnlAddLead.Name = "pnlAddLead"
        Me.pnlAddLead.Size = New System.Drawing.Size(376, 454)
        Me.pnlAddLead.TabIndex = 4
        '
        'btnconvert
        '
        Me.btnconvert.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnconvert.Enabled = False
        Me.btnconvert.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnconvert.Location = New System.Drawing.Point(128, 431)
        Me.btnconvert.Name = "btnconvert"
        Me.btnconvert.Size = New System.Drawing.Size(120, 20)
        Me.btnconvert.TabIndex = 11
        Me.btnconvert.Text = "Convert to job"
        '
        'lbltitle
        '
        Me.lbltitle.Location = New System.Drawing.Point(8, 70)
        Me.lbltitle.Name = "lbltitle"
        Me.lbltitle.Size = New System.Drawing.Size(104, 16)
        Me.lbltitle.TabIndex = 23
        Me.lbltitle.Text = "Lead Title"
        '
        'txttitle
        '
        Me.txttitle.Location = New System.Drawing.Point(120, 67)
        Me.txttitle.Name = "txttitle"
        Me.txttitle.Size = New System.Drawing.Size(248, 20)
        Me.txttitle.TabIndex = 1
        Me.txttitle.Text = ""
        '
        'grpDetails
        '
        Me.grpDetails.Controls.Add(Me.btndepartments)
        Me.grpDetails.Controls.Add(Me.cboProspect)
        Me.grpDetails.Controls.Add(Me.txtdepartment)
        Me.grpDetails.Controls.Add(Me.Label4)
        Me.grpDetails.Controls.Add(Me.lblOpenjournal)
        Me.grpDetails.Controls.Add(Me.btndeletelead)
        Me.grpDetails.Controls.Add(Me.lblAmount)
        Me.grpDetails.Controls.Add(Me.txtAmount)
        Me.grpDetails.Controls.Add(Me.dtpsniffed)
        Me.grpDetails.Controls.Add(Me.lblDateSniffed)
        Me.grpDetails.Controls.Add(Me.lblStatus)
        Me.grpDetails.Controls.Add(Me.txtDesc)
        Me.grpDetails.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpDetails.Location = New System.Drawing.Point(8, 88)
        Me.grpDetails.Name = "grpDetails"
        Me.grpDetails.Size = New System.Drawing.Size(360, 344)
        Me.grpDetails.TabIndex = 2
        Me.grpDetails.TabStop = False
        Me.grpDetails.Text = "Description"
        '
        'btndepartments
        '
        Me.btndepartments.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndepartments.Location = New System.Drawing.Point(112, 264)
        Me.btndepartments.Name = "btndepartments"
        Me.btndepartments.Size = New System.Drawing.Size(32, 20)
        Me.btndepartments.TabIndex = 6
        Me.btndepartments.Tag = "Add or edit departments"
        Me.btndepartments.Text = "A"
        '
        'cboProspect
        '
        Me.cboProspect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProspect.Location = New System.Drawing.Point(112, 238)
        Me.cboProspect.Name = "cboProspect"
        Me.cboProspect.Size = New System.Drawing.Size(240, 22)
        Me.cboProspect.TabIndex = 5
        '
        'txtdepartment
        '
        Me.txtdepartment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.txtdepartment.Items.AddRange(New Object() {"", "Survey", "GI", "RS", "BD"})
        Me.txtdepartment.Location = New System.Drawing.Point(144, 264)
        Me.txtdepartment.Name = "txtdepartment"
        Me.txtdepartment.Size = New System.Drawing.Size(208, 22)
        Me.txtdepartment.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 264)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 16)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "Department"
        '
        'lblOpenjournal
        '
        Me.lblOpenjournal.Location = New System.Drawing.Point(8, 212)
        Me.lblOpenjournal.Name = "lblOpenjournal"
        Me.lblOpenjournal.Size = New System.Drawing.Size(128, 16)
        Me.lblOpenjournal.TabIndex = 6
        Me.lblOpenjournal.TabStop = True
        Me.lblOpenjournal.Text = "Open  Journal"
        '
        'btndeletelead
        '
        Me.btndeletelead.BackColor = System.Drawing.Color.IndianRed
        Me.btndeletelead.Location = New System.Drawing.Point(183, 212)
        Me.btndeletelead.Name = "btndeletelead"
        Me.btndeletelead.Size = New System.Drawing.Size(168, 20)
        Me.btndeletelead.TabIndex = 4
        Me.btndeletelead.Text = "Delete this lead"
        '
        'lblAmount
        '
        Me.lblAmount.Location = New System.Drawing.Point(8, 312)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(104, 16)
        Me.lblAmount.TabIndex = 32
        Me.lblAmount.Text = "Amount"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(112, 312)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(240, 20)
        Me.txtAmount.TabIndex = 9
        Me.txtAmount.Text = ""
        '
        'dtpsniffed
        '
        Me.dtpsniffed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpsniffed.Location = New System.Drawing.Point(112, 288)
        Me.dtpsniffed.Name = "dtpsniffed"
        Me.dtpsniffed.Size = New System.Drawing.Size(240, 20)
        Me.dtpsniffed.TabIndex = 8
        Me.dtpsniffed.Value = New Date(2006, 2, 6, 9, 53, 19, 530)
        '
        'lblDateSniffed
        '
        Me.lblDateSniffed.Location = New System.Drawing.Point(8, 288)
        Me.lblDateSniffed.Name = "lblDateSniffed"
        Me.lblDateSniffed.Size = New System.Drawing.Size(104, 16)
        Me.lblDateSniffed.TabIndex = 28
        Me.lblDateSniffed.Text = "Date"
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(8, 236)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(104, 24)
        Me.lblStatus.TabIndex = 26
        Me.lblStatus.Text = "Status"
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(8, 16)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(344, 192)
        Me.txtDesc.TabIndex = 3
        Me.txtDesc.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblCompany)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(360, 56)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lblCompany
        '
        Me.lblCompany.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lblCompany.Location = New System.Drawing.Point(112, 14)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(240, 32)
        Me.lblCompany.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Company"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSave.Location = New System.Drawing.Point(8, 431)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(120, 20)
        Me.btnSave.TabIndex = 10
        Me.btnSave.Text = "Save"
        '
        'pnlviewjournal
        '
        Me.pnlviewjournal.Controls.Add(Me.btnsavejournal)
        Me.pnlviewjournal.Controls.Add(Me.rtbjournal)
        Me.pnlviewjournal.Location = New System.Drawing.Point(8, 8)
        Me.pnlviewjournal.Name = "pnlviewjournal"
        Me.pnlviewjournal.Size = New System.Drawing.Size(384, 456)
        Me.pnlviewjournal.TabIndex = 0
        '
        'btnsavejournal
        '
        Me.btnsavejournal.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnsavejournal.Location = New System.Drawing.Point(8, 6)
        Me.btnsavejournal.Name = "btnsavejournal"
        Me.btnsavejournal.Size = New System.Drawing.Size(120, 20)
        Me.btnsavejournal.TabIndex = 1
        Me.btnsavejournal.Text = "Save Journal"
        '
        'rtbjournal
        '
        Me.rtbjournal.Location = New System.Drawing.Point(8, 40)
        Me.rtbjournal.Name = "rtbjournal"
        Me.rtbjournal.Size = New System.Drawing.Size(368, 408)
        Me.rtbjournal.TabIndex = 0
        Me.rtbjournal.Text = ""
        '
        'tpgsendmail
        '
        Me.tpgsendmail.Controls.Add(Me.pnlsendmail)
        Me.tpgsendmail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tpgsendmail.Location = New System.Drawing.Point(4, 23)
        Me.tpgsendmail.Name = "tpgsendmail"
        Me.tpgsendmail.Size = New System.Drawing.Size(416, 485)
        Me.tpgsendmail.TabIndex = 0
        Me.tpgsendmail.Text = "Send Mail"
        Me.tpgsendmail.Visible = False
        '
        'pnlsendmail
        '
        Me.pnlsendmail.AutoScroll = True
        Me.pnlsendmail.BackColor = System.Drawing.SystemColors.Control
        Me.pnlsendmail.Controls.Add(Me.btnSendMail)
        Me.pnlsendmail.Controls.Add(Me.GroupBox5)
        Me.pnlsendmail.Controls.Add(Me.GroupBox2)
        Me.pnlsendmail.Controls.Add(Me.GroupBox4)
        Me.pnlsendmail.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlsendmail.ForeColor = System.Drawing.SystemColors.Control
        Me.pnlsendmail.Location = New System.Drawing.Point(6, 0)
        Me.pnlsendmail.Name = "pnlsendmail"
        Me.pnlsendmail.Size = New System.Drawing.Size(384, 472)
        Me.pnlsendmail.TabIndex = 6
        '
        'btnSendMail
        '
        Me.btnSendMail.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSendMail.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSendMail.ForeColor = System.Drawing.SystemColors.Desktop
        Me.btnSendMail.Location = New System.Drawing.Point(1, 448)
        Me.btnSendMail.Name = "btnSendMail"
        Me.btnSendMail.Size = New System.Drawing.Size(120, 20)
        Me.btnSendMail.TabIndex = 9
        Me.btnSendMail.Text = "Send"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.lstAttachment)
        Me.GroupBox5.Controls.Add(Me.btnMaillAttach)
        Me.GroupBox5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox5.Location = New System.Drawing.Point(8, 128)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(368, 136)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Attachments"
        '
        'lstAttachment
        '
        Me.lstAttachment.HorizontalScrollbar = True
        Me.lstAttachment.ItemHeight = 15
        Me.lstAttachment.Location = New System.Drawing.Point(8, 44)
        Me.lstAttachment.Name = "lstAttachment"
        Me.lstAttachment.Size = New System.Drawing.Size(352, 79)
        Me.lstAttachment.TabIndex = 6
        '
        'btnMaillAttach
        '
        Me.btnMaillAttach.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnMaillAttach.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnMaillAttach.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMaillAttach.ForeColor = System.Drawing.SystemColors.Desktop
        Me.btnMaillAttach.Location = New System.Drawing.Point(5, 17)
        Me.btnMaillAttach.Name = "btnMaillAttach"
        Me.btnMaillAttach.Size = New System.Drawing.Size(120, 20)
        Me.btnMaillAttach.TabIndex = 5
        Me.btnMaillAttach.Text = "Attach File"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtMailBody)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.GroupBox2.Location = New System.Drawing.Point(8, 264)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(369, 184)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Type Message"
        '
        'txtMailBody
        '
        Me.txtMailBody.Location = New System.Drawing.Point(8, 16)
        Me.txtMailBody.Name = "txtMailBody"
        Me.txtMailBody.Size = New System.Drawing.Size(352, 160)
        Me.txtMailBody.TabIndex = 8
        Me.txtMailBody.Text = ""
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtServer)
        Me.GroupBox4.Controls.Add(Me.lblServer)
        Me.GroupBox4.Controls.Add(Me.txtMailTo)
        Me.GroupBox4.Controls.Add(Me.txtMailSubject)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.txtMailCC)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.SystemColors.Control
        Me.GroupBox4.Location = New System.Drawing.Point(7, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(369, 120)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(88, 94)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(272, 20)
        Me.txtServer.TabIndex = 3
        Me.txtServer.Text = "192.168.1.200"
        '
        'lblServer
        '
        Me.lblServer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServer.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lblServer.Location = New System.Drawing.Point(16, 96)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(64, 16)
        Me.lblServer.TabIndex = 9
        Me.lblServer.Text = "Server"
        '
        'txtMailTo
        '
        Me.txtMailTo.ItemHeight = 14
        Me.txtMailTo.Location = New System.Drawing.Point(88, 14)
        Me.txtMailTo.Name = "txtMailTo"
        Me.txtMailTo.Size = New System.Drawing.Size(272, 22)
        Me.txtMailTo.TabIndex = 1
        '
        'txtMailSubject
        '
        Me.txtMailSubject.Location = New System.Drawing.Point(88, 71)
        Me.txtMailSubject.Name = "txtMailSubject"
        Me.txtMailSubject.Size = New System.Drawing.Size(272, 20)
        Me.txtMailSubject.TabIndex = 2
        Me.txtMailSubject.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Book Antiqua", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(16, 71)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Subject :"
        '
        'txtMailCC
        '
        Me.txtMailCC.Location = New System.Drawing.Point(88, 36)
        Me.txtMailCC.Multiline = True
        Me.txtMailCC.Name = "txtMailCC"
        Me.txtMailCC.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMailCC.Size = New System.Drawing.Size(272, 36)
        Me.txtMailCC.TabIndex = 2
        Me.txtMailCC.Text = ""
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Book Antiqua", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(24, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "CC:"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Book Antiqua", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(24, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 23)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "To:"
        '
        'tpgviewmail
        '
        Me.tpgviewmail.Controls.Add(Me.sbr1)
        Me.tpgviewmail.Controls.Add(Me.lv1)
        Me.tpgviewmail.Controls.Add(Me.cb1)
        Me.tpgviewmail.Controls.Add(Me.ExpTree1)
        Me.tpgviewmail.Controls.Add(Me.cmdRefresh)
        Me.tpgviewmail.Controls.Add(Me.cmdCTest)
        Me.tpgviewmail.Controls.Add(Me.cmdExit)
        Me.tpgviewmail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tpgviewmail.Location = New System.Drawing.Point(4, 23)
        Me.tpgviewmail.Name = "tpgviewmail"
        Me.tpgviewmail.Size = New System.Drawing.Size(416, 485)
        Me.tpgviewmail.TabIndex = 3
        Me.tpgviewmail.Text = "View Folder"
        '
        'sbr1
        '
        Me.sbr1.Location = New System.Drawing.Point(0, 461)
        Me.sbr1.Name = "sbr1"
        Me.sbr1.Size = New System.Drawing.Size(416, 24)
        Me.sbr1.TabIndex = 17
        '
        'lv1
        '
        Me.lv1.Location = New System.Drawing.Point(165, 32)
        Me.lv1.Name = "lv1"
        Me.lv1.Size = New System.Drawing.Size(224, 384)
        Me.lv1.TabIndex = 2
        '
        'cb1
        '
        Me.cb1.Location = New System.Drawing.Point(168, 8)
        Me.cb1.Name = "cb1"
        Me.cb1.Size = New System.Drawing.Size(224, 22)
        Me.cb1.TabIndex = 1
        Me.cb1.Text = "ComboBox1"
        '
        'ExpTree1
        '
        Me.ExpTree1.AllowDrop = True
        Me.ExpTree1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExpTree1.Location = New System.Drawing.Point(8, 8)
        Me.ExpTree1.Name = "ExpTree1"
        Me.ExpTree1.ShowRootLines = False
        Me.ExpTree1.Size = New System.Drawing.Size(152, 408)
        Me.ExpTree1.StartUpDirectory = ExpTreeLib.ExpTree.StartDir.Desktop
        Me.ExpTree1.TabIndex = 0
        '
        'cmdRefresh
        '
        Me.cmdRefresh.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cmdRefresh.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdRefresh.Location = New System.Drawing.Point(136, 424)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(120, 20)
        Me.cmdRefresh.TabIndex = 4
        Me.cmdRefresh.Text = "Refresh"
        '
        'cmdCTest
        '
        Me.cmdCTest.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cmdCTest.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdCTest.Location = New System.Drawing.Point(8, 424)
        Me.cmdCTest.Name = "cmdCTest"
        Me.cmdCTest.Size = New System.Drawing.Size(120, 20)
        Me.cmdCTest.TabIndex = 3
        Me.cmdCTest.Text = "View folder"
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.SystemColors.ControlLight
        Me.cmdExit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdExit.Location = New System.Drawing.Point(264, 424)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(120, 20)
        Me.cmdExit.TabIndex = 5
        Me.cmdExit.Text = "Exit"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tpgedit)
        Me.TabControl1.Controls.Add(Me.tpgviewmail)
        Me.TabControl1.Controls.Add(Me.tpgsendmail)
        Me.TabControl1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.ItemSize = New System.Drawing.Size(50, 19)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(424, 512)
        Me.TabControl1.TabIndex = 14
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmEditLead
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(426, 540)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnExit)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmEditLead"
        Me.Text = "Edit Lead"
        Me.tpgedit.ResumeLayout(False)
        Me.pnlAddLead.ResumeLayout(False)
        Me.grpDetails.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.pnlviewjournal.ResumeLayout(False)
        Me.tpgsendmail.ResumeLayout(False)
        Me.pnlsendmail.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.tpgviewmail.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub

#Region "private members"
    Private Delegate Sub mydelegate()
#End Region

#Region "editleads"
    Private Sub frmEditLead_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboProspect.Items.Add("Prospect")
            Me.cboProspect.Items.Add("Proposal")
            Me.cboProspect.Items.Add("PHO")
            Me.cboProspect.Items.Add("Failed")
            If Convert.IsDBNull(cname) = False Then
                Me.lblCompany.Text = cname
            End If
            If Convert.IsDBNull(desription) = False Then
                txtDesc.Text = desription
            End If
            If Convert.IsDBNull(title) = False Then
                txttitle.Text = title
            End If
            If Convert.IsDBNull(ddate) = False Then
                dtpsniffed.Text = ddate
            End If
            ' Me.cboProspect.DropDownStyle = ComboBoxStyle.DropDown
            Me.cboProspect.Text = cstatus
            loadmails()
        Catch ex As Exception

        Finally
            Me.txtAmount.Enabled = True
        End Try
        Try
            Dim _thread As Thread = New Thread(AddressOf ld)
            _thread.IsBackground = True
            _thread.Start()
        Catch ex As Exception

        End Try
        Try
            If myForms.mailserver.Trim.Length > 0 Then
                txtServer.Text = myForms.mailserver
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub loadmails()
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Try

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                Dim str
                str = "select e_mail1,e_mail2 from contact " _
                & " where client_no='" & clientno & "'"
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    While .EOF = False
                        Me.txtMailTo.Items.Add(.Fields("e_mail1").Value)
                        Me.txtMailTo.Items.Add(.Fields("e_mail2").Value)

                        .MoveNext()
                        Application.DoEvents()

                    End While
                End If
            End With
        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        Try
            If onactivateform = True Then
                Me.lblCompany.Text = cname
                Me.txtDesc.Text = desription
                Me.cboProspect.DropDownStyle = ComboBoxStyle.DropDown
                Me.cboProspect.Text = cstatus
                onactivateform = False
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnSave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Try
            If leadno = "" Then
                MsgBox("Lead has already been deleted", MsgBoxStyle.Information, "Save")
                Exit Try
            End If
            ''''''''''''---------------validation
            'If Me.txtAmount.Text = "" Then
            '    MessageBox.Show("Please input an amount", "Edit leads", _
            '    MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Try
            'End If
            If Me.txtdepartment.Text = "" Then
                MessageBox.Show("Please pick a department", "Edit leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.txtDesc.Text = "" Then
                MessageBox.Show("Please input description", "Edit leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.txttitle.Text = "" Then
                MessageBox.Show("Please input a title for the lead", "Edit leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.cboProspect.Text = "" Then
                MessageBox.Show("Please pick a status for this lead", "Edit leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            '--------------end
            Dim lno = leadno
            Dim cstatus = editclientstatus(clientno)
            Dim sdate As String
            sdate = dtpsniffed.Value.Year & "-" _
            & dtpsniffed.Value.Month & "-" _
            & dtpsniffed.Value.Day
            Dim strsql
            Call storedata()
            Try
                journalpath = journalpath.Replace("\", "|")
            Catch ex As Exception
            End Try
            '-------------------
            Dim arr() As String
            Dim strr As String
            Dim y As Integer
            txtDesc.Text = Me.txtDesc.Text.Trim()
            arr = txtDesc.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            '----------------------------------
            strsql = "update leads set" _
                      & " leads_no='" & lno & "'," _
                      & " client_no='" & clientno & "'," _
                      & " descrip='" & strr & "'," _
                      & " status='" & cboProspect.Text & "'," _
                      & " title='" & txttitle.Text & "'," _
                      & " journal='" & journalpath & "'," _
                      & " amount='" & txtAmount.Text & "'," _
                       & "department='" & txtdepartment.Text & "'," _
                      & " date_sniffed='" & sdate & "'" _
                      & " where leads_no= '" & lno & "'" _
                      & ";"
            strsql += " update clients set least_status='" & cstatus & "'"
            strsql += " where lower(client_no)='" & CStr(Me.clientno).ToLower() & "'"
            strsql += ";"
            connect.BeginTrans()
            connect.IsolationLevel = ADODB.IsolationLevelEnum.adXactSerializable
            connect.Execute(strsql)
            connect.CommitTrans()
            MessageBox.Show(Text:="Lead has been successfully changed", _
            buttons:=MessageBoxButtons.OK, caption:="Add Client", Icon:=MessageBoxIcon.Information)
            refreshleads = True
            refreshleadshome = True
            refreshjobs = True
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
        Try
            Dim tthread As System.Threading.Thread = New System.Threading.Thread(AddressOf threadjobs)
            Try
                If tthread.IsAlive = True Then
                    tthread.Abort()
                End If
            Catch ex As Exception

            End Try

            tthread.IsBackground = True
            tthread.Start()
        Catch ex As Exception

        End Try
        Try
            Dim tthread1 As System.Threading.Thread = New System.Threading.Thread(AddressOf myForms.Main.loadleads)
            Try
                If tthread1.IsAlive = True Then
                    tthread1.Abort()
                End If
            Catch ex As Exception

            End Try

            tthread1.IsBackground = True
            tthread1.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub threadjobs()
        Try
            myForms.CustomerForm3.Invoke(New mydelegate(AddressOf myForms.CustomerForm3.loadleads))

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            editleads = False
            Me.Dispose(True)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub lblViewJournal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Try
            pnlAddLead.Visible = False
            pnlviewjournal.Visible = True

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                Dim dstr = "select journal from  leads where leads_no='" & leadno & "'"
                .Open(dstr, connect)
                If .BOF = False And .EOF = False Then
                    journalpath = .Fields("journal").Value
                    journalpath = journalpath.Replace("|", "\")
                    rtbjournal.LoadFile(journalpath)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnsavejournal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsavejournal.Click
        Try
            pnlAddLead.Visible = True
            pnlviewjournal.Visible = False

        Catch ex As Exception

        End Try
    End Sub
    Private Sub storedata()
        Try
            File.Delete(journalpath)
            rtbjournal.SaveFile(journalpath)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub cboProspect_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProspect.SelectedValueChanged
        Try
            txtAmount.Enabled = False
            Select Case cboProspect.Text.ToLower()
                Case "prospect"
                    txtAmount.Enabled = True
                Case "suspect"
                    txtAmount.Enabled = False
                Case "pho"
                    txtAmount.Enabled = True
                    If cboProspect.Text.ToLower() = "pho" Then
                        btnconvert.Enabled = True
                    Else
                        btnconvert.Enabled = False
                    End If
                Case Else
                    txtAmount.Enabled = True
            End Select
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnconvert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnconvert.Click
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Try
            Dim issaved As Boolean = False
            If cboProspect.Text = "PHO" Then

                Dim rs As New ADODB.Recordset
                With rs
                    .CursorLocation = CursorLocationEnum.adUseClient
                    .CursorType = CursorTypeEnum.adOpenForwardOnly
                    Dim str As String
                    str = "select status from leads where leads_no"
                    str += "='" & leadno & "'"
                    .Open(str, connect)
                    If .BOF = False And .EOF = False Then
                        Dim nn As String = .Fields("status").Value
                        If nn.ToLower().Trim = "pho" Then
                            issaved = True
                        End If

                    End If
                End With
                rs.Close()
                If issaved = True Then
                    Dim isjobexist As Boolean = False
                    Dim rs1 As New ADODB.Recordset
                    With rs1
                        .CursorLocation = CursorLocationEnum.adUseClient
                        .CursorType = CursorTypeEnum.adOpenForwardOnly
                        Dim str As String
                        str = "select job_no from rcljobs where job_no"
                        str += "='" & leadno & "'"
                        .Open(str, connect)
                        If .BOF = False And .EOF = False Then
                            isjobexist = True
                        End If
                    End With
                    rs1.Close()
                    If isjobexist = False Then
                        Dim strsql As String
                        Dim mydate As String
                        mydate = dtpsniffed.Value.Year & "-" _
                        & dtpsniffed.Value.Month & "-" _
                        & dtpsniffed.Value.Day & " " _
                        & dtpsniffed.Value.Hour & ":" _
                        & dtpsniffed.Value.Minute & ":" _
                        & dtpsniffed.Value.Second
                        strsql += " insert into rcljobs (client_no,job_no,job_tittle,descrip,job_status," _
                        & "journal,amount,date_sniffed,sdate,department) values"
                        strsql += "('" & CStr(Me.clientno).ToLower() & "','" & leadno & "','" & txttitle.Text & "'," _
                        & "'" & txtDesc.Text & "','" & "Current" & "'," _
                        & " '" & journalpath & "','" & txtAmount.Text & "','" & mydate & "'," _
                        & " '" & Now & "','" & txtdepartment.Text & "')"
                        strsql += ";"
                        strsql += "delete from leads where leads_no='" & leadno & "'" & ";"
                        connect.BeginTrans()
                        connect.Execute(strsql)
                        connect.CommitTrans()
                        refreshjobs = True
                        MsgBox("Conversion successful", _
                        MsgBoxStyle.Information, "Convert lead to job")
                    Else
                        MsgBox("This job doesn't exist", MsgBoxStyle.Information, "Convert lead to job")
                    End If

                Else
                    MsgBox("Either the lead has been converted to a job or " _
                    & vbCrLf & "the lead has not been saved", MsgBoxStyle.Information _
                    , "Convert lead to job")
                End If
            Else

                MsgBox(" Only (PHO) leads can be converted to jobs" _
                        & "", MsgBoxStyle.Information _
                         , "Convert lead to job")
            End If
        Catch ex As Exception
            Try
                connect.RollbackTrans()
            Catch rt As Exception
                MsgBox(rt.Message.ToString())
            End Try
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btndeletelead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeletelead.Click
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Try
            If leadno = "" Then
                MsgBox("Lead has already been deleted", MsgBoxStyle.Information, "Save")
                Exit Try
            End If

            If MessageBox.Show("Are you sure you want to delete this lead" _
            , "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                connect.BeginTrans()
                Dim strsql As String
                Dim int1, int2, int3 As Integer
                int1 = CInt(leadno)
                int2 = CInt(leadno) - 1
                int3 = CInt(Me.clientno)
                strsql = "select clean_rcljobs(" & int1 & "," & int2 & "," & int3 & ");" 'use of user defined function
                strsql += "Delete from leads where leads_no="
                strsql += "'" & leadno & "';"
                connect.Execute(strsql)
                connect.CommitTrans()
                MsgBox("Deletion successful", MsgBoxStyle.Information)
                txtAmount.Text = ""
                txtDesc.Text = ""
                txttitle.Text = ""
                cboProspect.Text = ""
                leadno = "" '------------remember this
                refreshleads = True
                refreshleadshome = True

                '-------------- delete leads-------------
                'Dim nv As NameValueCollection
                'Dim myvar As String
                'nv = ConfigurationSettings.AppSettings()
                'myvar = nv("folderpath")
                ''str = Configuration.ConfigurationSettings.AppSettings("folderpath")
                'Dim mypath As String
                'Dim lno As String = leadno
                'mypath = myvar & "\" & lno.Substring(0, 4)
                'mypath += "\" & lno
                'Try
                '    'File.Delete(mypath & "\" & "*.*")
                '    'File.Delete(mypath)
                'Catch er As Exception
                '    'MsgBox(er.Message.ToString())
                'End Try
                'Try

                'Catch eg As Exception
                'End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
        Try
            Dim tthread As System.Threading.Thread = New System.Threading.Thread(AddressOf threadjobs)
            Try
                If tthread.IsAlive = True Then
                    tthread.Abort()
                End If
            Catch ex As Exception

            End Try

            tthread.IsBackground = True
            tthread.Start()
        Catch ex As Exception

        End Try
        Try
            Dim tthread1 As System.Threading.Thread = New System.Threading.Thread(AddressOf myForms.Main.loadleads)
            Try
                If tthread1.IsAlive = True Then
                    tthread1.Abort()
                End If
            Catch ex As Exception

            End Try

            tthread1.IsBackground = True
            tthread1.Start()
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "send mail"
    Private Sub checkdirectory()
        Try
            Dim myvar As String = "DSN=" & myForms.qfolderpath
            'str = Configuration.ConfigurationSettings.AppSettings("folderpath")

            Dim myfile, mypath
            mypath = myvar
            mypath = mypath & "\"
            mypath = mypath & clientno

            myfile = Dir(mypath, FileAttribute.Directory)
            If myfile <> "" Then
                mypath += "\" & leadno
                myfile = Dir(mypath, FileAttribute.Directory)
                If myfile <> "" Then
                    Me.storedata(mypath)

                Else
                    MkDir(mypath)
                    Me.storedata(mypath)
                End If

            Else
                MkDir(mypath)
                mypath += "\" & leadno
                MkDir(mypath)
                Me.storedata(mypath)
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub storedata(ByVal str As String)
        Try
            'Me.txtMailBody.SaveFile(str & Me.txtMailSubject.Text, System.Windows.Forms.RichTextBoxStreamType.RichText)
            Dim i, j
            Dim str1, myfilename As String
            Dim a()
            For i = 0 To Me.lstAttachment.Items.Count - 1



                str1 = lstAttachment.Items(i)
                a = str1.Split("\")
                j = a.GetUpperBound(0)
                myfilename = a(j)
                Try
                    File.Copy(str1, str & "\" & myfilename)
                Catch exd As Exception
                    MsgBox(exd.Message.ToString())
                End Try

            Next
            Dim dtp As New System.Windows.Forms.DateTimePicker
            myfilename.Replace(".txt", "")
            Me.txtMailBody.SaveFile(str & "\" & myclientno & "_" & dtp.Value.Year & dtp.Value.Month & dtp.Value.Day _
            & dtp.Value.Hour & dtp.Value.Minute & dtp.Value.Second & dtp.Value.Millisecond & ".txt")
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnMaillAttach_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMaillAttach.Click
        Try
            Dim Counter As Integer
            Dim ofd As New System.Windows.Forms.OpenFileDialog
            ofd.Multiselect = True
            ofd.CheckFileExists = True
            ofd.Title = "Select file(s) to attach"
            ofd.ShowDialog()
            For Counter = 0 To UBound(ofd.FileNames)
                lstAttachment.Items.Add(ofd.FileNames(Counter))
            Next
            'attachedfile = ofd.FileName
            'Me.txtMailAttachment.Text = attachedfile
        Catch ex As System.Exception

        End Try
    End Sub
    Private Sub btnSendMail_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendMail.Click
        Dim currentcursor As Cursor = Cursor.Current
        Dim kkk = 0
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim Counter As Integer
            If txtMailTo.Text.Trim = "" Then
                MsgBox("Enter the Recipient email address ...!!!", _
                         MsgBoxStyle.Information, "Send Email")
                Exit Sub
            End If
            If txtMailSubject.Text.Trim = "" Then
                MsgBox("Enter the Email subject ...!!!", _
                         MsgBoxStyle.Information, "Send Email")
                Exit Sub
            End If

            'Set the properties
            'Assign the SMTP server
            'obj.SmtpServer = "194.201.253.105"
            'obj.SmtpServer = "192.168.1.182"
            obj.SmtpServer.Insert(0, txtServer.Text.Trim)
            'obj.SmtpServer = txtServer.Text.Trim
            'Multiple recepients can be specified using ; as the delimeter
            'Address of the recipient
            If txtMailTo.Text.Trim.Length > 0 Then
                Mailmsg.To = txtMailTo.Text
            End If
            'Your From Address
            'You can also use a custom header Reply-To for a different replyto address
            'Mailmsg.From = "\" & txtFromDisplayName.Text & "\ <" & txtFrom.Text & ">"
            Mailmsg.From = "waruid@yahoo.com"
            'Specify the body format

            Mailmsg.BodyFormat = MailFormat.Html 'Send the mail in HTML Format

            'Mailmsg.BodyFormat = MailFormat.Text
            'If you want you can add a reply to header 
            'Mailmsg.Headers.Add("Reply-To", "Manoj@geinetech.net")
            'custom headersare added like this
            'Mailmsg.Headers.Add("Manoj", "TestHeader")
            'Mail Subject
            If txtMailSubject.Text.Trim.Length > 0 Then
                Mailmsg.Subject = txtMailSubject.Text
            End If
            'Attach the files one by one
            If lstAttachment.Items.Count > 0 Then
                For Counter = 0 To lstAttachment.Items.Count - 1
                    Attachment = New MailAttachment(lstAttachment.Items(Counter))
                    'Add it to the mail message
                    Mailmsg.Attachments.Add(Attachment)
                Next

                ' Attachment = New MailAttachment(txtMailAttachment.Text)
                'Add it to the mail message
                'Mailmsg.Attachments.Add(Attachment)
            End If
            'Mail Body
            If txtMailBody.Text.Trim.Length > 0 Then
                Mailmsg.Body = txtMailBody.Text
            End If
            'Call the send method to send the mail
            Application.DoEvents()
            If lstAttachment.Items.Count > 0 Then
                Me.checkdirectory()
            End If
            Mailmsg.Priority = MailPriority.High
            obj.Send(Mailmsg)

        Catch ex As Exception
            kkk = 1
            'MsgBox(ex.Message.ToString(), MsgBoxStyle.Information)
            DisplayError(ex)
        Finally

            Cursor.Current = currentcursor
            If kkk = 0 Then
                MessageBox.Show(Text:="Mail successfully sent", _
                caption:="Send Mail", buttons:=MessageBoxButtons.OK, _
            Icon:=MessageBoxIcon.Information)
            End If
        End Try
    End Sub
    Private Sub DisplayError(ByVal ex As Exception)
        Try
            MessageBox.Show(ex.GetType().ToString() & _
                    vbCrLf & vbCrLf & _
                    ex.Message & vbCrLf & vbCrLf & _
                    ex.StackTrace, _
                    "Error", _
                    MessageBoxButtons.AbortRetryIgnore, _
                    MessageBoxIcon.Stop)
        Catch vb As Exception
        End Try

    End Sub
#End Region

#Region "both"

    Public Property company() As String
        Get
            Return Me.lblCompany.Text
        End Get
        Set(ByVal Value As String)
            lblCompany.Text = Value
        End Set
    End Property
    Public Property descriptions() As String
        Get
            Return txtDesc.Text
        End Get
        Set(ByVal Value As String)
            txtDesc.Text = Value
        End Set
    End Property
    Public Property status() As String
        Get
            Return cboProspect.Text
        End Get
        Set(ByVal Value As String)
            cboProspect.Text = Value
        End Set

    End Property
#End Region

#Region "explorer"
#Region "VisibleChanged Event"
    Private Sub lv1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lv1.VisibleChanged
        Try
            If lv1.Visible Then
                SystemImageListManager.SetListViewImageList(lv1, True, False)
                SystemImageListManager.SetListViewImageList(lv1, False, False)
            End If
        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "   ExplorerTree Event Handling"
    Private Sub AfterNodeSelect(ByVal pathName As String, ByVal CSI As CShItem) Handles ExpTree1.ExpTreeNodeSelected
        Try
            Dim dirList As New ArrayList
            Dim fileList As New ArrayList
            Dim TotalItems As Integer
            LastSelectedCSI = CSI
            If CSI.DisplayName.Equals(CShItem.strMyComputer) Then
                dirList = CSI.GetDirectories 'avoid re-query since only has dirs
            Else
                dirList = CSI.GetDirectories
                fileList = CSI.GetFiles
            End If
            SetUpComboBox(CSI)
            TotalItems = dirList.Count + fileList.Count
            Event1.WaitOne()
            If TotalItems > 0 Then
                Dim item As CShItem
                dirList.Sort()
                fileList.Sort()
                Me.Text = pathName
                sbr1.Text = pathName & "                 " & _
                            dirList.Count & " Directories " & fileList.Count & " Files"
                Dim combList As New ArrayList(TotalItems)
                combList.AddRange(dirList)
                combList.AddRange(fileList)

                'Build the ListViewItems & add to lv1
                lv1.BeginUpdate()
                lv1.Items.Clear()
                For Each item In combList
                    Dim lvi As New ListViewItem(item.DisplayName)
                    With lvi
                        If Not item.IsDisk And item.IsFileSystem And Not item.IsFolder Then
                            If item.Length > 1024 Then
                                .SubItems.Add(Format(item.Length / 1024, "#,### KB"))
                            Else
                                .SubItems.Add(Format(item.Length, "##0 Bytes"))
                            End If
                        Else
                            .SubItems.Add("")
                        End If
                        .SubItems.Add(item.TypeName)
                        If item.IsDisk Then
                            .SubItems.Add("")
                        Else
                            If item.LastWriteTime = testTime Then '"#1/1/0001 12:00:00 AM#" is empty
                                .SubItems.Add("")
                            Else
                                .SubItems.Add(item.LastWriteTime)
                            End If
                        End If
                        '.ImageIndex = SystemImageListManager.GetIconIndex(item, False)
                        .Tag = item
                    End With
                    lv1.Items.Add(lvi)
                Next
                lv1.EndUpdate()
                LoadLV1Images()
            Else
                lv1.Items.Clear()
                sbr1.Text = pathName & " Has No Items"
            End If
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "   ListView and ComboBox Event Handling"
    Private BackList As ArrayList

    Private Sub lv1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lv1.MouseUp
        Try
            Dim lvi As ListViewItem = lv1.GetItemAt(e.X, e.Y)
            If IsNothing(lvi) Then Exit Sub
            If IsNothing(lv1.SelectedItems) OrElse lv1.SelectedItems.Count < 1 Then Exit Sub
            Dim item As CShItem = lv1.SelectedItems(0).Tag
            If item.IsFolder Then
                If e.Button = MouseButtons.Right Then
                    Event1.WaitOne()
                    SetUpComboBox(item)
                    ExpTree1.RootItem = item
                ElseIf e.Button = MouseButtons.Left Then
                    ExpTree1.ExpandANode(item)
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub SetUpComboBox(ByVal item As CShItem)
        Try
            BackList = New ArrayList
            With cb1
                .Items.Clear()
                .Text = ""
                Dim CSI As CShItem = item
                Do While Not IsNothing(CSI.Parent)
                    CSI = CSI.Parent
                    BackList.Add(CSI)
                    .Items.Add(CSI.DisplayName)
                Loop
                .SelectedIndex = -1
            End With
            lv1.Focus()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cb1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb1.SelectedIndexChanged
        Try
            With cb1
                If .SelectedIndex > -1 AndAlso _
                     .SelectedIndex < BackList.Count Then
                    Dim item As CShItem = BackList(.SelectedIndex)
                    BackList = New ArrayList
                    .Items.Clear()
                    ExpTree1.RootItem = item
                End If
            End With
        Catch ex As Exception

        End Try

    End Sub

#End Region

    '#Region "   View Menu Event Handling"
    '    Private Sub mnuViewLargeIcons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewLargeIcons.Click
    '        lv1.View = View.LargeIcon
    '    End Sub

    '    Private Sub mnuViewSmallIcons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewSmallIcons.Click
    '        lv1.View = View.SmallIcon
    '    End Sub

    '    Private Sub mnuViewList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewList.Click
    '        lv1.View = View.List
    '    End Sub

    '    Private Sub mnuViewDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewDetails.Click
    '        lv1.View = View.Details
    '    End Sub
    '#End Region
    Private Sub cmdCTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCTest.Click
        Try
            Dim connectstr As String
            Dim myvar As String = "value=" & myForms.qfolderpath
            connectstr = myvar
            connectstr += "\" & clientno & "\" & leadno

            Dim myfile, mypath
            mypath = connectstr
            'mypath = mypath & "\"
            'mypath = mypath & clientno

            myfile = Dir(mypath, FileAttribute.Directory)

            If myfile = "" Then
                MessageBox.Show("The leads folder whose path is as shown below doesn't exist" _
                & vbCrLf & "path" & "  " & mypath & "", "View Folder", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            End If

            Dim cDir As CShItem = GetCShItem(connectstr)
            If cDir.IsFolder Then
                ExpTree1.RootItem = cDir
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Me.Text = "Edit Lead"
        End Try


    End Sub
    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        Try
            If Not ExpTree1.SelectedItem Is Nothing Then
                ExpTree1.RefreshTree()
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Me.Text = "Edit Lead"
        End Try

    End Sub
#Region "   IconIndex Loading Thread"
    Private Sub LoadLV1Images()
        Try
            Dim ts As New ThreadStart(AddressOf DoLoadLv)
            Dim ot As New Thread(ts)
            ot.ApartmentState = ApartmentState.STA
            Event1.Reset()
            ot.Start()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub DoLoadLv()
        Try
            Dim lvi As ListViewItem
            For Each lvi In lv1.Items
                lvi.ImageIndex = SystemImageListManager.GetIconIndex(lvi.Tag, False)
            Next
            Event1.Set()
        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "   Various testing routines. Depend on Files/Dirs on Developmental system"
    ' The Routines in this region handle buttons that have been removed from the form
    '  with the obvious names.  The routines depend on Files and Dirs found on
    '  my development system.  To see how they work, add the buttons and change 
    '  the literal references to Files & Dirs that exist on your system
    'Private Sub cmdFilterTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilterTest.Click
    '    Dim filtAl As New ArrayList()
    '    Dim fList As New ArrayList()
    '    Dim baseItem As New CShItem("C:\Data\Checklists")
    '    fList = baseItem.GetFiles("*.doc")
    '    Dim Item As CShItem
    '    For Each Item In fList
    '        Dim xxx As New CShItem(Item.Path)
    '        Debug.WriteLine(xxx.Path)
    '        Debug.WriteLine(xxx.DisplayName)
    '        xxx.DebugDump()
    '        filtAl.Add(xxx)
    '    Next
    '    'Build the ListViewItems & add to lv1
    '    lv1.BeginUpdate()
    '    lv1.Items.Clear()
    '    For Each Item In fList
    '        Dim lvi As New ListViewItem(Item.DisplayName)
    '        With lvi
    '            If Not Item.IsDisk And Item.IsFileSystem And Not Item.IsFolder Then
    '                If Item.Length > 1024 Then
    '                    .SubItems.Add(Format(Item.Length / 1024, "#,### KB"))
    '                Else
    '                    .SubItems.Add(Format(Item.Length, "##0 Bytes"))
    '                End If
    '            Else
    '                .SubItems.Add("")
    '            End If
    '            .SubItems.Add(Item.TypeName)
    '            If Item.IsDisk Then
    '                .SubItems.Add("")
    '            Else
    '                If Item.LastWriteTime = testTime Then '"#1/1/0001 12:00:00 AM#" is empty
    '                    .SubItems.Add("")
    '                Else
    '                    .SubItems.Add(Item.LastWriteTime)
    '                End If
    '            End If
    '            .ImageIndex = SystemImageListManager.GetIconIndex(Item, False)
    '            .Tag = Item
    '        End With
    '        lv1.Items.Add(lvi)
    '    Next
    '    lv1.EndUpdate()
    'End Sub

    'Private Sub cmdExpandTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim testPath As String = "F:\Music\Clips\Brooks & Dunn\Borderline"
    '    ExpTree1.ExpandANode(testPath)
    'End Sub
#End Region
#End Region

#Region "validation"
    Private Sub txtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesc.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtDesc, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtDesc, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txttitle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttitle.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txttitle, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txttitle, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtAmount, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtAmount, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtdepartment_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdepartment.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtdepartment, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtdepartment, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

    Private Sub btndepartments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndepartments.Click
        Try
            Dim n As New frmdepartments
            n.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

#Region "loaddepartments"
    Public Sub loaddepart()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str = "select * from department"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                myForms.CustomerForm4.txtdepartment.Items.Clear()
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While rs.EOF = False
                        myForms.CustomerForm4.txtdepartment.Items.Add(.Fields("dept").Value)
                        .MoveNext()
                        Application.DoEvents()
                    End While
                End If
            End With
            Try
                rs.Close()

            Catch er34b As Exception

            End Try
            Try

                connect.Close()
            Catch er344 As Exception

            End Try
        Catch ex As Exception

        End Try
        Try
            System.GC.Collect()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ld()
        Try
            myForms.CustomerForm4.Invoke(New mydelegate(AddressOf loaddepart))
        Catch ex As Exception

        End Try
    End Sub
#End Region

End Class
