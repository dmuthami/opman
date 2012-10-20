Imports ADODB
Imports System.Text.StringBuilder
Imports System.String

Imports System.Threading
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Text
Public Class frmEditJob

#Region "members"
    Inherits System.Windows.Forms.Form
    Public myjobno, _client_no As String
    Public jobtitle As String
    Public contname As String
    Public jobstatus As String
    Public tecres, description As String
    Public Delegate Sub mydelegate()
    Public Delegate Sub mydelegate1()
    Dim hticasuals As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htiaccomodation As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htitravel As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htihired As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htiramequip As System.Windows.Forms.DataGrid.HitTestInfo
    Public totalincome, totalkost As String
    Public kcasual, kramani, khired As String
    Public kpersonnel, kaccomodation, ktravel As String
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
            editjob = False
            myForms.CustomerForm2 = Nothing
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
    Friend WithEvents lbClientName As System.Windows.Forms.Label
    Friend WithEvents lblClientNo As System.Windows.Forms.Label
    Friend WithEvents lblojobno As System.Windows.Forms.Label
    Friend WithEvents txtJobNo As System.Windows.Forms.TextBox
    Friend WithEvents lblJobNo As System.Windows.Forms.Label
    Friend WithEvents txtJobTitle As System.Windows.Forms.TextBox
    Friend WithEvents lblJobTitle As System.Windows.Forms.Label
    Friend WithEvents grpdesc As System.Windows.Forms.GroupBox
    Friend WithEvents txtdesc As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents pnlctrlpersonnel As System.Windows.Forms.Panel
    Friend WithEvents StiButton2 As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents pnltravel As System.Windows.Forms.Panel
    Friend WithEvents pnlaccomodation As System.Windows.Forms.Panel
    Friend WithEvents pnlcasuals As System.Windows.Forms.Panel
    Friend WithEvents lblCasual As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblCostAgreed As System.Windows.Forms.Label
    Friend WithEvents txtamount As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtContactName As System.Windows.Forms.TextBox
    Friend WithEvents lblContName As System.Windows.Forms.Label
    Friend WithEvents lblJobStatus As System.Windows.Forms.Label
    Friend WithEvents lblTechnicianResponsible As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblOpenjournal As System.Windows.Forms.LinkLabel
    Friend WithEvents btndeletelead As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents pnlhiredequip As System.Windows.Forms.Panel
    Friend WithEvents pnlramaniequip As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtdepartment As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnEditJobs As System.Windows.Forms.Button
    Friend WithEvents cboTechnicianresponsible As System.Windows.Forms.ComboBox
    Friend WithEvents cboJobstatus As System.Windows.Forms.ComboBox
    Friend WithEvents tbcjobs As System.Windows.Forms.TabControl
    Friend WithEvents tpgjobdetails As System.Windows.Forms.TabPage
    Friend WithEvents tpgequip As System.Windows.Forms.TabPage
    Friend WithEvents tpgfinance As System.Windows.Forms.TabPage
    Friend WithEvents tpgpersonnel As System.Windows.Forms.TabPage
    Friend WithEvents dtgpersonnel As System.Windows.Forms.DataGrid
    Friend WithEvents txtpersonelcost As VSEssentials.LabelTextBox
    Friend WithEvents dtgtravel As System.Windows.Forms.DataGrid
    Friend WithEvents dtgaccomodation As System.Windows.Forms.DataGrid
    Friend WithEvents dtgcasuals As System.Windows.Forms.DataGrid
    Friend WithEvents btncasuals As System.Windows.Forms.Button
    Friend WithEvents btntravel As System.Windows.Forms.Button
    Friend WithEvents btnaccomodation As System.Windows.Forms.Button
    Friend WithEvents txttravel As VSEssentials.LabelTextBox
    Friend WithEvents txtaccomodation As VSEssentials.LabelTextBox
    Friend WithEvents txtlabour As VSEssentials.LabelTextBox
    Friend WithEvents txthiredequip As VSEssentials.LabelTextBox
    Friend WithEvents dtghiredequip As System.Windows.Forms.DataGrid
    Friend WithEvents btnhired As System.Windows.Forms.Button
    Friend WithEvents txtramani As VSEssentials.LabelTextBox
    Friend WithEvents dtgramaniequip As System.Windows.Forms.DataGrid
    Friend WithEvents btnramaniequip As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtbudget As System.Windows.Forms.TextBox
    Friend WithEvents lblgrossmargin As System.Windows.Forms.Label
    Friend WithEvents btnprint As System.Windows.Forms.Button
    Friend WithEvents PrintPreviewDialog1 As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents btndelcasuals As System.Windows.Forms.Button
    Friend WithEvents btndelaccom As System.Windows.Forms.Button
    Friend WithEvents btndeltravel As System.Windows.Forms.Button
    Friend WithEvents btndelramequip As System.Windows.Forms.Button
    Friend WithEvents btndelhiredequip As System.Windows.Forms.Button
    Friend WithEvents btndepartments As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEditJob))
        Me.tbcjobs = New System.Windows.Forms.TabControl
        Me.tpgjobdetails = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btndepartments = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtbudget = New System.Windows.Forms.TextBox
        Me.cboJobstatus = New System.Windows.Forms.ComboBox
        Me.cboTechnicianresponsible = New System.Windows.Forms.ComboBox
        Me.txtdepartment = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.lblgrossmargin = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lblOpenjournal = New System.Windows.Forms.LinkLabel
        Me.btndeletelead = New System.Windows.Forms.Button
        Me.lblCostAgreed = New System.Windows.Forms.Label
        Me.txtamount = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtContactName = New System.Windows.Forms.TextBox
        Me.lblContName = New System.Windows.Forms.Label
        Me.lblJobStatus = New System.Windows.Forms.Label
        Me.lblTechnicianResponsible = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnEditJobs = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.lbClientName = New System.Windows.Forms.Label
        Me.lblClientNo = New System.Windows.Forms.Label
        Me.lblojobno = New System.Windows.Forms.Label
        Me.txtJobNo = New System.Windows.Forms.TextBox
        Me.lblJobNo = New System.Windows.Forms.Label
        Me.txtJobTitle = New System.Windows.Forms.TextBox
        Me.lblJobTitle = New System.Windows.Forms.Label
        Me.grpdesc = New System.Windows.Forms.GroupBox
        Me.txtdesc = New System.Windows.Forms.RichTextBox
        Me.tpgequip = New System.Windows.Forms.TabPage
        Me.pnlhiredequip = New System.Windows.Forms.Panel
        Me.btndelhiredequip = New System.Windows.Forms.Button
        Me.txthiredequip = New VSEssentials.LabelTextBox
        Me.dtghiredequip = New System.Windows.Forms.DataGrid
        Me.btnhired = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlramaniequip = New System.Windows.Forms.Panel
        Me.btndelramequip = New System.Windows.Forms.Button
        Me.txtramani = New VSEssentials.LabelTextBox
        Me.dtgramaniequip = New System.Windows.Forms.DataGrid
        Me.btnramaniequip = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.tpgfinance = New System.Windows.Forms.TabPage
        Me.pnltravel = New System.Windows.Forms.Panel
        Me.btndeltravel = New System.Windows.Forms.Button
        Me.txttravel = New VSEssentials.LabelTextBox
        Me.dtgtravel = New System.Windows.Forms.DataGrid
        Me.btntravel = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.pnlaccomodation = New System.Windows.Forms.Panel
        Me.btndelaccom = New System.Windows.Forms.Button
        Me.txtaccomodation = New VSEssentials.LabelTextBox
        Me.dtgaccomodation = New System.Windows.Forms.DataGrid
        Me.btnaccomodation = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnlcasuals = New System.Windows.Forms.Panel
        Me.btndelcasuals = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtlabour = New VSEssentials.LabelTextBox
        Me.dtgcasuals = New System.Windows.Forms.DataGrid
        Me.btncasuals = New System.Windows.Forms.Button
        Me.lblCasual = New System.Windows.Forms.Label
        Me.tpgpersonnel = New System.Windows.Forms.TabPage
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.txtpersonelcost = New VSEssentials.LabelTextBox
        Me.dtgpersonnel = New System.Windows.Forms.DataGrid
        Me.pnlctrlpersonnel = New System.Windows.Forms.Panel
        Me.StiButton2 = New System.Windows.Forms.Button
        Me.btnprint = New System.Windows.Forms.Button
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.tbcjobs.SuspendLayout()
        Me.tpgjobdetails.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.grpdesc.SuspendLayout()
        Me.tpgequip.SuspendLayout()
        Me.pnlhiredequip.SuspendLayout()
        CType(Me.dtghiredequip, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlramaniequip.SuspendLayout()
        CType(Me.dtgramaniequip, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpgfinance.SuspendLayout()
        Me.pnltravel.SuspendLayout()
        CType(Me.dtgtravel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlaccomodation.SuspendLayout()
        CType(Me.dtgaccomodation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlcasuals.SuspendLayout()
        CType(Me.dtgcasuals, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpgpersonnel.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.dtgpersonnel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlctrlpersonnel.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbcjobs
        '
        Me.tbcjobs.Controls.Add(Me.tpgjobdetails)
        Me.tbcjobs.Controls.Add(Me.tpgequip)
        Me.tbcjobs.Controls.Add(Me.tpgfinance)
        Me.tbcjobs.Controls.Add(Me.tpgpersonnel)
        Me.tbcjobs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbcjobs.Location = New System.Drawing.Point(0, 0)
        Me.tbcjobs.Name = "tbcjobs"
        Me.tbcjobs.SelectedIndex = 0
        Me.tbcjobs.Size = New System.Drawing.Size(602, 540)
        Me.tbcjobs.TabIndex = 68
        '
        'tpgjobdetails
        '
        Me.tpgjobdetails.Controls.Add(Me.GroupBox1)
        Me.tpgjobdetails.Controls.Add(Me.Panel1)
        Me.tpgjobdetails.Controls.Add(Me.lbClientName)
        Me.tpgjobdetails.Controls.Add(Me.lblClientNo)
        Me.tpgjobdetails.Controls.Add(Me.lblojobno)
        Me.tpgjobdetails.Controls.Add(Me.txtJobNo)
        Me.tpgjobdetails.Controls.Add(Me.lblJobNo)
        Me.tpgjobdetails.Controls.Add(Me.txtJobTitle)
        Me.tpgjobdetails.Controls.Add(Me.lblJobTitle)
        Me.tpgjobdetails.Controls.Add(Me.grpdesc)
        Me.tpgjobdetails.Location = New System.Drawing.Point(4, 23)
        Me.tpgjobdetails.Name = "tpgjobdetails"
        Me.tpgjobdetails.Size = New System.Drawing.Size(594, 513)
        Me.tpgjobdetails.TabIndex = 0
        Me.tpgjobdetails.Text = "Job details"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btndepartments)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.txtbudget)
        Me.GroupBox1.Controls.Add(Me.cboJobstatus)
        Me.GroupBox1.Controls.Add(Me.cboTechnicianresponsible)
        Me.GroupBox1.Controls.Add(Me.txtdepartment)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.lblCostAgreed)
        Me.GroupBox1.Controls.Add(Me.txtamount)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtContactName)
        Me.GroupBox1.Controls.Add(Me.lblContName)
        Me.GroupBox1.Controls.Add(Me.lblJobStatus)
        Me.GroupBox1.Controls.Add(Me.lblTechnicianResponsible)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 248)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(576, 224)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        '
        'btndepartments
        '
        Me.btndepartments.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndepartments.Location = New System.Drawing.Point(128, 89)
        Me.btndepartments.Name = "btndepartments"
        Me.btndepartments.Size = New System.Drawing.Size(32, 20)
        Me.btndepartments.TabIndex = 7
        Me.btndepartments.Text = "A"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(8, 186)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(152, 16)
        Me.Label9.TabIndex = 99
        Me.Label9.Text = "Budgetary cost"
        '
        'txtbudget
        '
        Me.txtbudget.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbudget.Location = New System.Drawing.Point(160, 184)
        Me.txtbudget.Name = "txtbudget"
        Me.txtbudget.Size = New System.Drawing.Size(232, 20)
        Me.txtbudget.TabIndex = 12
        Me.txtbudget.Text = ""
        '
        'cboJobstatus
        '
        Me.cboJobstatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboJobstatus.Location = New System.Drawing.Point(160, 136)
        Me.cboJobstatus.Name = "cboJobstatus"
        Me.cboJobstatus.Size = New System.Drawing.Size(232, 22)
        Me.cboJobstatus.TabIndex = 10
        '
        'cboTechnicianresponsible
        '
        Me.cboTechnicianresponsible.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTechnicianresponsible.Location = New System.Drawing.Point(160, 112)
        Me.cboTechnicianresponsible.Name = "cboTechnicianresponsible"
        Me.cboTechnicianresponsible.Size = New System.Drawing.Size(232, 22)
        Me.cboTechnicianresponsible.TabIndex = 9
        '
        'txtdepartment
        '
        Me.txtdepartment.Items.AddRange(New Object() {"", "Location Based Application", "Desktop Application", "Web Application"})
        Me.txtdepartment.Location = New System.Drawing.Point(160, 88)
        Me.txtdepartment.Name = "txtdepartment"
        Me.txtdepartment.Size = New System.Drawing.Size(232, 22)
        Me.txtdepartment.TabIndex = 8
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.lblgrossmargin)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(400, 64)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(168, 144)
        Me.GroupBox3.TabIndex = 94
        Me.GroupBox3.TabStop = False
        '
        'Label7
        '
        Me.Label7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(144, 16)
        Me.Label7.TabIndex = 93
        Me.Label7.Text = "Gross margin(%)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblgrossmargin
        '
        Me.lblgrossmargin.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblgrossmargin.BackColor = System.Drawing.Color.White
        Me.lblgrossmargin.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblgrossmargin.Location = New System.Drawing.Point(8, 32)
        Me.lblgrossmargin.Name = "lblgrossmargin"
        Me.lblgrossmargin.Size = New System.Drawing.Size(152, 104)
        Me.lblgrossmargin.TabIndex = 84
        Me.lblgrossmargin.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.lblOpenjournal)
        Me.GroupBox2.Controls.Add(Me.btndeletelead)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(8, 16)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(560, 48)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        '
        'lblOpenjournal
        '
        Me.lblOpenjournal.BackColor = System.Drawing.Color.Transparent
        Me.lblOpenjournal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpenjournal.Location = New System.Drawing.Point(8, 19)
        Me.lblOpenjournal.Name = "lblOpenjournal"
        Me.lblOpenjournal.Size = New System.Drawing.Size(136, 16)
        Me.lblOpenjournal.TabIndex = 86
        Me.lblOpenjournal.TabStop = True
        Me.lblOpenjournal.Text = "Open Journal"
        '
        'btndeletelead
        '
        Me.btndeletelead.BackColor = System.Drawing.Color.IndianRed
        Me.btndeletelead.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndeletelead.Location = New System.Drawing.Point(164, 16)
        Me.btndeletelead.Name = "btndeletelead"
        Me.btndeletelead.Size = New System.Drawing.Size(220, 20)
        Me.btndeletelead.TabIndex = 5
        Me.btndeletelead.Text = "Delete this Job"
        '
        'lblCostAgreed
        '
        Me.lblCostAgreed.BackColor = System.Drawing.Color.Transparent
        Me.lblCostAgreed.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCostAgreed.Location = New System.Drawing.Point(8, 161)
        Me.lblCostAgreed.Name = "lblCostAgreed"
        Me.lblCostAgreed.Size = New System.Drawing.Size(152, 16)
        Me.lblCostAgreed.TabIndex = 92
        Me.lblCostAgreed.Text = "Income"
        '
        'txtamount
        '
        Me.txtamount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtamount.Location = New System.Drawing.Point(160, 160)
        Me.txtamount.Name = "txtamount"
        Me.txtamount.Size = New System.Drawing.Size(232, 20)
        Me.txtamount.TabIndex = 11
        Me.txtamount.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 16)
        Me.Label4.TabIndex = 90
        Me.Label4.Text = "Department"
        '
        'txtContactName
        '
        Me.txtContactName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtContactName.Location = New System.Drawing.Point(160, 64)
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.Size = New System.Drawing.Size(232, 20)
        Me.txtContactName.TabIndex = 6
        Me.txtContactName.Text = ""
        '
        'lblContName
        '
        Me.lblContName.BackColor = System.Drawing.Color.Transparent
        Me.lblContName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContName.Location = New System.Drawing.Point(8, 72)
        Me.lblContName.Name = "lblContName"
        Me.lblContName.Size = New System.Drawing.Size(151, 16)
        Me.lblContName.TabIndex = 88
        Me.lblContName.Text = "Contact Name"
        '
        'lblJobStatus
        '
        Me.lblJobStatus.Font = New System.Drawing.Font("Bookman Old Style", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobStatus.Location = New System.Drawing.Point(8, 138)
        Me.lblJobStatus.Name = "lblJobStatus"
        Me.lblJobStatus.Size = New System.Drawing.Size(152, 16)
        Me.lblJobStatus.TabIndex = 87
        Me.lblJobStatus.Text = "Job Status"
        '
        'lblTechnicianResponsible
        '
        Me.lblTechnicianResponsible.BackColor = System.Drawing.Color.Transparent
        Me.lblTechnicianResponsible.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTechnicianResponsible.Location = New System.Drawing.Point(8, 114)
        Me.lblTechnicianResponsible.Name = "lblTechnicianResponsible"
        Me.lblTechnicianResponsible.Size = New System.Drawing.Size(152, 16)
        Me.lblTechnicianResponsible.TabIndex = 86
        Me.lblTechnicianResponsible.Text = "Technician Responsible"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnEditJobs)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 481)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(594, 32)
        Me.Panel1.TabIndex = 13
        '
        'btnEditJobs
        '
        Me.btnEditJobs.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnEditJobs.Location = New System.Drawing.Point(8, 6)
        Me.btnEditJobs.Name = "btnEditJobs"
        Me.btnEditJobs.Size = New System.Drawing.Size(88, 23)
        Me.btnEditJobs.TabIndex = 14
        Me.btnEditJobs.Text = "Save changes"
        '
        'btnClose
        '
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Location = New System.Drawing.Point(526, 5)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 23)
        Me.btnClose.TabIndex = 15
        Me.btnClose.Text = "Close"
        '
        'lbClientName
        '
        Me.lbClientName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbClientName.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lbClientName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbClientName.Location = New System.Drawing.Point(120, 8)
        Me.lbClientName.Name = "lbClientName"
        Me.lbClientName.Size = New System.Drawing.Size(464, 16)
        Me.lbClientName.TabIndex = 73
        '
        'lblClientNo
        '
        Me.lblClientNo.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lblClientNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientNo.Location = New System.Drawing.Point(16, 8)
        Me.lblClientNo.Name = "lblClientNo"
        Me.lblClientNo.Size = New System.Drawing.Size(96, 16)
        Me.lblClientNo.TabIndex = 72
        '
        'lblojobno
        '
        Me.lblojobno.BackColor = System.Drawing.SystemColors.Window
        Me.lblojobno.Location = New System.Drawing.Point(187, 28)
        Me.lblojobno.Name = "lblojobno"
        Me.lblojobno.Size = New System.Drawing.Size(140, 20)
        Me.lblojobno.TabIndex = 70
        '
        'txtJobNo
        '
        Me.txtJobNo.Enabled = False
        Me.txtJobNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtJobNo.Location = New System.Drawing.Point(104, 28)
        Me.txtJobNo.Name = "txtJobNo"
        Me.txtJobNo.Size = New System.Drawing.Size(80, 20)
        Me.txtJobNo.TabIndex = 71
        Me.txtJobNo.TabStop = False
        Me.txtJobNo.Text = ""
        '
        'lblJobNo
        '
        Me.lblJobNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobNo.Location = New System.Drawing.Point(16, 28)
        Me.lblJobNo.Name = "lblJobNo"
        Me.lblJobNo.Size = New System.Drawing.Size(80, 16)
        Me.lblJobNo.TabIndex = 69
        Me.lblJobNo.Text = "Job No"
        '
        'txtJobTitle
        '
        Me.txtJobTitle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtJobTitle.Location = New System.Drawing.Point(104, 52)
        Me.txtJobTitle.Name = "txtJobTitle"
        Me.txtJobTitle.Size = New System.Drawing.Size(224, 20)
        Me.txtJobTitle.TabIndex = 0
        Me.txtJobTitle.Text = ""
        '
        'lblJobTitle
        '
        Me.lblJobTitle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobTitle.Location = New System.Drawing.Point(16, 52)
        Me.lblJobTitle.Name = "lblJobTitle"
        Me.lblJobTitle.Size = New System.Drawing.Size(80, 16)
        Me.lblJobTitle.TabIndex = 68
        Me.lblJobTitle.Text = "Job Title"
        '
        'grpdesc
        '
        Me.grpdesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpdesc.Controls.Add(Me.txtdesc)
        Me.grpdesc.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpdesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpdesc.Location = New System.Drawing.Point(8, 77)
        Me.grpdesc.Name = "grpdesc"
        Me.grpdesc.Size = New System.Drawing.Size(576, 171)
        Me.grpdesc.TabIndex = 1
        Me.grpdesc.TabStop = False
        Me.grpdesc.Text = "Description"
        '
        'txtdesc
        '
        Me.txtdesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdesc.Location = New System.Drawing.Point(8, 16)
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(560, 144)
        Me.txtdesc.TabIndex = 2
        Me.txtdesc.Text = ""
        '
        'tpgequip
        '
        Me.tpgequip.Controls.Add(Me.pnlhiredequip)
        Me.tpgequip.Controls.Add(Me.pnlramaniequip)
        Me.tpgequip.Location = New System.Drawing.Point(4, 22)
        Me.tpgequip.Name = "tpgequip"
        Me.tpgequip.Size = New System.Drawing.Size(594, 514)
        Me.tpgequip.TabIndex = 2
        Me.tpgequip.Text = "Equipment"
        Me.tpgequip.Visible = False
        '
        'pnlhiredequip
        '
        Me.pnlhiredequip.Controls.Add(Me.btndelhiredequip)
        Me.pnlhiredequip.Controls.Add(Me.txthiredequip)
        Me.pnlhiredequip.Controls.Add(Me.dtghiredequip)
        Me.pnlhiredequip.Controls.Add(Me.btnhired)
        Me.pnlhiredequip.Controls.Add(Me.Label1)
        Me.pnlhiredequip.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlhiredequip.Location = New System.Drawing.Point(0, 312)
        Me.pnlhiredequip.Name = "pnlhiredequip"
        Me.pnlhiredequip.Size = New System.Drawing.Size(594, 202)
        Me.pnlhiredequip.TabIndex = 5
        '
        'btndelhiredequip
        '
        Me.btndelhiredequip.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelhiredequip.Location = New System.Drawing.Point(160, 1)
        Me.btndelhiredequip.Name = "btndelhiredequip"
        Me.btndelhiredequip.Size = New System.Drawing.Size(136, 20)
        Me.btndelhiredequip.TabIndex = 7
        Me.btndelhiredequip.Text = "Delete selected row"
        '
        'txthiredequip
        '
        Me.txthiredequip.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txthiredequip.LabelText = "Total:"
        Me.txthiredequip.Location = New System.Drawing.Point(453, 169)
        Me.txthiredequip.Name = "txthiredequip"
        Me.txthiredequip.Size = New System.Drawing.Size(144, 24)
        Me.txthiredequip.TabIndex = 9
        Me.txthiredequip.TextBoxText = ""
        '
        'dtghiredequip
        '
        Me.dtghiredequip.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtghiredequip.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dtghiredequip.CaptionVisible = False
        Me.dtghiredequip.DataMember = ""
        Me.dtghiredequip.FlatMode = True
        Me.dtghiredequip.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtghiredequip.Location = New System.Drawing.Point(4, 24)
        Me.dtghiredequip.Name = "dtghiredequip"
        Me.dtghiredequip.PreferredRowHeight = 20
        Me.dtghiredequip.ReadOnly = True
        Me.dtghiredequip.Size = New System.Drawing.Size(584, 137)
        Me.dtghiredequip.TabIndex = 8
        '
        'btnhired
        '
        Me.btnhired.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnhired.Location = New System.Drawing.Point(128, 3)
        Me.btnhired.Name = "btnhired"
        Me.btnhired.Size = New System.Drawing.Size(24, 16)
        Me.btnhired.TabIndex = 6
        Me.btnhired.Text = "+"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(4, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 11)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Hired equipment"
        '
        'pnlramaniequip
        '
        Me.pnlramaniequip.Controls.Add(Me.btndelramequip)
        Me.pnlramaniequip.Controls.Add(Me.txtramani)
        Me.pnlramaniequip.Controls.Add(Me.dtgramaniequip)
        Me.pnlramaniequip.Controls.Add(Me.btnramaniequip)
        Me.pnlramaniequip.Controls.Add(Me.Label5)
        Me.pnlramaniequip.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlramaniequip.Location = New System.Drawing.Point(0, 0)
        Me.pnlramaniequip.Name = "pnlramaniequip"
        Me.pnlramaniequip.Size = New System.Drawing.Size(594, 312)
        Me.pnlramaniequip.TabIndex = 0
        '
        'btndelramequip
        '
        Me.btndelramequip.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelramequip.Location = New System.Drawing.Point(160, 1)
        Me.btndelramequip.Name = "btndelramequip"
        Me.btndelramequip.Size = New System.Drawing.Size(136, 20)
        Me.btndelramequip.TabIndex = 2
        Me.btndelramequip.Text = "Delete selected row"
        '
        'txtramani
        '
        Me.txtramani.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtramani.LabelText = "Total:"
        Me.txtramani.Location = New System.Drawing.Point(453, 281)
        Me.txtramani.Name = "txtramani"
        Me.txtramani.Size = New System.Drawing.Size(144, 24)
        Me.txtramani.TabIndex = 4
        Me.txtramani.TextBoxText = ""
        '
        'dtgramaniequip
        '
        Me.dtgramaniequip.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgramaniequip.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dtgramaniequip.CaptionVisible = False
        Me.dtgramaniequip.DataMember = ""
        Me.dtgramaniequip.FlatMode = True
        Me.dtgramaniequip.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgramaniequip.Location = New System.Drawing.Point(5, 24)
        Me.dtgramaniequip.Name = "dtgramaniequip"
        Me.dtgramaniequip.PreferredRowHeight = 20
        Me.dtgramaniequip.ReadOnly = True
        Me.dtgramaniequip.Size = New System.Drawing.Size(584, 256)
        Me.dtgramaniequip.TabIndex = 3
        '
        'btnramaniequip
        '
        Me.btnramaniequip.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnramaniequip.Location = New System.Drawing.Point(128, 5)
        Me.btnramaniequip.Name = "btnramaniequip"
        Me.btnramaniequip.Size = New System.Drawing.Size(24, 16)
        Me.btnramaniequip.TabIndex = 1
        Me.btnramaniequip.Text = "+"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(7, 5)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(113, 11)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Ramani equipment"
        '
        'tpgfinance
        '
        Me.tpgfinance.Controls.Add(Me.pnltravel)
        Me.tpgfinance.Controls.Add(Me.pnlaccomodation)
        Me.tpgfinance.Controls.Add(Me.pnlcasuals)
        Me.tpgfinance.Location = New System.Drawing.Point(4, 22)
        Me.tpgfinance.Name = "tpgfinance"
        Me.tpgfinance.Size = New System.Drawing.Size(594, 514)
        Me.tpgfinance.TabIndex = 1
        Me.tpgfinance.Text = "Financial details"
        Me.tpgfinance.Visible = False
        '
        'pnltravel
        '
        Me.pnltravel.Controls.Add(Me.btndeltravel)
        Me.pnltravel.Controls.Add(Me.txttravel)
        Me.pnltravel.Controls.Add(Me.dtgtravel)
        Me.pnltravel.Controls.Add(Me.btntravel)
        Me.pnltravel.Controls.Add(Me.Label3)
        Me.pnltravel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnltravel.Location = New System.Drawing.Point(0, 362)
        Me.pnltravel.Name = "pnltravel"
        Me.pnltravel.Size = New System.Drawing.Size(594, 152)
        Me.pnltravel.TabIndex = 10
        '
        'btndeltravel
        '
        Me.btndeltravel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeltravel.Location = New System.Drawing.Point(122, 2)
        Me.btndeltravel.Name = "btndeltravel"
        Me.btndeltravel.Size = New System.Drawing.Size(136, 20)
        Me.btndeltravel.TabIndex = 12
        Me.btndeltravel.Text = "Delete selected row"
        '
        'txttravel
        '
        Me.txttravel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txttravel.LabelText = "Total:"
        Me.txttravel.Location = New System.Drawing.Point(456, 124)
        Me.txttravel.Name = "txttravel"
        Me.txttravel.Size = New System.Drawing.Size(128, 24)
        Me.txttravel.TabIndex = 14
        Me.txttravel.TextBoxText = ""
        '
        'dtgtravel
        '
        Me.dtgtravel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgtravel.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dtgtravel.CaptionVisible = False
        Me.dtgtravel.DataMember = ""
        Me.dtgtravel.FlatMode = True
        Me.dtgtravel.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgtravel.Location = New System.Drawing.Point(8, 24)
        Me.dtgtravel.Name = "dtgtravel"
        Me.dtgtravel.PreferredRowHeight = 20
        Me.dtgtravel.ReadOnly = True
        Me.dtgtravel.Size = New System.Drawing.Size(584, 100)
        Me.dtgtravel.TabIndex = 13
        '
        'btntravel
        '
        Me.btntravel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btntravel.Location = New System.Drawing.Point(94, 3)
        Me.btntravel.Name = "btntravel"
        Me.btntravel.Size = New System.Drawing.Size(24, 16)
        Me.btntravel.TabIndex = 11
        Me.btntravel.Text = "+"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(5, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 11)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Travel"
        '
        'pnlaccomodation
        '
        Me.pnlaccomodation.Controls.Add(Me.btndelaccom)
        Me.pnlaccomodation.Controls.Add(Me.txtaccomodation)
        Me.pnlaccomodation.Controls.Add(Me.dtgaccomodation)
        Me.pnlaccomodation.Controls.Add(Me.btnaccomodation)
        Me.pnlaccomodation.Controls.Add(Me.Label2)
        Me.pnlaccomodation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlaccomodation.Location = New System.Drawing.Point(0, 144)
        Me.pnlaccomodation.Name = "pnlaccomodation"
        Me.pnlaccomodation.Size = New System.Drawing.Size(594, 370)
        Me.pnlaccomodation.TabIndex = 5
        '
        'btndelaccom
        '
        Me.btndelaccom.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelaccom.Location = New System.Drawing.Point(122, 2)
        Me.btndelaccom.Name = "btndelaccom"
        Me.btndelaccom.Size = New System.Drawing.Size(136, 21)
        Me.btndelaccom.TabIndex = 7
        Me.btndelaccom.Text = "Delete selected row"
        '
        'txtaccomodation
        '
        Me.txtaccomodation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtaccomodation.LabelText = "Total:"
        Me.txtaccomodation.Location = New System.Drawing.Point(456, 192)
        Me.txtaccomodation.Name = "txtaccomodation"
        Me.txtaccomodation.Size = New System.Drawing.Size(128, 24)
        Me.txtaccomodation.TabIndex = 9
        Me.txtaccomodation.TextBoxText = ""
        '
        'dtgaccomodation
        '
        Me.dtgaccomodation.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dtgaccomodation.CaptionVisible = False
        Me.dtgaccomodation.DataMember = ""
        Me.dtgaccomodation.FlatMode = True
        Me.dtgaccomodation.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgaccomodation.Location = New System.Drawing.Point(8, 24)
        Me.dtgaccomodation.Name = "dtgaccomodation"
        Me.dtgaccomodation.PreferredRowHeight = 20
        Me.dtgaccomodation.ReadOnly = True
        Me.dtgaccomodation.Size = New System.Drawing.Size(584, 160)
        Me.dtgaccomodation.TabIndex = 8
        '
        'btnaccomodation
        '
        Me.btnaccomodation.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnaccomodation.Location = New System.Drawing.Point(94, 3)
        Me.btnaccomodation.Name = "btnaccomodation"
        Me.btnaccomodation.Size = New System.Drawing.Size(24, 16)
        Me.btnaccomodation.TabIndex = 6
        Me.btnaccomodation.Text = "+"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(4, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 11)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Accomodation"
        '
        'pnlcasuals
        '
        Me.pnlcasuals.Controls.Add(Me.btndelcasuals)
        Me.pnlcasuals.Controls.Add(Me.Label8)
        Me.pnlcasuals.Controls.Add(Me.txtlabour)
        Me.pnlcasuals.Controls.Add(Me.dtgcasuals)
        Me.pnlcasuals.Controls.Add(Me.btncasuals)
        Me.pnlcasuals.Controls.Add(Me.lblCasual)
        Me.pnlcasuals.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlcasuals.Location = New System.Drawing.Point(0, 0)
        Me.pnlcasuals.Name = "pnlcasuals"
        Me.pnlcasuals.Size = New System.Drawing.Size(594, 144)
        Me.pnlcasuals.TabIndex = 0
        '
        'btndelcasuals
        '
        Me.btndelcasuals.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelcasuals.Location = New System.Drawing.Point(123, 3)
        Me.btndelcasuals.Name = "btndelcasuals"
        Me.btndelcasuals.Size = New System.Drawing.Size(136, 20)
        Me.btndelcasuals.TabIndex = 2
        Me.btndelcasuals.Text = "Delete selected row"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(83, 11)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Casual labour"
        '
        'txtlabour
        '
        Me.txtlabour.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtlabour.LabelText = "Total:"
        Me.txtlabour.Location = New System.Drawing.Point(456, 117)
        Me.txtlabour.Name = "txtlabour"
        Me.txtlabour.Size = New System.Drawing.Size(128, 24)
        Me.txtlabour.TabIndex = 4
        Me.txtlabour.TextBoxText = ""
        '
        'dtgcasuals
        '
        Me.dtgcasuals.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgcasuals.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dtgcasuals.CaptionVisible = False
        Me.dtgcasuals.DataMember = ""
        Me.dtgcasuals.FlatMode = True
        Me.dtgcasuals.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgcasuals.Location = New System.Drawing.Point(8, 24)
        Me.dtgcasuals.Name = "dtgcasuals"
        Me.dtgcasuals.PreferredRowHeight = 20
        Me.dtgcasuals.ReadOnly = True
        Me.dtgcasuals.Size = New System.Drawing.Size(584, 88)
        Me.dtgcasuals.TabIndex = 3
        '
        'btncasuals
        '
        Me.btncasuals.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btncasuals.Location = New System.Drawing.Point(94, 5)
        Me.btncasuals.Name = "btncasuals"
        Me.btncasuals.Size = New System.Drawing.Size(24, 16)
        Me.btncasuals.TabIndex = 1
        Me.btncasuals.Text = "+"
        '
        'lblCasual
        '
        Me.lblCasual.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCasual.Location = New System.Drawing.Point(7, 5)
        Me.lblCasual.Name = "lblCasual"
        Me.lblCasual.Size = New System.Drawing.Size(81, 3)
        Me.lblCasual.TabIndex = 0
        Me.lblCasual.Text = "Casual labour"
        '
        'tpgpersonnel
        '
        Me.tpgpersonnel.Controls.Add(Me.Panel2)
        Me.tpgpersonnel.Controls.Add(Me.pnlctrlpersonnel)
        Me.tpgpersonnel.Location = New System.Drawing.Point(4, 22)
        Me.tpgpersonnel.Name = "tpgpersonnel"
        Me.tpgpersonnel.Size = New System.Drawing.Size(594, 514)
        Me.tpgpersonnel.TabIndex = 3
        Me.tpgpersonnel.Text = "Personnel"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtpersonelcost)
        Me.Panel2.Controls.Add(Me.dtgpersonnel)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 40)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(594, 474)
        Me.Panel2.TabIndex = 3
        '
        'txtpersonelcost
        '
        Me.txtpersonelcost.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtpersonelcost.LabelText = "Total cost:"
        Me.txtpersonelcost.Location = New System.Drawing.Point(448, 449)
        Me.txtpersonelcost.Name = "txtpersonelcost"
        Me.txtpersonelcost.Size = New System.Drawing.Size(144, 24)
        Me.txtpersonelcost.TabIndex = 5
        Me.txtpersonelcost.TextBoxText = ""
        '
        'dtgpersonnel
        '
        Me.dtgpersonnel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgpersonnel.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dtgpersonnel.CaptionVisible = False
        Me.dtgpersonnel.DataMember = ""
        Me.dtgpersonnel.FlatMode = True
        Me.dtgpersonnel.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgpersonnel.Location = New System.Drawing.Point(5, 16)
        Me.dtgpersonnel.Name = "dtgpersonnel"
        Me.dtgpersonnel.PreferredRowHeight = 20
        Me.dtgpersonnel.ReadOnly = True
        Me.dtgpersonnel.Size = New System.Drawing.Size(584, 425)
        Me.dtgpersonnel.TabIndex = 4
        '
        'pnlctrlpersonnel
        '
        Me.pnlctrlpersonnel.Controls.Add(Me.StiButton2)
        Me.pnlctrlpersonnel.Controls.Add(Me.btnprint)
        Me.pnlctrlpersonnel.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlctrlpersonnel.Location = New System.Drawing.Point(0, 0)
        Me.pnlctrlpersonnel.Name = "pnlctrlpersonnel"
        Me.pnlctrlpersonnel.Size = New System.Drawing.Size(594, 40)
        Me.pnlctrlpersonnel.TabIndex = 0
        '
        'StiButton2
        '
        Me.StiButton2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiButton2.Location = New System.Drawing.Point(106, 8)
        Me.StiButton2.Name = "StiButton2"
        Me.StiButton2.Size = New System.Drawing.Size(96, 23)
        Me.StiButton2.TabIndex = 2
        Me.StiButton2.Text = "Export to excel"
        Me.StiButton2.Visible = False
        '
        'btnprint
        '
        Me.btnprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnprint.Location = New System.Drawing.Point(8, 8)
        Me.btnprint.Name = "btnprint"
        Me.btnprint.Size = New System.Drawing.Size(96, 23)
        Me.btnprint.TabIndex = 1
        Me.btnprint.Text = "Print"
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Location = New System.Drawing.Point(7, 2)
        Me.PrintPreviewDialog1.MinimumSize = New System.Drawing.Size(375, 250)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.TransparencyKey = System.Drawing.Color.Empty
        Me.PrintPreviewDialog1.Visible = False
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmEditJob
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(602, 540)
        Me.Controls.Add(Me.tbcjobs)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmEditJob"
        Me.Text = "Change Job Details"
        Me.tbcjobs.ResumeLayout(False)
        Me.tpgjobdetails.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.grpdesc.ResumeLayout(False)
        Me.tpgequip.ResumeLayout(False)
        Me.pnlhiredequip.ResumeLayout(False)
        CType(Me.dtghiredequip, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlramaniequip.ResumeLayout(False)
        CType(Me.dtgramaniequip, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpgfinance.ResumeLayout(False)
        Me.pnltravel.ResumeLayout(False)
        CType(Me.dtgtravel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlaccomodation.ResumeLayout(False)
        CType(Me.dtgaccomodation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlcasuals.ResumeLayout(False)
        CType(Me.dtgcasuals, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpgpersonnel.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.dtgpersonnel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlctrlpersonnel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub

#Region "edit jobs"
    Private Sub loadtechnicians()
        Dim connect As New ADODB.Connection
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Try

            Dim cmd As New ADODB.Command
            Dim r As ADODB.Recordset
            With cmd
                .CommandType = CommandTypeEnum.adCmdText
                .ActiveConnection = connect
                Dim mystr, mystr1 As String
                'mystr1 = txtJobNo.Text.Trim.ToUpper
                mystr = "select name from seccheck"
                .CommandText = mystr
                r = .Execute
            End With
            With r
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Me.cboTechnicianresponsible.Items.Add(.Fields(0).Value)
                        .MoveNext()
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
    Private Sub GroupBox4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Private Function jobno() As String
        Dim connect As New ADODB.Connection
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Function
        End Try

        Try
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim str As String
            With cmd
                .ActiveConnection = connect
                .CommandType = ADODB.CommandTypeEnum.adCmdText

                str = " select max(job_no) from rcljobs"
                .CommandText = str
                rs = .Execute
            End With
            str = rs.Fields("max").Value
            Dim strno, str1 As String
            Dim i
            For i = 0 To str.Length - 1
                If IsNumeric(str.Substring(i, 1)) = True Then
                    strno = strno & str.Substring(i, 1)
                Else
                    str1 = str1 & str.Substring(i, 1)
                End If

            Next

            str = (CSng(strno) + 1).ToString
            str1 = str1.Insert(1, str)
            Return str1
            rs.Close()
            rs = Nothing
        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Sub frmEditJob_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.cboJobstatus.Items.Add("Current")
            Me.cboJobstatus.Items.Add("Delivered")
            Me.cboJobstatus.Items.Add("Completed")
            ' Me.cboJobstatus.Items.Add("Proposal")
            ' Me.cboJobstatus.Items.Add("Prospect")
            'Me.cboJobstatus.Items.Add("Suspect")
            cboJobstatus.Sorted = True
            'Me.txtJobNo.Text = jobno()
            Call loadtechnicians()
        Catch xc As Exception

        End Try
        Try
            Me._client_no = myclientno
        Catch shj As Exception

        End Try
        Try
            'cboJobstatus.DropDownStyle = ComboBoxStyle.DropDown
            Me.cboJobstatus.Text = jobstatus
            'cboTechnicianresponsible.DropDownStyle = ComboBoxStyle.DropDown
            Me.cboTechnicianresponsible.Text = tecres
            '--------
            Me.totalincome = Me.txtamount.Text
            mygross()
            '------
        Catch cv As Exception

        End Try

        '-------------automate client no
        If Me.Text = "Add jobs" Then
            Dim Tasks As New taskclass
            Tasks.clientno = myclientno
            Dim Threadme As New System.Threading.Thread( _
                AddressOf Tasks.newlnoinvoke)
            Threadme.IsBackground = True
            Threadme.Start()
        End If

        '--------------------
        Dim _thread As Thread = New Thread(AddressOf ld)
        _thread.IsBackground = True
        _thread.Start()

    End Sub
    Private Sub btnEditJobs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditJobs.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch zxcv As Exception

        End Try
        Try

            If Me.btnEditJobs.Text = "Save changes" Then
                Me.Invoke(New mydelegate(AddressOf editjobs))
            Else
                addjob()
            End If

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
    End Sub
    Private Sub threadjobs()
        Try
            myForms.CustomerForm3.Invoke(New mydelegate(AddressOf myForms.CustomerForm3.loadgridexistingjobs))

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            editjob = False
            myForms.CustomerForm2 = Nothing
            Me.Dispose(True)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub editjobs()
        Dim connect As New ADODB.Connection
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception
            Exit Sub
        End Try


        Try
            Cursor.Current = Cursors.WaitCursor
            ''''''''''''''''''''------------validation
            If Me.txtJobNo.Text = "" Or Me.txtJobNo.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="A job must have a job number", _
                caption:="Edit jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.txtdepartment.Text = "" Or Me.txtdepartment.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="Please select department", _
                caption:="Edit jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.txtJobTitle.Text = "" Or Me.txtJobTitle.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="A job must have a title", _
                caption:="Edit jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.txtamount.Text = "" Or Me.txtamount.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="Please supply amount", _
                caption:="Edit jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.cboJobstatus.Text = "" Or Me.cboJobstatus.Text.Length = 0 Then
                MessageBox.Show(Text:="Please select a job status", _
                caption:="Edit jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                Exit Try
            End If
            '---------------end of validation



            Dim rs As New ADODB.Recordset
            Dim cmd As New ADODB.Command
            With cmd
                .ActiveConnection = connect
                .CommandType = CommandTypeEnum.adCmdText
                Dim strd As String
                strd = "select job_no from rcljobs" _
                & " where job_no='" & myjobno.Trim.ToUpper & "' "
                .CommandText = strd
                rs = .Execute
            End With
            If rs.BOF = True And rs.EOF = True Then
                MessageBox.Show(Text:="Job number does not exist", _
                caption:="Save Changes failed", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                Exit Try
            End If
            rs.Close()
            rs = Nothing
            cmd.ActiveConnection = Nothing
            'Dim k() As String
            'k = Me.txtTechnicianResponsible.Lines
            'Dim mystr5 As String
            'mystr5 = rml(k)
            '// deals with dates
            'Dim a() As String
            Dim strsql, str As String
            'str = returndates()
            'a = str.Split("|")
            '------------stringbuilder function

            '------
            connect.BeginTrans()
            connect.IsolationLevel = IsolationLevelEnum.adXactSerializable

            connect.Execute(strb.ToString())
            connect.CommitTrans()
            myjobno = Me.txtJobNo.Text.Trim.ToUpper
            Me.totalincome = txtamount.Text.Trim()
            MessageBox.Show(Text:="Changes have been successfully made", _
            caption:="Save Changes", buttons:=MessageBoxButtons.OK, _
            Icon:=MessageBoxIcon.Information)
            refreshjobs = True
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
            Try
                connect.RollbackTrans()

            Catch fr As Exception

            End Try
        Finally
            'Me.txtJobNo.Text = ""
            'Me.txtJobTitle.Text = ""
            'Me.txtContactName.Text = ""
            'Me.txtJobNo.Focus()
            Cursor.Current = currentcursor
        End Try
        '---------refresh grosss margin
        Try
            myForms.CustomerForm2.mygross()
        Catch zx As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
        '----------------------
    End Sub
    Private Function strb() As StringBuilder
        Try
            strb = New StringBuilder(" update rcljobs set")
            strb.Append(" job_no='" & txtJobNo.Text.ToUpper & "',")
            strb.Append(" job_tittle='" & txtJobTitle.Text & "',")
            strb.Append(" job_status='" & cboJobstatus.Text & "',")
            strb.Append(" techres='" & cboTechnicianresponsible.Text & "',")
            strb.Append(" cont='" & txtContactName.Text & "',")

            Dim arr() As String
            Dim strr As String
            Dim y As Integer
            arr = txtdesc.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            strb.Append("descrip='" & strr & "',")
            strb.Append("department='" & txtdepartment.Text & "',")

            strb.Append("budgetarycost='" & Me.txtbudget.Text & "',")
            strb.Append("amount='" & txtamount.Text.Trim() & "'")
            strb.Append(" where job_no='" & myjobno & "'")
            'strb.Replace("'", "\'")
        Catch ex As Exception

        End Try
    End Function
    'Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
    '    Try
    '        ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
    '        ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
    '        If keyData = System.Windows.Forms.Keys.Return Then
    '            'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)
    '            If Me.TabIndex <> 3 Then
    '                Dim E As System.EventArgs
    '                'Me.Invoke(New mydelegate(AddressOf editjobs))
    '                'Call btnEditJobs_Click(Me, E)
    '            End If
    '            Return True ' True means we've processed the key
    '        Else
    '            Return MyBase.ProcessDialogKey(keyData)
    '        End If
    '    Catch ex As Exception
    '        'Trace.WriteLine(ex.ToString())
    '        MsgBox(ex.Message.ToString, , Title:="Return key")

    '    End Try
    'End Function
    Private Sub cboTechnicianresponsible_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            'System.Windows.Forms.Keys.Escape() = Keys.Escape
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btndeletelead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btndeletelead.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to delete jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch zxcv As Exception

        End Try
        Try
            If myjobno = "" Then
                MsgBox("There is no current job to be deleted", MsgBoxStyle.Information)
                Exit Try
            End If
            If MessageBox.Show("Are you sure you want to delete this job" _
                     , "Delete Job", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                Dim connectstr As String
                connectstr = "DSN=" & myForms.qconnstr
                'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
                Dim connect As New ADODB.Connection
                connect.Mode = ConnectModeEnum.adModeReadWrite
                connect.CursorLocation = CursorLocationEnum.adUseClient
                connect.ConnectionString = connectstr
                connect.Open()

                Dim strsql As String
                Dim int1, int2, int3 As Integer
                int1 = CInt(myjobno)
                int2 = CInt(myjobno) - 1
                int3 = CInt(Me._client_no)
                strsql = "select clean_rcljobs(" & int1 & "," & int2 & "," & int3 & ");" 'use of user defined function
                strsql += "delete from rcljobs where job_no='" & myjobno & "';"
                connect.BeginTrans()
                connect.Execute(strsql)
                connect.CommitTrans()

                myjobno = ""
                refreshjobs = True
                Me.txtamount.Text = ""
                Me.txtdesc.Text = ""
                Me.txtJobNo.Text = ""
                Me.txtJobTitle.Text = ""
                Me.cboJobstatus.Text = ""
                Me.cboTechnicianresponsible.Text = ""
                Me.lblojobno.Text = ""
                Me.txtContactName.Text = ""
            End If

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
    End Sub
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
                myForms.CustomerForm2.txtdepartment.Items.Clear()
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While rs.EOF = False
                        myForms.CustomerForm2.txtdepartment.Items.Add(.Fields("dept").Value)
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
            myForms.CustomerForm2.Invoke(New mydelegate(AddressOf loaddepart))
        Catch ex As Exception

        End Try
    End Sub
#End Region

#End Region

#Region "Add jobs"
    Private Sub addjob()
        Dim currentcursor As Cursor = Cursor.Current
        Dim isvalid As Boolean = True
        Try
            Cursor.Current = Cursors.WaitCursor
            ''''''''''''''''''''------------validation
            If Me.txtJobNo.Text = "" Or Me.txtJobNo.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="A job must have a job number", _
                caption:="Add jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                isvalid = False
                Exit Try
            End If
            If Me.txtdepartment.Text = "" Or Me.txtdepartment.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="Please select department", _
                caption:="Add jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                isvalid = False
                Exit Try
            End If
            If Me.txtJobTitle.Text = "" Or Me.txtJobTitle.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="A job must have a title", _
                caption:="Add jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                isvalid = False
                Exit Try
            End If
            If Me.txtamount.Text = "" Or Me.txtamount.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="Please supply amount", _
                caption:="Add jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                isvalid = False
                Exit Try
            End If
            '---------------end of validation
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

            Dim cmd As New ADODB.Command
            Dim r As ADODB.Recordset
            With cmd
                .CommandType = CommandTypeEnum.adCmdText
                .ActiveConnection = connect
                Dim mystr, mystr1 As String
                mystr1 = txtJobNo.Text.Trim.ToUpper
                mystr = "select job_no from rcljobs where job_no like '%" & mystr1 & "%'"
                .CommandText = mystr
                r = .Execute
            End With
            If r.BOF = False And r.EOF = False Then
                MessageBox.Show(Text:="A similar job number exists" & vbCrLf & "Input a different number", _
               caption:="Add jobs", buttons:=MessageBoxButtons.OK, _
               Icon:=MessageBoxIcon.Information)
                isvalid = False
                Exit Try
            End If

            If Me.cboJobstatus.Text = "" Or Me.cboJobstatus.Text.Length = 0 Then
                MessageBox.Show(Text:="Please select  job status", _
                caption:="Add jobs", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
                isvalid = False
                Exit Try
            End If
            Dim strsql, str As String
            ' txtJobNo.Text = newlno(myclientno)
            '-----------------
            Call checkdirectory()
            '--------------------
            connect.BeginTrans()
            connect.IsolationLevel = IsolationLevelEnum.adXactSerializable
            strsql = "insert into rcljobs"
            strsql = strsql & "(job_no,job_tittle,job_status," _
            & " client_no,cont,techres,descrip,amount,department, budgetarycost)"
            strsql = strsql & " values"
            strsql = strsql & "("
            strsql = strsql & "'" & txtJobNo.Text.ToUpper & "',"
            strsql = strsql & "'" & txtJobTitle.Text & "',"
            strsql = strsql & "'" & cboJobstatus.Text & "',"
            strsql = strsql & "'" & lblClientNo.Text & "',"
            strsql = strsql & "'" & txtContactName.Text & "',"
            strsql = strsql & "'" & cboTechnicianresponsible.Text & "',"
            strsql = strsql & "'" & txtdesc.Text & "',"
            strsql += "'" & txtamount.Text.Trim() & "',"
            strsql = strsql & "'" & txtdepartment.Text & "',"
            strsql = strsql & "'" & txtbudget.Text & "'"
            strsql = strsql & ");"
            strsql += "update clients set leads_no='" & txtJobNo.Text & "'"
            strsql += " where client_no='" & myclientno & "';"
            strsql += " insert into grossmargin (job_no) values ('" & txtJobNo.Text.ToUpper & "');"

            connect.Execute(strsql)
            connect.CommitTrans()
            MessageBox.Show(Text:="Job successfully added", _
            caption:="Add Jobs", buttons:=MessageBoxButtons.OK, _
            Icon:=MessageBoxIcon.Information)
            connect.Close()
            refreshjobs = True
        Catch ex As Exception

        Finally
            Cursor.Current = currentcursor

        End Try

        If isvalid = True Then
            'Me.txtJobNo.Text = ""
            txtJobTitle.Text = ""
            txtContactName.Text = ""
            txtamount.Text = ""
            cboJobstatus.SelectedIndex = -1
            cboTechnicianresponsible.SelectedIndex = -1
            txtdepartment.SelectedIndex = -1
            Me.txtbudget.Text = ""
            txtdesc.Clear()
            txtJobNo.Focus()

        End If
        Try
            Dim Tasks As New taskclass   '-------------automate job no

            Tasks.clientno = Me.lblClientNo.Text.Trim
            Dim Threadme As New System.Threading.Thread( _
                AddressOf Tasks.newlnoinvoke)
            Threadme.IsBackground = True
            Threadme.Start()

            Dim Threadleads As Thread = New System.Threading.Thread( _
                                                                         AddressOf loaddirectory)
            Threadleads.IsBackground = True
            Threadleads.Start()
        Catch jb As Exception
        End Try



    End Sub
    Private Sub loaddirectory()
        Try
            Me.Invoke(New mydelegate1(AddressOf checkdirectory))
        Catch ex As Exception

        End Try
    End Sub
    Private Sub checkdirectory()
        Try
            Dim myvar As String = "value=" & myForms.qfolderpath
            'str = Configuration.ConfigurationSettings.AppSettings("folderpath")
            Dim leadno As String = txtJobNo.Text
            Dim myfile, mypath
            mypath = myvar
            mypath = mypath & "\"
            mypath = mypath & myclientno

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
            ''Me.txtMailBody.SaveFile(str & Me.txtMailSubject.Text, System.Windows.Forms.RichTextBoxStreamType.RichText)
            Dim i, j
            Dim str1, myfilename As String
            Dim a()
            'For i = 0 To Me.lstAttachment.Items.Count - 1


            '    lstAttachment.SetSelected(i, True)
            '    str1 = lstAttachment.SelectedItems(i).ToString()
            '    a = str1.Split("\")
            '    j = a.GetUpperBound(0)
            '    myfilename = a(j)
            '    Try
            '        File.Copy(str1, str & "\" & myfilename)
            '    Catch ex As Exception
            '    End Try
            'Next
            Dim dtp As New System.Windows.Forms.DateTimePicker
            Dim journalpath As String
            Dim rtbjournal As New System.Windows.Forms.RichTextBox
            journalpath = str & "\" & txtJobNo.Text & "_" & dtp.Value.Year & dtp.Value.Month & dtp.Value.Day _
            & dtp.Value.Hour & dtp.Value.Minute & dtp.Value.Second & dtp.Value.Millisecond & ".txt"
            rtbjournal.SaveFile(journalpath)

        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Personnel"
    Private Sub loadpersonnel()
        Try
            Dim Tasks As New taskclass
            Tasks.strjobno = Me.txtJobNo.Text.Trim
            Dim Threadj5 As New System.Threading.Thread( _
                AddressOf Tasks.jobsinvoke)
            Threadj5.IsBackground = True
            Threadj5.Start()
        Catch wq As Exception

        End Try
    End Sub
    Private Sub print1()
        Try
            'i did'nt touch most of this code right here
            Dim aCol As Integer
            Dim aTblIndex As Integer = 0
            Dim aNumColumns As Integer
            Dim aColumnStyles As GridColumnStylesCollection
            '----------
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgpersonnel.DataSource
            Try
                '
                ' The dataset may have columns not displayed in the grid. Looping through 
                ' GridColumnStyles errors when the number of GridColumnStyles is exceeded. 
                ' The error is caught and the number of displayed columns is adjusted. The 
                ' GridColumnStyles collection doesn't expose a Count property nor does the 
                ' DataGrid expose the number of displayed columns.
                '
                '------------------remove unwanted columns
                aNumColumns = ds.Tables(aTblIndex).Columns.Count - 1
                Try
                    Dim strv As String
                    With ds.Tables(aTblIndex)
                        For aCol = 0 To aNumColumns
                            strv = .Columns(aCol).ColumnName
                            If .Columns(aCol).ColumnName = "namme" Or _
                            .Columns(aCol).ColumnName = "task" Or _
                            .Columns(aCol).ColumnName = "timespent" Or _
                            .Columns(aCol).ColumnName = "Cost" Then

                            Else
                                Try
                                    .Columns.Remove(strv)
                                    aCol = aCol - 1
                                    aNumColumns = aNumColumns - 1
                                Catch xc As Exception

                                End Try
                            End If

                        Next
                    End With
                Catch zx As Exception

                End Try
                ''pass a new tablestyle
                maddtablestylepersonnel(ds.Tables(0).TableName)
                '-----------------------
                '--------------------
                Dim intcount As Integer = ds.Tables(0).Rows.Count - 1
                Dim myrow As System.Data.DataRow = ds.Tables(0).NewRow
                ds.Tables(0).Rows.Add(myrow)
                intcount = ds.Tables(0).Rows.Count - 1
                Dim myrow1 As System.Data.DataRow = ds.Tables(0).NewRow
                ds.Tables(0).Rows.Add(myrow1)
                intcount = ds.Tables(0).Rows.Count - 1
                Try
                    Dim nb As Double = Convert.ToDouble(Me.txtpersonelcost.TextBoxText)
                    Dim f = Math.Round(Convert.ToDecimal(nb), 2)
                    Me.txtpersonelcost.TextBoxText = "Total : " & f
                Catch zx As Exception

                End Try
                ds.Tables(0).Rows(intcount).Item("Cost") = Me.txtpersonelcost.TextBoxText

                '---------------------------

                aColumnStyles = Me.dtgpersonnel.TableStyles(aTblIndex).GridColumnStyles
                aNumColumns = ds.Tables(aTblIndex).Columns.Count - 1

                With ds.Tables(aTblIndex)
                    For aCol = 0 To aNumColumns
                        .Columns(aCol).Caption = aColumnStyles.Item(aCol).HeaderText
                        .Columns.Item(aCol).ExtendedProperties.Clear()

                        If aColumnStyles.Item(aCol).Width = 0 Then
                            .Columns.Item(aCol).ExtendedProperties.Add("PrintWidth", -1)
                        Else
                            .Columns.Item(aCol).ExtendedProperties.Add( _
                                "PrintWidth", aColumnStyles.Item(aCol).Width)
                        End If
                    Next
                End With

            Catch Ex As System.ArgumentOutOfRangeException
                aNumColumns = aCol - 1

            Catch Ex As Exception
                Throw New Exception("Error setting column captions.", Ex)

            End Try

            '
            ' Call the PrintHandler.
            '
            Dim aPrintObj As New PrintHandler

            With aPrintObj
                .LineThreshold = 500
                .NumberOfColumns = aNumColumns
                .DataSetToPrint = ds
                .ReportTitle = "Ramani Communications"
                .DataSetToPrint = ds
                .TableIndex = aTblIndex
                'If blnPreview Then
                .PrintPreview()
                'Else
                '.Print()
                'End If
            End With

            aPrintObj = Nothing
        Catch za As Exception

        End Try
        '-------------reinitialize dataset
        Try
            Dim Tasks As New taskclass
            Tasks.strjobno = Me.txtJobNo.Text.Trim
            Dim Threadmn As New System.Threading.Thread( _
                AddressOf Tasks.jobsinvoke)
            Threadmn.IsBackground = True
            Threadmn.Start()
        Catch wq As Exception

        End Try
    End Sub
    Private GridPrinter As DataGridPrinter
    Private Sub btnprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnprint.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            myForms.Main.dgrid.DataSource = Nothing
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = dtgpersonnel.DataSource
            myForms.Main.dgrid.DataSource = ds
            Try
                ds.Tables(0).Columns.Remove("job_no")
                ds.Tables(0).Columns.Remove("id_no")
                ds.Tables(0).Columns.Remove("description")
                ds.Tables(0).Columns.Remove("ddate")
                ds.Tables(0).Columns.Remove("milliseconds")
                ds.Tables(0).Columns.Remove("ano")
                ds.Tables(0).Columns.Remove("hourly_rate")
            Catch xcvb As Exception

            End Try

            Dim MyTable As New DataTable
            MyTable = ds.Tables(0)
            myForms.Main.dgrid.DataSource = MyTable

            GridPrinter = New DataGridPrinter(myForms.Main.dgrid)
            With GridPrinter
                .HeaderText = "Time sheet "
                .HeaderHeightPercent = CInt(myForms.Main.NumericUpDown_HeaderHeightPercentage.Value)
                .FooterHeightPercent = CInt(myForms.Main.NumericUpDown_FooterHeightPercent.Value)
                .InterSectionSpacingPercent = CInt(myForms.Main.NumericUpDown_InterSectionSpacingPercent.Value)
                .HeaderPen = New Pen(CType(myForms.Main.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                .FooterPen = New Pen(CType(myForms.Main.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                .GridPen = New Pen(CType(myForms.Main.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                .HeaderBrush = CType(myForms.Main.ComboBox_HeaderBrush.SelectedItem, Brush)
                .EvenRowBrush = CType(myForms.Main.ComboBox_EvenBrush.SelectedItem, Brush)
                .OddRowBrush = CType(myForms.Main.ComboBox_OddRowBrush.SelectedItem, Brush)
                .FooterBrush = CType(myForms.Main.ComboBox_FooterBrush.SelectedItem, Brush)
                .ColumnHeaderBrush = CType(myForms.Main.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                .PagesAcross = 1


            End With

            With Me.PrintPreviewDialog1
                .Document = GridPrinter.PrintDocument
                If .ShowDialog = DialogResult.OK Then
                    GridPrinter.Print()
                End If
            End With

        Catch xz As Exception

        End Try
    End Sub
    Public Sub maddtablestylepersonnel(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.CustomerForm2.dtgpersonnel.Width - 20
            mywidth = mywidth / 4

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "task"
            myname.HeaderText = "Task"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn
            myname100.MappingName = "timespent"
            myname100.HeaderText = "Time taken(hrs)"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)

            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "namme"
            myno.HeaderText = "Done by"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)



            ' Add a second column style.
            Dim myname100x As New DataGridTextBoxColumn
            myname100x.MappingName = "Cost"
            myname100x.HeaderText = "Cost"
            myname100x.Width = mywidth
            ts1.GridColumnStyles.Add(myname100x)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.CustomerForm2.dtgpersonnel.TableStyles.Clear()
            myForms.CustomerForm2.dtgpersonnel.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "all"
    Private isjobdetails As Boolean = False
    Private ispersonnel As Boolean = False
    Private isfinance As Boolean = False
    Private isequip As Boolean = False
    Private Sub tbcjobs_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbcjobs.SelectedIndexChanged
        Try

            If tbcjobs.SelectedTab Is Me.tpgjobdetails Then
                Try

                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                      & ev.InnerException().ToString() & vbCrLf _
                      & ev.StackTrace.ToString())
                End Try

            ElseIf tbcjobs.SelectedTab Is Me.tpgequip Then
                Try
                    If isequip = False Then
                        loadequip()
                        isequip = True
                    End If


                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                    & ev.InnerException().ToString() & vbCrLf _
                     & ev.StackTrace.ToString())
                End Try
            ElseIf tbcjobs.SelectedTab Is Me.tpgfinance Then

                Try
                    If isfinance = False Then
                        loadfinance()
                        isfinance = True
                    End If


                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                    & ev.InnerException().ToString() & vbCrLf _
                     & ev.StackTrace.ToString())
                End Try
            ElseIf tbcjobs.SelectedTab Is Me.tpgpersonnel Then
                Try
                    If ispersonnel = False Then
                        loadpersonnel()
                        ispersonnel = True
                    End If

                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                    & ev.InnerException().ToString() & vbCrLf _
                    & ev.StackTrace.ToString())
                End Try



            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Function canmanipulatejobs() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strjobss.Split(",")
            If arr(1) = "1" Then
                canmanipulatejobs = True
            Else
                canmanipulatejobs = False
            End If
        Catch ex As Exception
            Try
                canmanipulatejobs = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "finances"
    Private Sub btncasuals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncasuals.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim cv As New frmaddcasual
            cv.jobno = Me.txtJobNo.Text.Trim
            cv.ShowDialog()
        Catch xc As Exception

        End Try

    End Sub
    Private Sub loadfinance()
        Try
            Dim Tasks As New taskclass
            Tasks.strjobno = Me.txtJobNo.Text
            Dim Threadf5 As New System.Threading.Thread( _
            AddressOf Tasks.casualsinvoke)
            Threadf5.IsBackground = True
            Threadf5.Start()

            Dim Threadfg As New System.Threading.Thread( _
            AddressOf Tasks.accomodationinvoke)
            Threadfg.IsBackground = True
            Threadfg.Start()

            Dim Threadfb As New System.Threading.Thread( _
            AddressOf Tasks.travelinvoke)
            Threadfb.IsBackground = True
            Threadfb.Start()
        Catch wq As Exception

        End Try
    End Sub
    Private Sub dtgaccomodation_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgaccomodation.DoubleClick
        Try
            If htiaccomodation.Type = Windows.Forms.DataGrid.HitTestType.Caption _
           Or htiaccomodation.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
           Or htiaccomodation.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
           Or htiaccomodation.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
           Or htiaccomodation.Type = Windows.Forms.DataGrid.HitTestType.None _
           Or htiaccomodation.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                 Then
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgaccomodation.DataSource
            Try
                myForms.accomodation.Text = "Edit accomodation"
                myForms.accomodation.jobno = txtJobNo.Text
                myForms.accomodation.btnadd.Text = "Edit"
                Try
                    myForms.accomodation.txtdesc.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("description")
                Catch sd As Exception
                End Try
                Try
                    myForms.accomodation.txtaccomodationcost.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("costincurred")
                Catch sd As Exception
                End Try
                Try
                    myForms.accomodation.txtname.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("namme")
                Catch sd As Exception
                End Try
                Try
                    myForms.accomodation.cboentry.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("entry")
                Catch sd As Exception
                End Try
                Try
                    myForms.accomodation.autono = ds.Tables(0).Rows(htiaccomodation.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch saw As Exception
                Try
                    Dim gd As New frmaddaccomodation
                    myForms.accomodation = gd
                    myForms.accomodation.Text = "Edit accomodation"
                    myForms.accomodation.jobno = txtJobNo.Text
                    myForms.accomodation.btnadd.Text = "Edit"
                    Try
                        myForms.accomodation.txtdesc.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("description")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.accomodation.txtaccomodationcost.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("costincurred")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.accomodation.txtname.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("namme")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.accomodation.cboentry.Text = ds.Tables(0).Rows(htiaccomodation.Row).Item("entry")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.accomodation.autono = ds.Tables(0).Rows(htiaccomodation.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.accomodation.Show()
                Catch sf As Exception

                End Try
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtgcasuals_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgcasuals.DoubleClick
        Try
            If hticasuals.Type = Windows.Forms.DataGrid.HitTestType.Caption _
           Or hticasuals.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
           Or hticasuals.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
           Or hticasuals.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
           Or hticasuals.Type = Windows.Forms.DataGrid.HitTestType.None _
           Or hticasuals.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                 Then
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgcasuals.DataSource
            Try
                myForms.casuals.Text = "Edit casuals"
                myForms.casuals.jobno = Me.txtJobNo.Text
                myForms.casuals.btnadd.Text = "Edit"
                'Try
                '    myForms.casuals.txtdesc.Text = ds.Tables(0).Rows(hticasuals.Row).Item("description")
                'Catch sd As Exception
                'End Try
                Try
                    myForms.casuals.txtname.Text = ds.Tables(0).Rows(hticasuals.Row).Item("namme")
                Catch sd As Exception
                End Try
                Try
                    myForms.casuals.txttask.Text = ds.Tables(0).Rows(hticasuals.Row).Item("task")
                Catch sd As Exception
                End Try
                Try
                    myForms.casuals.dtpdatehired.Text = ds.Tables(0).Rows(hticasuals.Row).Item("datehired")
                Catch sd As Exception
                End Try
                Try
                    myForms.casuals.txtwage.Text = ds.Tables(0).Rows(hticasuals.Row).Item("wagespaid")
                Catch sd As Exception
                End Try

                Try
                    myForms.casuals.autono = ds.Tables(0).Rows(hticasuals.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch saw As Exception
                Try
                    Dim gd As New frmaddcasual
                    myForms.casuals = gd
                    myForms.casuals.Text = "Edit casuals"
                    myForms.casuals.jobno = Me.txtJobNo.Text
                    myForms.casuals.btnadd.Text = "Edit"

                    'Try
                    '    myForms.casuals.txtdesc.Text = ds.Tables(0).Rows(hticasuals.Row).Item("description")
                    'Catch sd As Exception
                    'End Try
                    Try
                        myForms.casuals.txtname.Text = ds.Tables(0).Rows(hticasuals.Row).Item("namme")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.casuals.txttask.Text = ds.Tables(0).Rows(hticasuals.Row).Item("task")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.casuals.dtpdatehired.Text = ds.Tables(0).Rows(hticasuals.Row).Item("datehired")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.casuals.txtwage.Text = ds.Tables(0).Rows(hticasuals.Row).Item("wagespaid")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.casuals.autono = ds.Tables(0).Rows(hticasuals.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.casuals.Show()
                Catch sf As Exception

                End Try
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtgtravel_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgtravel.DoubleClick
        Try
            If htitravel.Type = Windows.Forms.DataGrid.HitTestType.Caption _
           Or htitravel.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
           Or htitravel.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
           Or htitravel.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
           Or htitravel.Type = Windows.Forms.DataGrid.HitTestType.None _
           Or htitravel.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                 Then
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgtravel.DataSource
            Try
                myForms.travel.Text = "Edit travel"
                myForms.travel.jobno = txtJobNo.Text
                myForms.travel.btnadd.Text = "Edit"
                Try
                    myForms.travel.txtdesc.Text = ds.Tables(0).Rows(htitravel.Row).Item("description")
                Catch sd As Exception
                End Try
                Try
                    myForms.travel.txttravelcost.Text = ds.Tables(0).Rows(htitravel.Row).Item("costincurred")
                Catch sd As Exception
                End Try
                Try
                    myForms.travel.txtkilometers.Text = ds.Tables(0).Rows(htitravel.Row).Item("kilometers")
                Catch sd As Exception
                End Try
                Try
                    myForms.travel.txtother.Text = ds.Tables(0).Rows(htitravel.Row).Item("othermodes")
                Catch sd As Exception
                End Try
                Try
                    myForms.travel.autono = ds.Tables(0).Rows(htitravel.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch saw As Exception
                Try
                    Dim gd As New frmtravel
                    myForms.travel = gd
                    myForms.travel.Text = "Edit travel"
                    myForms.travel.jobno = txtJobNo.Text
                    myForms.travel.btnadd.Text = "Edit"
                    Try
                        myForms.travel.txtdesc.Text = ds.Tables(0).Rows(htitravel.Row).Item("description")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.travel.txttravelcost.Text = ds.Tables(0).Rows(htitravel.Row).Item("costincurred")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.travel.txtkilometers.Text = ds.Tables(0).Rows(htitravel.Row).Item("kilometers")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.travel.txtother.Text = ds.Tables(0).Rows(htitravel.Row).Item("othermodes")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.travel.autono = ds.Tables(0).Rows(htitravel.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.travel.Show()
                Catch sf As Exception

                End Try
            End Try
        Catch ex As Exception
        End Try
    End Sub
    'If htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
    'AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
    'AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
    'AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
    'AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.None Then
    '    Try

    '    Catch er456 As Exception

    '    End Try

    'End If
    Private Sub dtgcasuals_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgcasuals.MouseDown
        Try
            hticasuals = dtgcasuals.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
        End Try
    End Sub
    Private Sub dtgaccomodation_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgaccomodation.MouseDown
        Try
            htiaccomodation = dtgaccomodation.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
        End Try
    End Sub
    Private Sub dtgtravel_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgtravel.MouseDown
        Try
            htitravel = dtgtravel.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
        End Try
    End Sub
    Private Sub btnaccomodation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaccomodation.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim cv As New frmaddaccomodation
            cv.jobno = Me.txtJobNo.Text.Trim
            cv.ShowDialog()
        Catch xc As Exception

        End Try
    End Sub
    Private Sub btntravel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btntravel.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim cv As New frmtravel
            cv.jobno = Me.txtJobNo.Text.Trim
            cv.ShowDialog()
        Catch xc As Exception

        End Try
    End Sub
    Private Sub btndelcasuals_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelcasuals.Click
        Try
            dtgcasuals.Select(hticasuals.Row)
            If MessageBox.Show("Do you wish to delete the current row?", "Deleting", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Information) = DialogResult.No Then
                Exit Try
            End If
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgcasuals.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hticasuals.Row).Item("ano")
            str = "delete from casuals where ano='" & sid & "'"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                Me.txtlabour.TextBoxText = ""
                myrow = ds.Tables(0).Rows(hticasuals.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
        Try
            Dim Tasks As New taskclass
            Tasks.strjobno = Me.txtJobNo.Text
            Dim Threadf5 As New System.Threading.Thread( _
            AddressOf Tasks.casualsinvoke)
            Threadf5.IsBackground = True
            Threadf5.Start()

        Catch cv As Exception

        End Try
    End Sub
    Private Sub btndelaccom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelaccom.Click
        Try
            dtgaccomodation.Select(htiaccomodation.Row)
            If MessageBox.Show("Do you wish to delete the current row?", "Deleting", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Information) = DialogResult.No Then
                Exit Try
            End If
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgaccomodation.DataSource
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htiaccomodation.Row).Item("ano")
            str = "delete from  accomodation where ano='" & sid & "'"

            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                Me.txtaccomodation.TextBoxText = ""
                myrow = ds.Tables(0).Rows(htiaccomodation.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
        Try
            Dim Tasks As New taskclass
            Tasks.strjobno = Me.txtJobNo.Text
            Dim Threadfg As New System.Threading.Thread( _
           AddressOf Tasks.accomodationinvoke)
            Threadfg.IsBackground = True
            Threadfg.Start()
        Catch cv As Exception

        End Try
    End Sub
    Private Sub btndeltravel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeltravel.Click
        Try
            dtgtravel.Select(htitravel.Row)
            If MessageBox.Show("Do you wish to delete the current row?", "Deleting", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Information) = DialogResult.No Then
                Exit Try
            End If
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgtravel.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htitravel.Row).Item("ano")
            str = "delete from travel where ano='" & sid & "'"

            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                Me.txttravel.TextBoxText = ""
                myrow = ds.Tables(0).Rows(htitravel.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
        Try
            Dim Tasks As New taskclass
            Tasks.strjobno = Me.txtJobNo.Text
            Dim Threadfb As New System.Threading.Thread( _
          AddressOf Tasks.travelinvoke)
            Threadfb.IsBackground = True
            Threadfb.Start()
        Catch cv As Exception

        End Try
    End Sub
#End Region

#Region "equipment"
    Private Sub loadequip()
        Try
            Dim Tasks As New taskclass
            Tasks.erjobno = txtJobNo.Text.Trim
            Dim Threadl5 As New System.Threading.Thread( _
            AddressOf Tasks.ramaniequipinvoke)
            Threadl5.IsBackground = True
            Threadl5.Start()

            Tasks.hiredjobno = txtJobNo.Text.Trim
            Dim Threadfgx As New System.Threading.Thread( _
            AddressOf Tasks.hiredequipinvoke)
            Threadfgx.IsBackground = True
            Threadfgx.Start()


        Catch wq As Exception

        End Try
    End Sub
    Private Sub btnhired_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnhired.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim vxc As New frmhiredequip
            vxc.jobno = Me.txtJobNo.Text
            vxc.ShowDialog()
        Catch xc As Exception

        End Try
    End Sub
    Private Sub btnramaniequip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnramaniequip.Click
        Try
            Dim x As Boolean = canmanipulatejobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate jobs. contact administrator", "Jobs", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim task As taskclass
            If task.hastojobloaded = False Then
            End If
            Try
                myForms.tojobs.Close()
                myForms.tojobs = Nothing
            Catch zx As Exception
            End Try
            Dim form As New frmtojobs
            form.StartPosition = FormStartPosition.CenterParent
            myForms.tojobs = form
            task.iaminjobs = True
            task.ijobno = Me.txtJobNo.Text
            myForms.tojobs.Show()
            task.hastojobloaded = True
        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtghiredequip_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtghiredequip.MouseDown
        Try
            htihired = dtghiredequip.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
        End Try
    End Sub
    Private Sub dtghiredequip_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtghiredequip.DoubleClick
        Try
            If htihired.Type = Windows.Forms.DataGrid.HitTestType.Caption _
                       Or htihired.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
                       Or htihired.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
                       Or htihired.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
                       Or htihired.Type = Windows.Forms.DataGrid.HitTestType.None _
                       Or htihired.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                    Then
                Exit Sub
            End If
            '--------------------------
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtghiredequip.DataSource
            Try
                myForms.hired.Text = "Edit hired equipments"
                myForms.hired.jobno = txtJobNo.Text
                myForms.hired.btnadd.Text = "Edit"
                Try
                    myForms.hired.txtdesc.Text = ds.Tables(0).Rows(htihired.Row).Item("description")
                Catch sd As Exception
                End Try
                Try
                    myForms.hired.txthourlyrate.Text = ds.Tables(0).Rows(htihired.Row).Item("hourly_rate")
                Catch sd As Exception
                End Try
                Try
                    myForms.hired.txtname.Text = ds.Tables(0).Rows(htihired.Row).Item("equipname")
                Catch sd As Exception
                End Try
                Try
                    myForms.hired.dtpreleasedate.Value = CDate(ds.Tables(0).Rows(htihired.Row).Item("releasedate"))
                Catch sd As Exception
                End Try
                Try
                    myForms.hired.dtpssigndate.Value = CDate(ds.Tables(0).Rows(htihired.Row).Item("assigndate"))
                Catch sd As Exception
                End Try
                Try
                    Dim cv As String = ds.Tables(0).Rows(htihired.Row).Item("assigndate")
                    Dim a() As String = cv.Split(" ")
                    cv = "5/12/2006 " & a(1)
                    myForms.hired.dtpssigntime.Value = CDate(cv)
                Catch sd As Exception
                End Try
                Try
                    Dim cv As String = ds.Tables(0).Rows(htihired.Row).Item("releasedate")
                    Dim a() As String = cv.Split(" ")
                    cv = "5/12/2006 " & a(1)
                    myForms.hired.dtpreleasetime.Value = CDate(cv)
                Catch sd As Exception
                End Try
                Try
                    myForms.hired.autono = ds.Tables(0).Rows(htihired.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch zx As Exception
                Try
                    Dim gd As New frmhiredequip
                    myForms.hired = gd
                    myForms.hired.Text = "Edit hired equipments"
                    myForms.hired.jobno = txtJobNo.Text
                    myForms.hired.btnadd.Text = "Edit"
                    Try
                        myForms.hired.txtdesc.Text = ds.Tables(0).Rows(htihired.Row).Item("description")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.hired.txthourlyrate.Text = ds.Tables(0).Rows(htihired.Row).Item("hourly_rate")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.hired.txtname.Text = ds.Tables(0).Rows(htihired.Row).Item("equipname")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.hired.dtpreleasedate.Value = CDate(ds.Tables(0).Rows(htihired.Row).Item("releasedate"))
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.hired.dtpssigndate.Value = CDate(ds.Tables(0).Rows(htihired.Row).Item("assigndate"))
                    Catch sd As Exception
                    End Try
                    Try
                        Dim cv As String = ds.Tables(0).Rows(htihired.Row).Item("assigndate")
                        Dim a() As String = cv.Split(" ")
                        cv = "5/12/2006 " & a(1)
                        myForms.hired.dtpssigntime.Value = CDate(cv)
                    Catch sd As Exception
                    End Try
                    Try
                        Dim cv As String = ds.Tables(0).Rows(htihired.Row).Item("releasedate")
                        Dim a() As String = cv.Split(" ")
                        cv = "5/12/2006 " & a(1)
                        myForms.hired.dtpreleasetime.Value = CDate(cv)
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.hired.autono = ds.Tables(0).Rows(htihired.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.hired.Show()
                Catch sf As Exception
                End Try
            End Try
            '-----------------------------
        Catch zxc As Exception
        End Try

    End Sub
    Private Sub btndelramequip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelramequip.Click
        Try
            dtgramaniequip.Select(htiramequip.Row)
            If MessageBox.Show("Do you wish to delete the current row?", "Deleting", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Information) = DialogResult.No Then
                Exit Try
            End If
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgramaniequip.DataSource
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htiramequip.Row).Item("ano")
            str = "delete from history_equip where ano='" & sid & "'"

            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                Me.txttravel.TextBoxText = ""
                myrow = ds.Tables(0).Rows(htiramequip.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
        Try
            Dim Tasks As New taskclass
            Tasks.erjobno = txtJobNo.Text.Trim
            Dim Threadl5 As New System.Threading.Thread( _
            AddressOf Tasks.ramaniequipinvoke)
            Threadl5.IsBackground = True
            Threadl5.Start()
        Catch cv As Exception

        End Try
    End Sub
    Private Sub btndelhiredequip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelhiredequip.Click
        Try
            dtghiredequip.Select(htihired.Row)
            If MessageBox.Show("Do you wish to delete the current row?", "Deleting", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Information) = DialogResult.No Then
                Exit Try
            End If
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtghiredequip.DataSource
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htihired.Row).Item("ano")
            str = "delete from hiredequip where ano='" & sid & "'"

            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                Me.txttravel.TextBoxText = ""
                myrow = ds.Tables(0).Rows(htihired.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
        Try
            Dim Tasks As New taskclass
            Tasks.hiredjobno = txtJobNo.Text.Trim
            Dim Threadfgx As New System.Threading.Thread( _
            AddressOf Tasks.hiredequipinvoke)
            Threadfgx.IsBackground = True
            Threadfgx.Start()

        Catch cv As Exception

        End Try
    End Sub
    Private Sub dtgramaniequip_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgramaniequip.MouseDown
        Try
            htiramequip = dtgramaniequip.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
        End Try
    End Sub
#End Region

#Region "grossmargin"
    Public Sub mygross()
        Try
            Dim task As taskclass
            task.strjobno = Me.txtJobNo.Text
            task.erjobno = txtJobNo.Text.Trim
            task.hiredjobno = txtJobNo.Text.Trim
            '------------
            Dim Threadzzz As New System.Threading.Thread( _
               AddressOf task.gross)
            Threadzzz.IsBackground = True
            Threadzzz.Start()
            '--------
        Catch zx As Exception

        End Try

    End Sub
#End Region

#Region "validation"

    Private Sub txtJobTitle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtJobTitle.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtJobTitle, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtJobTitle, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtdesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdesc.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtdesc, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtdesc, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtamount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtamount.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtamount, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtamount, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtContactName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtContactName.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtContactName, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtContactName, "")
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
    Private Sub txtbudget_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtbudget.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtbudget, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtbudget, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region


    Private Sub lblgrossmargin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblgrossmargin.Click

    End Sub

    Private Sub lblgrossmargin_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblgrossmargin.TextChanged
        Try
            Dim dvm As New DataViewManager
            dvm = myForms.Main.dtgJobs.DataSource
            dvm.DataSet.Tables(0).Rows(myForms.Main._jobno).Item("grossmargin") = Me.lblgrossmargin.Text
        Catch ex As Exception

        End Try
    End Sub
End Class


