Imports ADODB
Imports System.Text.StringBuilder
Imports System.String
Imports System.Threading

Imports System.Configuration
Imports System.Collections.Specialized

Public Class frmAddJobs
    Inherits System.Windows.Forms.Form
    Public Delegate Sub mydelegate()
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
            addjobs = False
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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblOpenjournal As System.Windows.Forms.LinkLabel
    Friend WithEvents btndeletelead As System.Windows.Forms.Button
    Friend WithEvents lblCostAgreed As System.Windows.Forms.Label
    Friend WithEvents txtamount As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtContactName As System.Windows.Forms.TextBox
    Friend WithEvents lblContName As System.Windows.Forms.Label
    Friend WithEvents lblJobStatus As System.Windows.Forms.Label
    Friend WithEvents lblTechnicianResponsible As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lbClientName As System.Windows.Forms.Label
    Friend WithEvents lblClientNo As System.Windows.Forms.Label
    Friend WithEvents lblojobno As System.Windows.Forms.Label
    Friend WithEvents txtJobNo As System.Windows.Forms.TextBox
    Friend WithEvents lblJobNo As System.Windows.Forms.Label
    Friend WithEvents txtJobTitle As System.Windows.Forms.TextBox
    Friend WithEvents lblJobTitle As System.Windows.Forms.Label
    Friend WithEvents grpdesc As System.Windows.Forms.GroupBox
    Friend WithEvents txtdesc As System.Windows.Forms.RichTextBox
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents pnlhiredequip As System.Windows.Forms.Panel
    Friend WithEvents LabelTextBox6 As VSEssentials.LabelTextBox
    Friend WithEvents DataGrid2 As System.Windows.Forms.DataGrid
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlramaniequip As System.Windows.Forms.Panel
    Friend WithEvents LabelTextBox5 As VSEssentials.LabelTextBox
    Friend WithEvents DataGrid5 As System.Windows.Forms.DataGrid
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents pnltravel As System.Windows.Forms.Panel
    Friend WithEvents LabelTextBox3 As VSEssentials.LabelTextBox
    Friend WithEvents DataGrid4 As System.Windows.Forms.DataGrid
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents pnlaccomodation As System.Windows.Forms.Panel
    Friend WithEvents DataGrid3 As System.Windows.Forms.DataGrid
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pnlcasuals As System.Windows.Forms.Panel
    Friend WithEvents LabelTextBox1 As VSEssentials.LabelTextBox
    Friend WithEvents dgMember As System.Windows.Forms.DataGrid
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblCasual As System.Windows.Forms.Label
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents LabelTextBox4 As VSEssentials.LabelTextBox
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents pnlctrlpersonnel As System.Windows.Forms.Panel
    Friend WithEvents StiButton2 As System.Windows.Forms.Button
    Friend WithEvents StiButton1 As System.Windows.Forms.Button
    Friend WithEvents LabelTextBox2 As VSEssentials.LabelTextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnAddJobs As System.Windows.Forms.Button
    Friend WithEvents txtdepartment As System.Windows.Forms.ComboBox
    Friend WithEvents cboTechnicianresponsible As System.Windows.Forms.ComboBox
    Friend WithEvents cboJobstatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtbudget As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddJobs))
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtbudget = New System.Windows.Forms.TextBox
        Me.cboJobstatus = New System.Windows.Forms.ComboBox
        Me.cboTechnicianresponsible = New System.Windows.Forms.ComboBox
        Me.txtdepartment = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
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
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnAddJobs = New System.Windows.Forms.Button
        Me.lbClientName = New System.Windows.Forms.Label
        Me.lblClientNo = New System.Windows.Forms.Label
        Me.lblojobno = New System.Windows.Forms.Label
        Me.txtJobNo = New System.Windows.Forms.TextBox
        Me.lblJobNo = New System.Windows.Forms.Label
        Me.txtJobTitle = New System.Windows.Forms.TextBox
        Me.lblJobTitle = New System.Windows.Forms.Label
        Me.grpdesc = New System.Windows.Forms.GroupBox
        Me.txtdesc = New System.Windows.Forms.RichTextBox
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.pnlhiredequip = New System.Windows.Forms.Panel
        Me.LabelTextBox6 = New VSEssentials.LabelTextBox
        Me.DataGrid2 = New System.Windows.Forms.DataGrid
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlramaniequip = New System.Windows.Forms.Panel
        Me.LabelTextBox5 = New VSEssentials.LabelTextBox
        Me.DataGrid5 = New System.Windows.Forms.DataGrid
        Me.Button5 = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.pnltravel = New System.Windows.Forms.Panel
        Me.LabelTextBox3 = New VSEssentials.LabelTextBox
        Me.DataGrid4 = New System.Windows.Forms.DataGrid
        Me.Button3 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.pnlaccomodation = New System.Windows.Forms.Panel
        Me.LabelTextBox2 = New VSEssentials.LabelTextBox
        Me.DataGrid3 = New System.Windows.Forms.DataGrid
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnlcasuals = New System.Windows.Forms.Panel
        Me.LabelTextBox1 = New VSEssentials.LabelTextBox
        Me.dgMember = New System.Windows.Forms.DataGrid
        Me.Button1 = New System.Windows.Forms.Button
        Me.lblCasual = New System.Windows.Forms.Label
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.LabelTextBox4 = New VSEssentials.LabelTextBox
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.pnlctrlpersonnel = New System.Windows.Forms.Panel
        Me.StiButton2 = New System.Windows.Forms.Button
        Me.StiButton1 = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.grpdesc.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.pnlhiredequip.SuspendLayout()
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlramaniequip.SuspendLayout()
        CType(Me.DataGrid5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.pnltravel.SuspendLayout()
        CType(Me.DataGrid4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlaccomodation.SuspendLayout()
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlcasuals.SuspendLayout()
        CType(Me.dgMember, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlctrlpersonnel.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(586, 500)
        Me.TabControl1.TabIndex = 69
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.Panel1)
        Me.TabPage1.Controls.Add(Me.lbClientName)
        Me.TabPage1.Controls.Add(Me.lblClientNo)
        Me.TabPage1.Controls.Add(Me.lblojobno)
        Me.TabPage1.Controls.Add(Me.txtJobNo)
        Me.TabPage1.Controls.Add(Me.lblJobNo)
        Me.TabPage1.Controls.Add(Me.txtJobTitle)
        Me.TabPage1.Controls.Add(Me.lblJobTitle)
        Me.TabPage1.Controls.Add(Me.grpdesc)
        Me.TabPage1.Location = New System.Drawing.Point(4, 21)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(578, 475)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Job details"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
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
        Me.GroupBox1.Location = New System.Drawing.Point(8, 216)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(568, 218)
        Me.GroupBox1.TabIndex = 83
        Me.GroupBox1.TabStop = False
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(8, 187)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(152, 16)
        Me.Label9.TabIndex = 101
        Me.Label9.Text = "Budgetary cost"
        '
        'txtbudget
        '
        Me.txtbudget.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbudget.Location = New System.Drawing.Point(160, 187)
        Me.txtbudget.Name = "txtbudget"
        Me.txtbudget.Size = New System.Drawing.Size(232, 20)
        Me.txtbudget.TabIndex = 100
        Me.txtbudget.Text = ""
        '
        'cboJobstatus
        '
        Me.cboJobstatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboJobstatus.Location = New System.Drawing.Point(160, 136)
        Me.cboJobstatus.Name = "cboJobstatus"
        Me.cboJobstatus.Size = New System.Drawing.Size(232, 20)
        Me.cboJobstatus.TabIndex = 98
        '
        'cboTechnicianresponsible
        '
        Me.cboTechnicianresponsible.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTechnicianresponsible.Location = New System.Drawing.Point(160, 112)
        Me.cboTechnicianresponsible.Name = "cboTechnicianresponsible"
        Me.cboTechnicianresponsible.Size = New System.Drawing.Size(232, 20)
        Me.cboTechnicianresponsible.TabIndex = 97
        '
        'txtdepartment
        '
        Me.txtdepartment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.txtdepartment.Items.AddRange(New Object() {"", "Survey", "GI", "RS", "BD"})
        Me.txtdepartment.Location = New System.Drawing.Point(160, 88)
        Me.txtdepartment.Name = "txtdepartment"
        Me.txtdepartment.Size = New System.Drawing.Size(232, 20)
        Me.txtdepartment.TabIndex = 96
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(400, 64)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(160, 146)
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
        Me.Label7.Size = New System.Drawing.Size(136, 16)
        Me.Label7.TabIndex = 93
        Me.Label7.Text = "Gross margin"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(144, 106)
        Me.Label6.TabIndex = 84
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
        Me.GroupBox2.Size = New System.Drawing.Size(552, 48)
        Me.GroupBox2.TabIndex = 93
        Me.GroupBox2.TabStop = False
        '
        'lblOpenjournal
        '
        Me.lblOpenjournal.BackColor = System.Drawing.Color.Transparent
        Me.lblOpenjournal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpenjournal.Location = New System.Drawing.Point(8, 16)
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
        Me.btndeletelead.Location = New System.Drawing.Point(152, 16)
        Me.btndeletelead.Name = "btndeletelead"
        Me.btndeletelead.Size = New System.Drawing.Size(232, 20)
        Me.btndeletelead.TabIndex = 87
        Me.btndeletelead.Text = "Delete this Job"
        '
        'lblCostAgreed
        '
        Me.lblCostAgreed.BackColor = System.Drawing.Color.Transparent
        Me.lblCostAgreed.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCostAgreed.Location = New System.Drawing.Point(8, 165)
        Me.lblCostAgreed.Name = "lblCostAgreed"
        Me.lblCostAgreed.Size = New System.Drawing.Size(152, 16)
        Me.lblCostAgreed.TabIndex = 92
        Me.lblCostAgreed.Text = "Income"
        '
        'txtamount
        '
        Me.txtamount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtamount.Location = New System.Drawing.Point(160, 161)
        Me.txtamount.Name = "txtamount"
        Me.txtamount.Size = New System.Drawing.Size(232, 20)
        Me.txtamount.TabIndex = 91
        Me.txtamount.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(151, 16)
        Me.Label4.TabIndex = 90
        Me.Label4.Text = "Department"
        '
        'txtContactName
        '
        Me.txtContactName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtContactName.Location = New System.Drawing.Point(160, 64)
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.Size = New System.Drawing.Size(232, 20)
        Me.txtContactName.TabIndex = 81
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
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Controls.Add(Me.btnAddJobs)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 443)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(578, 32)
        Me.Panel1.TabIndex = 76
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Location = New System.Drawing.Point(496, 5)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        '
        'btnAddJobs
        '
        Me.btnAddJobs.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddJobs.Location = New System.Drawing.Point(3, 3)
        Me.btnAddJobs.Name = "btnAddJobs"
        Me.btnAddJobs.TabIndex = 0
        Me.btnAddJobs.Text = "Add jobs"
        '
        'lbClientName
        '
        Me.lbClientName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbClientName.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lbClientName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbClientName.Location = New System.Drawing.Point(120, 8)
        Me.lbClientName.Name = "lbClientName"
        Me.lbClientName.Size = New System.Drawing.Size(448, 16)
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
        Me.lblojobno.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblojobno.BackColor = System.Drawing.SystemColors.Window
        Me.lblojobno.Location = New System.Drawing.Point(187, 28)
        Me.lblojobno.Name = "lblojobno"
        Me.lblojobno.Size = New System.Drawing.Size(140, 20)
        Me.lblojobno.TabIndex = 70
        '
        'txtJobNo
        '
        Me.txtJobNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
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
        Me.lblJobNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblJobNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobNo.Location = New System.Drawing.Point(16, 28)
        Me.lblJobNo.Name = "lblJobNo"
        Me.lblJobNo.Size = New System.Drawing.Size(80, 16)
        Me.lblJobNo.TabIndex = 69
        Me.lblJobNo.Text = "Job No"
        '
        'txtJobTitle
        '
        Me.txtJobTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtJobTitle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtJobTitle.Location = New System.Drawing.Point(104, 52)
        Me.txtJobTitle.Name = "txtJobTitle"
        Me.txtJobTitle.Size = New System.Drawing.Size(224, 20)
        Me.txtJobTitle.TabIndex = 67
        Me.txtJobTitle.Text = ""
        '
        'lblJobTitle
        '
        Me.lblJobTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
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
        Me.grpdesc.BackColor = System.Drawing.Color.Transparent
        Me.grpdesc.Controls.Add(Me.txtdesc)
        Me.grpdesc.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpdesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpdesc.Location = New System.Drawing.Point(8, 77)
        Me.grpdesc.Name = "grpdesc"
        Me.grpdesc.Size = New System.Drawing.Size(568, 133)
        Me.grpdesc.TabIndex = 39
        Me.grpdesc.TabStop = False
        Me.grpdesc.Text = "Description"
        '
        'txtdesc
        '
        Me.txtdesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtdesc.Location = New System.Drawing.Point(8, 16)
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(552, 106)
        Me.txtdesc.TabIndex = 3
        Me.txtdesc.Text = ""
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.pnlhiredequip)
        Me.TabPage3.Controls.Add(Me.pnlramaniequip)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(578, 474)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Equipment"
        Me.TabPage3.Visible = False
        '
        'pnlhiredequip
        '
        Me.pnlhiredequip.Controls.Add(Me.LabelTextBox6)
        Me.pnlhiredequip.Controls.Add(Me.DataGrid2)
        Me.pnlhiredequip.Controls.Add(Me.Button4)
        Me.pnlhiredequip.Controls.Add(Me.Label1)
        Me.pnlhiredequip.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlhiredequip.Location = New System.Drawing.Point(0, 312)
        Me.pnlhiredequip.Name = "pnlhiredequip"
        Me.pnlhiredequip.Size = New System.Drawing.Size(578, 162)
        Me.pnlhiredequip.TabIndex = 3
        '
        'LabelTextBox6
        '
        Me.LabelTextBox6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelTextBox6.LabelText = "Total:"
        Me.LabelTextBox6.Location = New System.Drawing.Point(437, 129)
        Me.LabelTextBox6.Name = "LabelTextBox6"
        Me.LabelTextBox6.Size = New System.Drawing.Size(144, 24)
        Me.LabelTextBox6.TabIndex = 14
        Me.LabelTextBox6.TextBoxText = ""
        '
        'DataGrid2
        '
        Me.DataGrid2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid2.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGrid2.CaptionVisible = False
        Me.DataGrid2.DataMember = ""
        Me.DataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid2.Location = New System.Drawing.Point(4, 24)
        Me.DataGrid2.Name = "DataGrid2"
        Me.DataGrid2.PreferredRowHeight = 20
        Me.DataGrid2.ReadOnly = True
        Me.DataGrid2.Size = New System.Drawing.Size(568, 97)
        Me.DataGrid2.TabIndex = 13
        '
        'Button4
        '
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Location = New System.Drawing.Point(128, 3)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(24, 16)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "+"
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
        Me.pnlramaniequip.Controls.Add(Me.LabelTextBox5)
        Me.pnlramaniequip.Controls.Add(Me.DataGrid5)
        Me.pnlramaniequip.Controls.Add(Me.Button5)
        Me.pnlramaniequip.Controls.Add(Me.Label5)
        Me.pnlramaniequip.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlramaniequip.Location = New System.Drawing.Point(0, 0)
        Me.pnlramaniequip.Name = "pnlramaniequip"
        Me.pnlramaniequip.Size = New System.Drawing.Size(578, 312)
        Me.pnlramaniequip.TabIndex = 2
        '
        'LabelTextBox5
        '
        Me.LabelTextBox5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelTextBox5.LabelText = "Total:"
        Me.LabelTextBox5.Location = New System.Drawing.Point(437, 281)
        Me.LabelTextBox5.Name = "LabelTextBox5"
        Me.LabelTextBox5.Size = New System.Drawing.Size(144, 24)
        Me.LabelTextBox5.TabIndex = 14
        Me.LabelTextBox5.TextBoxText = ""
        '
        'DataGrid5
        '
        Me.DataGrid5.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid5.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGrid5.CaptionVisible = False
        Me.DataGrid5.DataMember = ""
        Me.DataGrid5.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid5.Location = New System.Drawing.Point(5, 24)
        Me.DataGrid5.Name = "DataGrid5"
        Me.DataGrid5.PreferredRowHeight = 20
        Me.DataGrid5.ReadOnly = True
        Me.DataGrid5.Size = New System.Drawing.Size(568, 256)
        Me.DataGrid5.TabIndex = 13
        '
        'Button5
        '
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Location = New System.Drawing.Point(128, 5)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(24, 16)
        Me.Button5.TabIndex = 2
        Me.Button5.Text = "+"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(7, 5)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(113, 11)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Ramani equipment"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.pnltravel)
        Me.TabPage2.Controls.Add(Me.pnlaccomodation)
        Me.TabPage2.Controls.Add(Me.pnlcasuals)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(578, 474)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Financial details"
        Me.TabPage2.Visible = False
        '
        'pnltravel
        '
        Me.pnltravel.Controls.Add(Me.LabelTextBox3)
        Me.pnltravel.Controls.Add(Me.DataGrid4)
        Me.pnltravel.Controls.Add(Me.Button3)
        Me.pnltravel.Controls.Add(Me.Label3)
        Me.pnltravel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnltravel.Location = New System.Drawing.Point(0, 362)
        Me.pnltravel.Name = "pnltravel"
        Me.pnltravel.Size = New System.Drawing.Size(578, 112)
        Me.pnltravel.TabIndex = 2
        '
        'LabelTextBox3
        '
        Me.LabelTextBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelTextBox3.LabelText = "Total:"
        Me.LabelTextBox3.Location = New System.Drawing.Point(432, 84)
        Me.LabelTextBox3.Name = "LabelTextBox3"
        Me.LabelTextBox3.Size = New System.Drawing.Size(144, 24)
        Me.LabelTextBox3.TabIndex = 14
        Me.LabelTextBox3.TextBoxText = ""
        '
        'DataGrid4
        '
        Me.DataGrid4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid4.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGrid4.CaptionVisible = False
        Me.DataGrid4.DataMember = ""
        Me.DataGrid4.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid4.Location = New System.Drawing.Point(8, 24)
        Me.DataGrid4.Name = "DataGrid4"
        Me.DataGrid4.PreferredRowHeight = 20
        Me.DataGrid4.ReadOnly = True
        Me.DataGrid4.Size = New System.Drawing.Size(568, 60)
        Me.DataGrid4.TabIndex = 12
        '
        'Button3
        '
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Location = New System.Drawing.Point(94, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(24, 16)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "+"
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
        Me.pnlaccomodation.Controls.Add(Me.LabelTextBox2)
        Me.pnlaccomodation.Controls.Add(Me.DataGrid3)
        Me.pnlaccomodation.Controls.Add(Me.Button2)
        Me.pnlaccomodation.Controls.Add(Me.Label2)
        Me.pnlaccomodation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlaccomodation.Location = New System.Drawing.Point(0, 128)
        Me.pnlaccomodation.Name = "pnlaccomodation"
        Me.pnlaccomodation.Size = New System.Drawing.Size(578, 346)
        Me.pnlaccomodation.TabIndex = 1
        '
        'LabelTextBox2
        '
        Me.LabelTextBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelTextBox2.LabelText = "Total:"
        Me.LabelTextBox2.Location = New System.Drawing.Point(424, 209)
        Me.LabelTextBox2.Name = "LabelTextBox2"
        Me.LabelTextBox2.Size = New System.Drawing.Size(144, 24)
        Me.LabelTextBox2.TabIndex = 14
        Me.LabelTextBox2.TextBoxText = ""
        '
        'DataGrid3
        '
        Me.DataGrid3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid3.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGrid3.CaptionVisible = False
        Me.DataGrid3.DataMember = ""
        Me.DataGrid3.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid3.Location = New System.Drawing.Point(8, 24)
        Me.DataGrid3.Name = "DataGrid3"
        Me.DataGrid3.PreferredRowHeight = 20
        Me.DataGrid3.ReadOnly = True
        Me.DataGrid3.Size = New System.Drawing.Size(584, 177)
        Me.DataGrid3.TabIndex = 12
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(94, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(24, 16)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "+"
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
        Me.pnlcasuals.Controls.Add(Me.LabelTextBox1)
        Me.pnlcasuals.Controls.Add(Me.dgMember)
        Me.pnlcasuals.Controls.Add(Me.Button1)
        Me.pnlcasuals.Controls.Add(Me.lblCasual)
        Me.pnlcasuals.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlcasuals.Location = New System.Drawing.Point(0, 0)
        Me.pnlcasuals.Name = "pnlcasuals"
        Me.pnlcasuals.Size = New System.Drawing.Size(578, 128)
        Me.pnlcasuals.TabIndex = 0
        '
        'LabelTextBox1
        '
        Me.LabelTextBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelTextBox1.LabelText = "Total:"
        Me.LabelTextBox1.Location = New System.Drawing.Point(432, 101)
        Me.LabelTextBox1.Name = "LabelTextBox1"
        Me.LabelTextBox1.Size = New System.Drawing.Size(144, 24)
        Me.LabelTextBox1.TabIndex = 12
        Me.LabelTextBox1.TextBoxText = ""
        '
        'dgMember
        '
        Me.dgMember.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgMember.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dgMember.CaptionVisible = False
        Me.dgMember.DataMember = ""
        Me.dgMember.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgMember.Location = New System.Drawing.Point(8, 24)
        Me.dgMember.Name = "dgMember"
        Me.dgMember.PreferredRowHeight = 20
        Me.dgMember.ReadOnly = True
        Me.dgMember.Size = New System.Drawing.Size(568, 72)
        Me.dgMember.TabIndex = 11
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(94, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(24, 16)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "+"
        '
        'lblCasual
        '
        Me.lblCasual.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblCasual.Location = New System.Drawing.Point(7, 5)
        Me.lblCasual.Name = "lblCasual"
        Me.lblCasual.Size = New System.Drawing.Size(65, 0)
        Me.lblCasual.TabIndex = 0
        Me.lblCasual.Text = "Casual labour"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.Panel2)
        Me.TabPage4.Controls.Add(Me.pnlctrlpersonnel)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(578, 474)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Personnel"
        Me.TabPage4.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.LabelTextBox4)
        Me.Panel2.Controls.Add(Me.DataGrid1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 40)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(578, 434)
        Me.Panel2.TabIndex = 4
        '
        'LabelTextBox4
        '
        Me.LabelTextBox4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LabelTextBox4.LabelText = "Total:"
        Me.LabelTextBox4.Location = New System.Drawing.Point(432, 409)
        Me.LabelTextBox4.Name = "LabelTextBox4"
        Me.LabelTextBox4.Size = New System.Drawing.Size(144, 24)
        Me.LabelTextBox4.TabIndex = 14
        Me.LabelTextBox4.TextBoxText = ""
        '
        'DataGrid1
        '
        Me.DataGrid1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid1.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(5, 16)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.PreferredRowHeight = 20
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.Size = New System.Drawing.Size(568, 385)
        Me.DataGrid1.TabIndex = 13
        '
        'pnlctrlpersonnel
        '
        Me.pnlctrlpersonnel.Controls.Add(Me.StiButton2)
        Me.pnlctrlpersonnel.Controls.Add(Me.StiButton1)
        Me.pnlctrlpersonnel.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlctrlpersonnel.Location = New System.Drawing.Point(0, 0)
        Me.pnlctrlpersonnel.Name = "pnlctrlpersonnel"
        Me.pnlctrlpersonnel.Size = New System.Drawing.Size(578, 40)
        Me.pnlctrlpersonnel.TabIndex = 3
        '
        'StiButton2
        '
        Me.StiButton2.Location = New System.Drawing.Point(106, 8)
        Me.StiButton2.Name = "StiButton2"
        Me.StiButton2.Size = New System.Drawing.Size(96, 23)
        Me.StiButton2.TabIndex = 4
        Me.StiButton2.Text = "Export to excel"
        '
        'StiButton1
        '
        Me.StiButton1.Location = New System.Drawing.Point(8, 8)
        Me.StiButton1.Name = "StiButton1"
        Me.StiButton1.Size = New System.Drawing.Size(96, 23)
        Me.StiButton1.TabIndex = 3
        Me.StiButton1.Text = "Print"
        '
        'frmAddJobs
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(586, 500)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Garamond", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmAddJobs"
        Me.Text = " Add  Jobs"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.grpdesc.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.pnlhiredequip.ResumeLayout(False)
        CType(Me.DataGrid2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlramaniequip.ResumeLayout(False)
        CType(Me.DataGrid5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.pnltravel.ResumeLayout(False)
        CType(Me.DataGrid4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlaccomodation.ResumeLayout(False)
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlcasuals.ResumeLayout(False)
        CType(Me.dgMember, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
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
#Region "jobs"
    Private Sub frmAddJobs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblClientNo.Text = myclientno
        Me.lbClientName.Text = myclientname

        'Me.cboJobstatus.Items.Add("Pending")
        Me.cboJobstatus.Items.Add("Current")
        Me.cboJobstatus.Items.Add("Complete")
        Me.cboJobstatus.Items.Add("Delivered")
        'Me.cboJobstatus.Items.Add("Current")
        'Me.cboJobstatus.Items.Add("Invoiced")
        'Me.cboJobstatus.Items.Add("Completed")
        'Me.txtJobNo.Text = jobno()
        Call loadtechnicians()
        Me.txtJobNo.Text = newlno(myclientno)
    End Sub
    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub loadtechnicians()
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

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
    End Sub
    Private Function jobno1() As String
        Try
            'Dim cmd As New ADODB.Command()
            'Dim rs As New ADODB.Recordset()
            'Dim str As String
            'With cmd
            '    .ActiveConnection = connect
            '    .CommandType = ADODB.CommandTypeEnum.adCmdText

            '    str = " select max(job_no) from rcljobs"
            '    .CommandText = str
            '    rs = .Execute
            'End With
            'str = rs.Fields("max").Value
            'Dim strno, str1 As String
            'Dim i
            'For i = 0 To str.Length - 1
            '    If IsNumeric(str.Substring(i, 1)) = True Then
            '        strno = strno & str.Substring(i, 1)
            '    Else
            '        str1 = str1 & str.Substring(i, 1)
            '    End If

            'Next

            'str = (CSng(strno) + 1).ToString
            'str1 = str1.Insert(1, str)
            'Return str1
            'rs.Close()
            'rs = Nothing
        Catch ex As Exception

        End Try
    End Function
    Private Function jobno() As String
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
            Dim jno As String
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenKeyset
                Dim cno = myclientno & "1"
                Dim str = "select job_no from rcljobs where " _
                & "job_no like '" & cno & "%'" _
                & "order by job_no desc"
                .Open(str, connect)
                If .EOF = False And .EOF = False Then
                    jno = .Fields("job_no").Value
                    Dim l = jno.Trim.Length
                    Dim str33 As String = jno.Substring(5, l - 5)
                    Dim str34 As String = jno.Substring(0, 5)
                    str33 = (CLng(str33) + 1).ToString()
                    Select Case str33.Trim.Length
                        Case 1
                            str33 = "000" & str33
                        Case 2
                            str33 = "00" & str33
                        Case 3
                            str33 = "0" & str33
                        Case Else
                            str33 = str33
                    End Select
                    jno = str34 + str33

                Else
                    jno = Me.lblClientNo.Text & "1" & "0000"
                End If
            End With
            jobno = jno
        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Function validation() As Boolean
        Try
            Dim sdate As Date
            Dim edate As Date
            Dim deadlinedate As Date
            'sdate = Me.dtpStartDate.Value
            'edate = Me.dtpEndDate.Value
            'deadlinedate = Me.dtpDeadline.Value
            Dim today, today1, today2 As Long
            today = DateDiff(interval:=DateInterval.Day, date1:=Now, date2:=sdate)
            today1 = DateDiff(interval:=DateInterval.Day, date1:=Now, date2:=edate)
            today2 = DateDiff(interval:=DateInterval.Day, date1:=Now, date2:=deadlinedate)
            If sdate > edate Or sdate > deadlinedate Or deadlinedate > edate _
                Or today < 0 Or today1 < 0 _
                 Or today2 < 0 _
                Then
                validation = True
            Else
                validation = False
            End If
        Catch ex As Exception
            validation = False
        Finally

        End Try

    End Function
    Private Function returndates()
        Try
            Dim sdate As String
            Dim edate As String
            Dim deadlinedate As String
            'sdate = Me.dtpStartDate.Value
            'edate = Me.dtpEndDate.Value
            'deadlinedate = Me.dtpDeadline.Value

            Dim str1 As String


            'using the split function
            Dim a() As String
            'start date
            str1 = CStr(sdate)
            'str1 = str1.Substring(0, 10)
            'm = str1.Substring(0, 2)
            'd = str1.Substring(3, 2)
            'y = str1.Substring(6, 4)

            a = str1.Split(" ")
            str1 = a(0)
            a = str1.Split("/")
            returndates = returndates() & a(2) & "-" _
            & a(0) & "-" & a(1) & "|"

            'end date date
            str1 = CStr(edate)
            'str1 = str1.Substring(0, 10)
            'm = str1.Substring(0, 2)
            'd = str1.Substring(3, 2)
            'y = str1.Substring(6, 4)

            a = str1.Split(" ")
            str1 = a(0)
            a = str1.Split("/")
            returndates = returndates() & a(2) & "-" _
            & a(0) & "-" & a(1) & "|"

            'deadline date
            str1 = CStr(deadlinedate)
            'str1 = str1.Substring(0, 10)
            'm = str1.Substring(0, 2)
            'd = str1.Substring(3, 2)
            'y = str1.Substring(6, 4)

            a = str1.Split(" ")
            str1 = a(0)
            a = str1.Split("/")
            returndates = returndates() & a(2) & "-" _
            & a(0) & "-" & a(1) & "|"

        Catch ex As Exception

        End Try
    End Function
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Dispose(True)
    End Sub
    Private Sub addjob()
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()
        Try
            connect.Close()
        Catch ex As Exception
        End Try
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
            'Dim k() As String
            'k = Me.txtTechnicianResponsible.Lines
            'Dim mystr5 As String
            'mystr5 = rml(k)
            '// deals with dates
            'Dim a() As String
            Dim strsql, str As String
            'str = returndates()
            'a = str.Split("|")
            txtJobNo.Text = newlno(myclientno)
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
            cboJobstatus.Text = ""
            cboTechnicianresponsible.Text = ""
            txtdesc.Clear()
            txtJobNo.Focus()

            txtJobNo.Text = newlno(myclientno)
        End If


        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            addjobs = False
            Me.Dispose(True)
        Catch sd As Exception
        End Try
    End Sub
    Private Sub btnAddJobs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddJobs.Click
        Try
            Me.Invoke(New mydelegate(AddressOf addjob))
        Catch ex As Exception

        End Try
    End Sub
    Protected Overrides Sub Finalize()

        MyBase.Finalize()
    End Sub
    'Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
    '    Try
    '        ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
    '        ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
    '        If keyData = System.Windows.Forms.Keys.Return Then
    '            'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)
    '            Dim E As System.EventArgs

    '            'Call btnAddJobs_Click_1(Me, E)
    '            If Me.btnAddJobs.TabStop = True Then
    '                'Me.Invoke(New mydelegate(AddressOf addjob))
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
    Private Sub cboJobstatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
            MsgBox(ex.Message.ToString())
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
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
#End Region





End Class

