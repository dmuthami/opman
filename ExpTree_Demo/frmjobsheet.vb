
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports ADODB

Imports System.ArgumentNullException
Imports System.NullReferenceException
Imports System.ArgumentOutOfRangeException
Imports exporttoexcel
Imports System.Threading
Public Class frmjobsheet
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
            hasloadedjobsheet = False
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
    Friend WithEvents pnlgrid As System.Windows.Forms.Panel
    Friend WithEvents dtgtimesheet As System.Windows.Forms.DataGrid
    Friend WithEvents pnljobsheetcontrols As System.Windows.Forms.Panel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pnlemployeecontrols As System.Windows.Forms.Panel
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtmobileno As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtphoneno As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btnsave As System.Windows.Forms.Button
    Friend WithEvents txtidno As System.Windows.Forms.TextBox
    Friend WithEvents txthourlyrate As System.Windows.Forms.TextBox
    Friend WithEvents txtpostaladdress As System.Windows.Forms.TextBox
    Friend WithEvents cbogender As System.Windows.Forms.ComboBox
    Friend WithEvents txtemail As System.Windows.Forms.TextBox
    Friend WithEvents txtpin As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents GroupBox11 As System.Windows.Forms.GroupBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnpassport As System.Windows.Forms.Button
    Friend WithEvents txtbirthday As System.Windows.Forms.TextBox
    Friend WithEvents txtnextofkin As System.Windows.Forms.TextBox
    Friend WithEvents txtnssfno As System.Windows.Forms.TextBox
    Friend WithEvents txtcontractend As System.Windows.Forms.TextBox
    Friend WithEvents txtmedicalcover As System.Windows.Forms.TextBox
    Friend WithEvents txtnhifno As System.Windows.Forms.TextBox
    Friend WithEvents dtpdoe As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpdot As System.Windows.Forms.DateTimePicker
    Friend WithEvents cbotimeoff As System.Windows.Forms.ComboBox
    Friend WithEvents cbosickoff As System.Windows.Forms.ComboBox
    Friend WithEvents cbodayoff As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboleaves As System.Windows.Forms.ComboBox
    Friend WithEvents dtgleaves As System.Windows.Forms.DataGrid
    Friend WithEvents btnaddleave As System.Windows.Forms.Button
    Friend WithEvents dtgtimeoff As System.Windows.Forms.DataGrid
    Friend WithEvents dtgsickoff As System.Windows.Forms.DataGrid
    Friend WithEvents dtgdayoff As System.Windows.Forms.DataGrid
    Friend WithEvents btntimeoff As System.Windows.Forms.Button
    Friend WithEvents btnsickoff As System.Windows.Forms.Button
    Friend WithEvents btndayoff As System.Windows.Forms.Button
    Friend WithEvents pbimage As System.Windows.Forms.PictureBox
    Friend WithEvents tbcjobsheet As System.Windows.Forms.TabControl
    Friend WithEvents tpgemployees As System.Windows.Forms.TabPage
    Friend WithEvents tpgtimesheet As System.Windows.Forms.TabPage
    Friend WithEvents tpgmiscellanous As System.Windows.Forms.TabPage
    Friend WithEvents txtname As System.Windows.Forms.TextBox
    Friend WithEvents txtcomments As System.Windows.Forms.TextBox
    Friend WithEvents btnview As System.Windows.Forms.Button
    Friend WithEvents lblsdate As System.Windows.Forms.Label
    Friend WithEvents lblenddate As System.Windows.Forms.Label
    Friend WithEvents dtpedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpsdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents pnltop As System.Windows.Forms.Panel
    Friend WithEvents GroupBox12 As System.Windows.Forms.GroupBox
    Friend WithEvents btnsavetimesheet As System.Windows.Forms.Button
    Friend WithEvents btnaddentry As System.Windows.Forms.Button
    Friend WithEvents btndelete As System.Windows.Forms.Button
    Friend WithEvents GroupBox13 As System.Windows.Forms.GroupBox
    Friend WithEvents btnprint As System.Windows.Forms.Button
    Friend WithEvents btnexcel As System.Windows.Forms.Button
    Friend WithEvents btnleaves As System.Windows.Forms.Button
    Friend WithEvents btndeltimeoff As System.Windows.Forms.Button
    Friend WithEvents btndelsickoff As System.Windows.Forms.Button
    Friend WithEvents btndeldayoff As System.Windows.Forms.Button
    Friend WithEvents btndepartments As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmjobsheet))
        Me.tbcjobsheet = New System.Windows.Forms.TabControl
        Me.tpgemployees = New System.Windows.Forms.TabPage
        Me.pnlemployeecontrols = New System.Windows.Forms.Panel
        Me.btnpassport = New System.Windows.Forms.Button
        Me.btnsave = New System.Windows.Forms.Button
        Me.pnljobsheetcontrols = New System.Windows.Forms.Panel
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.btndepartments = New System.Windows.Forms.Button
        Me.txtname = New System.Windows.Forms.TextBox
        Me.txtmobileno = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtphoneno = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtbirthday = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.cbogender = New System.Windows.Forms.ComboBox
        Me.txtidno = New System.Windows.Forms.TextBox
        Me.txtemail = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.txtpin = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.txtpostaladdress = New System.Windows.Forms.TextBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.txtcontractend = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtnhifno = New System.Windows.Forms.TextBox
        Me.txtnssfno = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.txthourlyrate = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtnextofkin = New System.Windows.Forms.TextBox
        Me.dtpdot = New System.Windows.Forms.DateTimePicker
        Me.dtpdoe = New System.Windows.Forms.DateTimePicker
        Me.txtmedicalcover = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.pbimage = New System.Windows.Forms.PictureBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtcomments = New System.Windows.Forms.TextBox
        Me.tpgtimesheet = New System.Windows.Forms.TabPage
        Me.pnlgrid = New System.Windows.Forms.Panel
        Me.GroupBox13 = New System.Windows.Forms.GroupBox
        Me.btnexcel = New System.Windows.Forms.Button
        Me.btnprint = New System.Windows.Forms.Button
        Me.GroupBox12 = New System.Windows.Forms.GroupBox
        Me.btnsavetimesheet = New System.Windows.Forms.Button
        Me.btnaddentry = New System.Windows.Forms.Button
        Me.btndelete = New System.Windows.Forms.Button
        Me.dtgtimesheet = New System.Windows.Forms.DataGrid
        Me.pnltop = New System.Windows.Forms.Panel
        Me.btnview = New System.Windows.Forms.Button
        Me.lblsdate = New System.Windows.Forms.Label
        Me.lblenddate = New System.Windows.Forms.Label
        Me.dtpedate = New System.Windows.Forms.DateTimePicker
        Me.dtpsdate = New System.Windows.Forms.DateTimePicker
        Me.tpgmiscellanous = New System.Windows.Forms.TabPage
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboleaves = New System.Windows.Forms.ComboBox
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.btnleaves = New System.Windows.Forms.Button
        Me.dtgleaves = New System.Windows.Forms.DataGrid
        Me.btnaddleave = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbosickoff = New System.Windows.Forms.ComboBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.btndelsickoff = New System.Windows.Forms.Button
        Me.dtgsickoff = New System.Windows.Forms.DataGrid
        Me.btnsickoff = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cbotimeoff = New System.Windows.Forms.ComboBox
        Me.GroupBox9 = New System.Windows.Forms.GroupBox
        Me.btndeltimeoff = New System.Windows.Forms.Button
        Me.dtgtimeoff = New System.Windows.Forms.DataGrid
        Me.btntimeoff = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.GroupBox10 = New System.Windows.Forms.GroupBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cbodayoff = New System.Windows.Forms.ComboBox
        Me.GroupBox11 = New System.Windows.Forms.GroupBox
        Me.btndeldayoff = New System.Windows.Forms.Button
        Me.dtgdayoff = New System.Windows.Forms.DataGrid
        Me.btndayoff = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.tbcjobsheet.SuspendLayout()
        Me.tpgemployees.SuspendLayout()
        Me.pnlemployeecontrols.SuspendLayout()
        Me.pnljobsheetcontrols.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.tpgtimesheet.SuspendLayout()
        Me.pnlgrid.SuspendLayout()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        CType(Me.dtgtimesheet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnltop.SuspendLayout()
        Me.tpgmiscellanous.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        CType(Me.dtgleaves, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dtgsickoff, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        CType(Me.dtgtimeoff, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        Me.GroupBox11.SuspendLayout()
        CType(Me.dtgdayoff, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbcjobsheet
        '
        Me.tbcjobsheet.Controls.Add(Me.tpgemployees)
        Me.tbcjobsheet.Controls.Add(Me.tpgtimesheet)
        Me.tbcjobsheet.Controls.Add(Me.tpgmiscellanous)
        Me.tbcjobsheet.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbcjobsheet.Location = New System.Drawing.Point(0, 0)
        Me.tbcjobsheet.Name = "tbcjobsheet"
        Me.tbcjobsheet.SelectedIndex = 0
        Me.tbcjobsheet.Size = New System.Drawing.Size(802, 556)
        Me.tbcjobsheet.TabIndex = 0
        '
        'tpgemployees
        '
        Me.tpgemployees.Controls.Add(Me.pnlemployeecontrols)
        Me.tpgemployees.Controls.Add(Me.pnljobsheetcontrols)
        Me.tpgemployees.Location = New System.Drawing.Point(4, 23)
        Me.tpgemployees.Name = "tpgemployees"
        Me.tpgemployees.Size = New System.Drawing.Size(794, 529)
        Me.tpgemployees.TabIndex = 0
        Me.tpgemployees.Text = "Employee details"
        '
        'pnlemployeecontrols
        '
        Me.pnlemployeecontrols.Controls.Add(Me.btnpassport)
        Me.pnlemployeecontrols.Controls.Add(Me.btnsave)
        Me.pnlemployeecontrols.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlemployeecontrols.Location = New System.Drawing.Point(0, 0)
        Me.pnlemployeecontrols.Name = "pnlemployeecontrols"
        Me.pnlemployeecontrols.Size = New System.Drawing.Size(794, 32)
        Me.pnlemployeecontrols.TabIndex = 0
        '
        'btnpassport
        '
        Me.btnpassport.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnpassport.Location = New System.Drawing.Point(5, 3)
        Me.btnpassport.Name = "btnpassport"
        Me.btnpassport.Size = New System.Drawing.Size(123, 20)
        Me.btnpassport.TabIndex = 1
        Me.btnpassport.Text = "Browse for passport"
        '
        'btnsave
        '
        Me.btnsave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnsave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsave.Location = New System.Drawing.Point(688, 3)
        Me.btnsave.Name = "btnsave"
        Me.btnsave.Size = New System.Drawing.Size(104, 20)
        Me.btnsave.TabIndex = 2
        Me.btnsave.Text = "Save changes"
        '
        'pnljobsheetcontrols
        '
        Me.pnljobsheetcontrols.AutoScroll = True
        Me.pnljobsheetcontrols.Controls.Add(Me.GroupBox5)
        Me.pnljobsheetcontrols.Controls.Add(Me.GroupBox1)
        Me.pnljobsheetcontrols.Controls.Add(Me.GroupBox3)
        Me.pnljobsheetcontrols.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnljobsheetcontrols.Location = New System.Drawing.Point(0, 0)
        Me.pnljobsheetcontrols.Name = "pnljobsheetcontrols"
        Me.pnljobsheetcontrols.Size = New System.Drawing.Size(794, 529)
        Me.pnljobsheetcontrols.TabIndex = 5
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.btndepartments)
        Me.GroupBox5.Controls.Add(Me.txtname)
        Me.GroupBox5.Controls.Add(Me.txtmobileno)
        Me.GroupBox5.Controls.Add(Me.Label12)
        Me.GroupBox5.Controls.Add(Me.txtphoneno)
        Me.GroupBox5.Controls.Add(Me.Label13)
        Me.GroupBox5.Controls.Add(Me.txtbirthday)
        Me.GroupBox5.Controls.Add(Me.Label16)
        Me.GroupBox5.Controls.Add(Me.Label23)
        Me.GroupBox5.Controls.Add(Me.cbogender)
        Me.GroupBox5.Controls.Add(Me.txtidno)
        Me.GroupBox5.Controls.Add(Me.txtemail)
        Me.GroupBox5.Controls.Add(Me.Label25)
        Me.GroupBox5.Controls.Add(Me.txtpin)
        Me.GroupBox5.Controls.Add(Me.Label26)
        Me.GroupBox5.Controls.Add(Me.Label27)
        Me.GroupBox5.Controls.Add(Me.Label28)
        Me.GroupBox5.Controls.Add(Me.txtpostaladdress)
        Me.GroupBox5.Controls.Add(Me.Label30)
        Me.GroupBox5.Controls.Add(Me.txtcontractend)
        Me.GroupBox5.Controls.Add(Me.Label7)
        Me.GroupBox5.Controls.Add(Me.txtnhifno)
        Me.GroupBox5.Controls.Add(Me.txtnssfno)
        Me.GroupBox5.Controls.Add(Me.Label18)
        Me.GroupBox5.Controls.Add(Me.Label24)
        Me.GroupBox5.Controls.Add(Me.txthourlyrate)
        Me.GroupBox5.Controls.Add(Me.Label19)
        Me.GroupBox5.Controls.Add(Me.Label21)
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.txtnextofkin)
        Me.GroupBox5.Controls.Add(Me.dtpdot)
        Me.GroupBox5.Controls.Add(Me.dtpdoe)
        Me.GroupBox5.Controls.Add(Me.txtmedicalcover)
        Me.GroupBox5.Controls.Add(Me.Label20)
        Me.GroupBox5.Controls.Add(Me.Label6)
        Me.GroupBox5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox5.Location = New System.Drawing.Point(240, 39)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(552, 257)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        '
        'btndepartments
        '
        Me.btndepartments.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndepartments.Location = New System.Drawing.Point(118, 112)
        Me.btndepartments.Name = "btndepartments"
        Me.btndepartments.Size = New System.Drawing.Size(24, 20)
        Me.btndepartments.TabIndex = 13
        Me.btndepartments.Tag = "Add or edit job description"
        Me.btndepartments.Text = "A"
        '
        'txtname
        '
        Me.txtname.Location = New System.Drawing.Point(116, 16)
        Me.txtname.Name = "txtname"
        Me.txtname.Size = New System.Drawing.Size(188, 20)
        Me.txtname.TabIndex = 5
        Me.txtname.Text = ""
        '
        'txtmobileno
        '
        Me.txtmobileno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmobileno.Location = New System.Drawing.Point(116, 232)
        Me.txtmobileno.Name = "txtmobileno"
        Me.txtmobileno.Size = New System.Drawing.Size(188, 20)
        Me.txtmobileno.TabIndex = 22
        Me.txtmobileno.Text = ""
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label12.Location = New System.Drawing.Point(7, 232)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(98, 16)
        Me.Label12.TabIndex = 70
        Me.Label12.Text = "Mobile no"
        '
        'txtphoneno
        '
        Me.txtphoneno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtphoneno.Location = New System.Drawing.Point(116, 210)
        Me.txtphoneno.Name = "txtphoneno"
        Me.txtphoneno.Size = New System.Drawing.Size(188, 20)
        Me.txtphoneno.TabIndex = 20
        Me.txtphoneno.Text = ""
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label13.Location = New System.Drawing.Point(7, 210)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(98, 14)
        Me.Label13.TabIndex = 68
        Me.Label13.Text = "Phone no"
        '
        'txtbirthday
        '
        Me.txtbirthday.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbirthday.Location = New System.Drawing.Point(116, 88)
        Me.txtbirthday.Name = "txtbirthday"
        Me.txtbirthday.Size = New System.Drawing.Size(188, 20)
        Me.txtbirthday.TabIndex = 11
        Me.txtbirthday.Text = ""
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label16.Location = New System.Drawing.Point(4, 88)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(98, 16)
        Me.Label16.TabIndex = 54
        Me.Label16.Text = "Birthday"
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(6, 68)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(98, 12)
        Me.Label23.TabIndex = 4
        Me.Label23.Text = "Id no"
        '
        'cbogender
        '
        Me.cbogender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbogender.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbogender.Items.AddRange(New Object() {"Data developer", "Surveyor", "Processor", "Secretary", "Accountant"})
        Me.cbogender.Location = New System.Drawing.Point(144, 111)
        Me.cbogender.Name = "cbogender"
        Me.cbogender.Size = New System.Drawing.Size(160, 21)
        Me.cbogender.TabIndex = 14
        '
        'txtidno
        '
        Me.txtidno.Location = New System.Drawing.Point(116, 64)
        Me.txtidno.Name = "txtidno"
        Me.txtidno.Size = New System.Drawing.Size(188, 20)
        Me.txtidno.TabIndex = 9
        Me.txtidno.Text = ""
        '
        'txtemail
        '
        Me.txtemail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtemail.Location = New System.Drawing.Point(116, 135)
        Me.txtemail.Name = "txtemail"
        Me.txtemail.Size = New System.Drawing.Size(188, 20)
        Me.txtemail.TabIndex = 16
        Me.txtemail.Text = ""
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label25.Location = New System.Drawing.Point(6, 115)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(98, 16)
        Me.Label25.TabIndex = 44
        Me.Label25.Text = "Job description"
        '
        'txtpin
        '
        Me.txtpin.Location = New System.Drawing.Point(116, 40)
        Me.txtpin.Name = "txtpin"
        Me.txtpin.Size = New System.Drawing.Size(188, 20)
        Me.txtpin.TabIndex = 7
        Me.txtpin.Text = ""
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(6, 41)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(98, 16)
        Me.Label26.TabIndex = 2
        Me.Label26.Text = "Pin no"
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label27.Location = New System.Drawing.Point(5, 140)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(98, 16)
        Me.Label27.TabIndex = 36
        Me.Label27.Text = "E-mail address"
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label28.Location = New System.Drawing.Point(5, 168)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(98, 16)
        Me.Label28.TabIndex = 38
        Me.Label28.Text = "Postal address"
        '
        'txtpostaladdress
        '
        Me.txtpostaladdress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpostaladdress.Location = New System.Drawing.Point(116, 160)
        Me.txtpostaladdress.Multiline = True
        Me.txtpostaladdress.Name = "txtpostaladdress"
        Me.txtpostaladdress.Size = New System.Drawing.Size(188, 48)
        Me.txtpostaladdress.TabIndex = 18
        Me.txtpostaladdress.Text = ""
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(8, 16)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(88, 16)
        Me.Label30.TabIndex = 0
        Me.Label30.Text = "Name"
        '
        'txtcontractend
        '
        Me.txtcontractend.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcontractend.Location = New System.Drawing.Point(400, 19)
        Me.txtcontractend.Name = "txtcontractend"
        Me.txtcontractend.Size = New System.Drawing.Size(144, 20)
        Me.txtcontractend.TabIndex = 6
        Me.txtcontractend.Text = ""
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label7.Location = New System.Drawing.Point(304, 19)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "Contract End Date"
        '
        'txtnhifno
        '
        Me.txtnhifno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnhifno.Location = New System.Drawing.Point(400, 89)
        Me.txtnhifno.Name = "txtnhifno"
        Me.txtnhifno.Size = New System.Drawing.Size(144, 20)
        Me.txtnhifno.TabIndex = 12
        Me.txtnhifno.Text = ""
        '
        'txtnssfno
        '
        Me.txtnssfno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnssfno.Location = New System.Drawing.Point(400, 66)
        Me.txtnssfno.Name = "txtnssfno"
        Me.txtnssfno.Size = New System.Drawing.Size(144, 20)
        Me.txtnssfno.TabIndex = 10
        Me.txtnssfno.Text = ""
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label18.Location = New System.Drawing.Point(304, 66)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(88, 16)
        Me.Label18.TabIndex = 56
        Me.Label18.Text = "NSSF No"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(304, 42)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(88, 16)
        Me.Label24.TabIndex = 72
        Me.Label24.Text = "Hourly rate"
        '
        'txthourlyrate
        '
        Me.txthourlyrate.Location = New System.Drawing.Point(400, 42)
        Me.txthourlyrate.Name = "txthourlyrate"
        Me.txthourlyrate.Size = New System.Drawing.Size(144, 20)
        Me.txthourlyrate.TabIndex = 8
        Me.txthourlyrate.Text = ""
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label19.Location = New System.Drawing.Point(304, 91)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(88, 16)
        Me.Label19.TabIndex = 58
        Me.Label19.Text = "NHIF No"
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label21.Location = New System.Drawing.Point(305, 140)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(80, 16)
        Me.Label21.TabIndex = 62
        Me.Label21.Text = "Next of Kin"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(304, 163)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 32)
        Me.Label4.TabIndex = 68
        Me.Label4.Text = "Date of Employment"
        '
        'txtnextofkin
        '
        Me.txtnextofkin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnextofkin.Location = New System.Drawing.Point(400, 140)
        Me.txtnextofkin.Name = "txtnextofkin"
        Me.txtnextofkin.Size = New System.Drawing.Size(144, 20)
        Me.txtnextofkin.TabIndex = 17
        Me.txtnextofkin.Text = ""
        '
        'dtpdot
        '
        Me.dtpdot.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpdot.Location = New System.Drawing.Point(400, 193)
        Me.dtpdot.Name = "dtpdot"
        Me.dtpdot.Size = New System.Drawing.Size(144, 20)
        Me.dtpdot.TabIndex = 21
        '
        'dtpdoe
        '
        Me.dtpdoe.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpdoe.Location = New System.Drawing.Point(400, 164)
        Me.dtpdoe.Name = "dtpdoe"
        Me.dtpdoe.Size = New System.Drawing.Size(144, 20)
        Me.dtpdoe.TabIndex = 19
        '
        'txtmedicalcover
        '
        Me.txtmedicalcover.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmedicalcover.Location = New System.Drawing.Point(400, 114)
        Me.txtmedicalcover.Name = "txtmedicalcover"
        Me.txtmedicalcover.Size = New System.Drawing.Size(144, 20)
        Me.txtmedicalcover.TabIndex = 15
        Me.txtmedicalcover.Text = ""
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label20.Location = New System.Drawing.Point(304, 114)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(88, 16)
        Me.Label20.TabIndex = 60
        Me.Label20.Text = "Medical Cover "
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label6.Location = New System.Drawing.Point(305, 195)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 32)
        Me.Label6.TabIndex = 70
        Me.Label6.Text = "Date of Termination"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.pbimage)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 38)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(232, 258)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Passport photograph"
        '
        'pbimage
        '
        Me.pbimage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbimage.BackColor = System.Drawing.Color.Gray
        Me.pbimage.Location = New System.Drawing.Point(8, 16)
        Me.pbimage.Name = "pbimage"
        Me.pbimage.Size = New System.Drawing.Size(216, 227)
        Me.pbimage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbimage.TabIndex = 0
        Me.pbimage.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.txtcomments)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(8, 296)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(784, 224)
        Me.GroupBox3.TabIndex = 23
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Comments"
        '
        'txtcomments
        '
        Me.txtcomments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtcomments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcomments.Location = New System.Drawing.Point(8, 16)
        Me.txtcomments.Multiline = True
        Me.txtcomments.Name = "txtcomments"
        Me.txtcomments.Size = New System.Drawing.Size(768, 200)
        Me.txtcomments.TabIndex = 24
        Me.txtcomments.Text = ""
        '
        'tpgtimesheet
        '
        Me.tpgtimesheet.Controls.Add(Me.pnlgrid)
        Me.tpgtimesheet.Controls.Add(Me.pnltop)
        Me.tpgtimesheet.Location = New System.Drawing.Point(4, 22)
        Me.tpgtimesheet.Name = "tpgtimesheet"
        Me.tpgtimesheet.Size = New System.Drawing.Size(794, 530)
        Me.tpgtimesheet.TabIndex = 1
        Me.tpgtimesheet.Text = "Timesheet"
        '
        'pnlgrid
        '
        Me.pnlgrid.Controls.Add(Me.GroupBox13)
        Me.pnlgrid.Controls.Add(Me.GroupBox12)
        Me.pnlgrid.Controls.Add(Me.dtgtimesheet)
        Me.pnlgrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlgrid.Location = New System.Drawing.Point(0, 56)
        Me.pnlgrid.Name = "pnlgrid"
        Me.pnlgrid.Size = New System.Drawing.Size(794, 474)
        Me.pnlgrid.TabIndex = 23
        '
        'GroupBox13
        '
        Me.GroupBox13.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox13.Controls.Add(Me.btnexcel)
        Me.GroupBox13.Controls.Add(Me.btnprint)
        Me.GroupBox13.Location = New System.Drawing.Point(360, 1)
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.Size = New System.Drawing.Size(432, 47)
        Me.GroupBox13.TabIndex = 40
        Me.GroupBox13.TabStop = False
        Me.GroupBox13.Text = "Reporting"
        '
        'btnexcel
        '
        Me.btnexcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnexcel.Location = New System.Drawing.Point(111, 14)
        Me.btnexcel.Name = "btnexcel"
        Me.btnexcel.Size = New System.Drawing.Size(137, 23)
        Me.btnexcel.TabIndex = 40
        Me.btnexcel.Text = "Export to excel"
        '
        'btnprint
        '
        Me.btnprint.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnprint.Location = New System.Drawing.Point(5, 14)
        Me.btnprint.Name = "btnprint"
        Me.btnprint.Size = New System.Drawing.Size(104, 23)
        Me.btnprint.TabIndex = 39
        Me.btnprint.Text = "Print timesheet"
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.btnsavetimesheet)
        Me.GroupBox12.Controls.Add(Me.btnaddentry)
        Me.GroupBox12.Controls.Add(Me.btndelete)
        Me.GroupBox12.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Size = New System.Drawing.Size(352, 48)
        Me.GroupBox12.TabIndex = 39
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Actions"
        '
        'btnsavetimesheet
        '
        Me.btnsavetimesheet.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsavetimesheet.Location = New System.Drawing.Point(106, 14)
        Me.btnsavetimesheet.Name = "btnsavetimesheet"
        Me.btnsavetimesheet.Size = New System.Drawing.Size(112, 23)
        Me.btnsavetimesheet.TabIndex = 40
        Me.btnsavetimesheet.Text = "Save time entries"
        '
        'btnaddentry
        '
        Me.btnaddentry.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnaddentry.Location = New System.Drawing.Point(6, 14)
        Me.btnaddentry.Name = "btnaddentry"
        Me.btnaddentry.Size = New System.Drawing.Size(99, 23)
        Me.btnaddentry.TabIndex = 39
        Me.btnaddentry.Text = "Add time  entry"
        '
        'btndelete
        '
        Me.btndelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelete.Location = New System.Drawing.Point(218, 14)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(128, 23)
        Me.btndelete.TabIndex = 38
        Me.btndelete.Text = "Delete selected entry"
        '
        'dtgtimesheet
        '
        Me.dtgtimesheet.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgtimesheet.CaptionText = "Time sheet details"
        Me.dtgtimesheet.DataMember = ""
        Me.dtgtimesheet.FlatMode = True
        Me.dtgtimesheet.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgtimesheet.Location = New System.Drawing.Point(8, 48)
        Me.dtgtimesheet.Name = "dtgtimesheet"
        Me.dtgtimesheet.ReadOnly = True
        Me.dtgtimesheet.Size = New System.Drawing.Size(778, 417)
        Me.dtgtimesheet.TabIndex = 28
        '
        'pnltop
        '
        Me.pnltop.Controls.Add(Me.btnview)
        Me.pnltop.Controls.Add(Me.lblsdate)
        Me.pnltop.Controls.Add(Me.lblenddate)
        Me.pnltop.Controls.Add(Me.dtpedate)
        Me.pnltop.Controls.Add(Me.dtpsdate)
        Me.pnltop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnltop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnltop.Location = New System.Drawing.Point(0, 0)
        Me.pnltop.Name = "pnltop"
        Me.pnltop.Size = New System.Drawing.Size(794, 56)
        Me.pnltop.TabIndex = 22
        '
        'btnview
        '
        Me.btnview.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnview.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnview.Location = New System.Drawing.Point(185, 23)
        Me.btnview.Name = "btnview"
        Me.btnview.Size = New System.Drawing.Size(111, 20)
        Me.btnview.TabIndex = 33
        Me.btnview.Text = "View timesheet"
        '
        'lblsdate
        '
        Me.lblsdate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblsdate.Location = New System.Drawing.Point(6, 8)
        Me.lblsdate.Name = "lblsdate"
        Me.lblsdate.Size = New System.Drawing.Size(94, 16)
        Me.lblsdate.TabIndex = 24
        Me.lblsdate.Text = "Start date"
        '
        'lblenddate
        '
        Me.lblenddate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblenddate.Location = New System.Drawing.Point(102, 8)
        Me.lblenddate.Name = "lblenddate"
        Me.lblenddate.Size = New System.Drawing.Size(80, 16)
        Me.lblenddate.TabIndex = 22
        Me.lblenddate.Text = "End date"
        '
        'dtpedate
        '
        Me.dtpedate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpedate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpedate.Location = New System.Drawing.Point(100, 24)
        Me.dtpedate.Name = "dtpedate"
        Me.dtpedate.Size = New System.Drawing.Size(80, 20)
        Me.dtpedate.TabIndex = 21
        Me.dtpedate.Value = New Date(2006, 3, 22, 11, 59, 20, 312)
        '
        'dtpsdate
        '
        Me.dtpsdate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpsdate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpsdate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpsdate.Location = New System.Drawing.Point(6, 24)
        Me.dtpsdate.Name = "dtpsdate"
        Me.dtpsdate.Size = New System.Drawing.Size(94, 20)
        Me.dtpsdate.TabIndex = 20
        Me.dtpsdate.Value = New Date(2006, 3, 22, 11, 59, 20, 328)
        '
        'tpgmiscellanous
        '
        Me.tpgmiscellanous.Controls.Add(Me.TabControl2)
        Me.tpgmiscellanous.Location = New System.Drawing.Point(4, 22)
        Me.tpgmiscellanous.Name = "tpgmiscellanous"
        Me.tpgmiscellanous.Size = New System.Drawing.Size(794, 530)
        Me.tpgmiscellanous.TabIndex = 3
        Me.tpgmiscellanous.Text = "Miscellanous"
        '
        'TabControl2
        '
        Me.TabControl2.Alignment = System.Windows.Forms.TabAlignment.Bottom
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(0, 0)
        Me.TabControl2.Multiline = True
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(794, 530)
        Me.TabControl2.TabIndex = 7
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.GroupBox6)
        Me.TabPage3.Controls.Add(Me.GroupBox7)
        Me.TabPage3.Location = New System.Drawing.Point(4, 4)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(786, 502)
        Me.TabPage3.TabIndex = 0
        Me.TabPage3.Text = "Leaves"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Label3)
        Me.GroupBox6.Controls.Add(Me.cboleaves)
        Me.GroupBox6.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox6.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(786, 48)
        Me.GroupBox6.TabIndex = 18
        Me.GroupBox6.TabStop = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Pick an employee"
        '
        'cboleaves
        '
        Me.cboleaves.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboleaves.ItemHeight = 14
        Me.cboleaves.Location = New System.Drawing.Point(112, 16)
        Me.cboleaves.Name = "cboleaves"
        Me.cboleaves.Size = New System.Drawing.Size(272, 22)
        Me.cboleaves.TabIndex = 0
        '
        'GroupBox7
        '
        Me.GroupBox7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox7.Controls.Add(Me.btnleaves)
        Me.GroupBox7.Controls.Add(Me.dtgleaves)
        Me.GroupBox7.Controls.Add(Me.btnaddleave)
        Me.GroupBox7.Controls.Add(Me.Label5)
        Me.GroupBox7.Location = New System.Drawing.Point(1, 48)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(786, 447)
        Me.GroupBox7.TabIndex = 1
        Me.GroupBox7.TabStop = False
        '
        'btnleaves
        '
        Me.btnleaves.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnleaves.Location = New System.Drawing.Point(120, 8)
        Me.btnleaves.Name = "btnleaves"
        Me.btnleaves.Size = New System.Drawing.Size(160, 24)
        Me.btnleaves.TabIndex = 16
        Me.btnleaves.Text = "Delete selected entry"
        '
        'dtgleaves
        '
        Me.dtgleaves.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgleaves.CaptionText = "Leave details"
        Me.dtgleaves.DataMember = ""
        Me.dtgleaves.FlatMode = True
        Me.dtgleaves.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgleaves.Location = New System.Drawing.Point(2, 32)
        Me.dtgleaves.Name = "dtgleaves"
        Me.dtgleaves.PreferredRowHeight = 20
        Me.dtgleaves.ReadOnly = True
        Me.dtgleaves.Size = New System.Drawing.Size(777, 407)
        Me.dtgleaves.TabIndex = 15
        '
        'btnaddleave
        '
        Me.btnaddleave.BackColor = System.Drawing.SystemColors.Control
        Me.btnaddleave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnaddleave.Location = New System.Drawing.Point(87, 13)
        Me.btnaddleave.Name = "btnaddleave"
        Me.btnaddleave.Size = New System.Drawing.Size(24, 16)
        Me.btnaddleave.TabIndex = 14
        Me.btnaddleave.Text = "+"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Location = New System.Drawing.Point(7, 13)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 16)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Leave off"
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.GroupBox2)
        Me.TabPage5.Controls.Add(Me.GroupBox4)
        Me.TabPage5.Location = New System.Drawing.Point(4, 4)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(786, 502)
        Me.TabPage5.TabIndex = 1
        Me.TabPage5.Text = "Sick off"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.cbosickoff)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(786, 48)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Pick an employee"
        '
        'cbosickoff
        '
        Me.cbosickoff.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbosickoff.Location = New System.Drawing.Point(112, 16)
        Me.cbosickoff.Name = "cbosickoff"
        Me.cbosickoff.Size = New System.Drawing.Size(272, 22)
        Me.cbosickoff.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.btndelsickoff)
        Me.GroupBox4.Controls.Add(Me.dtgsickoff)
        Me.GroupBox4.Controls.Add(Me.btnsickoff)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Location = New System.Drawing.Point(0, 47)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(787, 455)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        '
        'btndelsickoff
        '
        Me.btndelsickoff.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelsickoff.Location = New System.Drawing.Point(122, 8)
        Me.btndelsickoff.Name = "btndelsickoff"
        Me.btndelsickoff.Size = New System.Drawing.Size(160, 24)
        Me.btndelsickoff.TabIndex = 17
        Me.btndelsickoff.Text = "Delete selected entry"
        '
        'dtgsickoff
        '
        Me.dtgsickoff.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgsickoff.CaptionText = "Sick off details"
        Me.dtgsickoff.DataMember = ""
        Me.dtgsickoff.FlatMode = True
        Me.dtgsickoff.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgsickoff.Location = New System.Drawing.Point(8, 32)
        Me.dtgsickoff.Name = "dtgsickoff"
        Me.dtgsickoff.PreferredRowHeight = 20
        Me.dtgsickoff.ReadOnly = True
        Me.dtgsickoff.Size = New System.Drawing.Size(771, 415)
        Me.dtgsickoff.TabIndex = 15
        '
        'btnsickoff
        '
        Me.btnsickoff.BackColor = System.Drawing.SystemColors.Control
        Me.btnsickoff.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnsickoff.Location = New System.Drawing.Point(93, 16)
        Me.btnsickoff.Name = "btnsickoff"
        Me.btnsickoff.Size = New System.Drawing.Size(24, 16)
        Me.btnsickoff.TabIndex = 14
        Me.btnsickoff.Text = "+"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Location = New System.Drawing.Point(8, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Sick off"
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.GroupBox8)
        Me.TabPage6.Controls.Add(Me.GroupBox9)
        Me.TabPage6.Location = New System.Drawing.Point(4, 4)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(786, 502)
        Me.TabPage6.TabIndex = 2
        Me.TabPage6.Text = "Time off"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.Label8)
        Me.GroupBox8.Controls.Add(Me.cbotimeoff)
        Me.GroupBox8.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox8.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(786, 48)
        Me.GroupBox8.TabIndex = 4
        Me.GroupBox8.TabStop = False
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Location = New System.Drawing.Point(8, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Pick an employee"
        '
        'cbotimeoff
        '
        Me.cbotimeoff.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbotimeoff.Location = New System.Drawing.Point(112, 16)
        Me.cbotimeoff.Name = "cbotimeoff"
        Me.cbotimeoff.Size = New System.Drawing.Size(272, 22)
        Me.cbotimeoff.TabIndex = 0
        '
        'GroupBox9
        '
        Me.GroupBox9.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox9.Controls.Add(Me.btndeltimeoff)
        Me.GroupBox9.Controls.Add(Me.dtgtimeoff)
        Me.GroupBox9.Controls.Add(Me.btntimeoff)
        Me.GroupBox9.Controls.Add(Me.Label9)
        Me.GroupBox9.Location = New System.Drawing.Point(0, 47)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(787, 455)
        Me.GroupBox9.TabIndex = 3
        Me.GroupBox9.TabStop = False
        '
        'btndeltimeoff
        '
        Me.btndeltimeoff.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeltimeoff.Location = New System.Drawing.Point(123, 8)
        Me.btndeltimeoff.Name = "btndeltimeoff"
        Me.btndeltimeoff.Size = New System.Drawing.Size(160, 24)
        Me.btndeltimeoff.TabIndex = 17
        Me.btndeltimeoff.Text = "Delete selected entry"
        '
        'dtgtimeoff
        '
        Me.dtgtimeoff.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgtimeoff.CaptionText = "Time off details"
        Me.dtgtimeoff.DataMember = ""
        Me.dtgtimeoff.FlatMode = True
        Me.dtgtimeoff.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgtimeoff.Location = New System.Drawing.Point(8, 32)
        Me.dtgtimeoff.Name = "dtgtimeoff"
        Me.dtgtimeoff.PreferredRowHeight = 20
        Me.dtgtimeoff.ReadOnly = True
        Me.dtgtimeoff.Size = New System.Drawing.Size(771, 415)
        Me.dtgtimeoff.TabIndex = 15
        '
        'btntimeoff
        '
        Me.btntimeoff.BackColor = System.Drawing.SystemColors.Control
        Me.btntimeoff.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btntimeoff.Location = New System.Drawing.Point(93, 16)
        Me.btntimeoff.Name = "btntimeoff"
        Me.btntimeoff.Size = New System.Drawing.Size(24, 16)
        Me.btntimeoff.TabIndex = 14
        Me.btntimeoff.Text = "+"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Location = New System.Drawing.Point(8, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(84, 16)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "Time off"
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.GroupBox10)
        Me.TabPage7.Controls.Add(Me.GroupBox11)
        Me.TabPage7.Location = New System.Drawing.Point(4, 4)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Size = New System.Drawing.Size(786, 502)
        Me.TabPage7.TabIndex = 3
        Me.TabPage7.Text = "Day off duty"
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.Label10)
        Me.GroupBox10.Controls.Add(Me.cbodayoff)
        Me.GroupBox10.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox10.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(786, 48)
        Me.GroupBox10.TabIndex = 4
        Me.GroupBox10.TabStop = False
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Location = New System.Drawing.Point(8, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "Pick an employee"
        '
        'cbodayoff
        '
        Me.cbodayoff.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbodayoff.Location = New System.Drawing.Point(112, 16)
        Me.cbodayoff.Name = "cbodayoff"
        Me.cbodayoff.Size = New System.Drawing.Size(272, 22)
        Me.cbodayoff.TabIndex = 0
        '
        'GroupBox11
        '
        Me.GroupBox11.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox11.Controls.Add(Me.btndeldayoff)
        Me.GroupBox11.Controls.Add(Me.dtgdayoff)
        Me.GroupBox11.Controls.Add(Me.btndayoff)
        Me.GroupBox11.Controls.Add(Me.Label11)
        Me.GroupBox11.Location = New System.Drawing.Point(0, 47)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Size = New System.Drawing.Size(787, 455)
        Me.GroupBox11.TabIndex = 3
        Me.GroupBox11.TabStop = False
        '
        'btndeldayoff
        '
        Me.btndeldayoff.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeldayoff.Location = New System.Drawing.Point(122, 8)
        Me.btndeldayoff.Name = "btndeldayoff"
        Me.btndeldayoff.Size = New System.Drawing.Size(160, 24)
        Me.btndeldayoff.TabIndex = 17
        Me.btndeldayoff.Text = "Delete selected entry"
        '
        'dtgdayoff
        '
        Me.dtgdayoff.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgdayoff.CaptionText = "Day off duty details"
        Me.dtgdayoff.DataMember = ""
        Me.dtgdayoff.FlatMode = True
        Me.dtgdayoff.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgdayoff.Location = New System.Drawing.Point(8, 32)
        Me.dtgdayoff.Name = "dtgdayoff"
        Me.dtgdayoff.PreferredRowHeight = 20
        Me.dtgdayoff.ReadOnly = True
        Me.dtgdayoff.Size = New System.Drawing.Size(771, 415)
        Me.dtgdayoff.TabIndex = 15
        '
        'btndayoff
        '
        Me.btndayoff.BackColor = System.Drawing.SystemColors.Control
        Me.btndayoff.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btndayoff.Location = New System.Drawing.Point(93, 16)
        Me.btndayoff.Name = "btndayoff"
        Me.btndayoff.Size = New System.Drawing.Size(24, 16)
        Me.btndayoff.TabIndex = 14
        Me.btndayoff.Text = "+"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Location = New System.Drawing.Point(8, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(84, 16)
        Me.Label11.TabIndex = 13
        Me.Label11.Text = "Day off"
        '
        'Timer1
        '
        Me.Timer1.Interval = 3000
        '
        'frmjobsheet
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(802, 556)
        Me.Controls.Add(Me.tbcjobsheet)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmjobsheet"
        Me.Text = "Personal Details"
        Me.tbcjobsheet.ResumeLayout(False)
        Me.tpgemployees.ResumeLayout(False)
        Me.pnlemployeecontrols.ResumeLayout(False)
        Me.pnljobsheetcontrols.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.tpgtimesheet.ResumeLayout(False)
        Me.pnlgrid.ResumeLayout(False)
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox12.ResumeLayout(False)
        CType(Me.dtgtimesheet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnltop.ResumeLayout(False)
        Me.tpgmiscellanous.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        CType(Me.dtgleaves, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.dtgsickoff, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        CType(Me.dtgtimeoff, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox11.ResumeLayout(False)
        CType(Me.dtgdayoff, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    <System.STAThread()> _
    Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub

#Region "public variables"
    '-------------global and local variables
    Private mid As String
    Public myid As String
    Public hasloaded As Boolean = False
    Public loadgridcontrols As Boolean = False
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htileaves As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htisickoff As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htitimeoff As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htidayoff As System.Windows.Forms.DataGrid.HitTestInfo

    Private mycomboid As String
    Private curr_row As Integer = -123
    Private udesc As Boolean = False
    Private utask As Boolean = False
    Private uidno As Boolean = False
    Private utimespent As Boolean = False
    Private uddate As Boolean = False
    Private ujobtitle As Boolean = False

    Private indexxcombo As Integer

    Private mydescription As String
    Private mytask As String
    Private myidno As String
    Private theiridno As String
    Private mytimespent As String
    Private myddate As String
    Private previousrow As New DataGridCell

    '--------------end of local and global variables
    Public WithEvents comboControl As System.Windows.Forms.ComboBox
    Public WithEvents comboid As System.Windows.Forms.ComboBox
    Public WithEvents combojob As System.Windows.Forms.ComboBox
    Public WithEvents datagridtextBox As DataGridTextBoxColumn
    Public WithEvents datagridtextBox1 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox2 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox3 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox4 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox5 As DataGridTextBoxColumn

    Public WithEvents cbomiscelanous As New System.Windows.Forms.ComboBox
    '------------------other controls okay

    '--------------
    Public WithEvents txttask As System.Windows.Forms.Button
    Public WithEvents txtdesc As System.Windows.Forms.Button
    Public WithEvents txttimespent As AMS.TextBox.MaskedTextBox
    Public WithEvents dtpddate As System.Windows.Forms.DateTimePicker
#End Region

#Region "private members"
    Dim Thread566 As System.Threading.Thread
    Dim Threadsickoff As System.Threading.Thread
    Public imagefilename As String
    Private GridPrinter As DataGridPrinter
    Public namme As String
#End Region

#Region "timesheet/employee details"
    Private Sub frmjobsheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Call loaddata()
            hasloaded = True
        Catch ex As Exception

        End Try
        Try
            Dim _thread As Thread = New Thread(AddressOf ld)
            _thread.IsBackground = True
            _thread.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub loaddata()
        Try
            Dim Tasks As New taskclass
            Tasks.mid = myid
            Dim Thread5 As New System.Threading.Thread( _
                AddressOf taskclass.jobsheetinvoke)
            Thread5.IsBackground = True
            Thread5.Start()

            ' load data for the controls in miscelanous
            Dim Thread55 As New System.Threading.Thread( _
                AddressOf taskclass.cbosinvoke)
            Thread55.IsBackground = True
            Thread55.Start()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Sub addtablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = dtgtimesheet.Width - 20
            mywidth = mywidth / 6

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            'Dim myno As New DataGridTextBoxColumn()
            'myno.MappingName = "job_no"
            'myno.HeaderText = "Job Number"
            'myno.Width = mywidth
            'ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.
            Dim myname1 As New DataGridBoolColumn
            myname1.MappingName = "Edit"
            myname1.HeaderText = "Edit"
            myname1.Width = mywidth
            myname1.AllowNull = False
            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "id_no"
            myname.HeaderText = "Id No"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)
            ' Add a second column style.
            Dim myname100z As New DataGridTextBoxColumn
            myname100z.MappingName = "job_tittle"
            myname100z.HeaderText = "Job Title"
            myname100z.Width = mywidth
            ts1.GridColumnStyles.Add(myname100z)
            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn
            myname100.MappingName = "task"
            myname100.HeaderText = "Task"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)
            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "description"
            mydesc.HeaderText = "Description"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn
            mydesc2.MappingName = "timespent"
            mydesc2.HeaderText = "Time Spent"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc3 As New DataGridTextBoxColumn
            mydesc3.MappingName = "ddate"
            mydesc3.HeaderText = "Date"
            mydesc3.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc3)



            ' Add the DataGridTableStyle objects to the collection.
            dtgtimesheet.TableStyles.Clear()
            dtgtimesheet.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub populategridcontrols(ByVal hti As System.Windows.Forms.DataGrid.HitTestInfo)
        Try
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgtimesheet.DataSource
            'If previousrow.ColumnNumber = 1 Then
            '    If uidno = True Then
            '        If comboid.Text.Trim.Length > 0 Then
            '            ' ds.Tables(0).Rows(previousrow.RowNumber).Item("name") = Me.comboControl.Text
            '            ds.Tables(0).Rows(previousrow.RowNumber).Item("id_no") = Me.comboid.Text
            '        End If
            '        uidno = False
            '    End If
            'End If
            If previousrow.ColumnNumber = 2 Then
                If ujobtitle = True Then
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("job_tittle") = Me.combojob.Text
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("job_no") = mycomboid
                    ujobtitle = False
                End If
            End If
            If previousrow.ColumnNumber = 3 Then
                If utask = True Then
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("task") = Me.txttask.Text
                    utask = False
                End If
            End If
            If previousrow.ColumnNumber = 4 Then
                If udesc = True Then
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("description") = Me.txtdesc.Text
                    udesc = False
                End If
            End If
            If previousrow.ColumnNumber = 5 Then
                If utimespent = True Then
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("timespent") = Me.txttimespent.Text
                    utimespent = False
                End If
            End If
            If previousrow.ColumnNumber = 6 Then
                If uddate = True Then
                    Dim sdate As String
                    sdate = dtpddate.Value.Year & "-" _
                     & dtpddate.Value.Month & "-" _
                     & dtpddate.Value.Day & " " _
                     & dtpddate.Value.Hour & ":" _
                     & dtpddate.Value.Minute & ":" _
                     & dtpddate.Value.Second
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("ddate") = sdate
                    'ds.Tables(0).Rows(previousrow.RowNumber).Item("milliseconds") = dtpddate.Value.Millisecond
                    uddate = False
                End If
            End If
            If hti.Column = 0 Then
                Try
                    Dim discontinuedColumn As Integer = 0
                    Dim pt As Point = Me.dtgtimesheet.PointToClient( _
                        Control.MousePosition)
                    'Dim hti As DataGrid.HitTestInfo = _
                    '    Me.dtgtimesheet.HitTest(pt)
                    Dim bmb As BindingManagerBase = _
                        Me.BindingContext(Me.dtgtimesheet.DataSource, _
                        Me.dtgtimesheet.DataMember)
                    If hti.Row < bmb.Count _
                       AndAlso hti.Type = DataGrid.HitTestType.Cell _
                       AndAlso hti.Column = discontinuedColumn Then
                        Me.dtgtimesheet(hti.Row, discontinuedColumn) = _
                           Not CBool(Me.dtgtimesheet(hti.Row, _
                                 discontinuedColumn))
                    ElseIf hti.Row < bmb.Count _
                       AndAlso hti.Type = DataGrid.HitTestType.Cell Then

                    End If

                Catch ex As Exception

                End Try
            End If
            'If hti.Column = 1 Then
            '    Try
            '        Me.comboid.Text = ds.Tables(0).Rows(hti.Row).Item("id_no")
            '        'Dim w As System.EventArgs
            '        'Me.comboid_SelectedIndexChanged(comboid, w)
            '        'comboControl.SelectedIndex = indexxcombo
            '        'Me.comboControl.Text = ds.Tables(0).Rows(hti.Row).Item("id_no")
            '        'Me.comboid.Text = ds.Tables(0).Rows(hti.Row).Item("id_no")
            '    Catch cf As Exception
            '    End Try
            '    uidno = True
            'End If
            If hti.Column = 2 Then
                Try
                    Me.combojob.Text = ds.Tables(0).Rows(hti.Row).Item("job_tittle")
                Catch cf As Exception
                End Try
                ujobtitle = True
            End If
            If hti.Column = 3 Then
                Try
                    Me.txttask.Text = ds.Tables(0).Rows(hti.Row).Item("task")
                Catch cf As Exception
                End Try
                utask = True
            End If
            If hti.Column = 4 Then
                Try
                    Me.txtdesc.Text = ds.Tables(0).Rows(hti.Row).Item("description")
                Catch cf As Exception
                End Try
                udesc = True
            End If
            If hti.Column = 5 Then
                Try
                    Me.txttimespent.Text = ds.Tables(0).Rows(hti.Row).Item("timespent")
                Catch cf As Exception
                End Try
                utimespent = True
            End If
            If hti.Column = 6 Then
                Try
                    Me.dtpddate.Value = CDate(ds.Tables(0).Rows(hti.Row).Item("ddate"))
                Catch cf As Exception
                End Try
                uddate = True
            End If
            curr_row = hti.Row
        Catch ex As Exception
        Finally
            Me.previousrow = dtgtimesheet.CurrentCell
        End Try
    End Sub
    Private Sub comboControl_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles comboControl.SelectedValueChanged

    End Sub
    Private Sub txttask_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txttask.TextChanged
        Try

        Catch ex As Exception

        End Try
    End Sub
    Private Sub txtdesc_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtdesc.TextChanged
        Try

        Catch ex As Exception

        End Try
    End Sub
    Private Sub txttimespent_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txttimespent.TextChanged
        Try

        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtpddate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpddate.ValueChanged
        Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtpddate_CloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpddate.CloseUp
        Try

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btncompute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            hasloadedjobsheet = False
            Me.Dispose(True)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub comboid_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles comboid.SelectedIndexChanged
        Try
            indexxcombo = Me.comboid.SelectedIndex
        Catch ex As Exception

        End Try
    End Sub
    Private Sub combojob_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles combojob.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.combojob.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.comboid.SelectedIndex = indexx
            Dim strp
            strp = comboid.Text
            mycomboid = strp
            Me.myidno = mycomboid
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub btnsick_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Timer1.Enabled = True
        Timer1.Start()

    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        'While Me.Pone.Height < Me.Height
        '    Me.Pone.Height += Me.Height / 3
        'End While
    End Sub
    Private Sub btnview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnview.Click
        Dim hasbound As Boolean = True
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim dtp2 As New System.Windows.Forms.DateTimePicker
            Dim sdate, edate As String

            sdate = dtpsdate.Value.Year & "-" _
             & dtpsdate.Value.Month & "-" _
             & dtpsdate.Value.Day & " " _
             & "00" & ":" _
             & "00" & ":" _
             & "00"

            dtp2.Value = dtpedate.Value
            edate = dtp2.Value.Year & "-" _
             & dtp2.Value.Month & "-" _
             & dtp2.Value.Day & " " _
             & "23" & ":" _
             & "59" & ":" _
             & "59"

            Dim str As String = "select " _
                   & "rcljobs.job_no,rcljobs.job_tittle,daily_time.* " _
                   & "" _
                   & " from rcljobs inner join daily_time on rcljobs.job_no = daily_time.job_no and " _
                   & " daily_time.ddate>='" & sdate & "'" _
                   & " and daily_time.ddate<'" & edate & "'" _
                   & " and daily_time.id_no='" & myid & "'" _
                  & "  order by daily_time.id_no asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)

                If .BOF = False And .EOF = False Then

                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim custDS As DataSet = New DataSet
                    custDA.Fill(custDS, rs, "leads")
                    Dim tname As String = custDS.Tables(0).TableName()
                    ' Create a column.
                    Dim myColumn = New System.Data.DataColumn
                    myColumn.DataType = Type.GetType("System.Boolean")
                    myColumn.ColumnName = "Edit"
                    myColumn.defaultvalue = False
                    custDS.Tables(0).Columns.Add(myColumn)
                    Dim myColumn1 = New System.Data.DataColumn
                    myColumn1.DataType = Type.GetType("System.String")
                    myColumn1.ColumnName = "isadded"
                    myColumn1.defaultvalue = "1"
                    custDS.Tables(0).Columns.Add(myColumn1)

                    'Dim mycount As Integer = custDS.Tables(0).Rows.Count
                    'Dim i
                    'For i = 0 To mycount - 1
                    '    Try
                    '        custDS.Tables(0).Rows(i).Item("Edit") = False
                    '        custDS.Tables(0).Rows(i).Item("isadded") = "1"
                    '    Catch er As Exception
                    '    End Try
                    '    Try
                    '        'custDS.Tables(0).Rows(i).Item("idno2") = custDS.Tables(0).Rows(i).Item("id_no")
                    '    Catch cv As Exception
                    '    End Try
                    'Next i


                    dtgtimesheet.DataSource = Nothing
                    dtgtimesheet.DataMember = Nothing
                    dtgtimesheet.SetDataBinding(custDS, tname)
                    Call addtablestyle(tname)
                    If Me.loadgridcontrols = False Then
                        Dim Tasks As New taskclass
                        Dim Thread5a As New System.Threading.Thread( _
                            AddressOf taskclass.gridinvoke)
                        Thread5a.Start()
                    End If
                Else
                    dtgtimesheet.DataSource = Nothing
                    hasbound = False
                End If
                '-------------load from archive timesheet
                'Dim str1 As String = "select " _
                '                 & "rcljobs.job_no,rcljobs.job_tittle,daily_time.id_no,daily_time.ddate," _
                '                 & "daily_time.job_no,daily_time.task,daily_time.description,daily_time.timespent,daily_time.ddate,daily_time.milliseconds" _
                '                 & " from rcljobs inner join daily_time on rcljobs.job_no = daily_time.job_no and " _
                '                 & "daily_time.ddate>='" & sdate & "'" _
                '                 & " and daily_time.ddate<'" & edate & "'" _
                '                 & " and daily_time.id_no='" & myid & "'" _
                '                & " order by daily_time.id_no asc"
                'Dim rs1 As New ADODB.Recordset()
                'With rs1
                '    .CursorLocation = CursorLocationEnum.adUseClient
                '    .CursorType = CursorTypeEnum.adOpenStatic
                '    .Open(str1, connect)
                '    If .BOF = False And .EOF = False Then
                '        Dim custDA1 As OleDbDataAdapter = New OleDbDataAdapter()
                '        Dim custDS1 As DataSet = New DataSet()
                '        custDA1.Fill(custDS1, rs1, "archive")
                '        'add new rows to the inital database
                '        If hasbound = True Then
                '            Dim ds As System.Data.DataSet = New System.Data.DataSet()
                '            ds = Me.dtgtimesheet.DataSource
                '            Dim mycount As Integer = custDS1.Tables(0).Rows.Count
                '            Dim i
                '            For i = 0 To mycount - 1
                '                Dim myrow As System.Data.DataRow = ds.Tables(0).NewRow
                '                myrow.Item("id_no") = custDS1.Tables(0).Rows(i).Item("id_no")
                '                myrow.Item("job_no") = custDS1.Tables(0).Rows(i).Item("job_no")
                '                myrow.Item("task") = custDS1.Tables(0).Rows(i).Item("task")
                '                myrow.Item("description") = custDS1.Tables(0).Rows(i).Item("description")
                '                myrow.Item("ddate") = custDS1.Tables(0).Rows(i).Item("ddate")
                '                myrow.Item("timespent") = custDS1.Tables(0).Rows(i).Item("timespent")
                '                myrow.Item("milliseconds") = custDS1.Tables(0).Rows(i).Item("milliseconds")
                '                myrow.Item("job_tittle") = custDS1.Tables(0).Rows(i).Item("job_tittle")
                '                ds.Tables(0).Rows.Add(myrow)
                '            Next i
                '            ds.Tables(0).AcceptChanges()
                '        Else
                '            dtgtimesheet.DataSource = Nothing
                '            dtgtimesheet.DataMember = Nothing
                '            dtgtimesheet.SetDataBinding(custDS1, custDS1.Tables(0).TableName)
                '            Call addtablestyle(custDS1.Tables(0).TableName)
                '        End If

                '    End If
                'End With
                'Try
                '    rs1.Close()
                'Catch ds As Exception
                'End Try

                '----------------end of load from archive timesheet
            End With
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Try
            If Me.txtidno.Text.Trim.Length = 0 Or _
            Me.txtname.Text.Trim.Length = 0 Then
                MessageBox.Show("Please supply a name and a valid identification number")
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
            '-------------check if id number already exists


            'Dim rs As New ADODB.Recordset()
            'Dim Str As String = " select id_no from personnel_info" _
            '& " where lower(id_no) like '" & txtidno.Text.Trim.ToLower() & "'"

            'With rs
            '    .CursorLocation = CursorLocationEnum.adUseClient
            '    .CursorType = CursorTypeEnum.adOpenStatic
            '    .Open(Str, connect)
            '    If .BOF = False And .EOF = False Then
            '        MessageBox.Show("Id no already exists", "Save", _
            '        MessageBoxButtons.OK, MessageBoxIcon.Information)
            '        txtidno.Text = myid
            '        Exit Try
            '    End If
            '    .Close()
            'End With
            '--------------end of sanity check
            Try

                imagefilename = imagefilename.Replace("\", "|")
            Catch xc As Exception
            End Try

            Dim sdate As String 'date of employment
            sdate = Me.dtpdoe.Value.Year & "-" _
             & dtpdoe.Value.Month & "-" _
             & dtpdoe.Value.Day & " " _
             & dtpdoe.Value.Hour & ":" _
             & dtpdoe.Value.Minute & ":" _
             & dtpdoe.Value.Second
            Dim sdate1 As String 'date of termination
            sdate1 = Me.dtpdot.Value.Year & "-" _
             & dtpdot.Value.Month & "-" _
             & dtpdot.Value.Day & " " _
             & dtpdot.Value.Hour & ":" _
             & dtpdot.Value.Minute & ":" _
             & dtpdot.Value.Second
            Dim strin As String = Me.txtcomments.Text.Trim()
            strin = strin.Replace("'", "\'")
            Me.txtcomments.Text = strin
            '-------------------
            Dim arr() As String
            Dim strr, strr2 As String
            Dim y As Integer
            txtcomments.Text = Me.txtcomments.Text.Trim()
            arr = txtcomments.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            '----------------------------------
            Dim strsql As String = "update  personnel_info set "
            strsql += " namme= '" & txtname.Text & "', id_no='" & txtidno.Text.Trim & "',hourly_rate='" & txthourlyrate.Text & "',gender='" & cbogender.Text & "',"
            strsql += " phone_no='" & txtphoneno.Text & "', mobile_no='" & txtmobileno.Text & "',postal_address='" & txtpostaladdress.Text & "',email='" & txtemail.Text & "',pin_no='" & txtpin.Text & "',"
            strsql += " birthday='" & Me.txtbirthday.Text & "', contract_end='" & Me.txtcontractend.Text & "',nssf_no='" & Me.txtnssfno.Text & "', " _
            & " nhif_no='" & Me.txtnhifno.Text & "',medical_cover='" & Me.txtmedicalcover.Text & "',imagefile='" & imagefilename & "',"
            strsql += " dateofemployment='" & sdate & "', nextofkin='" & Me.txtnextofkin.Text & "',dateoftermination='" & sdate1 & "',"
            strsql += " comments='" & strr & "'"
            strsql += " where id_no='" & myid & "';"
            strsql += " update daily_time set id_no='" & txtidno.Text.Trim & "' where id_no='" & myid & "';"
            strsql += " update seccheck set id_no='" & txtidno.Text.Trim & "' where id_no='" & myid & "';"
            connect.BeginTrans()
            connect.Execute(strsql)
            connect.CommitTrans()
            myid = txtidno.Text.Trim
            Dim Tasks As New taskclass
            Dim Thread788 As New System.Threading.Thread( _
                AddressOf taskclass.personnelinvoke)

            Thread788.Start()
            MessageBox.Show("Update are successful", "Employee details", _
            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Try
                connect.Close()
            Catch es As Exception

            End Try
        Catch ex As Exception

        Finally

        End Try
    End Sub
    Private Sub btnpassport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpassport.Click
        Try
            Dim MyImage As Bitmap
            Dim ofd As New System.Windows.Forms.OpenFileDialog
            ofd.Filter = " All image Files(*.BMP;*.JPG;*.GIF;*.JPEG;*.PNG;*.EMF;*.WMF)" _
            & "|*.BMP;*.JPG;*.GIF;*.JPEG;*.PNG;*.EMF;*.WMF"
            ofd.Multiselect = False
            ofd.ShowDialog()
            Dim hb As String = ofd.FileName
            imagefilename = hb
            ' Stretches the image to fit the pictureBox. 
            pbimage.SizeMode = PictureBoxSizeMode.StretchImage
            MyImage = New Bitmap(imagefilename)
            'pbimage.ClientSize = New Size(xSize, ySize)
            pbimage.Image = CType(MyImage, Image)

            'Me.pbimage.Image.FromFile(imagefilename)
            ' pbimage.Refresh()
        Catch sa As Exception

        End Try
    End Sub
    Private Sub dtgtimesheet_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgtimesheet.Click

    End Sub
    Private Sub dtgtimesheet_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgtimesheet.MouseDown
        Try
            hti = dtgtimesheet.HitTest(New Point(e.X, e.Y))
            '-----------hti.Row <> curr_row _ AndAlso 
            If hti.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
            AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
            AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
            AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
            AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.None Then
                Try
                    Call populategridcontrols(hti)
                Catch er456 As Exception

                End Try

            End If

        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgtimesheet_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgtimesheet.CurrentCellChanged
        Try
            'If previousrow.ColumnNumber > 0 Then
            '    Dim ds As System.Data.DataSet = New System.Data.DataSet()
            '    ds = Me.dtgtimesheet.DataSource
            '    If previousrow.ColumnNumber = 1 Then
            '        ds.Tables(0).Rows(previousrow.RowNumber).Item(1) = myidno
            '    End If
            '    If previousrow.ColumnNumber = 2 Then
            '        ds.Tables(0).Rows(previousrow.RowNumber).Item("task") = mytask
            '    End If
            '    If previousrow.ColumnNumber = 3 Then
            '        ds.Tables(0).Rows(previousrow.RowNumber).Item("description") = mydescription
            '    End If
            '    If previousrow.ColumnNumber = 4 Then
            '        ds.Tables(0).Rows(previousrow.RowNumber).Item("timespent") = mytimespent
            '    End If
            '    If previousrow.ColumnNumber = 5 Then
            '        Dim sdate As String = dtpddate.Value.Year & "-" _
            '                 & dtpddate.Value.Month & "-" _
            '                 & dtpddate.Value.Day & " " _
            '                 & dtpddate.Value.Hour & ":" _
            '                 & dtpddate.Value.Minute & ":" _
            '                 & dtpddate.Value.Second
            '        Me.myddate = sdate
            '        ds.Tables(0).Rows(previousrow.RowNumber).Item("ddate") = myddate
            '    End If
            'End If
            'If dtgtimesheet.CurrentCell.ColumnNumber = 1 Then
            '    Me.comboControl.Text = dtgtimesheet(dtgtimesheet.CurrentCell)
            'ElseIf dtgtimesheet.CurrentCell.ColumnNumber = 2 Then
            '    Me.txttask.Text = dtgtimesheet(dtgtimesheet.CurrentCell)
            'ElseIf dtgtimesheet.CurrentCell.ColumnNumber = 3 Then
            '    Me.txtdesc.Text = dtgtimesheet(dtgtimesheet.CurrentCell)
            'ElseIf dtgtimesheet.CurrentCell.ColumnNumber = 4 Then
            '    Me.txttimespent.Text = dtgtimesheet(dtgtimesheet.CurrentCell)
            'ElseIf dtgtimesheet.CurrentCell.ColumnNumber = 5 Then
            '    Me.dtpddate.Text = dtgtimesheet(dtgtimesheet.CurrentCell)
            'End If
        Catch ex As Exception

        Finally

            previousrow = Me.dtgtimesheet.CurrentCell
        End Try
    End Sub
    Private Sub txtnextofkin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtnextofkin.TextChanged

    End Sub
    Private Sub btnexcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexcel.Click
        Try
            Dim ds1 As System.Data.DataSet = New System.Data.DataSet
            Dim ds2 As System.Data.DataSet = New System.Data.DataSet
            ds2 = myForms.jobsheet.dtgtimesheet.DataSource
            ds1 = ds2.Copy
            myForms.Main.dgrid.DataSource = Nothing
            Dim MyTable As New DataTable
            MyTable = ds1.Tables(0)
            Try
                MyTable.Columns.Remove("job_no")
                MyTable.Columns.Remove("id_no")
                MyTable.Columns.Remove("ano")
                MyTable.Columns.Remove("job_no1")
                MyTable.Columns.Remove("milliseconds")
                MyTable.Columns.Remove("Edit")
                MyTable.Columns.Remove("isadded")
                MyTable.Columns.Remove("etime")
                MyTable.Columns.Remove("stime")
                MyTable.Columns.Remove("ddate1")

            Catch cv As Exception
            End Try
            Try
                MyTable.Columns("job_tittle").ColumnName = "Job title"
                MyTable.Columns("ddate").ColumnName = "Date"
                MyTable.Columns("task").ColumnName = "Task"
                MyTable.Columns("description").ColumnName = "Description"
                MyTable.Columns("timespent").ColumnName = "Timespent"
                MyTable.Columns("notes").ColumnName = "Notes"
            Catch cv As Exception
            End Try
            Try
                Dim sfd As System.Windows.Forms.SaveFileDialog _
                = New System.Windows.Forms.SaveFileDialog
                sfd.Filter = "Excel files (*.xls)|*.xls"
                sfd.CheckFileExists = False
                sfd.CheckPathExists = True
                sfd.ValidateNames = True
                sfd.ShowDialog()
                Dim m As String = sfd.FileName
                If m.Trim.Length > 0 Then
                    '------auto sum routine
                    ds1 = autosum(ds1)
                    '---------
                    exporttoexcel.exportexcel.exportToExcel(ds1, m)
                    MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch we As Exception
            End Try
        Catch ex As Exception
        End Try
        'Me.btnview_Click(Me, e)
        Try
            System.GC.Collect()
        Catch qw As Exception

        End Try
    End Sub
    Private Function autosum(ByVal ds1 As DataSet) As DataSet
        Try
            Dim int As Integer = ds1.Tables(0).Rows.Count
            Dim y As Integer = 0
            Dim _sum As Double
            For y = 0 To int - 1
                Application.DoEvents()
                Try
                    _sum = _sum + Convert.ToDouble(ds1.Tables(0).Rows(y).Item("Timespent"))
                Catch ex As Exception

                End Try

            Next
            _sum = Math.Round(_sum, 2)
            Dim myrow As System.Data.DataRow = ds1.Tables(0).NewRow
            myrow.Item("Timespent") = _sum
            ds1.Tables(0).Rows.Add(myrow)
            autosum = ds1
        Catch ex As Exception

        End Try
    End Function
    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
        Try
            dtgtimesheet.Select(hti.Row)
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
            ds = Me.dtgtimesheet.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hti.Row).Item("id_no")
            myseconds = ds.Tables(0).Rows(hti.Row).Item("milliseconds")
            str = "delete from daily_time where milliseconds='" & myseconds & "'"
            str += " and id_no='" & sid & "' and id_no='" & ds.Tables(0).Rows(y).Item("id_no") & "'"
            If ds.Tables(0).Rows(hti.Row).Item("isadded") = "1" Then
                Try
                    connect.BeginTrans()
                    connect.Execute(str)
                    connect.CommitTrans()
                    myrow = ds.Tables(0).Rows(hti.Row)
                    ds.Tables(0).Rows.Remove(myrow)
                Catch cv As Exception
                End Try
            Else
                myrow = ds.Tables(0).Rows(hti.Row)
                ds.Tables(0).Rows.Remove(myrow)
            End If
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnaddentry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddentry.Click
        Try
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgtimesheet.DataSource
            Dim f As Integer = ds.Tables(0).Rows.Count
            Dim myrow As System.Data.DataRow = ds.Tables(0).NewRow
            ds.Tables(0).Rows.Add(myrow)
            ds.Tables(0).Rows(f).Item("Edit") = False
            ds.Tables(0).Rows(f).Item("id_no") = myid
            Dim drt As New System.Windows.Forms.DateTimePicker
            drt.Value = Now
            Dim sdate As String
            sdate = drt.Value.Year & "-" _
             & drt.Value.Month & "-" _
             & drt.Value.Day & " " _
             & drt.Value.Hour & ":" _
             & drt.Value.Minute & ":" _
             & drt.Value.Second
            sdate += "|" & drt.Value.Millisecond
            ds.Tables(0).Rows(f).Item("milliseconds") = sdate
            ds.Tables(0).Rows(f).Item("isadded") = "0"
            'f = ds.Tables(0).Rows.Count
            'myrow = ds.Tables(0).Rows(f - 1)
            'ds.Tables(0).Rows.Remove(myrow)
            'dtgtimesheet.SetDataBinding(ds, ds.Tables(0).TableName)

            'myForms.jobsheet.datagridtextBox.TextBox.Controls.Remove(myForms.jobsheet.comboControl)

            'myForms.jobsheet.datagridtextBox1.TextBox.Controls.Remove(myForms.jobsheet.txttask)

            'myForms.jobsheet.datagridtextBox2.TextBox.Controls.Remove(myForms.jobsheet.txtdesc)

            'myForms.jobsheet.datagridtextBox3.TextBox.Controls.Remove(myForms.jobsheet.txttimespent)

            'myForms.jobsheet.datagridtextBox4.TextBox.Controls.Remove(myForms.jobsheet.dtpddate)
            'Dim Tasks As New taskclass()
            'Dim Thread5a As New System.Threading.Thread( _
            '    AddressOf taskclass.gridinvoke)

            'Thread5a.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnsavetimesheet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsavetimesheet.Click
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgtimesheet.DataSource
            ds.AcceptChanges()
            Dim f As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim strsql As String = ""
            For kappa = 0 To f - 1

                If ds.Tables(0).Rows(kappa).Item("isadded") = "0" Then
                    If Convert.ToString(ds.Tables(0).Rows(kappa).Item("job_tittle")).Length > 0 Then
                        'MsgBox("true")
                        strsql = "insert into daily_time (id_no,job_no,task,description,ddate,timespent,milliseconds) "
                        strsql += "values ('" & myid & "','" & ds.Tables(0).Rows(kappa).Item("job_no") & "','" & ds.Tables(0).Rows(kappa).Item("task") & "',"
                        strsql += "'" & ds.Tables(0).Rows(kappa).Item("description") & "','" & ds.Tables(0).Rows(kappa).Item("ddate") & "','" & ds.Tables(0).Rows(kappa).Item("timespent") & "','" & ds.Tables(0).Rows(kappa).Item("milliseconds") & "')"

                        Try
                            connect.BeginTrans()
                            connect.Execute(strsql)
                            connect.CommitTrans()
                            ds.Tables(0).Rows(kappa).Item("isadded") = "1"
                        Catch xc As Exception
                            Try
                                connect.RollbackTrans()
                            Catch er As Exception
                            End Try
                        End Try
                    End If

                Else
                    If ds.Tables(0).Rows(kappa).Item("Edit") = True Then
                        ' Convert.ToString(ds.Tables(0).Rows(kappa).Item("ddate")).Length < 1 _
                        ''''''''AndAlso 
                        If Convert.ToString(ds.Tables(0).Rows(kappa).Item("job_tittle")).Length > 0 Then

                            strsql = "update  daily_time set   "
                            strsql += " id_no='" & ds.Tables(0).Rows(kappa).Item("id_no") & "',job_no='" & ds.Tables(0).Rows(kappa).Item("job_no") & "'," _
                            & "task='" & ds.Tables(0).Rows(kappa).Item("task") & "',"
                            strsql += "description='" & ds.Tables(0).Rows(kappa).Item("description") & "'," _
                            & "ddate='" & ds.Tables(0).Rows(kappa).Item("ddate") & "'," _
                            & "timespent='" & ds.Tables(0).Rows(kappa).Item("timespent") & "'"


                            strsql += " where   ano='" & ds.Tables(0).Rows(kappa).Item("ano") & "'"
                            'strsql += " and   id_no='" & ds.Tables(0).Rows(kappa).Item("id_no") & "'"
                            Try
                                connect.BeginTrans()
                                connect.Execute(strsql)
                                connect.CommitTrans()
                            Catch xc As Exception
                                Try
                                    connect.RollbackTrans()
                                Catch er As Exception
                                End Try
                            End Try

                            Try
                                ds.Tables(0).Rows(kappa).Item("idno2") = ds.Tables(0).Rows(kappa).Item("id_no")
                            Catch er As Exception
                            End Try
                        End If

                    End If
                End If


                System.Windows.Forms.Application.DoEvents()

            Next
            MessageBox.Show("Successfully updated", "Timesheet", _
            MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnprint.Click
        Try
            Dim ds1 As System.Data.DataSet = New System.Data.DataSet
            ds1 = myForms.jobsheet.dtgtimesheet.DataSource
            myForms.Main.dgrid.DataSource = Nothing
            Dim MyTable As New DataTable
            MyTable = ds1.Tables(0)
            Try
                MyTable.Columns.Remove("job_no")
                MyTable.Columns.Remove("id_no")
                MyTable.Columns.Remove("ano")
                MyTable.Columns.Remove("job_no1")
                MyTable.Columns.Remove("milliseconds")
                MyTable.Columns.Remove("Edit")
                MyTable.Columns.Remove("ddate1")
                MyTable.Columns.Remove("isadded")
            Catch cv As Exception
            End Try
            Try
                MyTable.Columns("job_tittle").ColumnName = "Job title"
                MyTable.Columns("ddate").ColumnName = "Date"
                MyTable.Columns("task").ColumnName = "Task"
                MyTable.Columns("description").ColumnName = "Description"
                MyTable.Columns("timespent").ColumnName = "Timespent"
            Catch cv As Exception
            End Try
            myForms.Main.dgrid.DataSource = MyTable
            GridPrinter = New DataGridPrinter(myForms.Main.dgrid)
            '--------------page set up
            Try
                With myForms.Main.PageSetupDialog1
                    .Document = GridPrinter.PrintDocument
                    .ShowDialog()
                End With
            Catch xc As Exception
            End Try

            '----------------
            myForms.Main.TextBox2.Text = namme
            With GridPrinter
                .HeaderText = myForms.Main.TextBox2.Text
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
            With myForms.Main.PrintPreviewDialog1
                .Document = GridPrinter.PrintDocument
                If .ShowDialog = DialogResult.OK Then
                    GridPrinter.Print()
                End If
            End With
            Try
                Me.btnview_Click(Me, e)
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
        myForms.Main.Invalidate()
    End Sub

#Region "loaddepartments"

#Region "department members"
    Private Delegate Sub mydelegate()
#End Region

    Public Sub loaddepart()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str = "select * from jobdescription"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                myForms.jobsheet.cbogender.Items.Clear()
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While rs.EOF = False
                        myForms.jobsheet.cbogender.Items.Add(.Fields("description").Value)
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
            myForms.jobsheet.Invoke(New mydelegate(AddressOf loaddepart))
        Catch ex As Exception

        End Try
    End Sub
#End Region

#End Region

#Region "Miscellanous"
    Private Sub cboleaves_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboleaves.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cboleaves.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.cbomiscelanous.SelectedIndex = indexx
            mid = Me.cbomiscelanous.Text
            Try
                If Thread566.IsAlive = True Then
                    Try
                        Thread566.Abort()
                    Catch xc As Exception
                    End Try
                End If
            Catch we As Exception
            End Try

            Dim Tasks As New taskclass
            Tasks.mid = Me.cbomiscelanous.Text
            Thread566 = New System.Threading.Thread( _
                AddressOf Tasks.leavesinvoke)
            Thread566.IsBackground = True
            Thread566.Start()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub cbotimeoff_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbotimeoff.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cbotimeoff.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.cbomiscelanous.SelectedIndex = indexx
            mid = Me.cbomiscelanous.Text
            Try
                If Thread566.IsAlive = True Then
                    Try
                        Thread566.Abort()
                    Catch xc As Exception
                    End Try
                End If
            Catch we As Exception
            End Try

            Dim Tasks As New taskclass
            Tasks.mid = Me.cbomiscelanous.Text
            Thread566 = New System.Threading.Thread( _
                AddressOf Tasks.timeoffinvoke)
            Thread566.IsBackground = True
            Thread566.Start()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub cbosickoff_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbosickoff.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cbosickoff.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.cbomiscelanous.SelectedIndex = indexx
            mid = Me.cbomiscelanous.Text
            Try
                If Threadsickoff.IsAlive = True Then
                    Try
                        Threadsickoff.Abort()
                    Catch xc As Exception
                    End Try
                End If
            Catch we As Exception
            End Try

            Dim Tasks As New taskclass
            Tasks.mid = Me.cbomiscelanous.Text
            Threadsickoff = New System.Threading.Thread( _
                AddressOf Tasks.sickoffinvoke)
            Threadsickoff.IsBackground = True
            Threadsickoff.Start()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub cbodayoff_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbodayoff.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cbodayoff.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.cbomiscelanous.SelectedIndex = indexx
            mid = Me.cbomiscelanous.Text
            Try
                If Thread566.IsAlive = True Then
                    Try
                        Thread566.Abort()
                    Catch xc As Exception
                    End Try
                End If
            Catch we As Exception
            End Try

            Dim Tasks As New taskclass
            Tasks.mid = Me.cbomiscelanous.Text
            Thread566 = New System.Threading.Thread( _
                AddressOf Tasks.dayoffinvoke)
            Thread566.IsBackground = True
            Thread566.Start()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub btnaddleave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddleave.Click
        Try
            If Me.cboleaves.SelectedIndex = -1 Then
                MessageBox.Show("Please make a selection", "Add leave", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim gd As New frmaddsickleave
            gd.txtemployeename.Text = Me.cboleaves.Text
            gd.mid = mid
            gd.ShowDialog()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnsickoff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsickoff.Click
        Try
            If Me.cbosickoff.SelectedIndex = -1 Then
                MessageBox.Show("Please make a selection", "Add leave", _
               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim gd As New frmaddsickleave
            gd.txtemployeename.Text = Me.cbosickoff.Text
            gd.Text = "Add sick off"
            gd.mid = mid
            gd.ShowDialog()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btntimeoff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btntimeoff.Click
        Try
            If Me.cbotimeoff.SelectedIndex = -1 Then
                MessageBox.Show("Please make a selection", "Add leave", _
               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim gd As New frmtimeoff
            gd.txtname.Text = Me.cbotimeoff.Text
            gd.mid = mid
            gd.ShowDialog()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btndayoff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndayoff.Click
        Try
            If Me.cbodayoff.SelectedIndex = -1 Then
                MessageBox.Show("Please make a selection", "Add leave", _
               MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim gd As New frmtimeoff
            gd.txtname.Text = Me.cbodayoff.Text
            gd.Text = "Add Day off"
            gd.mid = mid
            gd.ShowDialog()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtgleaves_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgleaves.MouseDown
        Try
            htileaves = dtgleaves.HitTest(New Point(e.X, e.Y))
            '-----------hti.Row <> curr_row _ AndAlso 
            If htileaves.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
            AndAlso htileaves.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
            AndAlso htileaves.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
            AndAlso htileaves.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
            AndAlso htileaves.Type <> Windows.Forms.DataGrid.HitTestType.None Then
                Try

                Catch er456 As Exception

                End Try

            End If

        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgsickoff_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgsickoff.DoubleClick
        Try
            If htisickoff.Type = Windows.Forms.DataGrid.HitTestType.Caption _
           Or htisickoff.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
           Or htisickoff.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
           Or htisickoff.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
           Or htisickoff.Type = Windows.Forms.DataGrid.HitTestType.None _
           Or htisickoff.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                 Then
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgsickoff.DataSource
            Try
                myForms.sickleave.txtemployeename.Text = Me.cbosickoff.Text
                myForms.sickleave.Text = "Edit sick off"
                myForms.sickleave.mid = mid
                myForms.sickleave.btnadd.Text = "Edit"
                Try
                    myForms.sickleave.txtdesc.Text = ds.Tables(0).Rows(htisickoff.Row).Item("description")
                Catch sd As Exception
                End Try
                Try
                    myForms.sickleave.dtpsdate.Text = ds.Tables(0).Rows(htisickoff.Row).Item("sdate")
                Catch sd As Exception
                End Try
                Try
                    myForms.sickleave.dtpedate.Text = ds.Tables(0).Rows(htisickoff.Row).Item("edate")
                Catch sd As Exception
                End Try
                Try
                    myForms.sickleave.autono = ds.Tables(0).Rows(htisickoff.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch saw As Exception
                Try
                    Dim gd As New frmaddsickleave
                    myForms.sickleave = gd
                    myForms.sickleave.txtemployeename.Text = Me.cbosickoff.Text
                    myForms.sickleave.Text = "Edit sick off"
                    myForms.sickleave.mid = mid
                    myForms.sickleave.btnadd.Text = "Edit"
                    Try
                        myForms.sickleave.txtdesc.Text = ds.Tables(0).Rows(htisickoff.Row).Item("description")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.sickleave.dtpsdate.Text = ds.Tables(0).Rows(htisickoff.Row).Item("sdate")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.sickleave.dtpedate.Text = ds.Tables(0).Rows(htisickoff.Row).Item("edate")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.sickleave.autono = ds.Tables(0).Rows(htisickoff.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.sickleave.Show()
                Catch sf As Exception

                End Try
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtgdayoff_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgdayoff.DoubleClick
        Try
            If htidayoff.Type = Windows.Forms.DataGrid.HitTestType.Caption _
           Or htidayoff.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
           Or htidayoff.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
           Or htidayoff.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
           Or htidayoff.Type = Windows.Forms.DataGrid.HitTestType.None _
           Or htidayoff.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                 Then
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgdayoff.DataSource
            Try
                myForms.timeoff.txtname.Text = Me.cbodayoff.Text
                myForms.timeoff.Text = "Edit Day off"
                myForms.timeoff.mid = mid
                myForms.timeoff.btnadd.Text = "Edit"
                Try
                    myForms.timeoff.txtdesc.Text = ds.Tables(0).Rows(htidayoff.Row).Item("description")
                Catch sd As Exception
                End Try
                Try
                    myForms.timeoff.dtpdayoff.Text = ds.Tables(0).Rows(htidayoff.Row).Item("dateoff")
                Catch sd As Exception
                End Try
                Try
                    myForms.timeoff.dtptimeoff.Text = ds.Tables(0).Rows(htidayoff.Row).Item("timeoff")
                Catch sd As Exception
                End Try
                Try
                    myForms.timeoff.autono = ds.Tables(0).Rows(htidayoff.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch saw As Exception
                Try
                    Dim gd As New frmtimeoff
                    myForms.timeoff = gd
                    myForms.timeoff.txtname.Text = Me.cbodayoff.Text
                    myForms.timeoff.Text = "Edit Day off"
                    myForms.timeoff.mid = mid
                    myForms.timeoff.btnadd.Text = "Edit"
                    Try
                        myForms.timeoff.txtdesc.Text = ds.Tables(0).Rows(htidayoff.Row).Item("description")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.timeoff.dtpdayoff.Text = ds.Tables(0).Rows(htidayoff.Row).Item("dateoff")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.timeoff.dtptimeoff.Text = ds.Tables(0).Rows(htidayoff.Row).Item("timeoff")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.timeoff.autono = ds.Tables(0).Rows(htidayoff.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.timeoff.Show()
                Catch sf As Exception

                End Try
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtgtimeoff_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgtimeoff.MouseDown
        Try
            htitimeoff = dtgtimeoff.HitTest(New Point(e.X, e.Y))
            '-----------hti.Row <> curr_row _ AndAlso 
            If htitimeoff.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
            AndAlso htitimeoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
            AndAlso htitimeoff.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
            AndAlso htitimeoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
            AndAlso htitimeoff.Type <> Windows.Forms.DataGrid.HitTestType.None Then
                Try

                Catch er456 As Exception

                End Try

            End If

        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgsickoff_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgsickoff.MouseDown
        Try
            htisickoff = dtgsickoff.HitTest(New Point(e.X, e.Y))
            '-----------hti.Row <> curr_row _ AndAlso 
            If htisickoff.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
            AndAlso htisickoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
            AndAlso htisickoff.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
            AndAlso htisickoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
            AndAlso htisickoff.Type <> Windows.Forms.DataGrid.HitTestType.None Then
                Try

                Catch er456 As Exception

                End Try

            End If

        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgdayoff_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgdayoff.MouseDown
        Try
            htidayoff = dtgdayoff.HitTest(New Point(e.X, e.Y))
            '-----------hti.Row <> curr_row _ AndAlso 
            If htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
            AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
            AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
            AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
            AndAlso htidayoff.Type <> Windows.Forms.DataGrid.HitTestType.None Then
                Try

                Catch er456 As Exception

                End Try

            End If

        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgleaves_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgleaves.DoubleClick
        Try
            If htileaves.Type = Windows.Forms.DataGrid.HitTestType.Caption _
             Or htileaves.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
             Or htileaves.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
             Or htileaves.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
             Or htileaves.Type = Windows.Forms.DataGrid.HitTestType.None _
             Or htileaves.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                   Then
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgleaves.DataSource
            Try
                myForms.sickleave.txtemployeename.Text = Me.cboleaves.Text
                myForms.sickleave.Text = "Edit leave"
                myForms.sickleave.mid = mid
                myForms.sickleave.btnadd.Text = "Edit"
                Try
                    myForms.sickleave.txtdesc.Text = ds.Tables(0).Rows(htileaves.Row).Item("description")
                Catch sd As Exception
                End Try
                Try
                    myForms.sickleave.dtpsdate.Text = ds.Tables(0).Rows(htileaves.Row).Item("sdate")
                Catch sd As Exception
                End Try
                Try
                    myForms.sickleave.dtpedate.Text = ds.Tables(0).Rows(htileaves.Row).Item("edate")
                Catch sd As Exception
                End Try
                Try
                    myForms.sickleave.autono = ds.Tables(0).Rows(htileaves.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch saw As Exception
                Try
                    Dim gd As New frmaddsickleave
                    myForms.sickleave = gd
                    myForms.sickleave.txtemployeename.Text = Me.cboleaves.Text
                    myForms.sickleave.Text = "Edit leave"
                    myForms.sickleave.mid = mid
                    myForms.sickleave.btnadd.Text = "Edit"
                    Try
                        myForms.sickleave.txtdesc.Text = ds.Tables(0).Rows(htileaves.Row).Item("description")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.sickleave.dtpsdate.Text = ds.Tables(0).Rows(htileaves.Row).Item("sdate")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.sickleave.dtpedate.Text = ds.Tables(0).Rows(htileaves.Row).Item("edate")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.sickleave.autono = ds.Tables(0).Rows(htileaves.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.sickleave.Show()
                Catch za As Exception

                End Try
            End Try


        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtgtimeoff_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgtimeoff.DoubleClick
        Try
            If htitimeoff.Type = Windows.Forms.DataGrid.HitTestType.Caption _
           Or htitimeoff.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader _
           Or htitimeoff.Type = Windows.Forms.DataGrid.HitTestType.RowResize _
           Or htitimeoff.Type = Windows.Forms.DataGrid.HitTestType.ColumnResize _
           Or htitimeoff.Type = Windows.Forms.DataGrid.HitTestType.None _
           Or htitimeoff.Type = Windows.Forms.DataGrid.HitTestType.Cell _
                                                                 Then
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgtimeoff.DataSource
            Try
                myForms.timeoff.txtname.Text = Me.cbotimeoff.Text
                myForms.timeoff.Text = "Edit time off"
                myForms.timeoff.mid = mid
                myForms.timeoff.btnadd.Text = "Edit"
                Try
                    myForms.timeoff.txtdesc.Text = ds.Tables(0).Rows(htitimeoff.Row).Item("description")
                Catch sd As Exception
                End Try
                Try
                    myForms.timeoff.dtpdayoff.Text = ds.Tables(0).Rows(htitimeoff.Row).Item("dateoff")
                Catch sd As Exception
                End Try
                Try
                    myForms.timeoff.dtptimeoff.Text = ds.Tables(0).Rows(htitimeoff.Row).Item("timeoff")
                Catch sd As Exception
                End Try
                Try
                    myForms.timeoff.autono = ds.Tables(0).Rows(htitimeoff.Row).Item("ano")
                Catch sd As Exception
                End Try
            Catch saw As Exception
                Try
                    Dim gd As New frmtimeoff
                    myForms.timeoff = gd
                    myForms.timeoff.txtname.Text = Me.cbotimeoff.Text
                    myForms.timeoff.Text = "Edit time off"
                    myForms.timeoff.mid = mid
                    myForms.timeoff.btnadd.Text = "Edit"
                    Try
                        myForms.timeoff.txtdesc.Text = ds.Tables(0).Rows(htitimeoff.Row).Item("description")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.timeoff.dtpdayoff.Text = ds.Tables(0).Rows(htitimeoff.Row).Item("dateoff")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.timeoff.dtptimeoff.Text = ds.Tables(0).Rows(htitimeoff.Row).Item("timeoff")
                    Catch sd As Exception
                    End Try
                    Try
                        myForms.timeoff.autono = ds.Tables(0).Rows(htitimeoff.Row).Item("ano")
                    Catch sd As Exception
                    End Try
                    myForms.timeoff.Show()
                Catch sf As Exception

                End Try
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub tbcjobsheet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbcjobsheet.SelectedIndexChanged
        Try
            If tbcjobsheet.SelectedTab Is Me.tpgmiscellanous Then
                Try
                    'Dim task As New taskclass()
                    'task.loadname()
                    'Me.cboleaves.Text = task.globalnamme
                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                      & ev.InnerException().ToString() & vbCrLf _
                      & ev.StackTrace.ToString())
                End Try
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

#Region "delete miscellaneous"
    Private Sub btnleaves_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnleaves.Click
        Try
            myForms.jobsheet.dtgleaves.Select(htileaves.Row)
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
            ds = myForms.jobsheet.dtgleaves.DataSource
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htileaves.Row).Item("ano")
            str = "delete from leaves where ano='" & sid & "';"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(htileaves.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btndeltimeoff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeltimeoff.Click
        Try
            myForms.jobsheet.dtgtimeoff.Select(htitimeoff.Row)
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
            ds = myForms.jobsheet.dtgtimeoff.DataSource
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htitimeoff.Row).Item("ano")
            str = "delete from timeoff where ano='" & sid & "';"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(htitimeoff.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btndelsickoff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelsickoff.Click
        Try
            myForms.jobsheet.dtgsickoff.Select(htisickoff.Row)
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
            ds = myForms.jobsheet.dtgsickoff.DataSource
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htisickoff.Row).Item("ano")
            str = "delete from sickoff where ano='" & sid & "';"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(htisickoff.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btndeldayoff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeldayoff.Click
        Try
            myForms.jobsheet.dtgdayoff.Select(htidayoff.Row)
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
            ds = myForms.jobsheet.dtgdayoff.DataSource
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htidayoff.Row).Item("ano")
            str = "delete from dayoff where ano='" & sid & "';"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(htidayoff.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
#End Region

    Private Sub btndepartments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndepartments.Click
        Try
            Dim n As New frmjobdescription
            n.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub 'job description
End Class
