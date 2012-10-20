
Imports System
Imports System.Threading
Imports System.Drawing
Imports System.Windows.Forms
Imports ADODB
Imports System.ArgumentOutOfRangeException
Imports System.Exception

Imports System.Data
Imports System.Data.OleDb
Imports System.IO
'Imports ReportViewerLib
'Imports PinkieControls
'Imports SHDocVw


Public Class frmHome
    Inherits System.Windows.Forms.Form
    ' Cached GDI+ objects created in Form's constructor.
    ' Used in the event handlers for the custom column styles events.

    Private disabledBackBrush As Brush
    Private disabledTextBrush As Brush
    Private currentRowFont As Font
    Private currentRowBackBrush As Brush
    Public mythread As System.Threading.Thread

    Public seclevel As String
    ' Grid tootips-  fields used by the grid's mousemove events to manage.
    '   Row specific tips - initialized in Form_Load, used in dataGrid1_MouseMove.

    Private hitRow As Integer
    Private toolTip1 As System.Windows.Forms.ToolTip
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htijobs As System.Windows.Forms.DataGrid.HitTestInfo
    Private htileads As System.Windows.Forms.DataGrid.HitTestInfo
    Public Delegate Sub mydelegate()
    Private isloadleads As Boolean = False
    Private resizegrid As Integer = 0
    Private isfrmhomeloading As Boolean = True

    '-----------threads
    Private Threadleads As System.Threading.Thread
    Private Threadclients As System.Threading.Thread

    Public cboequipsearchtrue As New System.Windows.Forms.ComboBox()
    Private isdate As Boolean = False

#Region " Windows Form Designer generated code "
    Private tabArea As Rectangle
    Private tabTextArea As RectangleF
    <System.STAThread()> _
     Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.DoEvents()
        Application.Run(New frmHome)
    End Sub
    Public Sub New()
        MyBase.New()
        Dim handler As ThreadExceptionHandler = _
                   New ThreadExceptionHandler

        AddHandler Application.ThreadException, _
            AddressOf handler.Application_ThreadException
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Call PopulateColourList(Me.ComboBox_ColourBodyline)
        Call PopulateColourList(Me.ComboBox_ColourFooterLine)
        Call PopulateColourList(Me.ComboBox_ColourHeaderLine)

        Call PopulateBrushList(Me.ComboBox_EvenBrush)
        Call PopulateBrushList(Me.ComboBox_FooterBrush)
        Call PopulateBrushList(Me.ComboBox_HeaderBrush)
        Call PopulateBrushList(Me.ComboBox_OddRowBrush)
        Call PopulateBrushList(Me.ComboBox_ColumnHeaderBrush)
        '------------draw text on tab

        tabArea = tbcHome.GetTabRect(0)
        tabTextArea = RectangleF.op_Implicit(tbcHome.GetTabRect(0))

        ' Binds the event handler DrawOnTab to the DrawItem event 
        ' through the DrawItemEventHandler delegate.
        Me.SetStyle(ControlStyles.UserPaint, True)
        tbcHome.ItemSize = New Size(0, 15)
        AddHandler tbcHome.DrawItem, AddressOf DrawOnTab

        '---------------

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            Try
                isloading = False
                Call closeprogram()
            Catch ex As Exception
            End Try
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
    Friend WithEvents mnuMainMenu As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuView As System.Windows.Forms.MenuItem
    Friend WithEvents pnlleads As System.Windows.Forms.Panel
    Friend WithEvents dtgClients As System.Windows.Forms.DataGrid
    Friend WithEvents dtgLeads As System.Windows.Forms.DataGrid
    Friend WithEvents dtpedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpsdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents grpcontactgrid As System.Windows.Forms.GroupBox
    Friend WithEvents grpContactsearch As System.Windows.Forms.GroupBox
    Friend WithEvents grpleadsgrid As System.Windows.Forms.GroupBox
    Friend WithEvents grpLeadssearch As System.Windows.Forms.GroupBox
    Friend WithEvents pnlContacts As System.Windows.Forms.Panel
    Friend WithEvents lblleadenddate As System.Windows.Forms.Label
    Friend WithEvents lblleadstartdate As System.Windows.Forms.Label
    Friend WithEvents lblleadname As System.Windows.Forms.Label
    Friend WithEvents lblleadstatus As System.Windows.Forms.Label
    Friend WithEvents mnufileclose As System.Windows.Forms.MenuItem
    Friend WithEvents mnueditcontacts As System.Windows.Forms.MenuItem
    Friend WithEvents mnueditleads As System.Windows.Forms.MenuItem
    Friend WithEvents mnueditjobs As System.Windows.Forms.MenuItem
    Friend WithEvents mnuviewhome As System.Windows.Forms.MenuItem
    Friend WithEvents mnuviewcontacts As System.Windows.Forms.MenuItem
    Friend WithEvents mnuviewleads As System.Windows.Forms.MenuItem
    Friend WithEvents mnuviewjobs As System.Windows.Forms.MenuItem
    Friend WithEvents mnuviewequip As System.Windows.Forms.MenuItem
    Friend WithEvents mnuviewpersonnel As System.Windows.Forms.MenuItem
    Friend WithEvents mnuReports As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ToolBar2 As System.Windows.Forms.ToolBar
    Friend WithEvents tlbtimesheet As System.Windows.Forms.ToolBarButton
    Friend WithEvents pnlpersonnel As System.Windows.Forms.Panel
    Friend WithEvents pnlequip As System.Windows.Forms.Panel
    Friend WithEvents AxMaskEdBox1 As AxMSMask.AxMaskEdBox
    Friend WithEvents tlbadmin As System.Windows.Forms.ToolBarButton
    Friend WithEvents pnlequipcontrols As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnljobs As System.Windows.Forms.Panel
    Friend WithEvents grpjobsearch As System.Windows.Forms.GroupBox
    Friend WithEvents lbljobedate As System.Windows.Forms.Label
    Friend WithEvents lbljobsdate As System.Windows.Forms.Label
    Friend WithEvents dtpjobedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpjobsdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbljobstatus As System.Windows.Forms.Label
    Friend WithEvents grpjobs As System.Windows.Forms.GroupBox
    Friend WithEvents dtgJobs As System.Windows.Forms.DataGrid
    Friend WithEvents mnufilesettings As System.Windows.Forms.MenuItem
    Friend WithEvents mnufilesettingsadministrator As System.Windows.Forms.MenuItem
    Friend WithEvents tpgLeads As System.Windows.Forms.TabPage
    Friend WithEvents tpgClients As System.Windows.Forms.TabPage
    Friend WithEvents tpgJobs As System.Windows.Forms.TabPage
    Friend WithEvents tpgEquip As System.Windows.Forms.TabPage
    Friend WithEvents tpgPersonnel As System.Windows.Forms.TabPage
    Friend WithEvents tpgHome As System.Windows.Forms.TabPage
    Friend WithEvents tbcHome As System.Windows.Forms.TabControl
    Friend WithEvents cboequipsearch As System.Windows.Forms.ComboBox
    Friend WithEvents btnequipsearch As System.Windows.Forms.Button
    Friend WithEvents lblsearchparameter As System.Windows.Forms.Label
    Friend WithEvents cboleads As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtleads As System.Windows.Forms.TextBox
    Friend WithEvents btnleadssearch As System.Windows.Forms.Button
    Friend WithEvents btnleadspho As System.Windows.Forms.Button
    Friend WithEvents btnfailedlead As System.Windows.Forms.Button
    Friend WithEvents btnaddleads As System.Windows.Forms.Button
    Friend WithEvents btnShowAllleads As System.Windows.Forms.Button
    Friend WithEvents btnleadssuspect As System.Windows.Forms.Button
    Friend WithEvents btnleadsprospect As System.Windows.Forms.Button
    Friend WithEvents btnleadsproposal As System.Windows.Forms.Button
    Friend WithEvents btnsearchname As System.Windows.Forms.Button
    Friend WithEvents txtparams As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnaddnew As System.Windows.Forms.Button
    Friend WithEvents btnshowall As System.Windows.Forms.Button
    Friend WithEvents cbojobstatus As System.Windows.Forms.ComboBox
    Friend WithEvents txtjobcontactname As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboleadssearchfield As System.Windows.Forms.ComboBox
    Friend WithEvents cboclientssearchfield As System.Windows.Forms.ComboBox
    Friend WithEvents cbojobsearchfield As System.Windows.Forms.ComboBox
    Friend WithEvents btnjobsearch As System.Windows.Forms.Button
    Friend WithEvents btnjobshowall As System.Windows.Forms.Button
    Friend WithEvents btnjobdelivered As System.Windows.Forms.Button
    Friend WithEvents btnCompletedJobs As System.Windows.Forms.Button
    Friend WithEvents btnCurrentJobs As System.Windows.Forms.Button
    Friend WithEvents btngrossmargin As System.Windows.Forms.Button
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuallclients As System.Windows.Forms.MenuItem
    Friend WithEvents mnucurrentclients As System.Windows.Forms.MenuItem
    Friend WithEvents PrintPreviewDialog1 As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_ColourFooterLine As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_ColourBodyline As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_ColourHeaderLine As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_InterSectionSpacingPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_FooterHeightPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_HeaderHeightPercentage As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_PagesAcross As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox_ColumnHeaderBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_EvenBrush As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_OddRowBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_HeaderBrush As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_FooterBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
    Friend WithEvents mnucurrent As System.Windows.Forms.MenuItem
    Friend WithEvents mnucurrentviewofjobs As System.Windows.Forms.MenuItem
    Friend WithEvents mnucompletedjobs As System.Windows.Forms.MenuItem
    Friend WithEvents mnudeliveredjobs As System.Windows.Forms.MenuItem
    Friend WithEvents dgrid As System.Windows.Forms.DataGrid
    Friend WithEvents mnuleadssuspect As System.Windows.Forms.MenuItem
    Friend WithEvents mnuleadsprospect As System.Windows.Forms.MenuItem
    Friend WithEvents mnuleadsproposal As System.Windows.Forms.MenuItem
    Friend WithEvents mnuleadspho As System.Windows.Forms.MenuItem
    Friend WithEvents mnuleadsfailed As System.Windows.Forms.MenuItem
    Friend WithEvents mnuleadscurrent As System.Windows.Forms.MenuItem
    Friend WithEvents mnucur As System.Windows.Forms.MenuItem
    Friend WithEvents mnutss As System.Windows.Forms.MenuItem
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents tlbit As System.Windows.Forms.ToolBarButton
    Friend WithEvents msk As AxMSMask.AxMaskEdBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHome))
        Me.mnuMainMenu = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnufilesettings = New System.Windows.Forms.MenuItem
        Me.mnufilesettingsadministrator = New System.Windows.Forms.MenuItem
        Me.mnucur = New System.Windows.Forms.MenuItem
        Me.mnutss = New System.Windows.Forms.MenuItem
        Me.mnufileclose = New System.Windows.Forms.MenuItem
        Me.mnuEdit = New System.Windows.Forms.MenuItem
        Me.mnueditcontacts = New System.Windows.Forms.MenuItem
        Me.mnueditleads = New System.Windows.Forms.MenuItem
        Me.mnueditjobs = New System.Windows.Forms.MenuItem
        Me.mnuView = New System.Windows.Forms.MenuItem
        Me.mnuviewhome = New System.Windows.Forms.MenuItem
        Me.mnuviewcontacts = New System.Windows.Forms.MenuItem
        Me.mnuviewleads = New System.Windows.Forms.MenuItem
        Me.mnuviewjobs = New System.Windows.Forms.MenuItem
        Me.mnuviewequip = New System.Windows.Forms.MenuItem
        Me.mnuviewpersonnel = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuReports = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuleadssuspect = New System.Windows.Forms.MenuItem
        Me.mnuleadsprospect = New System.Windows.Forms.MenuItem
        Me.mnuleadsproposal = New System.Windows.Forms.MenuItem
        Me.mnuleadspho = New System.Windows.Forms.MenuItem
        Me.mnuleadsfailed = New System.Windows.Forms.MenuItem
        Me.mnuleadscurrent = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuallclients = New System.Windows.Forms.MenuItem
        Me.mnucurrentclients = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.mnucurrent = New System.Windows.Forms.MenuItem
        Me.mnucompletedjobs = New System.Windows.Forms.MenuItem
        Me.mnudeliveredjobs = New System.Windows.Forms.MenuItem
        Me.mnucurrentviewofjobs = New System.Windows.Forms.MenuItem
        Me.tbcHome = New System.Windows.Forms.TabControl
        Me.tpgHome = New System.Windows.Forms.TabPage
        Me.Label21 = New System.Windows.Forms.Label
        Me.ComboBox_FooterBrush = New System.Windows.Forms.ComboBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.ComboBox_ColourFooterLine = New System.Windows.Forms.ComboBox
        Me.ComboBox_ColourBodyline = New System.Windows.Forms.ComboBox
        Me.ComboBox_ColourHeaderLine = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.NumericUpDown_InterSectionSpacingPercent = New System.Windows.Forms.NumericUpDown
        Me.Label10 = New System.Windows.Forms.Label
        Me.NumericUpDown_FooterHeightPercent = New System.Windows.Forms.NumericUpDown
        Me.Label11 = New System.Windows.Forms.Label
        Me.NumericUpDown_HeaderHeightPercentage = New System.Windows.Forms.NumericUpDown
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.NumericUpDown_PagesAcross = New System.Windows.Forms.NumericUpDown
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ComboBox_ColumnHeaderBrush = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.ComboBox_EvenBrush = New System.Windows.Forms.ComboBox
        Me.ComboBox_OddRowBrush = New System.Windows.Forms.ComboBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.ComboBox_HeaderBrush = New System.Windows.Forms.ComboBox
        Me.dgrid = New System.Windows.Forms.DataGrid
        Me.tpgClients = New System.Windows.Forms.TabPage
        Me.pnlContacts = New System.Windows.Forms.Panel
        Me.grpcontactgrid = New System.Windows.Forms.GroupBox
        Me.btnshowall = New System.Windows.Forms.Button
        Me.btnaddnew = New System.Windows.Forms.Button
        Me.dtgClients = New System.Windows.Forms.DataGrid
        Me.grpContactsearch = New System.Windows.Forms.GroupBox
        Me.btnsearchname = New System.Windows.Forms.Button
        Me.cboclientssearchfield = New System.Windows.Forms.ComboBox
        Me.txtparams = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tpgLeads = New System.Windows.Forms.TabPage
        Me.pnlleads = New System.Windows.Forms.Panel
        Me.grpleadsgrid = New System.Windows.Forms.GroupBox
        Me.btnleadsproposal = New System.Windows.Forms.Button
        Me.btnleadsprospect = New System.Windows.Forms.Button
        Me.btnleadssuspect = New System.Windows.Forms.Button
        Me.btnShowAllleads = New System.Windows.Forms.Button
        Me.btnaddleads = New System.Windows.Forms.Button
        Me.btnfailedlead = New System.Windows.Forms.Button
        Me.btnleadspho = New System.Windows.Forms.Button
        Me.dtgLeads = New System.Windows.Forms.DataGrid
        Me.grpLeadssearch = New System.Windows.Forms.GroupBox
        Me.btnleadssearch = New System.Windows.Forms.Button
        Me.cboleadssearchfield = New System.Windows.Forms.ComboBox
        Me.txtleads = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboleads = New System.Windows.Forms.ComboBox
        Me.lblleadenddate = New System.Windows.Forms.Label
        Me.lblleadstartdate = New System.Windows.Forms.Label
        Me.lblleadname = New System.Windows.Forms.Label
        Me.lblleadstatus = New System.Windows.Forms.Label
        Me.dtpedate = New System.Windows.Forms.DateTimePicker
        Me.dtpsdate = New System.Windows.Forms.DateTimePicker
        Me.tpgJobs = New System.Windows.Forms.TabPage
        Me.pnljobs = New System.Windows.Forms.Panel
        Me.grpjobsearch = New System.Windows.Forms.GroupBox
        Me.btnjobsearch = New System.Windows.Forms.Button
        Me.cbojobsearchfield = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtjobcontactname = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cbojobstatus = New System.Windows.Forms.ComboBox
        Me.lbljobedate = New System.Windows.Forms.Label
        Me.lbljobsdate = New System.Windows.Forms.Label
        Me.dtpjobedate = New System.Windows.Forms.DateTimePicker
        Me.dtpjobsdate = New System.Windows.Forms.DateTimePicker
        Me.lbljobstatus = New System.Windows.Forms.Label
        Me.grpjobs = New System.Windows.Forms.GroupBox
        Me.btngrossmargin = New System.Windows.Forms.Button
        Me.btnCurrentJobs = New System.Windows.Forms.Button
        Me.dtgJobs = New System.Windows.Forms.DataGrid
        Me.btnCompletedJobs = New System.Windows.Forms.Button
        Me.btnjobdelivered = New System.Windows.Forms.Button
        Me.btnjobshowall = New System.Windows.Forms.Button
        Me.tpgEquip = New System.Windows.Forms.TabPage
        Me.pnlequip = New System.Windows.Forms.Panel
        Me.pnlequipcontrols = New System.Windows.Forms.Panel
        Me.btnequipsearch = New System.Windows.Forms.Button
        Me.lblsearchparameter = New System.Windows.Forms.Label
        Me.cboequipsearch = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.tpgPersonnel = New System.Windows.Forms.TabPage
        Me.pnlpersonnel = New System.Windows.Forms.Panel
        Me.ToolBar2 = New System.Windows.Forms.ToolBar
        Me.tlbtimesheet = New System.Windows.Forms.ToolBarButton
        Me.tlbadmin = New System.Windows.Forms.ToolBarButton
        Me.tlbit = New System.Windows.Forms.ToolBarButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.tbcHome.SuspendLayout()
        Me.tpgHome.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown_InterSectionSpacingPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_FooterHeightPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_HeaderHeightPercentage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.NumericUpDown_PagesAcross, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dgrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpgClients.SuspendLayout()
        Me.pnlContacts.SuspendLayout()
        Me.grpcontactgrid.SuspendLayout()
        CType(Me.dtgClients, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpContactsearch.SuspendLayout()
        Me.tpgLeads.SuspendLayout()
        Me.pnlleads.SuspendLayout()
        Me.grpleadsgrid.SuspendLayout()
        CType(Me.dtgLeads, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpLeadssearch.SuspendLayout()
        Me.tpgJobs.SuspendLayout()
        Me.pnljobs.SuspendLayout()
        Me.grpjobsearch.SuspendLayout()
        Me.grpjobs.SuspendLayout()
        CType(Me.dtgJobs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpgEquip.SuspendLayout()
        Me.pnlequipcontrols.SuspendLayout()
        Me.tpgPersonnel.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuMainMenu
        '
        Me.mnuMainMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.mnuView, Me.mnuReports})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnufilesettings, Me.mnufileclose})
        Me.mnuFile.Text = "File"
        '
        'mnufilesettings
        '
        Me.mnufilesettings.Index = 0
        Me.mnufilesettings.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnufilesettingsadministrator})
        Me.mnufilesettings.Text = "Settings"
        '
        'mnufilesettingsadministrator
        '
        Me.mnufilesettingsadministrator.Index = 0
        Me.mnufilesettingsadministrator.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnucur, Me.mnutss})
        Me.mnufilesettingsadministrator.Text = "Administrator"
        '
        'mnucur
        '
        Me.mnucur.Index = 0
        Me.mnucur.Text = "Configure user rights"
        '
        'mnutss
        '
        Me.mnutss.Index = 1
        Me.mnutss.Text = "Time sheet settings"
        Me.mnutss.Visible = False
        '
        'mnufileclose
        '
        Me.mnufileclose.Index = 1
        Me.mnufileclose.Text = "Close"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 1
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnueditcontacts, Me.mnueditleads, Me.mnueditjobs})
        Me.mnuEdit.Text = "Edit"
        '
        'mnueditcontacts
        '
        Me.mnueditcontacts.Index = 0
        Me.mnueditcontacts.Text = "Contacts"
        '
        'mnueditleads
        '
        Me.mnueditleads.Index = 1
        Me.mnueditleads.Text = "Leads"
        '
        'mnueditjobs
        '
        Me.mnueditjobs.Index = 2
        Me.mnueditjobs.Text = "Jobs"
        '
        'mnuView
        '
        Me.mnuView.Index = 2
        Me.mnuView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuviewhome, Me.mnuviewcontacts, Me.mnuviewleads, Me.mnuviewjobs, Me.mnuviewequip, Me.mnuviewpersonnel, Me.MenuItem1})
        Me.mnuView.Text = "View"
        '
        'mnuviewhome
        '
        Me.mnuviewhome.Index = 0
        Me.mnuviewhome.Text = "Home"
        '
        'mnuviewcontacts
        '
        Me.mnuviewcontacts.Index = 1
        Me.mnuviewcontacts.Text = "Clients"
        '
        'mnuviewleads
        '
        Me.mnuviewleads.Index = 2
        Me.mnuviewleads.Text = "Leads"
        '
        'mnuviewjobs
        '
        Me.mnuviewjobs.Index = 3
        Me.mnuviewjobs.Text = "Jobs"
        '
        'mnuviewequip
        '
        Me.mnuviewequip.Index = 4
        Me.mnuviewequip.Text = "Equipment"
        '
        'mnuviewpersonnel
        '
        Me.mnuviewpersonnel.Index = 5
        Me.mnuviewpersonnel.Text = "Personnel"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 6
        Me.MenuItem1.Text = "Reports"
        '
        'mnuReports
        '
        Me.mnuReports.Index = 3
        Me.mnuReports.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3, Me.MenuItem10})
        Me.mnuReports.Text = "Reports"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuleadssuspect, Me.mnuleadsprospect, Me.mnuleadsproposal, Me.mnuleadspho, Me.mnuleadsfailed, Me.mnuleadscurrent})
        Me.MenuItem2.Text = "Leads"
        '
        'mnuleadssuspect
        '
        Me.mnuleadssuspect.Index = 0
        Me.mnuleadssuspect.Text = "Suspect"
        '
        'mnuleadsprospect
        '
        Me.mnuleadsprospect.Index = 1
        Me.mnuleadsprospect.Text = "Prospect"
        '
        'mnuleadsproposal
        '
        Me.mnuleadsproposal.Index = 2
        Me.mnuleadsproposal.Text = "Proposal"
        '
        'mnuleadspho
        '
        Me.mnuleadspho.Index = 3
        Me.mnuleadspho.Text = "Pho"
        '
        'mnuleadsfailed
        '
        Me.mnuleadsfailed.Index = 4
        Me.mnuleadsfailed.Text = "Failed lead"
        '
        'mnuleadscurrent
        '
        Me.mnuleadscurrent.Index = 5
        Me.mnuleadscurrent.Text = "current view of leads"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuallclients, Me.mnucurrentclients})
        Me.MenuItem3.Text = "Clients"
        '
        'mnuallclients
        '
        Me.mnuallclients.Index = 0
        Me.mnuallclients.Text = "All clients"
        '
        'mnucurrentclients
        '
        Me.mnucurrentclients.Index = 1
        Me.mnucurrentclients.Text = "displayed currents only"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 2
        Me.MenuItem10.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnucurrent, Me.mnucompletedjobs, Me.mnudeliveredjobs, Me.mnucurrentviewofjobs})
        Me.MenuItem10.Text = "Jobs"
        '
        'mnucurrent
        '
        Me.mnucurrent.Index = 0
        Me.mnucurrent.Text = "Current"
        '
        'mnucompletedjobs
        '
        Me.mnucompletedjobs.Index = 1
        Me.mnucompletedjobs.Text = "Completed"
        '
        'mnudeliveredjobs
        '
        Me.mnudeliveredjobs.Index = 2
        Me.mnudeliveredjobs.Text = "Delivered"
        '
        'mnucurrentviewofjobs
        '
        Me.mnucurrentviewofjobs.Index = 3
        Me.mnucurrentviewofjobs.Text = "Current view of jobs"
        '
        'tbcHome
        '
        Me.tbcHome.Alignment = System.Windows.Forms.TabAlignment.Left
        Me.tbcHome.Controls.Add(Me.tpgHome)
        Me.tbcHome.Controls.Add(Me.tpgJobs)
        Me.tbcHome.Controls.Add(Me.tpgClients)
        Me.tbcHome.Controls.Add(Me.tpgLeads)
        Me.tbcHome.Controls.Add(Me.tpgEquip)
        Me.tbcHome.Controls.Add(Me.tpgPersonnel)
        Me.tbcHome.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbcHome.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.tbcHome.ItemSize = New System.Drawing.Size(19, 24)
        Me.tbcHome.Location = New System.Drawing.Point(0, 0)
        Me.tbcHome.Multiline = True
        Me.tbcHome.Name = "tbcHome"
        Me.tbcHome.SelectedIndex = 0
        Me.tbcHome.Size = New System.Drawing.Size(650, 615)
        Me.tbcHome.TabIndex = 0
        '
        'tpgHome
        '
        Me.tpgHome.BackgroundImage = CType(resources.GetObject("tpgHome.BackgroundImage"), System.Drawing.Image)
        Me.tpgHome.Controls.Add(Me.Label21)
        Me.tpgHome.Controls.Add(Me.ComboBox_FooterBrush)
        Me.tpgHome.Controls.Add(Me.GroupBox5)
        Me.tpgHome.Controls.Add(Me.Label16)
        Me.tpgHome.Controls.Add(Me.TextBox2)
        Me.tpgHome.Controls.Add(Me.ComboBox1)
        Me.tpgHome.Controls.Add(Me.GroupBox1)
        Me.tpgHome.Controls.Add(Me.GroupBox4)
        Me.tpgHome.Controls.Add(Me.GroupBox3)
        Me.tpgHome.Controls.Add(Me.dgrid)
        Me.tpgHome.Location = New System.Drawing.Point(28, 4)
        Me.tpgHome.Name = "tpgHome"
        Me.tpgHome.Size = New System.Drawing.Size(618, 607)
        Me.tpgHome.TabIndex = 5
        Me.tpgHome.Text = "Home"
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(64, 168)
        Me.Label21.Name = "Label21"
        Me.Label21.TabIndex = 43
        Me.Label21.Text = "Footer Brush"
        Me.Label21.Visible = False
        '
        'ComboBox_FooterBrush
        '
        Me.ComboBox_FooterBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_FooterBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_FooterBrush.Location = New System.Drawing.Point(168, 168)
        Me.ComboBox_FooterBrush.Name = "ComboBox_FooterBrush"
        Me.ComboBox_FooterBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_FooterBrush.TabIndex = 42
        Me.ComboBox_FooterBrush.Visible = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label15)
        Me.GroupBox5.Controls.Add(Me.Label7)
        Me.GroupBox5.Controls.Add(Me.Label8)
        Me.GroupBox5.Controls.Add(Me.ComboBox_ColourFooterLine)
        Me.GroupBox5.Controls.Add(Me.ComboBox_ColourBodyline)
        Me.GroupBox5.Controls.Add(Me.ComboBox_ColourHeaderLine)
        Me.GroupBox5.Controls.Add(Me.Label9)
        Me.GroupBox5.Location = New System.Drawing.Point(299, 287)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(272, 128)
        Me.GroupBox5.TabIndex = 39
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Grid line colours"
        Me.GroupBox5.Visible = False
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 16)
        Me.Label15.Name = "Label15"
        Me.Label15.TabIndex = 5
        Me.Label15.Text = "Header"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 23)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Footer line colour"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 40)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 23)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Header line colour"
        '
        'ComboBox_ColourFooterLine
        '
        Me.ComboBox_ColourFooterLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourFooterLine.Location = New System.Drawing.Point(128, 72)
        Me.ComboBox_ColourFooterLine.Name = "ComboBox_ColourFooterLine"
        Me.ComboBox_ColourFooterLine.Size = New System.Drawing.Size(121, 22)
        Me.ComboBox_ColourFooterLine.TabIndex = 9
        '
        'ComboBox_ColourBodyline
        '
        Me.ComboBox_ColourBodyline.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourBodyline.Location = New System.Drawing.Point(128, 96)
        Me.ComboBox_ColourBodyline.Name = "ComboBox_ColourBodyline"
        Me.ComboBox_ColourBodyline.Size = New System.Drawing.Size(121, 22)
        Me.ComboBox_ColourBodyline.TabIndex = 10
        '
        'ComboBox_ColourHeaderLine
        '
        Me.ComboBox_ColourHeaderLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourHeaderLine.Location = New System.Drawing.Point(128, 40)
        Me.ComboBox_ColourHeaderLine.Name = "ComboBox_ColourHeaderLine"
        Me.ComboBox_ColourHeaderLine.Size = New System.Drawing.Size(121, 22)
        Me.ComboBox_ColourHeaderLine.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 96)
        Me.Label9.Name = "Label9"
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "Body line colour"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(59, 215)
        Me.Label16.Name = "Label16"
        Me.Label16.TabIndex = 38
        Me.Label16.Text = "Page Heading"
        Me.Label16.Visible = False
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(163, 215)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(128, 22)
        Me.TextBox2.TabIndex = 37
        Me.TextBox2.Text = "Data Grid Print Test"
        Me.TextBox2.Visible = False
        '
        'ComboBox1
        '
        Me.ComboBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Location = New System.Drawing.Point(163, 191)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(127, 21)
        Me.ComboBox1.TabIndex = 41
        Me.ComboBox1.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown_InterSectionSpacingPercent)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown_FooterHeightPercent)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown_HeaderHeightPercentage)
        Me.GroupBox1.Location = New System.Drawing.Point(299, 191)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 96)
        Me.GroupBox1.TabIndex = 36
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Section Heights"
        Me.GroupBox1.Visible = False
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 16)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(104, 23)
        Me.Label14.TabIndex = 11
        Me.Label14.Text = "Header height"
        '
        'NumericUpDown_InterSectionSpacingPercent
        '
        Me.NumericUpDown_InterSectionSpacingPercent.Location = New System.Drawing.Point(136, 64)
        Me.NumericUpDown_InterSectionSpacingPercent.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NumericUpDown_InterSectionSpacingPercent.Name = "NumericUpDown_InterSectionSpacingPercent"
        Me.NumericUpDown_InterSectionSpacingPercent.TabIndex = 5
        Me.NumericUpDown_InterSectionSpacingPercent.Value = New Decimal(New Integer() {3, 0, 0, 0})
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 64)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(120, 23)
        Me.Label10.TabIndex = 4
        Me.Label10.Text = "Inter-section spacing"
        '
        'NumericUpDown_FooterHeightPercent
        '
        Me.NumericUpDown_FooterHeightPercent.Location = New System.Drawing.Point(136, 40)
        Me.NumericUpDown_FooterHeightPercent.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.NumericUpDown_FooterHeightPercent.Name = "NumericUpDown_FooterHeightPercent"
        Me.NumericUpDown_FooterHeightPercent.TabIndex = 3
        Me.NumericUpDown_FooterHeightPercent.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 40)
        Me.Label11.Name = "Label11"
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "Footer height"
        '
        'NumericUpDown_HeaderHeightPercentage
        '
        Me.NumericUpDown_HeaderHeightPercentage.Location = New System.Drawing.Point(136, 16)
        Me.NumericUpDown_HeaderHeightPercentage.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.NumericUpDown_HeaderHeightPercentage.Name = "NumericUpDown_HeaderHeightPercentage"
        Me.NumericUpDown_HeaderHeightPercentage.TabIndex = 1
        Me.NumericUpDown_HeaderHeightPercentage.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label13)
        Me.GroupBox4.Controls.Add(Me.NumericUpDown_PagesAcross)
        Me.GroupBox4.Location = New System.Drawing.Point(51, 359)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(240, 56)
        Me.GroupBox4.TabIndex = 35
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Page layout"
        Me.GroupBox4.Visible = False
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 16)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(192, 32)
        Me.Label13.TabIndex = 2
        Me.Label13.Text = "Minimum number of pages across to split the columns over"
        '
        'NumericUpDown_PagesAcross
        '
        Me.NumericUpDown_PagesAcross.Location = New System.Drawing.Point(206, 16)
        Me.NumericUpDown_PagesAcross.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.NumericUpDown_PagesAcross.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown_PagesAcross.Name = "NumericUpDown_PagesAcross"
        Me.NumericUpDown_PagesAcross.Size = New System.Drawing.Size(32, 20)
        Me.NumericUpDown_PagesAcross.TabIndex = 3
        Me.NumericUpDown_PagesAcross.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ComboBox_ColumnHeaderBrush)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.ComboBox_EvenBrush)
        Me.GroupBox3.Controls.Add(Me.ComboBox_OddRowBrush)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Controls.Add(Me.Label19)
        Me.GroupBox3.Controls.Add(Me.Label20)
        Me.GroupBox3.Controls.Add(Me.ComboBox_HeaderBrush)
        Me.GroupBox3.Location = New System.Drawing.Point(51, 239)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(240, 120)
        Me.GroupBox3.TabIndex = 34
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Grid background  colours"
        Me.GroupBox3.Visible = False
        '
        'ComboBox_ColumnHeaderBrush
        '
        Me.ComboBox_ColumnHeaderBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_ColumnHeaderBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColumnHeaderBrush.Location = New System.Drawing.Point(112, 88)
        Me.ComboBox_ColumnHeaderBrush.Name = "ComboBox_ColumnHeaderBrush"
        Me.ComboBox_ColumnHeaderBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_ColumnHeaderBrush.TabIndex = 15
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 88)
        Me.Label12.Name = "Label12"
        Me.Label12.TabIndex = 14
        Me.Label12.Text = "Columns"
        '
        'ComboBox_EvenBrush
        '
        Me.ComboBox_EvenBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_EvenBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_EvenBrush.Location = New System.Drawing.Point(112, 64)
        Me.ComboBox_EvenBrush.Name = "ComboBox_EvenBrush"
        Me.ComboBox_EvenBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_EvenBrush.TabIndex = 13
        '
        'ComboBox_OddRowBrush
        '
        Me.ComboBox_OddRowBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_OddRowBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_OddRowBrush.Location = New System.Drawing.Point(112, 40)
        Me.ComboBox_OddRowBrush.Name = "ComboBox_OddRowBrush"
        Me.ComboBox_OddRowBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_OddRowBrush.TabIndex = 12
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(8, 64)
        Me.Label18.Name = "Label18"
        Me.Label18.TabIndex = 11
        Me.Label18.Text = "Even rows"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(8, 32)
        Me.Label19.Name = "Label19"
        Me.Label19.TabIndex = 5
        Me.Label19.Text = "Header"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(8, 16)
        Me.Label20.Name = "Label20"
        Me.Label20.TabIndex = 10
        Me.Label20.Text = "Odd rows"
        '
        'ComboBox_HeaderBrush
        '
        Me.ComboBox_HeaderBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_HeaderBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_HeaderBrush.Location = New System.Drawing.Point(112, 16)
        Me.ComboBox_HeaderBrush.Name = "ComboBox_HeaderBrush"
        Me.ComboBox_HeaderBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_HeaderBrush.TabIndex = 8
        '
        'dgrid
        '
        Me.dgrid.DataMember = ""
        Me.dgrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgrid.Location = New System.Drawing.Point(440, 104)
        Me.dgrid.Name = "dgrid"
        Me.dgrid.TabIndex = 1
        Me.dgrid.Visible = False
        '
        'tpgClients
        '
        Me.tpgClients.BackgroundImage = CType(resources.GetObject("tpgClients.BackgroundImage"), System.Drawing.Image)
        Me.tpgClients.Controls.Add(Me.pnlContacts)
        Me.tpgClients.Location = New System.Drawing.Point(28, 4)
        Me.tpgClients.Name = "tpgClients"
        Me.tpgClients.Size = New System.Drawing.Size(618, 607)
        Me.tpgClients.TabIndex = 1
        Me.tpgClients.Text = "Clients"
        '
        'pnlContacts
        '
        Me.pnlContacts.Controls.Add(Me.grpcontactgrid)
        Me.pnlContacts.Controls.Add(Me.grpContactsearch)
        Me.pnlContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContacts.Location = New System.Drawing.Point(0, 0)
        Me.pnlContacts.Name = "pnlContacts"
        Me.pnlContacts.Size = New System.Drawing.Size(618, 607)
        Me.pnlContacts.TabIndex = 1
        '
        'grpcontactgrid
        '
        Me.grpcontactgrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpcontactgrid.Controls.Add(Me.btnshowall)
        Me.grpcontactgrid.Controls.Add(Me.btnaddnew)
        Me.grpcontactgrid.Controls.Add(Me.dtgClients)
        Me.grpcontactgrid.Location = New System.Drawing.Point(8, 80)
        Me.grpcontactgrid.Name = "grpcontactgrid"
        Me.grpcontactgrid.Size = New System.Drawing.Size(618, 512)
        Me.grpcontactgrid.TabIndex = 3
        Me.grpcontactgrid.TabStop = False
        '
        'btnshowall
        '
        Me.btnshowall.Location = New System.Drawing.Point(120, 13)
        Me.btnshowall.Name = "btnshowall"
        Me.btnshowall.Size = New System.Drawing.Size(72, 24)
        Me.btnshowall.TabIndex = 16
        Me.btnshowall.Text = "Show all"
        '
        'btnaddnew
        '
        Me.btnaddnew.Location = New System.Drawing.Point(7, 13)
        Me.btnaddnew.Name = "btnaddnew"
        Me.btnaddnew.Size = New System.Drawing.Size(112, 24)
        Me.btnaddnew.TabIndex = 15
        Me.btnaddnew.Text = "Add new contact"
        '
        'dtgClients
        '
        Me.dtgClients.AllowSorting = False
        Me.dtgClients.AlternatingBackColor = System.Drawing.SystemColors.WindowFrame
        Me.dtgClients.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgClients.CaptionText = "Clients"
        Me.dtgClients.DataMember = ""
        Me.dtgClients.GridLineColor = System.Drawing.Color.Transparent
        Me.dtgClients.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgClients.Location = New System.Drawing.Point(8, 40)
        Me.dtgClients.Name = "dtgClients"
        Me.dtgClients.ParentRowsBackColor = System.Drawing.Color.Thistle
        Me.dtgClients.ReadOnly = True
        Me.dtgClients.Size = New System.Drawing.Size(595, 464)
        Me.dtgClients.TabIndex = 5
        '
        'grpContactsearch
        '
        Me.grpContactsearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpContactsearch.Controls.Add(Me.btnsearchname)
        Me.grpContactsearch.Controls.Add(Me.cboclientssearchfield)
        Me.grpContactsearch.Controls.Add(Me.txtparams)
        Me.grpContactsearch.Controls.Add(Me.Label3)
        Me.grpContactsearch.Controls.Add(Me.Label4)
        Me.grpContactsearch.Location = New System.Drawing.Point(8, 0)
        Me.grpContactsearch.Name = "grpContactsearch"
        Me.grpContactsearch.Size = New System.Drawing.Size(603, 80)
        Me.grpContactsearch.TabIndex = 0
        Me.grpContactsearch.TabStop = False
        '
        'btnsearchname
        '
        Me.btnsearchname.Location = New System.Drawing.Point(256, 50)
        Me.btnsearchname.Name = "btnsearchname"
        Me.btnsearchname.Size = New System.Drawing.Size(88, 23)
        Me.btnsearchname.TabIndex = 23
        Me.btnsearchname.Text = "Search"
        '
        'cboclientssearchfield
        '
        Me.cboclientssearchfield.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboclientssearchfield.Items.AddRange(New Object() {"Client number", "Name", "Description", "Old number"})
        Me.cboclientssearchfield.Location = New System.Drawing.Point(144, 27)
        Me.cboclientssearchfield.Name = "cboclientssearchfield"
        Me.cboclientssearchfield.Size = New System.Drawing.Size(152, 22)
        Me.cboclientssearchfield.TabIndex = 22
        '
        'txtparams
        '
        Me.txtparams.Location = New System.Drawing.Point(305, 27)
        Me.txtparams.Name = "txtparams"
        Me.txtparams.Size = New System.Drawing.Size(183, 20)
        Me.txtparams.TabIndex = 21
        Me.txtparams.Text = ""
        '
        'Label3
        '
        Me.Label3.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label3.Location = New System.Drawing.Point(144, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 8)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Choose field"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label4.Location = New System.Drawing.Point(313, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(175, 8)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Type here"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tpgLeads
        '
        Me.tpgLeads.BackColor = System.Drawing.SystemColors.Control
        Me.tpgLeads.BackgroundImage = CType(resources.GetObject("tpgLeads.BackgroundImage"), System.Drawing.Image)
        Me.tpgLeads.Controls.Add(Me.pnlleads)
        Me.tpgLeads.Location = New System.Drawing.Point(28, 4)
        Me.tpgLeads.Name = "tpgLeads"
        Me.tpgLeads.Size = New System.Drawing.Size(618, 607)
        Me.tpgLeads.TabIndex = 0
        Me.tpgLeads.Text = "Leads     "
        '
        'pnlleads
        '
        Me.pnlleads.AutoScroll = True
        Me.pnlleads.BackColor = System.Drawing.SystemColors.Control
        Me.pnlleads.Controls.Add(Me.grpleadsgrid)
        Me.pnlleads.Controls.Add(Me.grpLeadssearch)
        Me.pnlleads.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlleads.ForeColor = System.Drawing.SystemColors.Control
        Me.pnlleads.Location = New System.Drawing.Point(0, 0)
        Me.pnlleads.Name = "pnlleads"
        Me.pnlleads.Size = New System.Drawing.Size(618, 607)
        Me.pnlleads.TabIndex = 5
        '
        'grpleadsgrid
        '
        Me.grpleadsgrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpleadsgrid.Controls.Add(Me.btnleadsproposal)
        Me.grpleadsgrid.Controls.Add(Me.btnleadsprospect)
        Me.grpleadsgrid.Controls.Add(Me.btnleadssuspect)
        Me.grpleadsgrid.Controls.Add(Me.btnShowAllleads)
        Me.grpleadsgrid.Controls.Add(Me.btnaddleads)
        Me.grpleadsgrid.Controls.Add(Me.btnfailedlead)
        Me.grpleadsgrid.Controls.Add(Me.btnleadspho)
        Me.grpleadsgrid.Controls.Add(Me.dtgLeads)
        Me.grpleadsgrid.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpleadsgrid.Location = New System.Drawing.Point(0, 88)
        Me.grpleadsgrid.Name = "grpleadsgrid"
        Me.grpleadsgrid.Size = New System.Drawing.Size(619, 511)
        Me.grpleadsgrid.TabIndex = 2
        Me.grpleadsgrid.TabStop = False
        Me.grpleadsgrid.Text = "Leads"
        '
        'btnleadsproposal
        '
        Me.btnleadsproposal.Location = New System.Drawing.Point(185, 42)
        Me.btnleadsproposal.Name = "btnleadsproposal"
        Me.btnleadsproposal.Size = New System.Drawing.Size(88, 23)
        Me.btnleadsproposal.TabIndex = 29
        Me.btnleadsproposal.Text = "Proposal"
        '
        'btnleadsprospect
        '
        Me.btnleadsprospect.Location = New System.Drawing.Point(96, 42)
        Me.btnleadsprospect.Name = "btnleadsprospect"
        Me.btnleadsprospect.Size = New System.Drawing.Size(88, 23)
        Me.btnleadsprospect.TabIndex = 28
        Me.btnleadsprospect.Text = "Prospect"
        '
        'btnleadssuspect
        '
        Me.btnleadssuspect.Location = New System.Drawing.Point(8, 42)
        Me.btnleadssuspect.Name = "btnleadssuspect"
        Me.btnleadssuspect.Size = New System.Drawing.Size(88, 23)
        Me.btnleadssuspect.TabIndex = 27
        Me.btnleadssuspect.Text = "Suspect"
        '
        'btnShowAllleads
        '
        Me.btnShowAllleads.Location = New System.Drawing.Point(97, 16)
        Me.btnShowAllleads.Name = "btnShowAllleads"
        Me.btnShowAllleads.Size = New System.Drawing.Size(88, 23)
        Me.btnShowAllleads.TabIndex = 26
        Me.btnShowAllleads.Text = "Show all"
        '
        'btnaddleads
        '
        Me.btnaddleads.Location = New System.Drawing.Point(8, 16)
        Me.btnaddleads.Name = "btnaddleads"
        Me.btnaddleads.Size = New System.Drawing.Size(88, 23)
        Me.btnaddleads.TabIndex = 25
        Me.btnaddleads.Text = "Add new lead"
        '
        'btnfailedlead
        '
        Me.btnfailedlead.Location = New System.Drawing.Point(366, 43)
        Me.btnfailedlead.Name = "btnfailedlead"
        Me.btnfailedlead.Size = New System.Drawing.Size(88, 23)
        Me.btnfailedlead.TabIndex = 24
        Me.btnfailedlead.Text = "Failed lead"
        '
        'btnleadspho
        '
        Me.btnleadspho.Location = New System.Drawing.Point(275, 43)
        Me.btnleadspho.Name = "btnleadspho"
        Me.btnleadspho.Size = New System.Drawing.Size(88, 23)
        Me.btnleadspho.TabIndex = 23
        Me.btnleadspho.Text = "Pho"
        '
        'dtgLeads
        '
        Me.dtgLeads.AllowSorting = False
        Me.dtgLeads.AlternatingBackColor = System.Drawing.SystemColors.Control
        Me.dtgLeads.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgLeads.BackColor = System.Drawing.SystemColors.Control
        Me.dtgLeads.CaptionForeColor = System.Drawing.SystemColors.Control
        Me.dtgLeads.CaptionText = "Leads"
        Me.dtgLeads.DataMember = ""
        Me.dtgLeads.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgLeads.Location = New System.Drawing.Point(8, 72)
        Me.dtgLeads.Name = "dtgLeads"
        Me.dtgLeads.PreferredRowHeight = 20
        Me.dtgLeads.ReadOnly = True
        Me.dtgLeads.Size = New System.Drawing.Size(603, 424)
        Me.dtgLeads.TabIndex = 1
        '
        'grpLeadssearch
        '
        Me.grpLeadssearch.Controls.Add(Me.btnleadssearch)
        Me.grpLeadssearch.Controls.Add(Me.cboleadssearchfield)
        Me.grpLeadssearch.Controls.Add(Me.txtleads)
        Me.grpLeadssearch.Controls.Add(Me.Label2)
        Me.grpLeadssearch.Controls.Add(Me.cboleads)
        Me.grpLeadssearch.Controls.Add(Me.lblleadenddate)
        Me.grpLeadssearch.Controls.Add(Me.lblleadstartdate)
        Me.grpLeadssearch.Controls.Add(Me.lblleadname)
        Me.grpLeadssearch.Controls.Add(Me.lblleadstatus)
        Me.grpLeadssearch.Controls.Add(Me.dtpedate)
        Me.grpLeadssearch.Controls.Add(Me.dtpsdate)
        Me.grpLeadssearch.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpLeadssearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpLeadssearch.Location = New System.Drawing.Point(0, 0)
        Me.grpLeadssearch.Name = "grpLeadssearch"
        Me.grpLeadssearch.Size = New System.Drawing.Size(618, 88)
        Me.grpLeadssearch.TabIndex = 1
        Me.grpLeadssearch.TabStop = False
        Me.grpLeadssearch.Text = "Define Search Parameters"
        '
        'btnleadssearch
        '
        Me.btnleadssearch.Location = New System.Drawing.Point(240, 56)
        Me.btnleadssearch.Name = "btnleadssearch"
        Me.btnleadssearch.Size = New System.Drawing.Size(88, 23)
        Me.btnleadssearch.TabIndex = 18
        Me.btnleadssearch.Text = "Search"
        '
        'cboleadssearchfield
        '
        Me.cboleadssearchfield.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboleadssearchfield.Items.AddRange(New Object() {"Client number", "Lead number", "Name", "Title", "Status"})
        Me.cboleadssearchfield.Location = New System.Drawing.Point(134, 32)
        Me.cboleadssearchfield.Name = "cboleadssearchfield"
        Me.cboleadssearchfield.Size = New System.Drawing.Size(120, 22)
        Me.cboleadssearchfield.TabIndex = 17
        '
        'txtleads
        '
        Me.txtleads.Location = New System.Drawing.Point(256, 32)
        Me.txtleads.Name = "txtleads"
        Me.txtleads.Size = New System.Drawing.Size(120, 20)
        Me.txtleads.TabIndex = 16
        Me.txtleads.Text = ""
        '
        'Label2
        '
        Me.Label2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label2.Location = New System.Drawing.Point(139, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 8)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Choose field"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cboleads
        '
        Me.cboleads.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboleads.Location = New System.Drawing.Point(8, 31)
        Me.cboleads.Name = "cboleads"
        Me.cboleads.Size = New System.Drawing.Size(120, 22)
        Me.cboleads.TabIndex = 13
        '
        'lblleadenddate
        '
        Me.lblleadenddate.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblleadenddate.Location = New System.Drawing.Point(498, 16)
        Me.lblleadenddate.Name = "lblleadenddate"
        Me.lblleadenddate.Size = New System.Drawing.Size(100, 8)
        Me.lblleadenddate.TabIndex = 10
        Me.lblleadenddate.Text = "End Date"
        Me.lblleadenddate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblleadstartdate
        '
        Me.lblleadstartdate.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblleadstartdate.Location = New System.Drawing.Point(378, 16)
        Me.lblleadstartdate.Name = "lblleadstartdate"
        Me.lblleadstartdate.Size = New System.Drawing.Size(100, 8)
        Me.lblleadstartdate.TabIndex = 9
        Me.lblleadstartdate.Text = "Start Date"
        Me.lblleadstartdate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblleadname
        '
        Me.lblleadname.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblleadname.Location = New System.Drawing.Point(264, 16)
        Me.lblleadname.Name = "lblleadname"
        Me.lblleadname.Size = New System.Drawing.Size(100, 8)
        Me.lblleadname.TabIndex = 8
        Me.lblleadname.Text = "Type here"
        Me.lblleadname.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblleadstatus
        '
        Me.lblleadstatus.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblleadstatus.Location = New System.Drawing.Point(8, 16)
        Me.lblleadstatus.Name = "lblleadstatus"
        Me.lblleadstatus.Size = New System.Drawing.Size(100, 8)
        Me.lblleadstatus.TabIndex = 7
        Me.lblleadstatus.Text = "Status"
        Me.lblleadstatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtpedate
        '
        Me.dtpedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpedate.Location = New System.Drawing.Point(498, 32)
        Me.dtpedate.Name = "dtpedate"
        Me.dtpedate.Size = New System.Drawing.Size(120, 20)
        Me.dtpedate.TabIndex = 5
        Me.dtpedate.Value = New Date(2006, 1, 12, 8, 13, 39, 14)
        '
        'dtpsdate
        '
        Me.dtpsdate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpsdate.Location = New System.Drawing.Point(378, 32)
        Me.dtpsdate.Name = "dtpsdate"
        Me.dtpsdate.Size = New System.Drawing.Size(120, 20)
        Me.dtpsdate.TabIndex = 4
        Me.dtpsdate.Value = New Date(2006, 1, 12, 8, 13, 39, 14)
        '
        'tpgJobs
        '
        Me.tpgJobs.BackgroundImage = CType(resources.GetObject("tpgJobs.BackgroundImage"), System.Drawing.Image)
        Me.tpgJobs.Controls.Add(Me.pnljobs)
        Me.tpgJobs.Location = New System.Drawing.Point(28, 4)
        Me.tpgJobs.Name = "tpgJobs"
        Me.tpgJobs.Size = New System.Drawing.Size(618, 607)
        Me.tpgJobs.TabIndex = 2
        Me.tpgJobs.Text = "Jobs      "
        '
        'pnljobs
        '
        Me.pnljobs.Controls.Add(Me.grpjobsearch)
        Me.pnljobs.Controls.Add(Me.grpjobs)
        Me.pnljobs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnljobs.Location = New System.Drawing.Point(0, 0)
        Me.pnljobs.Name = "pnljobs"
        Me.pnljobs.Size = New System.Drawing.Size(618, 607)
        Me.pnljobs.TabIndex = 0
        '
        'grpjobsearch
        '
        Me.grpjobsearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpjobsearch.Controls.Add(Me.btnjobsearch)
        Me.grpjobsearch.Controls.Add(Me.cbojobsearchfield)
        Me.grpjobsearch.Controls.Add(Me.Label6)
        Me.grpjobsearch.Controls.Add(Me.txtjobcontactname)
        Me.grpjobsearch.Controls.Add(Me.Label5)
        Me.grpjobsearch.Controls.Add(Me.cbojobstatus)
        Me.grpjobsearch.Controls.Add(Me.lbljobedate)
        Me.grpjobsearch.Controls.Add(Me.lbljobsdate)
        Me.grpjobsearch.Controls.Add(Me.dtpjobedate)
        Me.grpjobsearch.Controls.Add(Me.dtpjobsdate)
        Me.grpjobsearch.Controls.Add(Me.lbljobstatus)
        Me.grpjobsearch.Location = New System.Drawing.Point(0, 0)
        Me.grpjobsearch.Name = "grpjobsearch"
        Me.grpjobsearch.Size = New System.Drawing.Size(618, 88)
        Me.grpjobsearch.TabIndex = 4
        Me.grpjobsearch.TabStop = False
        Me.grpjobsearch.Text = "Search Parameters"
        '
        'btnjobsearch
        '
        Me.btnjobsearch.Location = New System.Drawing.Point(256, 56)
        Me.btnjobsearch.Name = "btnjobsearch"
        Me.btnjobsearch.Size = New System.Drawing.Size(88, 23)
        Me.btnjobsearch.TabIndex = 24
        Me.btnjobsearch.Text = "Search"
        '
        'cbojobsearchfield
        '
        Me.cbojobsearchfield.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbojobsearchfield.Items.AddRange(New Object() {"Client number", "Name", "Job number", "Job title", "Job status", "Technician responsible", "Amount", "Gross margin"})
        Me.cbojobsearchfield.Location = New System.Drawing.Point(128, 32)
        Me.cbojobsearchfield.Name = "cbojobsearchfield"
        Me.cbojobsearchfield.Size = New System.Drawing.Size(120, 22)
        Me.cbojobsearchfield.TabIndex = 23
        '
        'Label6
        '
        Me.Label6.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label6.Location = New System.Drawing.Point(136, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 8)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Choose field"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtjobcontactname
        '
        Me.txtjobcontactname.Location = New System.Drawing.Point(251, 32)
        Me.txtjobcontactname.Name = "txtjobcontactname"
        Me.txtjobcontactname.Size = New System.Drawing.Size(120, 20)
        Me.txtjobcontactname.TabIndex = 21
        Me.txtjobcontactname.Text = ""
        '
        'Label5
        '
        Me.Label5.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label5.Location = New System.Drawing.Point(259, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 8)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Type here"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbojobstatus
        '
        Me.cbojobstatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbojobstatus.Location = New System.Drawing.Point(6, 32)
        Me.cbojobstatus.Name = "cbojobstatus"
        Me.cbojobstatus.Size = New System.Drawing.Size(120, 22)
        Me.cbojobstatus.TabIndex = 19
        '
        'lbljobedate
        '
        Me.lbljobedate.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lbljobedate.Location = New System.Drawing.Point(496, 16)
        Me.lbljobedate.Name = "lbljobedate"
        Me.lbljobedate.Size = New System.Drawing.Size(111, 8)
        Me.lbljobedate.TabIndex = 8
        Me.lbljobedate.Text = "End date"
        Me.lbljobedate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbljobsdate
        '
        Me.lbljobsdate.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lbljobsdate.Location = New System.Drawing.Point(374, 16)
        Me.lbljobsdate.Name = "lbljobsdate"
        Me.lbljobsdate.Size = New System.Drawing.Size(107, 8)
        Me.lbljobsdate.TabIndex = 7
        Me.lbljobsdate.Text = "Start date"
        Me.lbljobsdate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtpjobedate
        '
        Me.dtpjobedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpjobedate.Location = New System.Drawing.Point(496, 32)
        Me.dtpjobedate.Name = "dtpjobedate"
        Me.dtpjobedate.Size = New System.Drawing.Size(120, 20)
        Me.dtpjobedate.TabIndex = 6
        '
        'dtpjobsdate
        '
        Me.dtpjobsdate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpjobsdate.Location = New System.Drawing.Point(374, 32)
        Me.dtpjobsdate.Name = "dtpjobsdate"
        Me.dtpjobsdate.Size = New System.Drawing.Size(120, 20)
        Me.dtpjobsdate.TabIndex = 5
        '
        'lbljobstatus
        '
        Me.lbljobstatus.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lbljobstatus.Location = New System.Drawing.Point(6, 16)
        Me.lbljobstatus.Name = "lbljobstatus"
        Me.lbljobstatus.Size = New System.Drawing.Size(100, 8)
        Me.lbljobstatus.TabIndex = 0
        Me.lbljobstatus.Text = "Job Status"
        Me.lbljobstatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'grpjobs
        '
        Me.grpjobs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpjobs.Controls.Add(Me.btngrossmargin)
        Me.grpjobs.Controls.Add(Me.btnCurrentJobs)
        Me.grpjobs.Controls.Add(Me.dtgJobs)
        Me.grpjobs.Controls.Add(Me.btnCompletedJobs)
        Me.grpjobs.Controls.Add(Me.btnjobdelivered)
        Me.grpjobs.Controls.Add(Me.btnjobshowall)
        Me.grpjobs.Location = New System.Drawing.Point(0, 88)
        Me.grpjobs.Name = "grpjobs"
        Me.grpjobs.Size = New System.Drawing.Size(611, 511)
        Me.grpjobs.TabIndex = 3
        Me.grpjobs.TabStop = False
        '
        'btngrossmargin
        '
        Me.btngrossmargin.Location = New System.Drawing.Point(315, 11)
        Me.btngrossmargin.Name = "btngrossmargin"
        Me.btngrossmargin.Size = New System.Drawing.Size(136, 23)
        Me.btngrossmargin.TabIndex = 28
        Me.btngrossmargin.Text = "Refresh gross margin"
        '
        'btnCurrentJobs
        '
        Me.btnCurrentJobs.Location = New System.Drawing.Point(8, 11)
        Me.btnCurrentJobs.Name = "btnCurrentJobs"
        Me.btnCurrentJobs.TabIndex = 5
        Me.btnCurrentJobs.Text = "Current"
        '
        'dtgJobs
        '
        Me.dtgJobs.AllowSorting = False
        Me.dtgJobs.AlternatingBackColor = System.Drawing.SystemColors.WindowFrame
        Me.dtgJobs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgJobs.CaptionText = "Jobs"
        Me.dtgJobs.DataMember = ""
        Me.dtgJobs.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgJobs.Location = New System.Drawing.Point(8, 40)
        Me.dtgJobs.Name = "dtgJobs"
        Me.dtgJobs.ReadOnly = True
        Me.dtgJobs.Size = New System.Drawing.Size(597, 456)
        Me.dtgJobs.TabIndex = 0
        '
        'btnCompletedJobs
        '
        Me.btnCompletedJobs.Location = New System.Drawing.Point(84, 11)
        Me.btnCompletedJobs.Name = "btnCompletedJobs"
        Me.btnCompletedJobs.TabIndex = 27
        Me.btnCompletedJobs.Text = "Completed"
        '
        'btnjobdelivered
        '
        Me.btnjobdelivered.Location = New System.Drawing.Point(159, 11)
        Me.btnjobdelivered.Name = "btnjobdelivered"
        Me.btnjobdelivered.TabIndex = 26
        Me.btnjobdelivered.Text = "Delivered"
        '
        'btnjobshowall
        '
        Me.btnjobshowall.Location = New System.Drawing.Point(237, 11)
        Me.btnjobshowall.Name = "btnjobshowall"
        Me.btnjobshowall.TabIndex = 25
        Me.btnjobshowall.Text = "Show all"
        '
        'tpgEquip
        '
        Me.tpgEquip.BackgroundImage = CType(resources.GetObject("tpgEquip.BackgroundImage"), System.Drawing.Image)
        Me.tpgEquip.Controls.Add(Me.pnlequip)
        Me.tpgEquip.Controls.Add(Me.pnlequipcontrols)
        Me.tpgEquip.Location = New System.Drawing.Point(28, 4)
        Me.tpgEquip.Name = "tpgEquip"
        Me.tpgEquip.Size = New System.Drawing.Size(618, 607)
        Me.tpgEquip.TabIndex = 3
        Me.tpgEquip.Text = "Equipment"
        '
        'pnlequip
        '
        Me.pnlequip.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlequip.Location = New System.Drawing.Point(0, 80)
        Me.pnlequip.Name = "pnlequip"
        Me.pnlequip.Size = New System.Drawing.Size(618, 527)
        Me.pnlequip.TabIndex = 1
        '
        'pnlequipcontrols
        '
        Me.pnlequipcontrols.Controls.Add(Me.btnequipsearch)
        Me.pnlequipcontrols.Controls.Add(Me.lblsearchparameter)
        Me.pnlequipcontrols.Controls.Add(Me.cboequipsearch)
        Me.pnlequipcontrols.Controls.Add(Me.Label1)
        Me.pnlequipcontrols.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlequipcontrols.Location = New System.Drawing.Point(0, 0)
        Me.pnlequipcontrols.Name = "pnlequipcontrols"
        Me.pnlequipcontrols.Size = New System.Drawing.Size(618, 80)
        Me.pnlequipcontrols.TabIndex = 2
        '
        'btnequipsearch
        '
        Me.btnequipsearch.Location = New System.Drawing.Point(266, 55)
        Me.btnequipsearch.Name = "btnequipsearch"
        Me.btnequipsearch.Size = New System.Drawing.Size(96, 23)
        Me.btnequipsearch.TabIndex = 9
        Me.btnequipsearch.Text = "Search"
        '
        'lblsearchparameter
        '
        Me.lblsearchparameter.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblsearchparameter.Location = New System.Drawing.Point(319, 10)
        Me.lblsearchparameter.Name = "lblsearchparameter"
        Me.lblsearchparameter.Size = New System.Drawing.Size(206, 16)
        Me.lblsearchparameter.TabIndex = 8
        Me.lblsearchparameter.Text = "Type search parameter"
        Me.lblsearchparameter.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cboequipsearch
        '
        Me.cboequipsearch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboequipsearch.ItemHeight = 15
        Me.cboequipsearch.Location = New System.Drawing.Point(120, 31)
        Me.cboequipsearch.Name = "cboequipsearch"
        Me.cboequipsearch.Size = New System.Drawing.Size(192, 22)
        Me.cboequipsearch.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label1.Location = New System.Drawing.Point(128, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(184, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Choose search field"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tpgPersonnel
        '
        Me.tpgPersonnel.BackgroundImage = CType(resources.GetObject("tpgPersonnel.BackgroundImage"), System.Drawing.Image)
        Me.tpgPersonnel.Controls.Add(Me.pnlpersonnel)
        Me.tpgPersonnel.Controls.Add(Me.ToolBar2)
        Me.tpgPersonnel.Location = New System.Drawing.Point(28, 4)
        Me.tpgPersonnel.Name = "tpgPersonnel"
        Me.tpgPersonnel.Size = New System.Drawing.Size(618, 607)
        Me.tpgPersonnel.TabIndex = 4
        Me.tpgPersonnel.Text = "Personnel "
        '
        'pnlpersonnel
        '
        Me.pnlpersonnel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlpersonnel.Location = New System.Drawing.Point(0, 56)
        Me.pnlpersonnel.Name = "pnlpersonnel"
        Me.pnlpersonnel.Size = New System.Drawing.Size(618, 551)
        Me.pnlpersonnel.TabIndex = 5
        '
        'ToolBar2
        '
        Me.ToolBar2.AutoSize = False
        Me.ToolBar2.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tlbtimesheet, Me.tlbadmin, Me.tlbit})
        Me.ToolBar2.ButtonSize = New System.Drawing.Size(67, 50)
        Me.ToolBar2.DropDownArrows = True
        Me.ToolBar2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ToolBar2.Location = New System.Drawing.Point(0, 0)
        Me.ToolBar2.Name = "ToolBar2"
        Me.ToolBar2.ShowToolTips = True
        Me.ToolBar2.Size = New System.Drawing.Size(618, 56)
        Me.ToolBar2.TabIndex = 4
        '
        'tlbtimesheet
        '
        Me.tlbtimesheet.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tlbtimesheet.Text = "Timesheet"
        '
        'tlbadmin
        '
        Me.tlbadmin.Enabled = False
        Me.tlbadmin.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tlbadmin.Text = "Admin"
        '
        'tlbit
        '
        Me.tlbit.Text = "IT issues"
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Location = New System.Drawing.Point(242, 16)
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
        'frmHome
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.BackColor = System.Drawing.Color.Orange
        Me.ClientSize = New System.Drawing.Size(650, 615)
        Me.Controls.Add(Me.tbcHome)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.mnuMainMenu
        Me.Name = "frmHome"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.tbcHome.ResumeLayout(False)
        Me.tpgHome.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.NumericUpDown_InterSectionSpacingPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_FooterHeightPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_HeaderHeightPercentage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.NumericUpDown_PagesAcross, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.dgrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpgClients.ResumeLayout(False)
        Me.pnlContacts.ResumeLayout(False)
        Me.grpcontactgrid.ResumeLayout(False)
        CType(Me.dtgClients, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpContactsearch.ResumeLayout(False)
        Me.tpgLeads.ResumeLayout(False)
        Me.pnlleads.ResumeLayout(False)
        Me.grpleadsgrid.ResumeLayout(False)
        CType(Me.dtgLeads, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpLeadssearch.ResumeLayout(False)
        Me.tpgJobs.ResumeLayout(False)
        Me.pnljobs.ResumeLayout(False)
        Me.grpjobsearch.ResumeLayout(False)
        Me.grpjobs.ResumeLayout(False)
        CType(Me.dtgJobs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpgEquip.ResumeLayout(False)
        Me.pnlequipcontrols.ResumeLayout(False)
        Me.tpgPersonnel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "public members"
    Public strclientss, strleadss, strjobss, strequipss, strpersonnelss, strvalss As String
    Public currentdate As String
    '-------jobs know which button was ckicked
    Dim which As String = ""
    '0 is search
    '1 is current '2 is completed '3 is delivered '4 is show all
    '----------
    Private Delegate Sub ddelegate()
    Public reportoption As String = "0"
#End Region

#Region "Clients"
    Private Function returnhittest(ByVal hit As System.Windows.Forms.DataGrid.HitTestType) As String
        Try
            Select Case hit
                Case System.Windows.Forms.DataGrid.HitTestType.Cell

                    Dim mycell As New DataGridCell()

                    mycell.RowNumber = Me.dtgClients.CurrentRowIndex
                    mycell.ColumnNumber = 0

                    returnhittest = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell()

                    mycell.RowNumber = Me.dtgClients.CurrentRowIndex
                    mycell.ColumnNumber = 0

                    returnhittest = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case Else
                    returnhittest = ""
            End Select
        Catch ex As Exception
            MessageBox.Show(Text:="Error:" & ex.Message.ToString, _
            caption:="Throwing an exception", _
            Icon:=MessageBoxIcon.Information, _
            buttons:=MessageBoxButtons.OK)

        End Try
    End Function
    Private Sub frmHome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Me.Invalidate(True)
            'Dim otter As New ThreadStart(AddressOf dbconnect)
            'mythread = New System.Threading.Thread(otter)
            'mythread.Start()
            resizegrid = 1
            'Dim threadstart As Thread(  ThreadStart(AddressOf Me.loadgrid))
            myfrmAddClientsform = 0
            Me.txtparams.Focus()
            jobsform = False
            'Dim t As Integer
            't = Me.Controls.GetChildIndex(btnClose)

            '-------------------administration of rights
            Try
                Dim myarray() As String
                strvalss = seclevel
                myarray = strvalss.Split(":")
                strclientss = myarray(0)
                strjobss = myarray(1)
                strleadss = myarray(2)
                strequipss = myarray(3)
                strpersonnelss = myarray(4)
            Catch xc As Exception

            End Try
            '------------------------
            '-------------load current date
            Try
                Dim dm As New datemanipulation
                Dim Threadm As New System.Threading.Thread( _
                    AddressOf dm.curdateinvoke)
                Threadm.Start()
            Catch xc As Exception

            End Try


            '----------------------
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Private Sub loadgrid()
        Try
            myForms.Main.Invoke(New mydelegate(AddressOf loadgrid2))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub loadgrid2()
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Dim currentCursor As Cursor = Cursor.Current
        Try
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select client_no, Name ,description, oclient_no " _
            & " from clients order by client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "clients")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgClients.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------


            Call AddCustomDataTableStyle()

        Catch t As Exception

        Finally
            'statusBar1.Text = "Done"
            Cursor.Current = currentCursor


        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    'Private Sub btnsearchno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim currentcursor As Cursor = Cursor.Current
    '    Try
    '        Dim connectstr As String
    '        Dim nv As NameValueCollection
    '        nv = ConfigurationSettings.AppSettings()
    '        connectstr = nv("connectionstring")
    '        Dim str As String
    '        Cursor.Current = Cursors.WaitCursor
    '        str = Trim(Me.txtparams.Text)

    '        If str.Trim.Length = 0 Then
    '            Exit Try
    '        End If

    '        '-----------------try this dave
    '        Call dbconnect()
    '        Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
    '        Dim custDS As DataSet = New DataSet()
    '        Dim adors As New ADODB.Recordset()
    '        Dim str23 As String = "select client_no, Name, description, oclient_no from clients" _
    '        & " where lower(client_no) like '%" & LCase(str) & "%' order by client_no"
    '        '--------------oledbdataadapter.fill--------------
    '        adors.Open(str23, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)
    '        custDA.Fill(custDS, adors, "clients")
    '        Dim tname As String = custDS.Tables(0).TableName()
    '        Me.dtgClients.SetDataBinding(custDS, tname)
    '        connect.Close()
    '        '--------------------this is quite cool---------------------------------------------------------------------


    '        Call AddCustomDataTableStyle()
    '    Catch t As Exception

    '    Finally
    '        Cursor.Current = currentcursor

    '    End Try
    'End Sub
    Private Sub dtgClients_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgClients.DoubleClick
        Try
            Dim results As String
            results = returnhittest(hti.Type)
            If results <> "" Then
                If jobsform = False Then
                    Dim mycell As New DataGridCell()
                    Dim a() As String
                    a = results.Split("|")
                    mycell.RowNumber = CInt(a(0))
                    mycell.ColumnNumber = CInt(a(1))
                    myclientno = Me.dtgClients(mycell)
                    mycell.ColumnNumber = 1
                    myclientname = Me.dtgClients(mycell)
                    Dim myform As New frmMe()
                    myForms.CustomerForm3 = myform
                    myForms.CustomerForm3.Show()
                    jobsform = True
                Else
                    Dim mycell As New DataGridCell()
                    Dim a() As String
                    a = results.Split("|")
                    mycell.RowNumber = CInt(a(0))
                    mycell.ColumnNumber = CInt(a(1))
                    myclientno = dtgClients(mycell)
                    mycell.ColumnNumber = 1
                    myclientname = dtgClients(mycell)
                    myForms.CustomerForm3.lblClientName.Text = myclientname
                    myForms.CustomerForm3.lblClientNo.Text = myclientno
                    myForms.CustomerForm3.loadgridexistingjobs()
                    myForms.CustomerForm3.loadgridcontact()
                    myForms.CustomerForm3.loadleads()
                End If
            End If
            'Me.Dispose(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString)
        End Try
    End Sub
    Private Sub dtgClients_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub
    Private Sub AddCustomDataTableStyle()
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = "clients"
            Dim mywidth, mywidth1 As Integer
            mywidth = Me.dtgClients.Width - 10
            mywidth = mywidth / 3
            mywidth1 = mywidth / 3
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "client_no"
            myno.HeaderText = "Client Number"
            myno.Width = mywidth1 + 12
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "name"
            myname.HeaderText = "Name"
            myname.Width = mywidth + mywidth1 - 6
            ts1.GridColumnStyles.Add(myname)


            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn()
            mydesc.MappingName = "description"
            mydesc.HeaderText = "Description"
            mydesc.Width = mywidth + mywidth1 - 6
            ts1.GridColumnStyles.Add(mydesc)

            Dim myno2 As New DataGridTextBoxColumn()
            myno2.MappingName = "oclient_no"
            myno2.HeaderText = "Old Number"
            myno2.Width = mywidth1 + 12
            ts1.GridColumnStyles.Add(myno2)

            '' Add a second column style.
            'Dim mydesc1 As New DataGridTextBoxColumn()
            'mydesc1.MappingName = "least_status"
            'mydesc1.HeaderText = "Least Status"
            'mydesc1.Width = mywidth + mywidth1 - 6
            'ts1.GridColumnStyles.Add(mydesc1)
            ' Add the DataGridTableStyle objects to the collection.

            dtgClients.TableStyles.Clear()
            ts1.AllowSorting = False
            dtgClients.TableStyles.Add(ts1)

        Catch ex As Exception

        End Try

    End Sub 'AddCustomDataTableStyle
    Private Sub dtgClients_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgClients.MouseDown

        Try
            hti = Me.dtgClients.HitTest(New Point(e.X, e.Y))
            dtgClients.AllowSorting = False
        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally

        End Try
    End Sub
    Private Sub btnexit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        isloading = False
        'Call closeprogram()
    End Sub
    Protected Overrides Sub Finalize()
        isloading = False
        ' Call closeprogram()
        MyBase.Finalize()
    End Sub
    Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
        Try
            ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
            ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
            If Me.tbcHome.SelectedTab Is Me.tpgClients Then
                If keyData = System.Windows.Forms.Keys.Return Then

                    Dim E As System.EventArgs

                    Call Me.btnsearchname_Click(Me, E)

                    Return True ' True means we've processed the key
                    'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)

                Else
                    Return MyBase.ProcessDialogKey(keyData)
                End If
            End If

        Catch ex As Exception
            'Trace.WriteLine(ex.ToString())
            MsgBox(ex.Message.ToString, , Title:="Return key")

        End Try
    End Function
    Private Sub showmeall()
        Dim connect As New ADODB.Connection()
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

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor

            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select * from clients where lower(name)" _
            & " like '" & LCase(str) & "%' order by client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "clients")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgClients.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------


            Call AddCustomDataTableStyle()
        Catch t As Exception

        Finally
            Cursor.Current = currentcursor

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        Try
            If refreshclients = True Then
                'Me.Invoke(New mydelegate(AddressOf initializegrid))
                Me.dtgClients.Invoke(New mydelegate(AddressOf loadgrid))
                refreshclients = False
            End If
            If refreshleadshome = True Then
                cboleads.Items.Clear()
                loadleads()
                refreshleadshome = False
            End If
            'Try
            '    If isloadleads = False Then
            '        loadleads()
            '        isloadleads = True
            '    End If

            '    refreshleads = True
            '    Me.OnActivated(e)
            'Catch ex As Exception

            'End Try
        Catch ex As Exception

        End Try
    End Sub
    Private Sub initializegrid()
        Try
            'CType(Me.dtgClients, System.ComponentModel.ISupportInitialize).BeginInit()
            'Me.dtgClients.DataMember = ""
            'Me.dtgClients.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            'Me.dtgClients.HeaderForeColor = System.Drawing.SystemColors.ControlText
            'Me.dtgClients.Location = New System.Drawing.Point(8, 40)
            'Me.dtgClients.Name = "dtgClients"
            'Me.dtgClients.PreferredColumnWidth = 300
            'Me.dtgClients.ReadOnly = True
            'Me.dtgClients.Size = New System.Drawing.Size(936, 240)
            'Me.dtgClients.TabIndex = 3
            'Me.dtgClients.TabStop = False
            'Me.grpClients.Controls.AddRange(New System.Windows.Forms.Control() {Me.dtgClients})
            'CType(Me.dtgClients, System.ComponentModel.ISupportInitialize).EndInit()
            'grpClients.Width = Me.Width - 20
            'dtgClients.Width = Me.Width - 35

            'grpClients.Height = Me.Height - GroupBox1.Height - 60
            'dtgClients.Height = grpClients.Height - 50
        Catch er As Exception

        End Try
    End Sub
    Private Sub btnshowall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnshowall.Click
        Try
            Call showmeall()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnsearchname_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearchname.Click
        Dim connect As New ADODB.Connection()
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

        Dim currentcursor As Cursor = Cursor.Current
        Try

            Dim str As String
            Cursor.Current = Cursors.WaitCursor
            str = Trim(Me.txtparams.Text)
            If str.Trim.Length = 0 Then
                Exit Try
            End If

            Dim dfv As String
            Select Case Me.cboclientssearchfield.Text.Trim()
                Case "Client number"
                    dfv = "client_no"
                Case "Name"
                    dfv = "name"
                Case "Description"
                    dfv = "description"
                Case Else
                    'Old number
                    dfv = "oclient_no"
            End Select
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str24 As String = "select client_no, Name, description, oclient_no " _
            & " from clients where lower(" & dfv & ") like '%" & LCase(str) & "%' order by name"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str24, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "clients")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgClients.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------


            Call AddCustomDataTableStyle()
        Catch t As Exception

        Finally
            Cursor.Current = currentcursor

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnaddnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddnew.Click
        Try
            Dim x As Boolean = canmanipulateclients()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate clients contact administrator", "Clients", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim newClient As New frmAddClients()
            If myfrmAddClientsform = 0 Then
                newClient.Show()
            End If
        Catch ex As Exception

        End Try
    End Sub
    'Private Sub tpgClients_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tpgClients.VisibleChanged
    '    Try
    '        If tpgClients.Visible = True Then
    '            dtgClients.Invoke(New mydelegate(AddressOf Me.loadgrid))
    '        Else
    '            dtgClients.SetDataBinding("", "")
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Public Function canmanipulateclients() As Boolean
        Try
            Dim arr() As String
            arr = strclientss.Split(",")
            If arr(1) = "1" Then
                canmanipulateclients = True
            Else
                canmanipulateclients = False
            End If
        Catch ex As Exception
            Try
                canmanipulateclients = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "leads"
    Private ptname As String
    Public Delegate Sub mydelegateleads()
    Private Sub loadleads2()
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Try
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                       & "leads.leads_no,leads.client_no,clients.name,leads.title,leads.status," _
                           & "leads.date_sniffed,leads.amount,leads.department" _
                           & " from leads inner join clients on leads.client_no = clients.client_no" _
                           & " where lower(leads.status) <> '" & "job" & "'" _
                           & " order by leads.status"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            ptname = tname
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------
            addcustleadstablestyle(tname)
            Me.cboleads.Items.Add("Suspect")
            Me.cboleads.Items.Add("Prospect")
            Me.cboleads.Items.Add("Proposal")
            Me.cboleads.Items.Add("PHO")
            Me.cboleads.Items.Add("Failed Leads")
            'adoConn.Close()
            'With adors
            '    .CursorLocation = CursorLocationEnum.adUseClient
            '    .CursorType = CursorTypeEnum.adOpenKeyset
            '    Dim str As String = "select " _
            '           & "leads.client_no,leads.leads_no,leads.descrip,leads.date_sniffed,leads.title," _
            '               & "leads.status,leads.amount,clients.name" _
            '               & " from leads inner join clients on leads.client_no = clients.client_no" _
            '               & " where lower(leads.status) <> '" & "job" & "'" _
            '               & " order by leads.status"
            '    .Open(str, connect)
            'End With

            'Me.dtgLeads.DataSource = adors
            'dtgLeads.DataMember = adors

            'Dim tname As String = "Mytable"
            'addcustleadstablestyle(tname)
            'Me.cboleads.Items.Add("Suspect")
            'Me.cboleads.Items.Add("Prospect")
            'Me.cboleads.Items.Add("Proposal")
            'Me.cboleads.Items.Add("PHO")
            'Me.cboleads.Items.Add("Failed Leads")
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & _
                     ex.InnerException.ToString() & ex.StackTrace.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub loadleads()
        Try
            myForms.Main.Invoke(New mydelegateleads(AddressOf loadleads2))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub addcustleadstablestyle(ByVal tname As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = dtgLeads.Width - 20
            mywidth = mywidth / 6

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "leads_no"
            myno.HeaderText = "Lead Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "client_no"
            myname.HeaderText = "Client Number"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "name"
            myname100.HeaderText = "Name"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)
            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn()
            mydesc.MappingName = "title"
            mydesc.HeaderText = "Title"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn()
            mydesc2.MappingName = "status"
            mydesc2.HeaderText = "Status"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc200 As New DataGridTextBoxColumn()
            mydesc200.MappingName = "date_sniffed"
            mydesc200.HeaderText = "Date"
            mydesc200.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200)

            ' Add a second column style.
            Dim mydesc2000 As New DataGridTextBoxColumn()
            mydesc2000.MappingName = "amount"
            mydesc2000.HeaderText = "Amount"
            mydesc2000.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2000)
            ' Add a second column style.
            Dim mydesc200v As New DataGridTextBoxColumn()
            mydesc200v.MappingName = "department"
            mydesc200v.HeaderText = "Department"
            mydesc200v.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200v)
            ' Add the DataGridTableStyle objects to the collection.
            dtgLeads.TableStyles.Clear()
            ts1.AllowSorting = False
            dtgLeads.TableStyles.Add(ts1)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub dtgLeads_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgLeads.DoubleClick
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim leadno, cno, status, descrip, name, ddate, title
            Dim amount, department
            Dim results As String
            results = returnhittestleads(htileads.Type)
            If results <> "" Then
                Dim mycell As New DataGridCell()
                Dim a() As String
                a = results.Split("|")
                mycell.RowNumber = CInt(a(0))
                mycell.ColumnNumber = CInt(a(1))
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    leadno = Me.dtgLeads(mycell)
                End If
                mycell.ColumnNumber = 1
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    cno = Me.dtgLeads(mycell)
                End If
                mycell.ColumnNumber = 2
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    name = Me.dtgLeads(mycell)
                End If
                mycell.ColumnNumber = 3
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    title = Me.dtgLeads(mycell)
                End If
                mycell.ColumnNumber = 4
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    status = Me.dtgLeads(mycell)
                End If
                mycell.ColumnNumber = 5
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    ddate = Me.dtgLeads(mycell)
                End If
                mycell.ColumnNumber = 6
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    amount = Me.dtgLeads(mycell)
                End If
                mycell.ColumnNumber = 7
                If Convert.IsDBNull(dtgLeads(mycell)) = False Then
                    department = Me.dtgLeads(mycell)
                End If

                Dim rs As New ADODB.Recordset()
                With rs
                    .CursorLocation = CursorLocationEnum.adUseClient
                    .CursorType = CursorTypeEnum.adOpenForwardOnly
                    Dim dstr = "select descrip from  leads where leads_no='" & leadno & "'"
                    .Open(dstr, connect)
                    If .BOF = False And .EOF = False Then
                        descrip = .Fields("descrip").Value
                        'amount = .Fields("amount").Value
                        'department = .Fields("department").Value
                    End If
                End With
                rs.Close()
                If editleads = False Then
                    Dim form As New frmEditLead()
                    form.clientno = cno
                    form.cstatus = status
                    'form.cname = lblClientName.Text
                    form.leadno = leadno
                    form.desription = descrip
                    form.cname = name
                    form.ddate = ddate

                    form.amount = amount
                    form.title = title
                    myForms.CustomerForm4 = form
                    Try
                        myForms.CustomerForm4.txtdepartment.Text = department
                    Catch sd As Exception

                    End Try
                    Try

                    Catch es As Exception
                        Try

                        Catch eg As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.dtpsniffed.Value = CDate(ddate)
                    Catch es As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.cboProspect.Text = status
                    Catch es As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.txtDesc.Text = descrip
                    Catch es As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.txtAmount.Text = amount
                    Catch es As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.txttitle.Text = title
                    Catch es As Exception

                    End Try

                    'myForms.CustomerForm4.txtAmount.Enabled = False
                    myForms.CustomerForm4.Show()
                    'form.Show()
                    'myForms.CustomerForm = form
                    editleads = True
                Else

                    myForms.CustomerForm4.clientno = cno
                    myForms.CustomerForm4.cstatus = status
                    myForms.CustomerForm4.cname = name
                    myForms.CustomerForm4.leadno = leadno
                    myForms.CustomerForm4.desription = descrip
                    myForms.CustomerForm4.ddate = ddate
                    myForms.CustomerForm4.title = title
                    myForms.CustomerForm4.amount = amount
                    'myForms.CustomerForm4.txtAmount.Enabled = False
                    Try

                    Catch er As Exception
                    End Try
                    Try
                        myForms.CustomerForm4.txtAmount.Text = amount
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.txtAmount.Text = ""
                        Catch eg As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.txttitle.Text = title
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.txttitle.Text = ""
                        Catch eg As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.txtDesc.Text = descrip
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.txtDesc.Text = ""
                        Catch eg As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.cboProspect.Text = status
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.cboProspect.Text = ""
                        Catch eg As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.dtpsniffed.Value = CDate(ddate)
                    Catch ex1 As Exception
                        Try
                            myForms.CustomerForm4.dtpsniffed.Text = ""
                        Catch eg As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.lblCompany.Text = name

                    Catch er45 As Exception

                    End Try
                    editleads = True
                    'form.descriptions = descrip
                    'form.status = status
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Cursor.Current = currentcursor
        End Try

        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Function getname(ByVal cno) As String
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Function
        End Try

        Try
            Dim rs As New ADODB.Recordset()
            With rs
                .Open(Source:="select name  from clients  where client_no ='" & cno & "'", _
                activeconnection:=connect, cursortype:=CursorTypeEnum.adOpenForwardOnly)
                If .BOF = False And .EOF = False Then
                    Return .Fields("name").Value
                End If
            End With

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Sub dtgLeads_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgLeads.MouseDown
        Try
            htileads = Me.dtgLeads.HitTest(New Point(e.X, e.Y))
            dtgLeads.AllowSorting = False
            If htileads.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader Then

            End If
        Catch ex As System.ArgumentOutOfRangeException
            MsgBox(ex.Message.ToString())

        Finally

        End Try
    End Sub
    Private Function returnhittestleads(ByVal hit As System.Windows.Forms.DataGrid.HitTestType) As String
        Try
            Select Case hit
                Case System.Windows.Forms.DataGrid.HitTestType.Cell

                    Dim mycell As New DataGridCell()
                    mycell.RowNumber = Me.dtgLeads.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittestleads = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell()
                    mycell.RowNumber = Me.dtgLeads.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittestleads = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case Else
                    returnhittestleads = ""
            End Select
        Catch ex As Exception
            MessageBox.Show(Text:="Error:" & ex.Message.ToString, _
            caption:="Throwing an exception", _
            Icon:=MessageBoxIcon.Information, _
            buttons:=MessageBoxButtons.OK)

        End Try
    End Function
    Private Sub addsearchtablestyle(ByVal tname As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = dtgLeads.Width
            mywidth = mywidth / 7

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "leads_no"
            myno.HeaderText = "Lead Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "client_no"
            myname.HeaderText = "Client Number"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            Dim mydesc100 As New DataGridTextBoxColumn()
            mydesc100.MappingName = "name"
            mydesc100.HeaderText = "Name"
            mydesc100.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc100)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn()
            mydesc.MappingName = "title"
            mydesc.HeaderText = "Title"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn()
            mydesc2.MappingName = "status"
            mydesc2.HeaderText = "Status"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc200 As New DataGridTextBoxColumn()
            mydesc200.MappingName = "date_sniffed"
            mydesc200.HeaderText = "Date"
            mydesc200.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200)

            ' Add a second column style.
            Dim mydesc2000 As New DataGridTextBoxColumn()
            mydesc2000.MappingName = "amount"
            mydesc2000.HeaderText = "Amount"
            mydesc2000.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2000)

            ' Add a second column style.
            Dim mydesc200v As New DataGridTextBoxColumn()
            mydesc200v.MappingName = "department"
            mydesc200v.HeaderText = "Department"
            mydesc200v.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200v)

            ' Add the DataGridTableStyle objects to the collection.
            dtgLeads.TableStyles.Clear()
            ts1.AllowSorting = False
            dtgLeads.TableStyles.Add(ts1)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnShowAllleads_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowAllleads.Click
        Try
            Me.cboleads.Items.Clear()
            loadleads()

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
            MessageBox.Show(ex.Message.ToString() & _
            ex.InnerException.ToString() & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub btnleadspho_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnleadspho.Click
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Try

            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                       & "leads.client_no,leads.leads_no,clients.name,leads.title,leads.status," _
                       & "leads.date_sniffed,leads.amount,leads.department" _
                       & " from leads inner join clients on leads.client_no = clients.client_no and " _
                       & " lower(leads.status) like " _
                       & "'%" & "pho" & "%'" _
                       & " order by leads.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addsearchtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnleadssearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnleadssearch.Click
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Try

            Dim sdate, edate As String
            sdate = dtpsdate.Value.Year & "-" _
            & dtpsdate.Value.Month & "-" _
            & dtpsdate.Value.Day & " " _
            & dtpsdate.Value.Hour & ":" _
            & dtpsdate.Value.Minute & ":" _
            & dtpsdate.Value.Second
            edate = dtpedate.Value.Year & "-" _
             & dtpedate.Value.Month & "-" _
            & dtpedate.Value.Day & " " _
            & dtpedate.Value.Hour & ":" _
            & dtpedate.Value.Minute & ":" _
            & dtpedate.Value.Second

            '------------configure search parameter
            Dim dfv As String
            Select Case Me.cboleadssearchfield.Text.Trim()

                Case "Client number"
                    dfv = "clients.client_no"
                Case "Lead number"
                    dfv = "leads.leads_no"
                Case "Name"
                    dfv = "clients.name"
                Case "Title"
                    dfv = "leads.title"
                Case Else
                    'Status()
                    dfv = "leads.status"
            End Select
            Dim dfv2 As String
            Select Case Me.cboleads.SelectedIndex
                Case 0
                    dfv2 = "sus"
                Case 1
                    dfv2 = "pros"
                Case 2
                    dfv2 = "prop"
                Case 3
                    dfv2 = "ph"
                Case 4
                    dfv2 = "fa"
                Case Else
                    'Status()
                    dfv2 = ""
            End Select
            '----------end
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                       & "leads.client_no,leads.leads_no,clients.name,leads.title,leads.status," _
                       & "leads.date_sniffed,leads.amount,leads.department" _
                       & " from leads inner join clients on leads.client_no = clients.client_no and " _
                       & " lower(leads.status) like " _
                       & "'%" & dfv2.ToLower() & "%'" _
                       & " and  lower(" & dfv & ") like '%" & txtleads.Text.Trim.ToLower() & "%'" _
                       & " and date_sniffed > '" & sdate & "' and date_sniffed < '" & edate & "'" _
                       & " order by leads.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addsearchtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnleadssuspect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnleadssuspect.Click
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Try

            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                       & "leads.client_no,leads.leads_no,clients.name,leads.title,leads.status," _
                       & "leads.date_sniffed,leads.amount,leads.department" _
                       & " from leads inner join clients on leads.client_no = clients.client_no and " _
                       & " lower(leads.status) like " _
                       & "'%" & "suspect" & "%'" _
                       & " order by leads.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addsearchtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnleadsprospect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnleadsprospect.Click
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Try

            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                       & "leads.client_no,leads.leads_no,clients.name,leads.title,leads.status," _
                       & "leads.date_sniffed,leads.amount,leads.department" _
                       & " from leads inner join clients on leads.client_no = clients.client_no and " _
                       & " lower(leads.status) like " _
                       & "'%" & "prospect" & "%'" _
                       & " order by leads.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addsearchtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnleadsproposal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnleadsproposal.Click
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Try

            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                       & "leads.client_no,leads.leads_no,clients.name,leads.title,leads.status," _
                       & "leads.date_sniffed,leads.amount,leads.department" _
                       & " from leads inner join clients on leads.client_no = clients.client_no and " _
                       & " lower(leads.status) like " _
                       & "'%" & "proposal" & "%'" _
                       & " order by leads.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addsearchtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnaddleads_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddleads.Click
        Try
            Dim x As Boolean = canmanipulateleads()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate leads contact administrator", "Leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If addleads = False Then
                Dim form As New frmAddLead()
                form.grpnormal.Visible = False
                form.pnlSearch.Visible = True
                'form.clientno = Me.lblClientNo.Text.ToString()
                'form.cname = Me.lblClientName.Text.ToString()
                'form.Show()
                myForms.CustomerForm = form
                myForms.CustomerForm.Show()
                addleads = True
            Else

                myForms.CustomerForm.grpnormal.Visible = False
                myForms.CustomerForm.pnlSearch.Visible = True
                'form.clientno = Me.lblClientNo.Text.ToString()
                'form.cname = Me.lblClientName.Text.ToString()
            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnfailedlead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnfailedlead.Click
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Try

            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                       & "leads.client_no,leads.leads_no,clients.name,leads.title,leads.status," _
                       & "leads.date_sniffed,leads.amount,leads.department" _
                       & " from leads inner join clients on leads.client_no = clients.client_no and " _
                       & " lower(leads.status) like " _
                       & "'%" & "failed" & "%'" _
                       & " order by leads.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addsearchtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Function canmanipulateleads() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strleadss.Split(",")
            If arr(1) = "1" Then
                canmanipulateleads = True
            Else
                canmanipulateleads = False
            End If
        Catch ex As Exception
            Try
                canmanipulateleads = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "jobs"
    Private jname
    Private Sub addcurrjobtablestyle(ByVal tname As String)
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = Me.dtgJobs.Width
            mywidth = mywidth / 8
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Try

                ' Add a second column style.
                Dim mydesc1 As New DataGridTextBoxColumn()
                mydesc1.MappingName = "client_no"
                mydesc1.HeaderText = "Client Number"
                mydesc1.Width = mywidth
                ts1.GridColumnStyles.Add(mydesc1)

            Catch bcv As Exception
            End Try

            ' Add a second column style.
            Dim mydesc4 As New DataGridTextBoxColumn()
            mydesc4.MappingName = "name"
            mydesc4.HeaderText = "Name"
            mydesc4.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc4)

            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "job_no"
            myno.HeaderText = "Job Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "job_tittle"
            myname.HeaderText = "Job Title"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn()
            mydesc.MappingName = "job_status"
            mydesc.HeaderText = "Job Status"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            Try
                ' Add a second column style.
                Dim mydesc6 As New DataGridTextBoxColumn()
                mydesc6.MappingName = "techres"
                mydesc6.HeaderText = "Technician Responsible"
                mydesc6.Width = mywidth
                ts1.GridColumnStyles.Add(mydesc6)
            Catch xc As Exception

            End Try


            ' Add a second column style.
            Dim mydesc66 As New DataGridTextBoxColumn()
            mydesc66.MappingName = "amount"
            mydesc66.HeaderText = "Amount"
            mydesc66.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc66)

            ' Add a second column style.
            Dim mydesc66x As New DataGridTextBoxColumn()
            mydesc66x.MappingName = "grossmargin"
            mydesc66x.HeaderText = "Gross Margin(%)"
            mydesc66x.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc66x)

            ' Add the DataGridTableStyle objects to the collection.
            dtgJobs.TableStyles.Clear()
            ts1.AllowSorting = False
            dtgJobs.TableStyles.Add(ts1)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
    End Sub
    Private Function returnhittest3(ByVal hit As System.Windows.Forms.DataGrid.HitTestType) As String
        Try
            Select Case hit
                Case System.Windows.Forms.DataGrid.HitTestType.Cell

                    Dim mycell As New DataGridCell()
                    mycell.RowNumber = dtgJobs.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittest3 = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell()
                    mycell.RowNumber = dtgJobs.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittest3 = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case Else
                    returnhittest3 = ""
            End Select
        Catch ex As Exception
            MessageBox.Show(Text:="Error:" & ex.Message.ToString, _
            caption:="Throwing an exception", _
            Icon:=MessageBoxIcon.Information, _
            buttons:=MessageBoxButtons.OK)

        End Try
    End Function
    Private Sub dtgJobs_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgJobs.MouseDown
        Try
            htijobs = dtgJobs.HitTest(New Point(e.X, e.Y))
            'dtgJobs.AllowSorting = False
            'If htijobs.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader Then
            '    Dim dvm As New DataViewManager()
            '    dvm = dtgJobs.DataSource
            '    dvm.DataViewSettings(0).Sort = dvm.DataSet.Tables(0).Columns(htijobs.Column).ColumnName
            '    dtgJobs.DataSource = dvm
            'End If
        Catch ex As System.ArgumentOutOfRangeException
            MsgBox(ex.Message.ToString())
        Finally
        End Try
    End Sub
    Private Sub dtgJobs_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgJobs.DoubleClick

        Try
            Dim x As Boolean = canviewjobs()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate clients contact administrator", "Clients", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch jnm As Exception
        End Try
        Dim cnnstr As String
         cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection()
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim results As String
            Dim int As Integer = dtgJobs.CurrentRowIndex()

            results = returnhittest3(htijobs.Type)
            If results <> "" Then
                Dim mycell As New DataGridCell()
                Dim a() As String
                a = results.Split("|")
                mycell.RowNumber = CInt(a(0))
                mycell.ColumnNumber = CInt(a(1))
                '----------
                Dim ds As System.Data.DataViewManager = New System.Data.DataViewManager()
                Dim dvm As New DataViewManager()
                dvm = Me.dtgJobs.DataSource
                Dim jobno As String
                Dim ojobno As String
                Dim jobtitle As String
                Dim contname, ddate As String
                Dim jobstatus, amount As String
                Dim tecres, desc, department As String
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("job_no")) = False Then
                        jobno = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("job_no")
                    End If
                Catch vb As Exception

                End Try

                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("job_tittle")) = False Then
                        jobtitle = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("job_tittle")
                    End If

                Catch vb As Exception

                End Try
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("cont")) = False Then
                        contname = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("cont")
                    End If
                Catch vb As Exception

                End Try

                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("job_status")) = False Then
                        jobstatus = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("job_status")
                    End If
                Catch vb As Exception

                End Try
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("techres")) = False Then
                        tecres = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("techres")
                    End If
                Catch vb As Exception

                End Try
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("ojob_no")) = False Then
                        ojobno = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("ojob_no")
                        ojobno = "Old job no:" & " " & ojobno
                    Else
                        ojobno = "Old job no:" & " " & ojobno
                    End If
                Catch vb As Exception

                End Try
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("department")) = False Then
                        department = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("department")
                    End If
                Catch vb As Exception

                End Try
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("descrip")) = False Then
                        desc = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("descrip")
                    End If
                Catch vb As Exception

                End Try
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("amount")) = False Then
                        amount = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("amount")
                    End If
                Catch vb As Exception

                End Try
                Try
                    If Convert.IsDBNull(dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("sdate")) = False Then
                        ddate = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("sdate")
                    End If
                Catch vb As Exception

                End Try




                If editjob = False Then

                    Dim form As New frmEditJob()
                    form.jobtitle = jobtitle
                    form.myjobno = jobno
                    form.contname = contname
                    form.tecres = tecres
                    form.jobstatus = jobstatus
                    form.description = desc
                    myForms.CustomerForm2 = form

                    'form.Show()
                    myForms.CustomerForm2.txtJobTitle.Text = jobtitle
                    myForms.CustomerForm2.txtJobNo.Text = jobno
                    myForms.CustomerForm2.txtContactName.Text = contname
                    myForms.CustomerForm2.cboTechnicianresponsible.Text = tecres
                    myForms.CustomerForm2.cboJobstatus.Text = jobstatus
                    myForms.CustomerForm2.txtdesc.Text = desc
                    myForms.CustomerForm2.txtamount.Text = amount
                    myForms.CustomerForm2.lblojobno.Text = ojobno
                    myForms.CustomerForm2.txtdepartment.Text = department
                    Try
                        myForms.CustomerForm2.txtbudget.Text = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("budgetarycost")
                    Catch assf As Exception
                    End Try
                    myForms.CustomerForm2.Show()
                    editjob = True
                Else
                    myForms.CustomerForm2.jobtitle = jobtitle
                    myForms.CustomerForm2.myjobno = jobno
                    myForms.CustomerForm2.contname = contname
                    myForms.CustomerForm2.tecres = tecres
                    myForms.CustomerForm2.jobstatus = jobstatus
                    myForms.CustomerForm2.description = desc

                    myForms.CustomerForm2.txtJobTitle.Text = jobtitle
                    myForms.CustomerForm2.txtJobNo.Text = jobno
                    myForms.CustomerForm2.txtContactName.Text = contname
                    myForms.CustomerForm2.cboTechnicianresponsible.Text = tecres
                    myForms.CustomerForm2.cboJobstatus.Text = jobstatus
                    myForms.CustomerForm2.txtdesc.Text = desc
                    myForms.CustomerForm2.txtamount.Text = amount
                    myForms.CustomerForm2.lblojobno.Text = ojobno
                    myForms.CustomerForm2.txtdepartment.Text = department
                    Try
                        myForms.CustomerForm2.txtbudget.Text = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("budgetarycost")
                    Catch assf As Exception
                    End Try
                    editjob = True
                End If
                Try
                    myForms.CustomerForm2.lbClientName.Text = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("name")
                Catch assf As Exception
                End Try
                Try
                    myForms.CustomerForm2.lblClientNo.Text = dvm.DataSet.Tables(0).Rows(mycell.RowNumber).Item("client_no")
                Catch assf As Exception
                End Try
            End If
            Try
                connect.Close()
            Catch ex500 As Exception
            End Try
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Cursor.Current = currentcursor
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnCurrentJobs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCurrentJobs.Click
        Try
            Dim tcj As Thread = New System.Threading.Thread( _
            AddressOf cjinvoke)
            tcj.IsBackground = True
            tcj.Start()
        Catch xa As Exception

        End Try
    End Sub
    Public Sub cjinvoke()
        Try
            myForms.Main.Invoke(New ddelegate(AddressOf cj))
        Catch sa As Exception

        End Try
    End Sub
    Private Sub cj()
        Dim connect As New ADODB.Connection()
        which = "1"
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Try
                dtgJobs.DataSource = Nothing
            Catch xc As Exception

            End Try
            Cursor.Current = Cursors.WaitCursor
            dtgJobs.CaptionText = "Current jobs"
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As System.Data.DataSet = New System.Data.DataSet()
            Dim adors As New ADODB.Recordset()
            '  "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status,rcljobs.techres," _
            '& " rcljobs.amount,rcljobs.grossmargin,rcljobs.costagreed,rcljobs.ojob_no,rcljobs.descrip,rcljobs.date_sniffed,rcljobs.journal,rcljobs.department,rcljobs.budgetarycost " _
            ' 
            Dim str As String = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status,rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin,rcljobs.costagreed,rcljobs.ojob_no,rcljobs.descrip,rcljobs.date_sniffed,rcljobs.journal,rcljobs.department,rcljobs.budgetarycost " _
            & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
            & " lower(rcljobs.job_status) like" _
            & "'%" & LCase("current") & "%' order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            custDS.Tables(0).Columns("job_no").Unique = True
            Dim tname As String = custDS.Tables(0).TableName()
            Dim dv As DataViewManager = New DataViewManager()

            dv.DataSet = custDS
            Me.dtgJobs.SetDataBinding(dv, tname)

            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------


            addcurrjobtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnjobsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnjobsearch.Click
        Try
            Dim tjd As Thread = New System.Threading.Thread( _
            AddressOf jseinvoke)
            tjd.IsBackground = True
            tjd.Start()
        Catch xa As Exception

        End Try
    End Sub
    Public Sub jseinvoke()
        Try
            myForms.Main.Invoke(New ddelegate(AddressOf jse))
        Catch sa As Exception

        End Try
    End Sub
    Public Sub jse()
        Dim connect As New ADODB.Connection()
        which = "0"
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            '-----------------try this dave
            Dim dfv As String
            Select Case Me.cbojobsearchfield.Text.Trim()
                Case "Client number"
                    dfv = "clients.client_no"
                Case "Job title"
                    dfv = "rcljobs.job_tittle"
                Case "Job status"
                    dfv = "rcljobs.job_status"
                Case "Job number"
                    dfv = "rcljobs.job_no"
                Case "Technician responsible"
                    dfv = "rcljobs.techres"
                Case "Amount"
                    dfv = "rcljobs.amount"
                Case "Gross margin"
                    dfv = "rcljobs.grossmargin"
                Case Else
                    dfv = "clients.name"
            End Select


            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status,rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin,rcljobs.costagreed,rcljobs.ojob_no,rcljobs.descrip,rcljobs.date_sniffed,rcljobs.journal,rcljobs.department,rcljobs.budgetarycost " _
                       & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
                       & " lower(rcljobs.job_status) like " _
                       & "'%" & Me.cbojobstatus.Text.Trim.ToLower() & "%'" _
                       & " where  lower(" & dfv & " ) like '%" & txtjobcontactname.Text.Trim.ToLower() & "%'" _
                       & "" _
                       & " order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leadscv")
            Dim tname As String = custDS.Tables(0).TableName()
            Dim dv As DataViewManager = New DataViewManager()

            dv.DataSet = custDS
            Me.dtgJobs.SetDataBinding(dv, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addcurrjobtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnjobshowall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnjobshowall.Click
        Try
            Dim tjd As Thread = New System.Threading.Thread( _
            AddressOf jsinvoke)
            tjd.IsBackground = True
            tjd.Start()
        Catch xa As Exception

        End Try
    End Sub
    Public Sub jsinvoke()
        Try
            myForms.Main.Invoke(New ddelegate(AddressOf js))
        Catch sa As Exception

        End Try
    End Sub
    Public Sub js()
        Dim connect As New ADODB.Connection()
        which = "4"
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()
        Catch jb As Exception

            Exit Sub
        End Try

        Try
            dtgJobs.CaptionText = "Current jobs"
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status,rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin,rcljobs.costagreed,rcljobs.ojob_no,rcljobs.descrip,rcljobs.date_sniffed,rcljobs.journal,rcljobs.department,rcljobs.budgetarycost " _
            & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
            & " lower(rcljobs.job_status) <>" _
            & "'%" & LCase("tt") & "%' order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Dim dv As DataViewManager = New DataViewManager()

            dv.DataSet = custDS
            Me.dtgJobs.SetDataBinding(dv, tname)
            Try
                connect.Close()
            Catch cv As Exception
            End Try

            '--------------------this is quite cool---------------------------------------------------------------------



            addcurrjobtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnjobdelivered_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnjobdelivered.Click
        Try
            Dim tjd As Thread = New System.Threading.Thread( _
            AddressOf jdinvoke)
            tjd.IsBackground = True
            tjd.Start()
        Catch xa As Exception

        End Try
    End Sub
    Public Sub jdinvoke()
        Try
            myForms.Main.Invoke(New ddelegate(AddressOf jd))
        Catch sa As Exception

        End Try
    End Sub
    Public Sub jd()
        Dim connect As New ADODB.Connection()
        which = "3"
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            dtgJobs.CaptionText = "Current jobs"
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status,rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin,rcljobs.costagreed,rcljobs.ojob_no,rcljobs.descrip,rcljobs.date_sniffed,rcljobs.journal,rcljobs.department,rcljobs.budgetarycost " _
            & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
            & " lower(rcljobs.job_status) like" _
            & "'%" & LCase("delivered") & "%' order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Dim dv As DataViewManager = New DataViewManager()

            dv.DataSet = custDS
            Me.dtgJobs.SetDataBinding(dv, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------


            addcurrjobtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btngrossmargin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btngrossmargin.Click
        Dim cur As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim task As New taskclass()
            task.wwhich = which
            Dim tjd As Thread = New System.Threading.Thread( _
           AddressOf task.loopgross)
            tjd.IsBackground = True
            tjd.Start()
        Catch xc As Exception

        Finally
            Cursor.Current = cur
        End Try
    End Sub
    Private Sub btnCompletedJobs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompletedJobs.Click
        Try
            Dim tjd As Thread = New System.Threading.Thread( _
            AddressOf jcinvoke)
            tjd.IsBackground = True
            tjd.Start()
        Catch xa As Exception

        End Try
    End Sub
    Public Sub jcinvoke()
        Try
            myForms.Main.Invoke(New ddelegate(AddressOf jc))
        Catch sa As Exception

        End Try
    End Sub
    Public Sub jc()
        Dim connect As New ADODB.Connection()
        which = "2"
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Try
                dtgJobs.DataSource = Nothing
            Catch xc As Exception

            End Try
            Cursor.Current = Cursors.WaitCursor
            dtgJobs.CaptionText = "Current jobs"
            '-----------------try this dave
            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            ' 
            ' 
            Dim str As String = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status,rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin,rcljobs.costagreed,rcljobs.ojob_no,rcljobs.descrip,rcljobs.date_sniffed,rcljobs.journal,rcljobs.department,rcljobs.budgetarycost " _
            & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
            & " lower(rcljobs.job_status) like" _
            & "'%" & LCase("complete") & "%' order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Dim dv As DataViewManager = New DataViewManager()

            dv.DataSet = custDS
            Me.dtgJobs.SetDataBinding(dv, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------


            addcurrjobtablestyle(tname)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Function canviewjobs() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strjobss.Split(",")
            If arr(0) = "1" Then
                canviewjobs = True
            Else
                canviewjobs = False
            End If
        Catch ex As Exception
            Try
                canviewjobs = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "home"

#End Region

#Region "equip"
    Private Sub tpgEquip_VisibleChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpgEquip.VisibleChanged
        Try

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    '-----------------------------------------------------------------------
    'Private Sub LoadFormIntoPanel()
    '    Dim MyForm2 As New Form2()
    '    MyForm2.Size = pnlContent.Size
    '    MyForm2.TopLevel = False
    '    MyForm2.Parent = pnlContent
    '    MyForm2.Show()
    'End Sub
    '-----------------------------------------------------------------------------------
    Private Sub loadequip()
        Try
            Dim ts As New frminventories()
            myForms.equipments = ts
            myForms.equipments.Size = pnlequip.Size
            myForms.equipments.TopLevel = False
            myForms.equipments.Parent = pnlequip
            myForms.equipments.Dock = DockStyle.Fill
            myForms.equipments.Show()
        Catch we As Exception

        End Try
        '------------massive thread 


        '-----------end of massive thread
    End Sub
    Private Sub btnequipsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnequipsearch.Click
        Try
            Dim str As String
            If isdate = False Then
                str = "select equip_info.*, equip_finances.equip_id as id2,equip_finances.hourly_rate" _
                                           & " from equip_info inner join equip_finances on" _
                                           & " equip_info.equip_id=equip_finances.equip_id "
                str += " where equip_info." & cboequipsearchtrue.Text.Trim & " like '%" & msk.FormattedText.Trim & "%'"
            Else
                Dim a() As String
                Dim inputstring As String = Me.msk.FormattedText
                a = inputstring.Split("and")
                Dim c, d As String
                c = inputstring.Substring(8, 10)
                d = inputstring.Substring(22, 10)
                Dim a1(), a2() As String
                a2 = d.Split("/")

                a1 = c.Split("/")
                Dim sdate, edate As String
                sdate = a1(2) & "-" _
                             & a1(1) & "-" _
                             & a1(0) & " " _
                             & "00" & ":" _
                             & "00" & ":" _
                             & "00"
                edate = a2(2) & "-" _
                        & a2(1) & "-" _
                        & a2(0) & " " _
                        & "23" & ":" _
                        & "59" & ":" _
                        & "59"
                str = "select equip_info.*, equip_finances.equip_id as id2,equip_finances.hourly_rate" _
                                                       & " from equip_info inner join equip_finances on" _
                                                       & " equip_info.equip_id=equip_finances.equip_id "
                str += " where equip_info." & cboequipsearchtrue.Text.Trim & " >= '" & sdate & "'"
                str += " and equip_info." & cboequipsearchtrue.Text.Trim & " <= '" & edate & "'"

            End If

            Dim tasks As taskclass
            tasks.issearch = True
            tasks.strsearch = str

            Dim Threade1 As New System.Threading.Thread( _
                AddressOf tasks.equipinvoke)
            Threade1.Start()
        Catch qw As Exception

        End Try
    End Sub
    Private Sub cboequipsearch_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboequipsearch.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cboequipsearch.SelectedIndex
            Me.msk.Mask = ""
            Me.msk.Text = ""
            isdate = False
            Me.lblsearchparameter.Text = "Type search parameter"
            If indexx = -1 Then
                Exit Try
            ElseIf indexx = 5 Then
                Me.msk.Mask = "between ##/##/####\" & "and ##/##/####"
                Me.lblsearchparameter.Text = "Type search parameter(dd/mm/yyyy)"
                isdate = True
            End If
            Me.cboequipsearchtrue.SelectedIndex = indexx
            Dim strp
            strp = cboequipsearchtrue.Text

        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
#End Region

#Region "personnel"
    'Private Sub tpgPersonnel_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tpgPersonnel.VisibleChanged
    '    Try

    '    Catch ex As Exception
    '        MsgBox(ex.Message.ToString())
    '    End Try
    'End Sub
    Private Sub loadtimesheet()
        Try
            ToolBar2.Buttons(0).Pushed = True
            Dim ts As New frmtime()
            myForms.timesheet = ts
            myForms.timesheet.Size = pnlpersonnel.Size
            myForms.timesheet.TopLevel = False
            myForms.timesheet.Parent = pnlpersonnel
            myForms.timesheet.Dock = DockStyle.Fill
            myForms.timesheet.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
                 & ex.InnerException().ToString() & vbCrLf _
                 & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub ToolBar2_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar2.ButtonClick
        Try
            Dim x As Boolean = myForms.Main.canmanipulatepersonnel()
            If ToolBar2.Buttons.IndexOf(e.Button) = 1 Then
                If x = False Then
                    MessageBox.Show("Not allowed to manipulate personnel contact administrator", "Personnel", _
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
            End If

        Catch xcv As Exception

        End Try
        Dim currentCursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Select Case ToolBar2.Buttons.IndexOf(e.Button)
                Case 0
                    Try
                        myForms.adminform.Close()
                        myForms.adminform = Nothing
                    Catch ex As Exception
                    End Try
                    Try
                        myForms.itissues.Close()
                        myForms.itissues = Nothing
                    Catch er As Exception
                    End Try
                    Try
                        myForms.qjobs.Close()
                        myForms.qjobs = Nothing
                    Catch er As Exception
                    End Try
                    ToolBar2.Buttons(0).Pushed = True
                    ToolBar2.Buttons(1).Pushed = False
                    ToolBar2.Buttons(2).Pushed = False
                    Dim ts As New frmtime()
                    myForms.timesheet = ts
                    myForms.timesheet.Size = pnlpersonnel.Size
                    myForms.timesheet.TopLevel = False
                    myForms.timesheet.Parent = pnlpersonnel
                    myForms.timesheet.Dock = DockStyle.Fill
                    myForms.timesheet.Show()
                    myForms.timesheet.BringToFront()
                Case 1
                    Try
                        myForms.timesheet.Close()
                        myForms.timesheet = Nothing
                    Catch er As Exception
                    End Try
                    Try
                        myForms.itissues.Close()
                        myForms.itissues = Nothing
                    Catch er As Exception
                    End Try
                    Try
                        myForms.qjobs.Close()
                        myForms.qjobs = Nothing
                    Catch er As Exception
                    End Try
                    ToolBar2.Buttons(1).Pushed = True
                    ToolBar2.Buttons(0).Pushed = False
                    ToolBar2.Buttons(2).Pushed = False
                    Dim admin As New frmpersonneladmin()
                    myForms.adminform = admin
                    myForms.adminform.Size = pnlpersonnel.Size
                    myForms.adminform.TopLevel = False
                    myForms.adminform.Parent = pnlpersonnel
                    myForms.adminform.Dock = DockStyle.Fill
                    myForms.adminform.Show()
                    myForms.adminform.BringToFront()
                Case 2 ' it issues and query jobs
                    Try
                        myForms.timesheet.Close()
                        myForms.timesheet = Nothing
                    Catch er As Exception
                    End Try
                    Try
                        myForms.adminform.Close()
                        myForms.adminform = Nothing
                    Catch ex As Exception
                    End Try
                    Try
                        myForms.qjobs.Close()
                        myForms.qjobs = Nothing
                    Catch er As Exception
                    End Try
                    ToolBar2.Buttons(2).Pushed = True
                    ToolBar2.Buttons(1).Pushed = False
                    ToolBar2.Buttons(0).Pushed = False
                    Dim ad As New frmit()
                    myForms.itissues = ad
                    myForms.itissues.Size = pnlpersonnel.Size
                    myForms.itissues.TopLevel = False
                    myForms.itissues.Parent = pnlpersonnel
                    myForms.itissues.Dock = DockStyle.Fill
                    myForms.itissues.Show()
                    myForms.itissues.BringToFront()
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
                          & ex.InnerException().ToString() & vbCrLf _
                          & ex.StackTrace.ToString())
        Finally
            Cursor.Current = currentCursor
        End Try

    End Sub
#End Region

#Region "all forms:resize,tabselect"
    '#Region " resize code"
    '    Private mypoint As System.Drawing.Point
    '    Private pttxtparams As System.Drawing.Point
    '    Private ptbtnsearchno As System.Drawing.Point
    '    Private ptbtnsearchname As System.Drawing.Point
    '    'leads control-----------------
    '    Private ptlblstatus As System.Drawing.Point
    '    Private ptlblname As System.Drawing.Point
    '    Private ptlblstartdate As System.Drawing.Point
    '    Private ptlblenddate As System.Drawing.Point
    '    Private ptcbostatus As System.Drawing.Point
    '    Private pttxtname As System.Drawing.Point
    '    Private ptdtpstartdate As System.Drawing.Point
    '    Private ptdtpenddate As System.Drawing.Point
    '    Private ptbtnleadssearch As System.Drawing.Point
    '    '-------size
    '    Private szlblstatus As System.Drawing.Size
    '    Private szlblname As System.Drawing.Size
    '    Private szlblstartdate As System.Drawing.Size
    '    Private szlblenddate As System.Drawing.Size
    '    Private szcbostatus As System.Drawing.Size
    '    Private sztxtname As System.Drawing.Size
    '    Private szdtpstartdate As System.Drawing.Size
    '    Private szdtpenddate As System.Drawing.Size
    '    'jobs------------
    '    Private ptlbljobcontactname As System.Drawing.Point
    '    Private ptlbljobstatus As System.Drawing.Point
    '    Private ptcbojobstatus As System.Drawing.Point
    '    Private ptltxtjobcontactname As System.Drawing.Point
    '    Private ptbtnjobsearch As System.Drawing.Point
    '    ' size-----------------------
    '    Private szlbljobcontactname As System.Drawing.Size
    '    Private szlbljobstatus As System.Drawing.Size
    '    Private szcbojobstatus As System.Drawing.Size
    '    Private sztxtjobcontactname As System.Drawing.Size
    '    Private szbtnjobsearch As System.Drawing.Size
    '    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
    '        Try
    '            Select Case Me.WindowState

    '                Case FormWindowState.Maximized

    '                    Me.tbcHome.Width = Me.Width - 10
    '                    Me.tbcHome.Height = Me.Height - 10
    '                    'code for resizing contacts-----------------------
    '                    Me.pnlContacts.Width = Me.tbcHome.Width - 10
    '                    Me.grpcontactgrid.Width = Me.pnlContacts.Width - 10
    '                    Me.grpContactsearch.Width = Me.pnlContacts.Width - 10
    '                    Me.dtgClients.Width = grpcontactgrid.Width - 30
    '                    Me.pnlContacts.Height = Me.Height - 56
    '                    grpcontactgrid.Height = _
    '                    pnlContacts.Height - (grpContactsearch.Height)
    '                    dtgClients.Height = grpcontactgrid.Height - 50
    '                    AddCustomDataTableStyle() 're add custom table style
    '                    'reposition controls on the contacts search grid
    '                    Dim intposlblsearch = ((Me.grpcontactgrid.Width / 2) - (Me.lblContactsearch.Width / 2)) 'label
    '                    Me.lblContactsearch.Location = New System.Drawing.Point(intposlblsearch, mypoint.Y)
    '                    Dim intpostxtparam = ((Me.grpcontactgrid.Width / 2) - (Me.txtparams.Width / 2)) 'textbox
    '                    Me.txtparams.Location = New System.Drawing.Point(intpostxtparam, pttxtparams.Y)

    '                    Dim intposbtnsearchno = ((Me.grpcontactgrid.Width / 2) - (Me.btnsearchno.Width + 30))   'search no
    '                    Me.btnsearchno.Location = New System.Drawing.Point(intposbtnsearchno, ptbtnsearchno.Y)

    '                    Dim intposbtnsearchname = ((Me.grpcontactgrid.Width / 2) + (Me.btnsearchname.Width / 2) - 60) 'search name
    '                    Me.btnsearchname.Location = New System.Drawing.Point(intposbtnsearchname, ptbtnsearchname.Y)
    '                    '---------------------------------------------------------
    '                    'code for resizing leads----------------------------------
    '                    Me.pnlleads.Width = Me.tbcHome.Width - 10
    '                    Me.grpLeadssearch.Width = Me.pnlleads.Width - 30
    '                    Me.grpleadsgrid.Width = Me.pnlleads.Width - 30
    '                    Me.dtgLeads.Width = grpleadsgrid.Width - 30
    '                    Me.pnlleads.Height = Me.Height - 56
    '                    grpleadsgrid.Height = _
    '                                        pnlleads.Height - (grpLeadssearch.Height)
    '                    dtgLeads.Height = grpleadsgrid.Height - 50
    '                    addcustleadstablestyle(ptname)
    '                    'reposition controls on the leads
    '                    Dim fff = ((grpLeadssearch.Width / 4))
    '                    Dim intposlblname = ((grpLeadssearch.Width / 2) _
    '                    - ((grpLeadssearch.Width / 4)))
    '                    Me.lblleadname.Location = New System.Drawing.Point(intposlblname, ptlblname.Y)
    '                    lblleadname.Width = fff
    '                    Me.txtleads.Location = New System.Drawing.Point(lblleadname.Location.X, lblleadname.Location.Y + lblleadname.Height)
    '                    txtleads.Width = fff

    '                    Me.lblleadstatus.Location = New System.Drawing.Point(grpLeadssearch.Location.X, ptlblstatus.Y)
    '                    lblleadstatus.Width = fff
    '                    Me.cboleads.Location = New System.Drawing.Point(lblleadstatus.Location.X, lblleadstatus.Location.Y + lblleadstatus.Height)
    '                    cboleads.Width = fff

    '                    Dim intposlblsdate = ((grpLeadssearch.Width / 2))
    '                    Me.lblleadstartdate.Location = New System.Drawing.Point(intposlblsdate, ptlblstartdate.Y)
    '                    lblleadstartdate.Width = fff
    '                    Me.dtpsdate.Location = New System.Drawing.Point(lblleadstartdate.Location.X, lblleadstartdate.Location.Y + lblleadstartdate.Height)
    '                    dtpsdate.Width = fff

    '                    Me.lblleadenddate.Location = New System.Drawing.Point(intposlblsdate + fff, ptlblenddate.Y)
    '                    lblleadenddate.Width = fff
    '                    Me.dtpedate.Location = New System.Drawing.Point(lblleadenddate.Location.X, lblleadenddate.Location.Y + lblleadenddate.Height)
    '                    dtpedate.Width = fff

    '                    Dim intposlblleadssearch = ((grpLeadssearch.Width / 2)) - (btnleadssearch.Width / 2)
    '                    Me.btnleadssearch.Location = New System.Drawing.Point(intposlblleadssearch, ptbtnleadssearch.Y)
    '                    '---------------------------------------------------------------
    '                    'code for resizing jobs
    '                    Me.pnljobs.Width = Me.tbcHome.Width
    '                    Me.grpjobs.Width = Me.tbcHome.Width - 10
    '                    Me.grpjobsearch.Width = Me.tbcHome.Width - 10
    '                    Me.dtgJobs.Width = grpjobs.Width - 50
    '                    pnljobs.Height = Me.Height - 10
    '                    Me.grpjobs.Height = pnljobs.Height - (grpjobsearch.Height + 80)
    '                    dtgJobs.Height = grpjobs.Height - 50
    '                    addcurrjobtablestyle(jname)
    '                    Dim intposlbljobstatus As Integer = (Me.grpjobsearch.Width / 2) - (Me.lbljobstatus.Width)
    '                    Me.lbljobstatus.Location = New System.Drawing.Point(intposlbljobstatus, ptlbljobstatus.Y)

    '                    Dim intposlbljobscontactname As Integer = (Me.grpjobsearch.Width / 2) + 0
    '                    Me.lbljobcontactname.Location = New System.Drawing.Point(intposlbljobscontactname, ptlbljobcontactname.Y)

    '                    Dim intposcbojobstatus As Integer = (Me.grpjobsearch.Width / 2) - cbojobstatus.Width
    '                    Me.cbojobstatus.Location = New System.Drawing.Point(intposcbojobstatus, ptcbojobstatus.Y)

    '                    Dim intpostxtjobcontactname As Integer = (Me.grpjobsearch.Width / 2) - 0
    '                    Me.txtjobcontactname.Location = New System.Drawing.Point(intpostxtjobcontactname, ptltxtjobcontactname.Y)

    '                    Dim intposbtnjobsearch As Integer = (Me.grpjobsearch.Width / 2) - (btnjobsearch.Width / 2)
    '                    Me.btnjobsearch.Location = New System.Drawing.Point(intposbtnjobsearch, ptbtnjobsearch.Y)
    '                    '----------------------------------------------------------
    '                Case FormWindowState.Normal

    '                    Me.tbcHome.Width = Me.Width - 10
    '                    Me.tbcHome.Height = Me.Height - 10
    '                    'code for resizing contacts-----------------------
    '                    Me.pnlContacts.Width = Me.tbcHome.Width - 10
    '                    Me.grpcontactgrid.Width = Me.pnlContacts.Width - 10
    '                    Me.grpContactsearch.Width = Me.pnlContacts.Width - 10
    '                    Me.dtgClients.Width = grpcontactgrid.Width - 30
    '                    Me.pnlContacts.Height = Me.Height - 56
    '                    grpcontactgrid.Height = _
    '                    pnlContacts.Height - (grpContactsearch.Height)
    '                    dtgClients.Height = grpcontactgrid.Height - 50

    '                    'reposition controls on the contacts search grid
    '                    If Not (mypoint.X = 0 And mypoint.Y = 0) Then
    '                        Me.lblContactsearch.Location = New System.Drawing.Point(mypoint.X, mypoint.Y)
    '                    Else
    '                        mypoint = Me.lblContactsearch.Location
    '                    End If
    '                    If Not (pttxtparams.X = 0 And pttxtparams.Y = 0) Then
    '                        Me.txtparams.Location = New System.Drawing.Point(pttxtparams.X, pttxtparams.Y)
    '                    Else
    '                        pttxtparams = Me.txtparams.Location
    '                    End If
    '                    If Not (ptbtnsearchno.X = 0 And ptbtnsearchno.Y = 0) Then
    '                        Me.btnsearchno.Location = New System.Drawing.Point(ptbtnsearchno.X, ptbtnsearchno.Y)
    '                    Else
    '                        ptbtnsearchno = Me.btnsearchno.Location
    '                    End If
    '                    If Not (ptbtnsearchname.X = 0 And ptbtnsearchname.Y = 0) Then
    '                        Me.btnsearchname.Location = New System.Drawing.Point(ptbtnsearchname.X, ptbtnsearchname.Y)
    '                    Else
    '                        ptbtnsearchname = Me.btnsearchname.Location
    '                    End If



    '                    'Me.lblContactsearch.Location = New System.Drawing.Point(intposlblsearch, mypoint.Y)
    '                    '---------------------------------------------------------

    '                    'code for resizing leads----------------------------------
    '                    Me.pnlleads.Width = Me.tbcHome.Width - 10
    '                    Me.grpLeadssearch.Width = Me.pnlleads.Width - 30
    '                    Me.grpleadsgrid.Width = Me.pnlleads.Width - 30
    '                    Me.dtgLeads.Width = grpleadsgrid.Width - 30
    '                    Me.pnlleads.Height = Me.Height - 56
    '                    grpleadsgrid.Height = _
    '                                        pnlleads.Height - (grpLeadssearch.Height)
    '                    dtgLeads.Height = grpleadsgrid.Height - 50
    '                    addcustleadstablestyle(ptname)
    '                    'reposition controls on the leads page

    '                    If Not (ptlblstatus.X = 0 And ptlblstatus.Y = 0) Then
    '                        Me.lblleadstatus.Location = New System.Drawing.Point(ptlblstatus.X, ptlblstatus.Y)
    '                        Me.lblleadstatus.Width = szlblstatus.Width

    '                    Else
    '                        ptlblstatus = Me.lblleadstatus.Location
    '                        szlblstatus = Me.lblleadstatus.Size
    '                    End If
    '                    If Not (ptlblname.X = 0 And ptlblname.Y = 0) Then
    '                        Me.lblleadname.Location = New System.Drawing.Point(ptlblname.X, ptlblname.Y)
    '                        lblleadname.Width = szlblname.Width
    '                    Else
    '                        ptlblname = Me.lblleadname.Location
    '                        szlblname = Me.lblleadname.Size
    '                    End If
    '                    If Not (ptlblstartdate.X = 0 And ptlblstartdate.Y = 0) Then
    '                        lblleadstartdate.Location = New System.Drawing.Point(ptlblstartdate.X, ptlblstartdate.Y)
    '                        lblleadstartdate.Width = szlblstartdate.Width
    '                    Else
    '                        ptlblstartdate = lblleadstartdate.Location
    '                        szlblstartdate = lblleadstartdate.Size
    '                    End If
    '                    If Not (ptlblenddate.X = 0 And ptlblenddate.Y = 0) Then
    '                        Me.lblleadenddate.Location = New System.Drawing.Point(ptlblenddate.X, ptlblenddate.Y)
    '                        lblleadenddate.Width = szlblenddate.Width
    '                    Else
    '                        ptlblenddate = Me.lblleadenddate.Location
    '                        szlblenddate = Me.lblleadenddate.Size
    '                    End If
    '                    '///////////////
    '                    If Not (ptcbostatus.X = 0 And ptcbostatus.Y = 0) Then
    '                        Me.cboleads.Location = New System.Drawing.Point(ptcbostatus.X, ptcbostatus.Y)
    '                        cboleads.Width = szcbostatus.Width
    '                    Else
    '                        ptcbostatus = Me.cboleads.Location
    '                        szcbostatus = Me.cboleads.Size
    '                    End If
    '                    If Not (pttxtname.X = 0 And pttxtname.Y = 0) Then
    '                        Me.txtleads.Location = New System.Drawing.Point(pttxtname.X, pttxtname.Y)
    '                        txtleads.Width = sztxtname.Width
    '                    Else
    '                        pttxtname = Me.txtleads.Location
    '                        sztxtname = Me.txtleads.Size
    '                    End If
    '                    If Not (ptdtpstartdate.X = 0 And ptdtpstartdate.Y = 0) Then
    '                        Me.dtpsdate.Location = New System.Drawing.Point(ptdtpstartdate.X, ptdtpstartdate.Y)
    '                        dtpsdate.Width = szdtpstartdate.Width
    '                    Else
    '                        ptdtpstartdate = Me.dtpsdate.Location
    '                        szdtpstartdate = Me.dtpsdate.Size
    '                    End If
    '                    If Not (Me.ptdtpenddate.X = 0 And ptdtpenddate.Y = 0) Then
    '                        Me.dtpedate.Location = New System.Drawing.Point(ptdtpenddate.X, ptdtpenddate.Y)
    '                        dtpedate.Width = szdtpenddate.Width
    '                    Else
    '                        ptdtpenddate = Me.dtpedate.Location
    '                        szdtpenddate = Me.dtpedate.Size
    '                    End If

    '                    If Not (Me.ptbtnleadssearch.X = 0 And ptbtnleadssearch.Y = 0) Then
    '                        Me.btnleadssearch.Location = New System.Drawing.Point(ptbtnleadssearch.X, ptbtnleadssearch.Y)
    '                    Else
    '                        ptbtnleadssearch = Me.btnleadssearch.Location

    '                    End If
    '                    '------------------------------------------------------------

    '                    '---------------------------------------------------------------
    '                    'code for resizing jobs
    '                    Me.pnljobs.Width = Me.tbcHome.Width
    '                    Me.grpjobs.Width = Me.tbcHome.Width - 10
    '                    Me.grpjobsearch.Width = Me.tbcHome.Width - 10
    '                    Me.dtgJobs.Width = grpjobs.Width - 50
    '                    pnljobs.Height = Me.Height - 10
    '                    Me.grpjobs.Height = pnljobs.Height - (grpjobsearch.Height + 80)
    '                    dtgJobs.Height = grpjobs.Height - 50
    '                    addcurrjobtablestyle(jname)
    '                    If Not (ptlbljobcontactname.X = 0 And ptlbljobcontactname.Y = 0) Then
    '                        Me.lbljobcontactname.Location = New System.Drawing.Point(ptlbljobcontactname.X, ptlbljobcontactname.Y)
    '                        lbljobcontactname.Width = szcbostatus.Width
    '                    Else
    '                        ptlbljobcontactname = Me.lbljobcontactname.Location
    '                        szlbljobcontactname = Me.lbljobcontactname.Size
    '                    End If
    '                    If Not (ptlbljobstatus.X = 0 And ptlbljobstatus.Y = 0) Then
    '                        Me.lbljobstatus.Location = New System.Drawing.Point(ptlbljobstatus.X, ptlbljobstatus.Y)
    '                        lbljobstatus.Width = szlbljobstatus.Width
    '                    Else
    '                        ptlbljobstatus = Me.lbljobstatus.Location
    '                        szlbljobstatus = Me.lbljobstatus.Size
    '                    End If
    '                    If Not (ptcbojobstatus.X = 0 And ptcbojobstatus.Y = 0) Then
    '                        Me.cbojobstatus.Location = New System.Drawing.Point(ptcbojobstatus.X, ptcbojobstatus.Y)
    '                        cbojobstatus.Width = szcbojobstatus.Width
    '                    Else
    '                        ptcbojobstatus = Me.cbojobstatus.Location
    '                        szcbojobstatus = Me.cbojobstatus.Size
    '                    End If
    '                    If Not (Me.ptltxtjobcontactname.X = 0 And ptltxtjobcontactname.Y = 0) Then
    '                        Me.txtjobcontactname.Location = New System.Drawing.Point(ptltxtjobcontactname.X, ptltxtjobcontactname.Y)
    '                        txtjobcontactname.Width = sztxtjobcontactname.Width
    '                    Else
    '                        ptltxtjobcontactname = Me.txtjobcontactname.Location
    '                        sztxtjobcontactname = Me.txtjobcontactname.Size
    '                    End If

    '                    If Not (Me.ptbtnjobsearch.X = 0 And ptbtnjobsearch.Y = 0) Then
    '                        Me.btnjobsearch.Location = New System.Drawing.Point(ptbtnjobsearch.X, ptbtnjobsearch.Y)
    '                    Else
    '                        ptbtnjobsearch = Me.btnjobsearch.Location

    '                    End If

    '                    '----------------------------------------------------------
    '            End Select

    '        Catch ex As Exception
    '            MsgBox(ex.Message.ToString())
    '        End Try
    '    End Sub
    '#End Region
    Private Sub tbcHome_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbcHome.SelectedIndexChanged
        Try
            Try
                'dtgClients.DataMember = Nothing
                'dtgClients.DataSource = Nothing
                'dtgLeads.DataMember = Nothing
                'dtgLeads.DataSource = Nothing
                'dtgJobs.DataMember = Nothing
                'dtgJobs.DataSource = Nothing
            Catch er As Exception
                MsgBox(er.Message.ToString() & vbCrLf _
          & er.InnerException().ToString() & vbCrLf _
          & er.StackTrace.ToString())
            End Try
            ' myForms.Main.Text = "Home"
            If tbcHome.SelectedTab Is tpgClients Then
                Try
                    If Threadclients Is Nothing = False Then
                        Try
                            Threadclients.Abort()
                        Catch we As Exception
                        End Try
                    End If
                    Threadclients = New System.Threading.Thread( _
                                                                       AddressOf loadgrid)
                    Threadclients.IsBackground = True
                    Threadclients.Start()
                    myForms.Main.Text = "Contacts"
                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                      & ev.InnerException().ToString() & vbCrLf _
                      & ev.StackTrace.ToString())
                End Try

            ElseIf tbcHome.SelectedTab Is tpgHome Then
                'Try
                '    myForms.Main.Text = "Home"
                'Catch ev As Exception
                '    MsgBox(ev.Message.ToString() & vbCrLf _
                '   & ev.InnerException().ToString() & vbCrLf _
                '   & ev.StackTrace.ToString())
                'End Try

            ElseIf tbcHome.SelectedTab Is tpgLeads Then

                Try
                    dtpedate.Value = Now
                    Me.cboleads.Items.Clear()
                    If Threadleads Is Nothing = False Then
                        Threadleads.Abort()
                    End If
                    Threadleads = New System.Threading.Thread( _
                                                               AddressOf loadleads)
                    Threadleads.IsBackground = True
                    Threadleads.Start()

                    myForms.Main.Text = "Leads"
                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                    & ev.InnerException().ToString() & vbCrLf _
                     & ev.StackTrace.ToString())
                End Try
            ElseIf tbcHome.SelectedTab Is tpgJobs Then
                Try
                    Me.cbojobstatus.Items.Clear()
                    cbojobstatus.Items.Add("Current")
                    cbojobstatus.Items.Add("complete")
                    cbojobstatus.Items.Add("delivered")
                    'dtgClients.SetDataBinding("", "")
                    'dtgLeads.SetDataBinding("", "")
                    myForms.Main.Text = "Jobs"
                Catch ev As Exception
                    MsgBox(ev.Message.ToString() & vbCrLf _
                    & ev.InnerException().ToString() & vbCrLf _
                    & ev.StackTrace.ToString())
                End Try
            ElseIf tbcHome.SelectedTab Is tpgPersonnel Then
                Try
                    myForms.Main.Text = "Personnel"
                    If myForms.timesheet Is Nothing Then
                        Call loadtimesheet()

                    End If

                Catch we As Exception

                End Try

            ElseIf tbcHome.SelectedTab Is tpgEquip Then
                Try
                    myForms.Main.Text = "Equipment"
                    If myForms.equipments Is Nothing Then
                        Call loadequip()
                    End If

                Catch we As Exception

                End Try


            End If

        Catch ex As Exception
            MsgBox(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Function canmanipulateequip() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strleadss.Split(",")
            If arr(1) = "1" Then
                canmanipulateequip = True
            Else
                canmanipulateequip = False
            End If
        Catch ex As Exception
            Try
                canmanipulateequip = False
            Catch exc As Exception

            End Try
        End Try
    End Function
    Public Function canmanipulatepersonnel() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strpersonnelss.Split(",")
            If arr(1) = "1" Then
                canmanipulatepersonnel = True
            Else
                canmanipulatepersonnel = False
            End If
        Catch ex As Exception
            Try
                canmanipulatepersonnel = False
            Catch exc As Exception

            End Try
        End Try
    End Function
    Public Function canmanipulateit() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strpersonnelss.Split(",")
            If arr(3) = "1" Then
                canmanipulateit = True
            Else
                canmanipulateit = False
            End If
        Catch ex As Exception
            Try
                canmanipulateit = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "menu bar"
    Private Sub mnufileclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnufileclose.Click
        Try
            isloading = False
            closeprogram()
            Me.Dispose(True)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub mnueditcontacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnueditcontacts.Click
        Try

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub mnueditleads_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnueditleads.Click
        Try

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub mnueditjobs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnueditjobs.Click
        Try

        Catch ex As Exception

        End Try
    End Sub
    Private Sub mnuviewhome_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuviewhome.Click
        Try
            Me.tbcHome.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub mnuviewleads_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuviewleads.Click
        Try
            Me.tbcHome.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub mnuviewcontacts_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuviewcontacts.Click
        Try
            Me.tbcHome.SelectedIndex = 2
        Catch ex As Exception

        End Try
    End Sub
    Private Sub mnuviewjobs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuviewjobs.Click
        Try
            Me.tbcHome.SelectedIndex = 3
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub mnuviewequip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuviewequip.Click
        Try
            Me.tbcHome.SelectedIndex = 4
        Catch ex As Exception

        End Try
    End Sub
    Private Sub mnuviewpersonnel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuviewpersonnel.Click
        Try
            Me.tbcHome.SelectedIndex = 5
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub mnuReports_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuReports.Click
        Try
            'If frmReports.isrunning = False Then
            '    frmReports.Main()
            'End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub mnufilesettingsadministrator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnufilesettingsadministrator.Click

    End Sub
    Private Sub mnutodaysdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim connectstr As String
         connectstr = "DSN=" & myForms.qconnstr
        'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection()
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = connectstr
        connect.Open()
        Try
            Dim dtpp As New System.Windows.Forms.DateTimePicker()
            dtpp.Value = Now
            Dim sdate2 As String
            sdate2 = dtpp.Value.Year & "-" _
                        & dtpp.Value.Month & "-" _
                        & dtpp.Value.Day & " " _
                        & "00" & ":" _
                        & "00" & ":" _
                        & "00"
            Dim strsql As String
            strsql = " delete from storedate;"
            strsql += " insert into storedate (curdate) values ('" & sdate2 & "');"
            Try
                connect.BeginTrans()
                connect.Execute(strsql)
                connect.CommitTrans()
            Catch er As Exception

            End Try
        Catch we As Exception

        End Try
        Try
            connect.Close()
        Catch qw As Exception

        End Try
    End Sub
    Private Sub mnucur_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnucur.Click 'confer user rights
        Try
            Dim gr As New frmadmin()
            myForms.admin = gr
            myForms.admin.Show()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub mnutss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnutss.Click 'update time sheet settings
        Try
            Dim grv As New frmtimesheetmanager()
            grv.strpersonnel = myForms.Main.strpersonnelss
            grv.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "support code"
    'Private Sub LoadFormIntoPanel()
    '    Dim MyForm2 As New Form2()
    '    MyForm2.Size = pnlContent.Size
    '    MyForm2.TopLevel = False
    '    MyForm2.Parent = pnlContent
    '    MyForm2.Show()
    'End Sub
    'Class TasksClass
    '    Friend StrArg As String
    '    Friend RetVal As Boolean
    '    Sub SomeTask()
    '        ' Use the StrArg field as an argument.
    '        MsgBox("The StrArg contains the string " & StrArg)
    '        RetVal = True ' Set a return value in the return argument.
    '    End Sub
    'End Class
    ' To use the class, set the properties or fields that store parameters,
    ' and then asynchronously call the methods as needed.
    'Sub DoWork()
    '    Dim Tasks As New taskclass()
    '    Dim Thread1 As New System.Threading.Thread( _
    '        AddressOf taskclass.SomeTask)
    '    Tasks.StrArg = "Some Arg" ' Set a field that is used as an argument
    '    Thread1.Start() ' Start the new thread.
    '    Thread1.Join() ' Wait for thread 1 to finish.
    '    ' Display the return value.
    '    MsgBox("Thread 1 returned the value " & Tasks.RetVal)
    'End Sub
#End Region

#Region "error handling"
    Private Sub AddWithUnhandledException()
        ' txtName.Text = "Kevin"
        Throw New InvalidOperationException( _
            "Invalid operation.")
    End Sub
    Private Sub DisplayError(ByVal ex As Exception)
        MessageBox.Show(ex.GetType().ToString() & _
            vbCrLf & vbCrLf & _
            ex.Message & vbCrLf & vbCrLf & _
            ex.StackTrace, _
            "Error", _
            MessageBoxButtons.AbortRetryIgnore, _
            MessageBoxIcon.Stop)
    End Sub
    Friend Class ThreadExceptionHandler

        '''
        ''' Handles the thread exception.
        '''
        Public Sub Application_ThreadException( _
            ByVal sender As System.Object, _
            ByVal e As ThreadExceptionEventArgs)

            Try
                ' Exit the program if the user clicks Abort.
                Dim result As DialogResult = _
                    ShowThreadExceptionDialog(e.Exception)

                If (result = System.Windows.Forms.DialogResult.Abort) Then
                    Application.Exit()
                End If
            Catch
                ' Fatal error, terminate program
                Try
                    MessageBox.Show("Fatal Error", _
                        "Fatal Error", _
                        MessageBoxButtons.OK, _
                        MessageBoxIcon.Stop)
                Finally
                    Application.Exit()
                End Try
            End Try
        End Sub

        '''
        ''' Creates and displays the error message.
        '''
        Private Function ShowThreadExceptionDialog( _
            ByVal ex As Exception) As DialogResult

            Dim errorMessage As String = _
                "Unhandled Exception:" _
                & vbCrLf & vbCrLf & _
                ex.Message & vbCrLf & vbCrLf & _
                ex.GetType().ToString() & vbCrLf & vbCrLf & _
                "Stack Trace:" & vbCrLf & _
                ex.StackTrace

            Return MessageBox.Show(errorMessage, _
                "Application Error", _
                MessageBoxButtons.AbortRetryIgnore, _
                MessageBoxIcon.Stop)
        End Function

    End Class ' ThreadExceptionHandler
#End Region

#Region "Reports"

#Region "Private members"
    Private GridPrinter As DataGridPrinter
#End Region

#Region "rclients"
    Private Sub mnuallclients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuallclients.Click
        Try
            Dim x As Boolean = canviewclientsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view clients reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Call loadgridallclients()
            GridPrinter = New DataGridPrinter(Me.dtgClients)
            Dim ds As New System.Data.DataSet()
            Dim ds1 As New System.Data.DataSet()
            Dim tbl As New DataTable()
            ds1 = dtgClients.DataSource
            ds = ds1.Copy
            '---------------------report options
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export
                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            ds.Tables(0).Columns("client_no").ColumnName = "Client Number"
                            ds.Tables(0).Columns("name").ColumnName = "Name"
                            ds.Tables(0).Columns("description").ColumnName = "Description"
                            ds.Tables(0).Columns("oclient_no").ColumnName = "Old Client Number"
                            ds.Tables(0).Columns.Remove("ano")
                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(ds, m)
                        ds1 = Me.dtgClients.DataSource
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        loadgridallclients()
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = "ALL CLIENTS"

                With GridPrinter
                    .HeaderText = Me.TextBox2.Text

                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = CInt(Me.NumericUpDown_PagesAcross.Value)


                End With
                'With GridPrinter
                '    .HeaderText = "Some text"
                '    .HeaderHeightPercent = 10
                '    .FooterHeightPercent = 5
                '    .InterSectionSpacingPercent = 2
                '    .PagesAcross = 1
                '    '\\ Set any other properties to 
                '    'affect the look of the grid...
                'End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If



        Catch ex As Exception

        End Try

    End Sub
    Private Sub mnucurrentclients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnucurrentclients.Click
        Try
            Dim x As Boolean = canviewclientsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view clients reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            Dim ds1 As System.Data.DataSet = New System.Data.DataSet()
            ds1 = Me.dtgClients.DataSource
            ds = ds1.Copy
            If ds Is Nothing Then
                MessageBox.Show("No data is available", "Current jobs" _
                , MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            '---------------------report options
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export

                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            ds.Tables(0).Columns("client_no").ColumnName = "Client Number"
                            ds.Tables(0).Columns("name").ColumnName = "Name"
                            ds.Tables(0).Columns("description").ColumnName = "Description"
                            ds.Tables(0).Columns("oclient_no").ColumnName = "Old Client Number"
                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(ds, m)
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ds.Tables(0).Columns("Client Number").ColumnName = "client_no"
                        ds.Tables(0).Columns("Name").ColumnName = "name"
                        ds.Tables(0).Columns("Description").ColumnName = "description"
                        ds.Tables(0).Columns("Old Client Number").ColumnName = "oclient_no"
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print

                dtgClients.DataSource = Nothing
                dtgClients.DataSource = MyTable
                GridPrinter = New DataGridPrinter(Me.dtgClients)
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = "CURRENT CLIENTS"
                With GridPrinter
                    .HeaderText = Me.TextBox2.Text

                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = CInt(Me.NumericUpDown_PagesAcross.Value)
                End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If

        Catch xc As Exception
        End Try

        Try
            GridPrinter = Nothing
        Catch zx As Exception
        End Try
    End Sub
    Private Sub loadgridallclients()
        Dim cnnstr As String
         cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection()
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Dim currentCursor As Cursor = Cursor.Current
        Try
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim custDS As DataSet = New DataSet()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select *  from clients order by client_no"

            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "clients")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgClients.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            Call AddCustomDataTableStyle()
            '---------------remove some rows
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            ds = Me.dtgClients.DataSource
            ds.Tables(0).Columns.Remove("leads_no")
            ds.Tables(0).Columns.Remove("least_status")
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            dtgClients.DataSource = ds
            '------------
        Catch t As Exception

        Finally
            'statusBar1.Text = "Done"
            Cursor.Current = currentCursor


        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Function canviewclientsreport() As Boolean
        Try
            Dim arr() As String
            arr = strclientss.Split(",")
            If arr(2) = "1" Then
                canviewclientsreport = True
            Else
                canviewclientsreport = False
            End If
        Catch ex As Exception
            Try
                canviewclientsreport = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "jobs"
    Private Sub mnucurrent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnucurrent.Click
        Try
            Dim x As Boolean = canviewjobsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view job reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            dgrid.DataSource = Nothing
            dgrid.DataSource = dscurrentjobs()
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            ds = dgrid.DataSource
            dgrid.DataSource = Nothing
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            dgrid.DataSource = MyTable
            '---------------------report options
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export
                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            'ds.Tables(0).Columns("client_no").ColumnName = "Client Number"
                            'ds.Tables(0).Columns("name").ColumnName = "Name"
                            'ds.Tables(0).Columns("description").ColumnName = "Description"
                            'ds.Tables(0).Columns("oclient_no").ColumnName = "Old Client Number"
                            'ds.Tables(0).Columns.Remove("ano")
                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(ds, m)
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print
                GridPrinter = New DataGridPrinter(dgrid)
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = "CURRENT JOBS"
                With GridPrinter
                    .HeaderText = Me.TextBox2.Text
                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = 1


                End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If

        Catch ex As Exception

        End Try
        myForms.Main.Invalidate()
    End Sub
    Private Function dscurrentjobs() As System.Data.DataSet
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()
        Catch sd As Exception
            Exit Function
        End Try
        Dim custDS As DataSet = New DataSet()
        Try
            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle, rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin" _
            & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
            & " lower(rcljobs.job_status) like" _
            & "'%" & LCase("current") & "%' order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)
            custDA.Fill(custDS, adors, "leads20")
            Dim tname As String = custDS.Tables(0).TableName()
            custDS.Tables(0).Columns(0).ColumnName = "Client number"
            custDS.Tables(0).Columns(1).ColumnName = "Name"
            custDS.Tables(0).Columns(2).ColumnName = "Job number"
            custDS.Tables(0).Columns(3).ColumnName = "Job title"
            custDS.Tables(0).Columns(4).ColumnName = "Technician responsible"
            custDS.Tables(0).Columns(5).ColumnName = "Amount"
            custDS.Tables(0).Columns(6).ColumnName = "Gross margin"

            Me.dgrid.SetDataBinding(custDS, tname)


            '---------format width

            connect.Close()
            Return custDS
        Catch xc As Exception
            Return custDS
        Finally

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Sub mnucompletedjobs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnucompletedjobs.Click
        Try
            Dim x As Boolean = canviewjobsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view job reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            dgrid.DataSource = Nothing
            dgrid.DataSource = dscompletedjobs()
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            ds = dgrid.DataSource
            dgrid.DataSource = Nothing
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            dgrid.DataSource = MyTable
            '---------------------report options
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export
                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            'ds.Tables(0).Columns("client_no").ColumnName = "Client Number"
                            'ds.Tables(0).Columns("name").ColumnName = "Name"
                            'ds.Tables(0).Columns("description").ColumnName = "Description"
                            'ds.Tables(0).Columns("oclient_no").ColumnName = "Old Client Number"
                            'ds.Tables(0).Columns.Remove("ano")
                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(ds, m)
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print
                GridPrinter = New DataGridPrinter(dgrid)
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = "COMPLETED JOBS"
                With GridPrinter
                    .HeaderText = Me.TextBox2.Text
                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = 1


                End With
                'With GridPrinter
                '    .HeaderText = "Some text"
                '    .HeaderHeightPercent = 10
                '    .FooterHeightPercent = 5
                '    .InterSectionSpacingPercent = 2
                '    .PagesAcross = 1
                '    '\\ Set any other properties to 
                '    'affect the look of the grid...
                'End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If

        Catch ex As Exception

        End Try
        myForms.Main.Invalidate()
    End Sub
    Private Function dscompletedjobs() As System.Data.DataSet
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()
        Catch sd As Exception
            Exit Function
        End Try

        Dim custDS As DataSet = New DataSet()
        Try

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()

            Dim adors As New ADODB.Recordset()
            Dim str As String = "select rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin" _
            & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
            & " lower(rcljobs.job_status) like" _
            & "'%" & LCase("complete") & "%' order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)
            custDA.Fill(custDS, adors, "complete")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dgrid.SetDataBinding(custDS, tname)
            custDS.Tables(0).Columns(0).ColumnName = "Client number"
            custDS.Tables(0).Columns(1).ColumnName = "Name"
            custDS.Tables(0).Columns(2).ColumnName = "Job number"
            custDS.Tables(0).Columns(3).ColumnName = "Job title"
            custDS.Tables(0).Columns(4).ColumnName = "Technician responsible"
            custDS.Tables(0).Columns(5).ColumnName = "Amount"
            custDS.Tables(0).Columns(6).ColumnName = "Gross margin"
            connect.Close()
            Return custDS
        Catch xc As Exception
            Return custDS
        Finally

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Sub mnudeliveredjobs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnudeliveredjobs.Click
        Try
            Dim x As Boolean = canviewjobsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view job reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            dgrid.DataSource = Nothing
            dgrid.DataSource = dsdeliveredjobs()
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            ds = dgrid.DataSource
            dgrid.DataSource = Nothing
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            dgrid.DataSource = MyTable
            '---------------------report options
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export
                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            'ds.Tables(0).Columns("client_no").ColumnName = "Client Number"
                            'ds.Tables(0).Columns("name").ColumnName = "Name"
                            'ds.Tables(0).Columns("description").ColumnName = "Description"
                            'ds.Tables(0).Columns("oclient_no").ColumnName = "Old Client Number"
                            'ds.Tables(0).Columns.Remove("ano")
                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(ds, m)
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print
                GridPrinter = New DataGridPrinter(dgrid)
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = "DELIVERED JOBS"
                With GridPrinter
                    .HeaderText = Me.TextBox2.Text
                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = 1


                End With
                'With GridPrinter
                '    .HeaderText = "Some text"
                '    .HeaderHeightPercent = 10
                '    .FooterHeightPercent = 5
                '    .InterSectionSpacingPercent = 2
                '    .PagesAcross = 1
                '    '\\ Set any other properties to 
                '    'affect the look of the grid...
                'End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If

        Catch ex As Exception

        End Try
        myForms.Main.Invalidate()
    End Sub
    Private Function dsdeliveredjobs() As System.Data.DataSet
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()
        Catch sd As Exception
            Exit Function
        End Try

        Dim custDS As DataSet = New DataSet()
        Try

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()

            Dim adors As New ADODB.Recordset()
            Dim str As String = "select rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle, rcljobs.techres," _
            & " rcljobs.amount,rcljobs.grossmargin" _
            & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
            & " lower(rcljobs.job_status) like" _
            & "'%" & LCase("delivered") & "%' order by rcljobs.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)
            custDA.Fill(custDS, adors, "complete")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dgrid.SetDataBinding(custDS, tname)
            custDS.Tables(0).Columns(0).ColumnName = "Client number"
            custDS.Tables(0).Columns(1).ColumnName = "Name"
            custDS.Tables(0).Columns(2).ColumnName = "Job number"
            custDS.Tables(0).Columns(3).ColumnName = "Job title"
            custDS.Tables(0).Columns(4).ColumnName = "Technician responsible"
            custDS.Tables(0).Columns(5).ColumnName = "Amount"
            custDS.Tables(0).Columns(6).ColumnName = "Gross margin"
            connect.Close()
            Return custDS
        Catch xc As Exception
            Return custDS
        Finally

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Sub mnucurrentviewofjobs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnucurrentviewofjobs.Click
        Try
            Dim x As Boolean = canviewjobsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view job reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim dv2 As DataViewManager = New DataViewManager()
            Dim dv As DataViewManager = New DataViewManager()

            dv2 = dtgJobs.DataSource
            dv.DataSet = dv2.DataSet.Copy
            If dv.DataSet Is Nothing Then
                MessageBox.Show("No data is available", "Current jobs" _
                , MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Try
                dv.DataSet.Tables(0).Columns.Remove("costagreed")
                dv.DataSet.Tables(0).Columns.Remove("ojob_no")
                dv.DataSet.Tables(0).Columns.Remove("descrip")
                dv.DataSet.Tables(0).Columns.Remove("date_sniffed")
                dv.DataSet.Tables(0).Columns.Remove("journal")
                dv.DataSet.Tables(0).Columns.Remove("department")
                dv.DataSet.Tables(0).Columns.Remove("budgetarycost")

            Catch xcvb As Exception

            End Try
            'Dim MyTable As New DataTable()
            'MyTable = dv.DataSet.Tables(0)
            'dtgJobs.DataSource = Nothing
            'dtgJobs.DataSource = MyTable

            '---------------------report options
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export
                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            dv.DataSet.Tables(0).Columns("client_no").ColumnName = "Client Number"
                            dv.DataSet.Tables(0).Columns("name").ColumnName = "Name"
                            dv.DataSet.Tables(0).Columns("job_no").ColumnName = "Job Number"
                            dv.DataSet.Tables(0).Columns("job_tittle").ColumnName = "Job Title"
                            dv.DataSet.Tables(0).Columns("job_status").ColumnName = "Job Status"
                            dv.DataSet.Tables(0).Columns("techres").ColumnName = "Technician Responsible"
                            dv.DataSet.Tables(0).Columns("amount").ColumnName = "Amount"
                            dv.DataSet.Tables(0).Columns("grossmargin").ColumnName = "Gross Margin"
                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(dv.DataSet, m)
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Try
                            Dim tjd As Thread = New System.Threading.Thread( _
                            AddressOf jseinvoke)
                            tjd.IsBackground = True
                            tjd.Start()
                        Catch xa As Exception

                        End Try
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print
                myForms.Main.dgrid.DataSource = dv.DataSet.Tables(0)
                GridPrinter = New DataGridPrinter(myForms.Main.dgrid)
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = "CURRENT JOBS"
                With GridPrinter
                    .HeaderText = Me.TextBox2.Text

                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = CInt(Me.NumericUpDown_PagesAcross.Value)
                End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If

        Catch xc As Exception
        End Try

        Try
            GridPrinter = Nothing
        Catch zx As Exception
        End Try
        myForms.Main.Invalidate()
    End Sub
    Private Function canviewjobsreport() As Boolean
        Try
            Dim arr() As String
            arr = strjobss.Split(",")
            If arr(2) = "1" Then
                canviewjobsreport = True
            Else
                canviewjobsreport = False
            End If
        Catch ex As Exception
            Try
                canviewjobsreport = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "leads"
    Private strtitle As String
    Private Function dssuspectleads(ByVal strr As String) As System.Data.DataSet
        Dim connect As New ADODB.Connection()
        Try
            Dim cnnstr As String
            cnnstr= "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()
        Catch sd As Exception
            Exit Function
        End Try
        Dim custDS As DataSet = New DataSet()
        Try

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
            Dim adors As New ADODB.Recordset()
            Dim str As String = "select " _
                        & "leads.leads_no,clients.name,leads.title," _
                        & "leads.date_sniffed,leads.amount,leads. department" _
                        & " from leads inner join clients on leads.client_no = clients.client_no and " _
                        & " lower(leads.status) like " _
                        & "'%" & strr & "%'" _
                        & " order by leads.client_no"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)
            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dgrid.SetDataBinding(custDS, tname)
            custDS.Tables(0).Columns(0).ColumnName = "Lead number"
            custDS.Tables(0).Columns(1).ColumnName = "Name"
            custDS.Tables(0).Columns(2).ColumnName = "Lead Title"
            custDS.Tables(0).Columns(3).ColumnName = "Date"
            custDS.Tables(0).Columns(4).ColumnName = "Amount"
            custDS.Tables(0).Columns(5).ColumnName = "Department"

            '---------format width

            connect.Close()
            Return custDS
        Catch xc As Exception
            Return custDS
        Finally

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Sub loadview(ByVal mystr As String)
        Try
            dgrid.DataSource = Nothing
            dgrid.DataSource = dssuspectleads(mystr)
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            ds = dgrid.DataSource
            dgrid.DataSource = Nothing
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            dgrid.DataSource = MyTable

            '---------------------report options
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export
                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            'ds.Tables(0).Columns("client_no").ColumnName = "Client Number"
                            'ds.Tables(0).Columns("name").ColumnName = "Name"
                            'ds.Tables(0).Columns("description").ColumnName = "Description"
                            'ds.Tables(0).Columns("oclient_no").ColumnName = "Old Client Number"
                            'ds.Tables(0).Columns.Remove("ano")
                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(ds, m)
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print
                GridPrinter = New DataGridPrinter(dgrid)
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = strtitle
                With GridPrinter
                    .HeaderText = Me.TextBox2.Text
                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = 1


                End With
                'With GridPrinter
                '    .HeaderText = "Some text"
                '    .HeaderHeightPercent = 10
                '    .FooterHeightPercent = 5
                '    .InterSectionSpacingPercent = 2
                '    .PagesAcross = 1
                '    '\\ Set any other properties to 
                '    'affect the look of the grid...
                'End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If

        Catch ex As Exception

        End Try
        myForms.Main.Invalidate()
    End Sub
    Private Sub mnuleadssuspect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuleadssuspect.Click
        Try
            Dim x As Boolean = canviewleadsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print leads or view reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            strtitle = "SUSPECT LEADS"
            loadview("suspect")
        Catch xc As Exception
        End Try
    End Sub
    Private Sub mnuleadsprospect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuleadsprospect.Click
        Try
            Dim x As Boolean = canviewleadsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view leads reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            strtitle = "PROSPECT LEADS"
            loadview("prospect")
        Catch xc As Exception
        End Try
    End Sub
    Private Sub mnuleadsproposal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuleadsproposal.Click
        Try
            Dim x As Boolean = canviewleadsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print leads or view reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            strtitle = "PROPOSAL LEADS"
            loadview("proposal")
        Catch xc As Exception
        End Try
    End Sub
    Private Sub mnuleadspho_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuleadspho.Click
        Try
            Dim x As Boolean = canviewleadsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or or view view leads reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            strtitle = "PHO LEADS"
            loadview("pho")
        Catch xc As Exception
        End Try
    End Sub
    Private Sub mnuleadsfailed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuleadsfailed.Click
        Try
            Dim x As Boolean = canviewleadsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view leads reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            strtitle = "FAILED LEADS"
            loadview("failed")
        Catch xc As Exception
        End Try
    End Sub
    Private Sub mnuleadscurrent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuleadscurrent.Click
        Try
            Dim x As Boolean = canviewleadsreport()
            If x = False Then
                MessageBox.Show("Not allowed to print or view leads reports", "Reports", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            Dim ds1 As System.Data.DataSet = New System.Data.DataSet()
            ds1 = Me.dtgLeads.DataSource
            ds = ds1.Copy
            If ds Is Nothing Then
                MessageBox.Show("No data is available", "Current jobs" _
                , MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            'Dim MyTable As New DataTable()
            'MyTable = ds.Tables(0)
            'dtgLeads.DataSource = Nothing
            'dtgLeads.DataSource = MyTable
            Try
                ds.Tables(0).Columns.Remove("client_no")
                'dv.DataSet.Tables(0).Columns.Remove("ojob_no")
                'dv.DataSet.Tables(0).Columns.Remove("descrip")
                'dv.DataSet.Tables(0).Columns.Remove("date_sniffed")
                'dv.DataSet.Tables(0).Columns.Remove("journal")
                'dv.DataSet.Tables(0).Columns.Remove("department")
                'dv.DataSet.Tables(0).Columns.Remove("budgetarycost")

            Catch xcvb As Exception

            End Try
          
            '---------------------report options
            Dim gj As New frmreportoptions()
            gj.ShowDialog()
            If Me.reportoption = "0" Then
                '---------export
                Try
                    Dim sfd As System.Windows.Forms.SaveFileDialog _
                    = New System.Windows.Forms.SaveFileDialog()
                    sfd.Filter = "Excel files (*.xls)|*.xls"
                    sfd.CheckFileExists = False
                    sfd.CheckPathExists = True
                    sfd.ValidateNames = True
                    sfd.ShowDialog()
                    Dim m As String = sfd.FileName
                    If m.Trim.Length > 0 Then
                        Try
                            ds.Tables(0).Columns("leads_no").ColumnName = "Leads  Number"
                            ds.Tables(0).Columns("name").ColumnName = "Name"
                            ds.Tables(0).Columns("title").ColumnName = "Title"
                            ds.Tables(0).Columns("status").ColumnName = "Status"
                            ds.Tables(0).Columns("date_sniffed").ColumnName = "Date"
                            ds.Tables(0).Columns("amount").ColumnName = "Amount"
                            ds.Tables(0).Columns("department").ColumnName = "Department"
                            'dv.DataSet.Tables(0).Columns("grossmargin").ColumnName = "Gross Margin"


                        Catch cv As Exception
                        End Try
                        exporttoexcel.exportexcel.exportToExcel(ds, m)
                        MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Try
                            Dim tjd As Thread = New System.Threading.Thread( _
                            AddressOf jseinvoke)
                            tjd.IsBackground = True
                            tjd.Start()
                        Catch xa As Exception

                        End Try
                    Else
                        MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch we As Exception
                End Try
            Else
                '--------------------------print 
                myForms.Main.dgrid.DataSource = ds.Tables(0)
                GridPrinter = New DataGridPrinter(myForms.Main.dgrid)
                '--------------page set up
                Try
                    With Me.PageSetupDialog1
                        .Document = GridPrinter.PrintDocument
                        .ShowDialog()
                    End With
                Catch xc As Exception
                End Try

                '----------------
                TextBox2.Text = "CURRENT LEADS"
                With GridPrinter
                    .HeaderText = Me.TextBox2.Text

                    .HeaderHeightPercent = CInt(Me.NumericUpDown_HeaderHeightPercentage.Value)
                    .FooterHeightPercent = CInt(Me.NumericUpDown_FooterHeightPercent.Value)
                    .InterSectionSpacingPercent = CInt(Me.NumericUpDown_InterSectionSpacingPercent.Value)
                    .HeaderPen = New Pen(CType(Me.ComboBox_ColourHeaderLine.SelectedItem, System.Drawing.Color))
                    .FooterPen = New Pen(CType(Me.ComboBox_ColourFooterLine.SelectedItem, System.Drawing.Color))
                    .GridPen = New Pen(CType(Me.ComboBox_ColourBodyline.SelectedItem, System.Drawing.Color))
                    .HeaderBrush = CType(Me.ComboBox_HeaderBrush.SelectedItem, Brush)
                    .EvenRowBrush = CType(Me.ComboBox_EvenBrush.SelectedItem, Brush)
                    .OddRowBrush = CType(Me.ComboBox_OddRowBrush.SelectedItem, Brush)
                    .FooterBrush = CType(Me.ComboBox_FooterBrush.SelectedItem, Brush)
                    .ColumnHeaderBrush = CType(Me.ComboBox_ColumnHeaderBrush.SelectedItem, Brush)
                    .PagesAcross = CInt(Me.NumericUpDown_PagesAcross.Value)
                End With
                With Me.PrintPreviewDialog1
                    .Document = GridPrinter.PrintDocument
                    If .ShowDialog = DialogResult.OK Then
                        GridPrinter.Print()
                    End If
                End With
            End If

        Catch xc As Exception
        End Try

        Try
            GridPrinter = Nothing
        Catch zx As Exception
        End Try
        myForms.Main.Invalidate()
    End Sub
    Private Function canviewleadsreport() As Boolean
        Try
            Dim arr() As String
            arr = strleadss.Split(",")
            If arr(2) = "1" Then
                canviewleadsreport = True
            Else
                canviewleadsreport = False
            End If
        Catch ex As Exception
            Try
                canviewleadsreport = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "print controls"
    Private Sub ComboBox_EvenBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_EvenBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_EvenBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub ComboBox_OddRowBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_OddRowBrush.DrawItem

        e.Graphics.FillRectangle(CType(ComboBox_OddRowBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub ComboBox_HeaderBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_HeaderBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_HeaderBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub ComboBox_ColumnHeaderBrush_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_ColumnHeaderBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_ColumnHeaderBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub ComboBox_FooterBrush_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_FooterBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_FooterBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub NumericUpDown_PagesAcross_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub GroupBox1_Enter_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox_ColumnHeaderBrush_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub NumericUpDown_HeaderHeightPercentage_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox_EvenBrush_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox_HeaderBrush_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub NumericUpDown_FooterHeightPercent_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox_ColourHeaderLine_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox_ColourFooterLine_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Private Sub ComboBox_OddRowBrush_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ComboBox_ColourBodyline_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub NumericUpDown_InterSectionSpacingPercent_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
#End Region

#Region "Private methods"
    Private Sub PopulateColourList(ByVal combo As ComboBox)

        combo.Items.Clear()
        combo.Items.Add(System.Drawing.Color.AliceBlue)
        combo.Items.Add(System.Drawing.Color.Aqua)
        combo.Items.Add(System.Drawing.Color.Azure)
        combo.Items.Add(System.Drawing.Color.Beige)
        combo.Items.Add(System.Drawing.Color.Black)
        combo.Items.Add(System.Drawing.Color.Blue)
        combo.Items.Add(System.Drawing.Color.Green)
        combo.Items.Add(System.Drawing.Color.Red)
        combo.SelectedIndex = 0
    End Sub

    Private Sub PopulateBrushList(ByVal Combo As ComboBox)
        Combo.Items.Clear()
        Combo.Items.Add(System.Drawing.Brushes.White)
        Combo.Items.Add(System.Drawing.Brushes.Beige)
        Combo.Items.Add(System.Drawing.Brushes.Bisque)
        Combo.Items.Add(System.Drawing.Brushes.BlanchedAlmond)
        Combo.Items.Add(System.Drawing.Brushes.Blue)
        Combo.Items.Add(System.Drawing.Brushes.BlueViolet)
        Combo.Items.Add(System.Drawing.Brushes.Brown)
        Combo.Items.Add(System.Drawing.Brushes.BurlyWood)
        Combo.Items.Add(System.Drawing.Brushes.CadetBlue)
        Combo.Items.Add(System.Drawing.Brushes.Chartreuse)
        Combo.Items.Add(System.Drawing.Brushes.Chocolate)
        Combo.Items.Add(System.Drawing.Brushes.Coral)
        Combo.Items.Add(System.Drawing.Brushes.CornflowerBlue)
        Combo.Items.Add(System.Drawing.Brushes.Cornsilk)
        Combo.Items.Add(System.Drawing.Brushes.Crimson)
        Combo.Items.Add(System.Drawing.Brushes.Cyan)
        Combo.SelectedIndex = 0
    End Sub
#End Region

#End Region

#Region "validation"
    Private Sub txtleads_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtleads.KeyPress
        Try
            Dim vt As New validation()
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtleads, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtleads, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtparams_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtparams.KeyPress
        Try
            Dim vt As New validation()
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtparams, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtparams, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

#Region "appconfig"
    ' Windows-1252
#End Region

#Region "tab control properties"
    Private Sub DrawOnTab(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        Dim g As Graphics = e.Graphics
        Dim p As New Pen(Color.Blue)
        Dim font As New font("Arial", 6.0F)
        Dim brush As New SolidBrush(Color.Black)

        g.DrawRectangle(p, tabArea)
        g.DrawString("Home", font, brush, tabTextArea)

        '    'This line of code will help you to change the apperance like size,name,style.
        '    Dim f As font
        '    'For background color
        '    Dim backBrush As brush
        '    'For forground color
        '    Dim foreBrush As brush

        '    'This construct will hell you to deside which tab page have current focus
        '    'to change the style.
        'If e.Index = Me.tbcHome.SelectedIndex Then

        '    'This line of code will help you to change the apperance like size,name,style.
        '    'f = New Font(e.Font, FontStyle.Bold)
        '    f = New font(e.Font, FontStyle.Bold)

        '    backBrush = New System.Drawing.SolidBrush(Color.DarkGray)
        '    foreBrush = Brushes.White

        'Else

        '    f = e.Font
        '    backBrush = New SolidBrush(e.BackColor)
        '    foreBrush = New SolidBrush(e.ForeColor)


        '    'To set the alignment of the caption.
        '    Dim tabName As String = Me.tbcHome.TabPages(e.Index).Text
        '    Dim sf As New StringFormat
        '    sf.Alignment = StringAlignment.Center
        '    'sf.LineAlignment = StringAlignment.Near
        '    'Thsi will help you to fill the interior portion of
        '    'selected tabpage.
        '    e.Graphics.FillRectangle(backBrush, e.Bounds)
        '    Dim r As Rectangle = e.Bounds
        '    r = New Rectangle(r.X, r.Y + 3, r.Width, r.Height - 3)
        '    'Dim r1 As RectangleF = r
        '    'Dim rc As New RectangleConverter

        '    'e.Graphics.DrawString(tabName, f, foreBrush, r, sf)
        '    e.Graphics.DrawString(tabName, f, foreBrush, r.X, r.Y, sf)
        '    sf.Dispose()
        '    If e.Index = Me.tbcHome.SelectedIndex Then

        '        f.Dispose()
        '        backBrush.Dispose()

        '    Else

        '        backBrush.Dispose()
        '        foreBrush.Dispose()

        '    End If
        'End If
			

    End Sub

#End Region
    Private Sub btnsearchname_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsearchname.Enter

    End Sub

    Private Sub mnuFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFile.Click

    End Sub

    Private Sub tbcHome_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tbcHome.DragOver

    End Sub

    Private Sub tbcHome_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles tbcHome.DrawItem
        'Try

        '    Me.DrawOnTab(sender, e)
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub tpgClients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpgClients.Click

    End Sub

    Private Sub tpgHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpgHome.Click

    End Sub

    Private Sub tpgPersonnel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpgPersonnel.Click

    End Sub

    Private Sub tpgJobs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpgJobs.Click

    End Sub

    Private Sub tpgEquip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tpgEquip.Click

    End Sub
End Class
