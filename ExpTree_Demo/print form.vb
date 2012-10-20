
Imports System.Data
Imports System.Data.OleDb
Imports ADODB
Public Class print_form
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

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

        '\\ Populate teh data grids with some bumpf
        Dim MyTable As New DataTable()
        MyTable.Columns.Add(New DataColumn("Team", GetType(String)))
        MyTable.Columns.Add(New DataColumn("Played", GetType(Integer)))
        MyTable.Columns.Add(New DataColumn("Goals For", GetType(Integer)))
        MyTable.Columns.Add(New DataColumn("Goals Against", GetType(Integer)))
        MyTable.Columns.Add(New DataColumn("Points", GetType(Integer)))


        Me.DataGrid1.DataSource = MyTable
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
    Friend WithEvents MainMenu_App As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem_File As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem_File_PageSetup As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem_File_print As System.Windows.Forms.MenuItem
    Friend WithEvents PrintPreviewDialog1 As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents PageSetupDialog1 As System.Windows.Forms.PageSetupDialog
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents PrintPreviewDialog2 As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Mnuviewallclients As System.Windows.Forms.MenuItem
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_ColourFooterLine As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_ColourBodyline As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_ColourHeaderLine As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_InterSectionSpacingPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_FooterHeightPercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_HeaderHeightPercentage As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_PagesAcross As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox_ColumnHeaderBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_OddRowBrush As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_HeaderBrush As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_EvenBrush As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_FooterBrush As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(print_form))
        Me.MainMenu_App = New System.Windows.Forms.MainMenu()
        Me.MenuItem_File = New System.Windows.Forms.MenuItem()
        Me.MenuItem_File_PageSetup = New System.Windows.Forms.MenuItem()
        Me.MenuItem_File_print = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.Mnuviewallclients = New System.Windows.Forms.MenuItem()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.PageSetupDialog1 = New System.Windows.Forms.PageSetupDialog()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.PrintPreviewDialog2 = New System.Windows.Forms.PrintPreviewDialog()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox_ColourFooterLine = New System.Windows.Forms.ComboBox()
        Me.ComboBox_ColourBodyline = New System.Windows.Forms.ComboBox()
        Me.ComboBox_ColourHeaderLine = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.NumericUpDown_InterSectionSpacingPercent = New System.Windows.Forms.NumericUpDown()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.NumericUpDown_FooterHeightPercent = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.NumericUpDown_HeaderHeightPercentage = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.NumericUpDown_PagesAcross = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.ComboBox_ColumnHeaderBrush = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ComboBox_OddRowBrush = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ComboBox_HeaderBrush = New System.Windows.Forms.ComboBox()
        Me.ComboBox_EvenBrush = New System.Windows.Forms.ComboBox()
        Me.ComboBox_FooterBrush = New System.Windows.Forms.ComboBox()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown_InterSectionSpacingPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_FooterHeightPercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_HeaderHeightPercentage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.NumericUpDown_PagesAcross, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu_App
        '
        Me.MainMenu_App.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem_File, Me.MenuItem1})
        '
        'MenuItem_File
        '
        Me.MenuItem_File.Index = 0
        Me.MenuItem_File.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem_File_PageSetup, Me.MenuItem_File_print})
        Me.MenuItem_File.Text = "&File"
        '
        'MenuItem_File_PageSetup
        '
        Me.MenuItem_File_PageSetup.Index = 0
        Me.MenuItem_File_PageSetup.Text = "Page &Setup"
        '
        'MenuItem_File_print
        '
        Me.MenuItem_File_print.Index = 1
        Me.MenuItem_File_print.Text = "&Print"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.Mnuviewallclients})
        Me.MenuItem1.Text = "Clients"
        '
        'Mnuviewallclients
        '
        Me.Mnuviewallclients.Index = 0
        Me.Mnuviewallclients.Text = "View all clients"
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Location = New System.Drawing.Point(218, 3)
        Me.PrintPreviewDialog1.MaximumSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Opacity = 1
        Me.PrintPreviewDialog1.TransparencyKey = System.Drawing.Color.Empty
        Me.PrintPreviewDialog1.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label7})
        Me.GroupBox2.Location = New System.Drawing.Point(266, -130)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 69)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Grid line colours"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Header"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(-54, -26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 26)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Page Heading"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(50, -26)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.TabIndex = 12
        Me.TextBox1.Text = "Data Grid Print Test"
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(8, 240)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(520, 136)
        Me.DataGrid1.TabIndex = 11
        '
        'PrintPreviewDialog2
        '
        Me.PrintPreviewDialog2.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog2.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog2.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog2.Enabled = True
        Me.PrintPreviewDialog2.Icon = CType(resources.GetObject("PrintPreviewDialog2.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog2.Location = New System.Drawing.Point(354, 1)
        Me.PrintPreviewDialog2.MaximumSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog2.Name = "PrintPreviewDialog2"
        Me.PrintPreviewDialog2.Opacity = 1
        Me.PrintPreviewDialog2.TransparencyKey = System.Drawing.Color.Empty
        Me.PrintPreviewDialog2.Visible = False
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(-54, -138)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(100, 28)
        Me.Label9.TabIndex = 16
        Me.Label9.Text = "Footer Brush"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label15, Me.Label6, Me.Label2, Me.ComboBox_ColourFooterLine, Me.ComboBox_ColourBodyline, Me.ComboBox_ColourHeaderLine, Me.Label5})
        Me.GroupBox5.Location = New System.Drawing.Point(256, 104)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(272, 128)
        Me.GroupBox5.TabIndex = 31
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Grid line colours"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 16)
        Me.Label15.Name = "Label15"
        Me.Label15.TabIndex = 5
        Me.Label15.Text = "Header"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 72)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 23)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Footer line colour"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Header line colour"
        '
        'ComboBox_ColourFooterLine
        '
        Me.ComboBox_ColourFooterLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourFooterLine.Location = New System.Drawing.Point(128, 72)
        Me.ComboBox_ColourFooterLine.Name = "ComboBox_ColourFooterLine"
        Me.ComboBox_ColourFooterLine.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_ColourFooterLine.TabIndex = 9
        '
        'ComboBox_ColourBodyline
        '
        Me.ComboBox_ColourBodyline.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourBodyline.Location = New System.Drawing.Point(128, 96)
        Me.ComboBox_ColourBodyline.Name = "ComboBox_ColourBodyline"
        Me.ComboBox_ColourBodyline.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_ColourBodyline.TabIndex = 10
        '
        'ComboBox_ColourHeaderLine
        '
        Me.ComboBox_ColourHeaderLine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ColourHeaderLine.Location = New System.Drawing.Point(128, 40)
        Me.ComboBox_ColourHeaderLine.Name = "ComboBox_ColourHeaderLine"
        Me.ComboBox_ColourHeaderLine.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_ColourHeaderLine.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 96)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Body line colour"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(16, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.TabIndex = 30
        Me.Label16.Text = "Page Heading"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(120, 32)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(128, 20)
        Me.TextBox2.TabIndex = 29
        Me.TextBox2.Text = "Data Grid Print Test"
        '
        'ComboBox1
        '
        Me.ComboBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Location = New System.Drawing.Point(120, 8)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(127, 21)
        Me.ComboBox1.TabIndex = 33
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(16, 8)
        Me.Label17.Name = "Label17"
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Footer Brush"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label14, Me.NumericUpDown_InterSectionSpacingPercent, Me.Label4, Me.NumericUpDown_FooterHeightPercent, Me.Label3, Me.NumericUpDown_HeaderHeightPercentage})
        Me.GroupBox1.Location = New System.Drawing.Point(256, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(272, 96)
        Me.GroupBox1.TabIndex = 28
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Section Heights"
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
        Me.NumericUpDown_InterSectionSpacingPercent.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 23)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Inter-section spacing"
        '
        'NumericUpDown_FooterHeightPercent
        '
        Me.NumericUpDown_FooterHeightPercent.Location = New System.Drawing.Point(136, 40)
        Me.NumericUpDown_FooterHeightPercent.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.NumericUpDown_FooterHeightPercent.Name = "NumericUpDown_FooterHeightPercent"
        Me.NumericUpDown_FooterHeightPercent.TabIndex = 3
        Me.NumericUpDown_FooterHeightPercent.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Footer height"
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
        Me.GroupBox4.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label13, Me.NumericUpDown_PagesAcross})
        Me.GroupBox4.Location = New System.Drawing.Point(8, 176)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(240, 56)
        Me.GroupBox4.TabIndex = 27
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Page layout"
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
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.ComboBox_EvenBrush, Me.ComboBox_ColumnHeaderBrush, Me.Label12, Me.ComboBox_OddRowBrush, Me.Label8, Me.Label10, Me.Label11, Me.ComboBox_HeaderBrush})
        Me.GroupBox3.Location = New System.Drawing.Point(8, 56)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(240, 120)
        Me.GroupBox3.TabIndex = 26
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Grid background  colours"
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
        'ComboBox_OddRowBrush
        '
        Me.ComboBox_OddRowBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_OddRowBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_OddRowBrush.Location = New System.Drawing.Point(112, 40)
        Me.ComboBox_OddRowBrush.Name = "ComboBox_OddRowBrush"
        Me.ComboBox_OddRowBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_OddRowBrush.TabIndex = 12
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 64)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Even rows"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 32)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Header"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "Odd rows"
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
        'ComboBox_EvenBrush
        '
        Me.ComboBox_EvenBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_EvenBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_EvenBrush.Location = New System.Drawing.Point(112, 64)
        Me.ComboBox_EvenBrush.Name = "ComboBox_EvenBrush"
        Me.ComboBox_EvenBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_EvenBrush.TabIndex = 16
        '
        'ComboBox_FooterBrush
        '
        Me.ComboBox_FooterBrush.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBox_FooterBrush.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_FooterBrush.Location = New System.Drawing.Point(72, 120)
        Me.ComboBox_FooterBrush.Name = "ComboBox_FooterBrush"
        Me.ComboBox_FooterBrush.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox_FooterBrush.TabIndex = 18
        '
        'print_form
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(536, 381)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox5, Me.Label16, Me.TextBox2, Me.ComboBox1, Me.Label17, Me.GroupBox1, Me.GroupBox4, Me.GroupBox3, Me.GroupBox2, Me.Label1, Me.TextBox1, Me.DataGrid1, Me.ComboBox_FooterBrush, Me.Label9})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu_App
        Me.Name = "print_form"
        Me.Text = "print_form"
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.NumericUpDown_InterSectionSpacingPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_FooterHeightPercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_HeaderHeightPercentage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.NumericUpDown_PagesAcross, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Private members"
    Private GridPrinter As DataGridPrinter
#End Region

    Private Sub print_form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

#Region "Menu handlers"

    Private Sub MenuItem_File_PageSetup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem_File_PageSetup.Click

        If GridPrinter Is Nothing Then
            GridPrinter = New DataGridPrinter(Me.DataGrid1)
        End If

        With Me.PageSetupDialog1
            .Document = GridPrinter.PrintDocument
            .ShowDialog()
        End With

    End Sub

    Private Sub MenuItem_File_print_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem_File_print.Click
        If GridPrinter Is Nothing Then
            GridPrinter = New DataGridPrinter(Me.DataGrid1)
        End If

        With GridPrinter
            .HeaderText = Me.TextBox1.Text

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

    Private Sub ComboBox_EvenBrush_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
        e.Graphics.FillRectangle(CType(ComboBox_EvenBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub ComboBox_FooterBrush_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ComboBox_FooterBrush.DrawItem
        e.Graphics.FillRectangle(CType(ComboBox_FooterBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub ComboBox_OddRowBrush_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)

        e.Graphics.FillRectangle(CType(ComboBox_OddRowBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)

    End Sub
    Private Sub ComboBox_HeaderBrush_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
        e.Graphics.FillRectangle(CType(ComboBox_HeaderBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub ComboBox_ColumnHeaderBrush_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
        e.Graphics.FillRectangle(CType(ComboBox_ColumnHeaderBrush.Items(e.Index), Brush), e.Bounds)
        e.Graphics.DrawRectangle(System.Drawing.Pens.Black, e.Bounds)
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
    Private Sub DataGrid2_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub
    Private Sub Mnuviewallclients_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Mnuviewallclients.Click
        Try
            loadgrid2()
        Catch sa As Exception


        End Try
    End Sub

#Region "clients"
    Private Sub loadgrid2()
        Dim currentCursor As Cursor = Cursor.Current
        Try
            '-----------------try this dave
            Dim connectstr As String
             connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

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
            Me.DataGrid1.SetDataBinding(custDS, tname)
            Try
                connect.Close()
            Catch xc As Exception

            End Try

            '--------------------this is quite cool---------------------------------------------------------------------


            Call AddCustomDataTableStyle()
            '---------------remove some rows
            Dim ds As System.Data.DataSet = New System.Data.DataSet()
            ds = Me.DataGrid1.DataSource
            ds.Tables(0).Columns.Remove("leads_no")
            ds.Tables(0).Columns.Remove("least_status")
            Dim MyTable As New DataTable()
            MyTable = ds.Tables(0)
            DataGrid1.DataSource = MyTable
            '------------
        Catch t As Exception

        Finally
            'statusBar1.Text = "Done"
            Cursor.Current = currentCursor


        End Try
    End Sub
    Private Sub AddCustomDataTableStyle()
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = "clients"
            Dim mywidth, mywidth1 As Integer
            mywidth = Me.DataGrid1.Width - 10
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
            DataGrid1.TableStyles.Clear()
            DataGrid1.TableStyles.Add(ts1)
        Catch ex As Exception

        End Try

    End Sub 'AddCustomDataTableStyle
#End Region

End Class
