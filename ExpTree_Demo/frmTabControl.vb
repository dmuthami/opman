'------------------------------------------------------------------------------
'/ <copyright from='1997' to='2001' company='Microsoft Corporation'>
'/    Copyright (c) Microsoft Corporation. All Rights Reserved.
'/
'/    This source code is intended only as a supplement to Microsoft
'/    Development Tools and/or on-line documentation.  See these other
'/    materials for detailed information regarding Microsoft code samples.
'/
'/ </copyright>
'------------------------------------------------------------------------------
Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Resources

' <doc>
' <desc>
'     This class demonstrates the TabControl control.
'
'     Typically the contents of each page are encapsulated
'     in a UserControl however for this simple example the
'     contents of all pages are defined directly in this
'     form.
'
'     TabPages can either be added at designtime or runtime
'     The main MiscUI file shows an example of how to add pages
'     dynamically at runtime
'
' </desc>
' </doc>
'
Public Class TabControlCtl
    Inherits System.Windows.Forms.Form

    Public Sub New()

        MyBase.New()

        TabControlCtl = Me

        'This call is required by the Windows Forms Designer.
        InitializeComponent()

        appearanceComboBox.SelectedIndex = 0
        alignmentComboBox.SelectedIndex = 0
        sizeModeComboBox.SelectedIndex = 0
        tabControl1.ImageList = Nothing

    End Sub

    'Form overrides dispose to clean up the component list.
    'Public Overloads Overrides Sub Dispose()
    '    MyBase.Dispose()
    '    components.Dispose()
    'End Sub

    Protected Sub checkBox1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles checkBox1.Click
        Me.tabControl1.Multiline = checkBox1.Checked
        alignmentComboBox_SelectedIndexChanged(Nothing, EventArgs.Empty)
    End Sub

    Protected Sub checkBox2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles checkBox2.Click
        Me.tabControl1.HotTrack = checkBox2.Checked
    End Sub

    Protected Sub checkBox3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles checkBox3.Click
        If checkBox3.Checked Then
            tabControl1.ImageList = imageList1
        Else
            tabControl1.ImageList = Nothing
        End If

    End Sub

    Protected Sub appearanceComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles appearanceComboBox.SelectedIndexChanged
        Dim index As Integer = appearanceComboBox.SelectedIndex
        If index = 0 Then
            tabControl1.Appearance = TabAppearance.Normal
        Else
            If index = 1 Then
                tabControl1.Appearance = TabAppearance.Buttons
            Else
                tabControl1.Appearance = TabAppearance.FlatButtons
            End If
        End If
        tabControl1.PerformLayout()

    End Sub

    Protected Sub alignmentComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles alignmentComboBox.SelectedIndexChanged
        Dim index As Integer = alignmentComboBox.SelectedIndex
        If index = 0 Then
            tabControl1.Alignment = TabAlignment.Top
        Else
            If index = 1 Then
                tabControl1.Alignment = TabAlignment.Bottom
            Else
                If index = 2 Then
                    tabControl1.Alignment = TabAlignment.Left
                Else
                    tabControl1.Alignment = TabAlignment.Right
                End If

            End If
        End If
    End Sub

    Protected Sub sizeModeComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles sizeModeComboBox.SelectedIndexChanged
        Dim index As Integer = sizeModeComboBox.SelectedIndex
        If index = 0 Then
            tabControl1.SizeMode = TabSizeMode.Normal
        Else
            If index = 1 Then
                tabControl1.SizeMode = TabSizeMode.FillToRight
            Else
                tabControl1.SizeMode = TabSizeMode.Fixed

            End If
        End If

    End Sub

    Protected Sub trackBar_Scroll(ByVal sender As Object, ByVal e As EventArgs) Handles TrackBar.Scroll
        tabControl1.Width = trackBar.Value
    End Sub
    Private components As System.ComponentModel.IContainer

#Region " Windows Form Designer generated code "

    'Required by the Windows Form Designer
    Protected WithEvents tp5DomainUpDown1 As System.Windows.Forms.DomainUpDown
    Protected WithEvents tp5GroupBox1 As System.Windows.Forms.GroupBox
    Protected WithEvents tp4ComboBox1 As System.Windows.Forms.ComboBox
    Protected WithEvents tp4UpDown1 As System.Windows.Forms.NumericUpDown
    Protected WithEvents tp4GroupBox1 As System.Windows.Forms.GroupBox
    Protected WithEvents tp3ComboBox1 As System.Windows.Forms.ComboBox
    Protected WithEvents tp3RadioButton1 As System.Windows.Forms.RadioButton
    Protected WithEvents tp3RadioButton2 As System.Windows.Forms.RadioButton
    Protected WithEvents tp3ComboBox2 As System.Windows.Forms.ComboBox
    Protected WithEvents tp3Label1 As System.Windows.Forms.Label
    Protected WithEvents tp3Button1 As System.Windows.Forms.Button
    Protected WithEvents tp3GroupBox1 As System.Windows.Forms.GroupBox
    Protected WithEvents tabPage5 As System.Windows.Forms.TabPage
    Protected WithEvents tabPage4 As System.Windows.Forms.TabPage
    Protected WithEvents tabPage3 As System.Windows.Forms.TabPage
    Protected WithEvents tp2ComboBox1 As System.Windows.Forms.ComboBox
    Protected WithEvents tp2RadioButton1 As System.Windows.Forms.RadioButton
    Protected WithEvents tp2RadioButton2 As System.Windows.Forms.RadioButton
    Protected WithEvents tp2GroupBox1 As System.Windows.Forms.GroupBox
    Protected WithEvents tabPage2 As System.Windows.Forms.TabPage
    Protected WithEvents tabPage1 As System.Windows.Forms.TabPage
    Protected WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Protected WithEvents appearanceComboBox As System.Windows.Forms.ComboBox
    Protected WithEvents checkBox1 As System.Windows.Forms.CheckBox
    Protected WithEvents tabControl1 As System.Windows.Forms.TabControl
    Protected WithEvents alignmentComboBox As System.Windows.Forms.ComboBox
    Protected WithEvents checkBox2 As System.Windows.Forms.CheckBox
    Protected WithEvents sizeModeComboBox As System.Windows.Forms.ComboBox
    Protected WithEvents label1 As System.Windows.Forms.Label
    Protected WithEvents label2 As System.Windows.Forms.Label
    Protected WithEvents label3 As System.Windows.Forms.Label
    Protected WithEvents trackBar As System.Windows.Forms.TrackBar
    Protected WithEvents label4 As System.Windows.Forms.Label
    Protected WithEvents pictureBox1 As System.Windows.Forms.PictureBox
    Protected WithEvents toolTip1 As System.Windows.Forms.ToolTip
    Protected WithEvents imageList1 As System.Windows.Forms.ImageList
    Protected WithEvents checkBox3 As System.Windows.Forms.CheckBox
    Protected WithEvents tp1ComboBox1 As System.Windows.Forms.ComboBox
    Protected WithEvents tp1Label1 As System.Windows.Forms.Label
    Protected WithEvents tp1Label2 As System.Windows.Forms.Label
    Protected WithEvents tp1Button1 As System.Windows.Forms.Button
    Protected WithEvents tp1GroupBox1 As System.Windows.Forms.GroupBox

    Private WithEvents TabControlCtl As System.Windows.Forms.Form

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.tp3RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.tp5DomainUpDown1 = New System.Windows.Forms.DomainUpDown()
        Me.alignmentComboBox = New System.Windows.Forms.ComboBox()
        Me.tp1Label2 = New System.Windows.Forms.Label()
        Me.tp1Label1 = New System.Windows.Forms.Label()
        Me.tp1GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tp1Button1 = New System.Windows.Forms.Button()
        Me.tp1ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.tp3RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.trackBar = New System.Windows.Forms.TrackBar()
        Me.tp2ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.tp2RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.label4 = New System.Windows.Forms.Label()
        Me.imageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.tp4GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tp4ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.tp4UpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.sizeModeComboBox = New System.Windows.Forms.ComboBox()
        Me.toolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.checkBox2 = New System.Windows.Forms.CheckBox()
        Me.appearanceComboBox = New System.Windows.Forms.ComboBox()
        Me.checkBox1 = New System.Windows.Forms.CheckBox()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.checkBox3 = New System.Windows.Forms.CheckBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.tabPage1 = New System.Windows.Forms.TabPage()
        Me.tp2RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.tp3ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.pictureBox1 = New System.Windows.Forms.PictureBox()
        Me.tabPage4 = New System.Windows.Forms.TabPage()
        Me.tp5GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tp2GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tp3ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.tp3Button1 = New System.Windows.Forms.Button()
        Me.tabControl1 = New System.Windows.Forms.TabControl()
        Me.tabPage3 = New System.Windows.Forms.TabPage()
        Me.tp3GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tp3Label1 = New System.Windows.Forms.Label()
        Me.tabPage2 = New System.Windows.Forms.TabPage()
        Me.tabPage5 = New System.Windows.Forms.TabPage()
        Me.tp1GroupBox1.SuspendLayout()
        CType(Me.trackBar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp4GroupBox1.SuspendLayout()
        CType(Me.tp4UpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.groupBox1.SuspendLayout()
        Me.tabPage1.SuspendLayout()
        Me.tabPage4.SuspendLayout()
        Me.tp5GroupBox1.SuspendLayout()
        Me.tp2GroupBox1.SuspendLayout()
        Me.tabControl1.SuspendLayout()
        Me.tabPage3.SuspendLayout()
        Me.tp3GroupBox1.SuspendLayout()
        Me.tabPage2.SuspendLayout()
        Me.tabPage5.SuspendLayout()
        Me.SuspendLayout()
        '
        'label2
        '
        Me.label2.Location = New System.Drawing.Point(16, 48)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(64, 16)
        Me.label2.TabIndex = 8
        Me.label2.Text = "Alignment"
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(16, 16)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(72, 16)
        Me.label1.TabIndex = 7
        Me.label1.Text = "Appearance"
        '
        'tp3RadioButton1
        '
        Me.tp3RadioButton1.Enabled = False
        Me.tp3RadioButton1.Location = New System.Drawing.Point(8, 24)
        Me.tp3RadioButton1.Name = "tp3RadioButton1"
        Me.tp3RadioButton1.Size = New System.Drawing.Size(136, 23)
        Me.tp3RadioButton1.TabIndex = 4
        Me.tp3RadioButton1.Text = "&Single Instrument"
        '
        'tp5DomainUpDown1
        '
        Me.tp5DomainUpDown1.AccessibleName = "DomainUpDown"
        Me.tp5DomainUpDown1.AccessibleRole = System.Windows.Forms.AccessibleRole.ComboBox
        Me.tp5DomainUpDown1.Enabled = False
        Me.tp5DomainUpDown1.Location = New System.Drawing.Point(24, 32)
        Me.tp5DomainUpDown1.Name = "tp5DomainUpDown1"
        Me.tp5DomainUpDown1.Size = New System.Drawing.Size(112, 20)
        Me.tp5DomainUpDown1.TabIndex = 0
        Me.tp5DomainUpDown1.Text = "11:01:35 AM"
        '
        'alignmentComboBox
        '
        Me.alignmentComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.alignmentComboBox.Items.AddRange(New Object() {"Top", "Bottom", "Left", "Right"})
        Me.alignmentComboBox.Location = New System.Drawing.Point(128, 48)
        Me.alignmentComboBox.Name = "alignmentComboBox"
        Me.alignmentComboBox.Size = New System.Drawing.Size(104, 21)
        Me.alignmentComboBox.TabIndex = 4
        Me.toolTip1.SetToolTip(Me.alignmentComboBox, "Determines whether the tabs appear on the top, bottom,left or, right side of the " & _
        "control.")
        '
        'tp1Label2
        '
        Me.tp1Label2.Location = New System.Drawing.Point(24, 88)
        Me.tp1Label2.Name = "tp1Label2"
        Me.tp1Label2.Size = New System.Drawing.Size(176, 16)
        Me.tp1Label2.TabIndex = 1
        Me.tp1Label2.Text = "Select Advanced Options:"
        '
        'tp1Label1
        '
        Me.tp1Label1.Location = New System.Drawing.Point(24, 24)
        Me.tp1Label1.Name = "tp1Label1"
        Me.tp1Label1.Size = New System.Drawing.Size(100, 16)
        Me.tp1Label1.TabIndex = 2
        Me.tp1Label1.Text = "Preferred device:"
        '
        'tp1GroupBox1
        '
        Me.tp1GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp1Label1, Me.tp1Label2, Me.tp1Button1, Me.tp1ComboBox1})
        Me.tp1GroupBox1.Location = New System.Drawing.Point(12, 16)
        Me.tp1GroupBox1.Name = "tp1GroupBox1"
        Me.tp1GroupBox1.Size = New System.Drawing.Size(202, 144)
        Me.tp1GroupBox1.TabIndex = 0
        Me.tp1GroupBox1.TabStop = False
        Me.tp1GroupBox1.Text = "Playback"
        '
        'tp1Button1
        '
        Me.tp1Button1.Enabled = False
        Me.tp1Button1.Location = New System.Drawing.Point(24, 112)
        Me.tp1Button1.Name = "tp1Button1"
        Me.tp1Button1.Size = New System.Drawing.Size(160, 23)
        Me.tp1Button1.TabIndex = 0
        Me.tp1Button1.Text = "Advanced &Properties"
        '
        'tp1ComboBox1
        '
        Me.tp1ComboBox1.Enabled = False
        Me.tp1ComboBox1.Location = New System.Drawing.Point(24, 48)
        Me.tp1ComboBox1.Name = "tp1ComboBox1"
        Me.tp1ComboBox1.Size = New System.Drawing.Size(160, 21)
        Me.tp1ComboBox1.TabIndex = 3
        Me.tp1ComboBox1.Text = "(Use any available device)"
        '
        'tp3RadioButton2
        '
        Me.tp3RadioButton2.Enabled = False
        Me.tp3RadioButton2.Location = New System.Drawing.Point(8, 80)
        Me.tp3RadioButton2.Name = "tp3RadioButton2"
        Me.tp3RadioButton2.Size = New System.Drawing.Size(168, 23)
        Me.tp3RadioButton2.TabIndex = 3
        Me.tp3RadioButton2.Text = "&Custom Configuration"
        '
        'trackBar
        '
        Me.trackBar.BackColor = System.Drawing.SystemColors.Control
        Me.trackBar.Location = New System.Drawing.Point(16, 192)
        Me.trackBar.Maximum = 220
        Me.trackBar.Name = "trackBar"
        Me.trackBar.Size = New System.Drawing.Size(216, 45)
        Me.trackBar.TabIndex = 2
        Me.trackBar.TabStop = False
        Me.trackBar.Text = "TrackBar"
        Me.trackBar.TickFrequency = 10
        Me.trackBar.Value = 220
        '
        'tp2ComboBox1
        '
        Me.tp2ComboBox1.Enabled = False
        Me.tp2ComboBox1.Location = New System.Drawing.Point(32, 80)
        Me.tp2ComboBox1.Name = "tp2ComboBox1"
        Me.tp2ComboBox1.Size = New System.Drawing.Size(160, 21)
        Me.tp2ComboBox1.TabIndex = 2
        Me.tp2ComboBox1.Text = "Original Size"
        '
        'tp2RadioButton2
        '
        Me.tp2RadioButton2.Enabled = False
        Me.tp2RadioButton2.Location = New System.Drawing.Point(32, 48)
        Me.tp2RadioButton2.Name = "tp2RadioButton2"
        Me.tp2RadioButton2.Size = New System.Drawing.Size(100, 23)
        Me.tp2RadioButton2.TabIndex = 0
        Me.tp2RadioButton2.Text = "&Full Screen"
        '
        'label4
        '
        Me.label4.Location = New System.Drawing.Point(16, 168)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(120, 16)
        Me.label4.TabIndex = 3
        Me.label4.Text = "Tab Control Width:"
        '
        'imageList1
        '
        Me.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.imageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.imageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'tp4GroupBox1
        '
        Me.tp4GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp4ComboBox1, Me.tp4UpDown1})
        Me.tp4GroupBox1.Location = New System.Drawing.Point(12, 16)
        Me.tp4GroupBox1.Name = "tp4GroupBox1"
        Me.tp4GroupBox1.Size = New System.Drawing.Size(202, 88)
        Me.tp4GroupBox1.TabIndex = 0
        Me.tp4GroupBox1.TabStop = False
        Me.tp4GroupBox1.Text = "Date"
        '
        'tp4ComboBox1
        '
        Me.tp4ComboBox1.Enabled = False
        Me.tp4ComboBox1.Location = New System.Drawing.Point(16, 32)
        Me.tp4ComboBox1.Name = "tp4ComboBox1"
        Me.tp4ComboBox1.Size = New System.Drawing.Size(56, 21)
        Me.tp4ComboBox1.TabIndex = 1
        Me.tp4ComboBox1.Text = "June"
        '
        'tp4UpDown1
        '
        Me.tp4UpDown1.DecimalPlaces = 2
        Me.tp4UpDown1.Enabled = False
        Me.tp4UpDown1.Location = New System.Drawing.Point(88, 32)
        Me.tp4UpDown1.Maximum = New Decimal(New Integer() {0, 0, 0, 0})
        Me.tp4UpDown1.Name = "tp4UpDown1"
        Me.tp4UpDown1.Size = New System.Drawing.Size(64, 20)
        Me.tp4UpDown1.TabIndex = 0
        Me.tp4UpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'sizeModeComboBox
        '
        Me.sizeModeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.sizeModeComboBox.Items.AddRange(New Object() {"Normal", "Fill to Right", "Fixed"})
        Me.sizeModeComboBox.Location = New System.Drawing.Point(128, 80)
        Me.sizeModeComboBox.Name = "sizeModeComboBox"
        Me.sizeModeComboBox.Size = New System.Drawing.Size(104, 21)
        Me.sizeModeComboBox.TabIndex = 6
        Me.toolTip1.SetToolTip(Me.sizeModeComboBox, "Indicates how tabs are sized.")
        '
        'checkBox2
        '
        Me.checkBox2.Location = New System.Drawing.Point(16, 136)
        Me.checkBox2.Name = "checkBox2"
        Me.checkBox2.Size = New System.Drawing.Size(100, 23)
        Me.checkBox2.TabIndex = 5
        Me.checkBox2.Text = "HotTrack"
        Me.toolTip1.SetToolTip(Me.checkBox2, "Indicates whether the tabs visually change as the mouse passes")
        '
        'appearanceComboBox
        '
        Me.appearanceComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.appearanceComboBox.Items.AddRange(New Object() {"Normal", "Buttons", "Flat Buttons"})
        Me.appearanceComboBox.Location = New System.Drawing.Point(128, 16)
        Me.appearanceComboBox.Name = "appearanceComboBox"
        Me.appearanceComboBox.Size = New System.Drawing.Size(104, 21)
        Me.appearanceComboBox.TabIndex = 1
        Me.toolTip1.SetToolTip(Me.appearanceComboBox, "Indicates whether the tabs are painted as buttons or regular tabs.")
        '
        'checkBox1
        '
        Me.checkBox1.Location = New System.Drawing.Point(16, 112)
        Me.checkBox1.Name = "checkBox1"
        Me.checkBox1.Size = New System.Drawing.Size(88, 16)
        Me.checkBox1.TabIndex = 0
        Me.checkBox1.Text = "Multiline"
        Me.toolTip1.SetToolTip(Me.checkBox1, "Indicates if more than one row of tabs is allowed.")
        '
        'groupBox1
        '
        Me.groupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.checkBox3, Me.label4, Me.trackBar, Me.label3, Me.label2, Me.label1, Me.sizeModeComboBox, Me.checkBox2, Me.alignmentComboBox, Me.appearanceComboBox, Me.checkBox1})
        Me.groupBox1.Location = New System.Drawing.Point(280, 16)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(248, 264)
        Me.groupBox1.TabIndex = 1
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "TabControl"
        '
        'checkBox3
        '
        Me.checkBox3.Location = New System.Drawing.Point(128, 112)
        Me.checkBox3.Name = "checkBox3"
        Me.checkBox3.Size = New System.Drawing.Size(72, 16)
        Me.checkBox3.TabIndex = 10
        Me.checkBox3.Text = "Images"
        '
        'label3
        '
        Me.label3.Location = New System.Drawing.Point(16, 80)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(80, 16)
        Me.label3.TabIndex = 9
        Me.label3.Text = "SizeMode"
        '
        'tabPage1
        '
        Me.tabPage1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp1GroupBox1})
        Me.tabPage1.ImageIndex = 0
        Me.tabPage1.Location = New System.Drawing.Point(4, 23)
        Me.tabPage1.Name = "tabPage1"
        Me.tabPage1.Size = New System.Drawing.Size(224, 193)
        Me.tabPage1.TabIndex = 0
        Me.tabPage1.Text = "Audio"
        '
        'tp2RadioButton1
        '
        Me.tp2RadioButton1.Enabled = False
        Me.tp2RadioButton1.Location = New System.Drawing.Point(32, 24)
        Me.tp2RadioButton1.Name = "tp2RadioButton1"
        Me.tp2RadioButton1.Size = New System.Drawing.Size(100, 23)
        Me.tp2RadioButton1.TabIndex = 1
        Me.tp2RadioButton1.Text = "&Window"
        '
        'tp3ComboBox2
        '
        Me.tp3ComboBox2.Enabled = False
        Me.tp3ComboBox2.Location = New System.Drawing.Point(24, 120)
        Me.tp3ComboBox2.Name = "tp3ComboBox2"
        Me.tp3ComboBox2.Size = New System.Drawing.Size(96, 21)
        Me.tp3ComboBox2.TabIndex = 2
        Me.tp3ComboBox2.Text = "Default"
        '
        'pictureBox1
        '
        Me.pictureBox1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pictureBox1.Location = New System.Drawing.Point(8, 24)
        Me.pictureBox1.Name = "pictureBox1"
        Me.pictureBox1.Size = New System.Drawing.Size(264, 256)
        Me.pictureBox1.TabIndex = 2
        Me.pictureBox1.TabStop = False
        Me.pictureBox1.Text = "PictureBox"
        '
        'tabPage4
        '
        Me.tabPage4.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp4GroupBox1})
        Me.tabPage4.ImageIndex = 3
        Me.tabPage4.Location = New System.Drawing.Point(4, 23)
        Me.tabPage4.Name = "tabPage4"
        Me.tabPage4.Size = New System.Drawing.Size(224, 193)
        Me.tabPage4.TabIndex = 3
        Me.tabPage4.Text = "Date"
        '
        'tp5GroupBox1
        '
        Me.tp5GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp5DomainUpDown1})
        Me.tp5GroupBox1.Location = New System.Drawing.Point(12, 16)
        Me.tp5GroupBox1.Name = "tp5GroupBox1"
        Me.tp5GroupBox1.Size = New System.Drawing.Size(202, 88)
        Me.tp5GroupBox1.TabIndex = 0
        Me.tp5GroupBox1.TabStop = False
        Me.tp5GroupBox1.Text = "Time"
        '
        'tp2GroupBox1
        '
        Me.tp2GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp2ComboBox1, Me.tp2RadioButton1, Me.tp2RadioButton2})
        Me.tp2GroupBox1.Location = New System.Drawing.Point(12, 16)
        Me.tp2GroupBox1.Name = "tp2GroupBox1"
        Me.tp2GroupBox1.Size = New System.Drawing.Size(202, 128)
        Me.tp2GroupBox1.TabIndex = 0
        Me.tp2GroupBox1.TabStop = False
        Me.tp2GroupBox1.Text = "Show video in:"
        '
        'tp3ComboBox1
        '
        Me.tp3ComboBox1.Enabled = False
        Me.tp3ComboBox1.Location = New System.Drawing.Point(24, 48)
        Me.tp3ComboBox1.Name = "tp3ComboBox1"
        Me.tp3ComboBox1.Size = New System.Drawing.Size(160, 21)
        Me.tp3ComboBox1.TabIndex = 5
        Me.tp3ComboBox1.Text = "Creative Music Synth [220]"
        '
        'tp3Button1
        '
        Me.tp3Button1.Enabled = False
        Me.tp3Button1.Location = New System.Drawing.Point(122, 120)
        Me.tp3Button1.Name = "tp3Button1"
        Me.tp3Button1.Size = New System.Drawing.Size(74, 24)
        Me.tp3Button1.TabIndex = 0
        Me.tp3Button1.Text = "&Configure"
        '
        'tabControl1
        '
        Me.tabControl1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tabPage1, Me.tabPage2, Me.tabPage3, Me.tabPage4, Me.tabPage5})
        Me.tabControl1.ImageList = Me.imageList1
        Me.tabControl1.Location = New System.Drawing.Point(24, 32)
        Me.tabControl1.Name = "tabControl1"
        Me.tabControl1.SelectedIndex = 0
        Me.tabControl1.Size = New System.Drawing.Size(232, 220)
        Me.tabControl1.TabIndex = 0
        Me.tabControl1.Text = "TabControl"
        '
        'tabPage3
        '
        Me.tabPage3.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp3GroupBox1})
        Me.tabPage3.ImageIndex = 2
        Me.tabPage3.Location = New System.Drawing.Point(4, 23)
        Me.tabPage3.Name = "tabPage3"
        Me.tabPage3.Size = New System.Drawing.Size(224, 193)
        Me.tabPage3.TabIndex = 2
        Me.tabPage3.Text = "MIDI"
        '
        'tp3GroupBox1
        '
        Me.tp3GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp3ComboBox2, Me.tp3Label1, Me.tp3RadioButton2, Me.tp3Button1, Me.tp3ComboBox1, Me.tp3RadioButton1})
        Me.tp3GroupBox1.Location = New System.Drawing.Point(12, 16)
        Me.tp3GroupBox1.Name = "tp3GroupBox1"
        Me.tp3GroupBox1.Size = New System.Drawing.Size(202, 160)
        Me.tp3GroupBox1.TabIndex = 0
        Me.tp3GroupBox1.TabStop = False
        Me.tp3GroupBox1.Text = "MIDI Output"
        '
        'tp3Label1
        '
        Me.tp3Label1.Location = New System.Drawing.Point(24, 104)
        Me.tp3Label1.Name = "tp3Label1"
        Me.tp3Label1.Size = New System.Drawing.Size(100, 16)
        Me.tp3Label1.TabIndex = 1
        Me.tp3Label1.Text = "MIDI Scheme:"
        '
        'tabPage2
        '
        Me.tabPage2.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp2GroupBox1})
        Me.tabPage2.ImageIndex = 1
        Me.tabPage2.Location = New System.Drawing.Point(4, 23)
        Me.tabPage2.Name = "tabPage2"
        Me.tabPage2.Size = New System.Drawing.Size(224, 193)
        Me.tabPage2.TabIndex = 1
        Me.tabPage2.Text = "Video"
        '
        'tabPage5
        '
        Me.tabPage5.Controls.AddRange(New System.Windows.Forms.Control() {Me.tp5GroupBox1})
        Me.tabPage5.ImageIndex = 4
        Me.tabPage5.Location = New System.Drawing.Point(4, 23)
        Me.tabPage5.Name = "tabPage5"
        Me.tabPage5.Size = New System.Drawing.Size(224, 193)
        Me.tabPage5.TabIndex = 4
        Me.tabPage5.Text = "Time"
        '
        'TabControlCtl
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(546, 293)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.tabControl1, Me.pictureBox1, Me.groupBox1})
        Me.Name = "TabControlCtl"
        Me.Text = "TabControl"
        Me.tp1GroupBox1.ResumeLayout(False)
        CType(Me.trackBar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp4GroupBox1.ResumeLayout(False)
        CType(Me.tp4UpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.groupBox1.ResumeLayout(False)
        Me.tabPage1.ResumeLayout(False)
        Me.tabPage4.ResumeLayout(False)
        Me.tp5GroupBox1.ResumeLayout(False)
        Me.tp2GroupBox1.ResumeLayout(False)
        Me.tabControl1.ResumeLayout(False)
        Me.tabPage3.ResumeLayout(False)
        Me.tp3GroupBox1.ResumeLayout(False)
        Me.tabPage2.ResumeLayout(False)
        Me.tabPage5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    '' The main entry point for the application.
    '<STAThread()> Public Shared Sub Main()
    '    Application.Run(New TabControlCtl())
    'End Sub
End Class



