Public Class frmtimesheetmanager
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
    Friend WithEvents pnltimesheet As System.Windows.Forms.Panel
    Friend WithEvents chkudt As System.Windows.Forms.CheckBox
    Friend WithEvents chkulmt As System.Windows.Forms.CheckBox
    Friend WithEvents grpdts As System.Windows.Forms.GroupBox
    Friend WithEvents chkpdd As System.Windows.Forms.CheckBox
    Friend WithEvents chkpdm As System.Windows.Forms.CheckBox
    Friend WithEvents nud As System.Windows.Forms.NumericUpDown
    Friend WithEvents grppd As System.Windows.Forms.GroupBox
    Friend WithEvents dtpdays As System.Windows.Forms.DateTimePicker
    Friend WithEvents lstdates As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnsave As VSEssentials.VSHotButton
    Friend WithEvents btnrci As VSEssentials.VSHotButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnltimesheet = New System.Windows.Forms.Panel
        Me.grpdts = New System.Windows.Forms.GroupBox
        Me.grppd = New System.Windows.Forms.GroupBox
        Me.btnrci = New VSEssentials.VSHotButton
        Me.btnsave = New VSEssentials.VSHotButton
        Me.lstdates = New System.Windows.Forms.CheckedListBox
        Me.dtpdays = New System.Windows.Forms.DateTimePicker
        Me.nud = New System.Windows.Forms.NumericUpDown
        Me.chkpdm = New System.Windows.Forms.CheckBox
        Me.chkpdd = New System.Windows.Forms.CheckBox
        Me.chkulmt = New System.Windows.Forms.CheckBox
        Me.chkudt = New System.Windows.Forms.CheckBox
        Me.pnltimesheet.SuspendLayout()
        Me.grpdts.SuspendLayout()
        Me.grppd.SuspendLayout()
        CType(Me.nud, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnltimesheet
        '
        Me.pnltimesheet.Controls.Add(Me.grpdts)
        Me.pnltimesheet.Controls.Add(Me.chkulmt)
        Me.pnltimesheet.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnltimesheet.Location = New System.Drawing.Point(0, 0)
        Me.pnltimesheet.Name = "pnltimesheet"
        Me.pnltimesheet.Size = New System.Drawing.Size(328, 338)
        Me.pnltimesheet.TabIndex = 0
        '
        'grpdts
        '
        Me.grpdts.Controls.Add(Me.grppd)
        Me.grpdts.Controls.Add(Me.nud)
        Me.grpdts.Controls.Add(Me.chkpdm)
        Me.grpdts.Controls.Add(Me.chkpdd)
        Me.grpdts.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpdts.Location = New System.Drawing.Point(8, 26)
        Me.grpdts.Name = "grpdts"
        Me.grpdts.Size = New System.Drawing.Size(376, 310)
        Me.grpdts.TabIndex = 2
        Me.grpdts.TabStop = False
        Me.grpdts.Text = "Database time setting"
        '
        'grppd
        '
        Me.grppd.Controls.Add(Me.btnrci)
        Me.grppd.Controls.Add(Me.btnsave)
        Me.grppd.Controls.Add(Me.lstdates)
        Me.grppd.Controls.Add(Me.dtpdays)
        Me.grppd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grppd.Location = New System.Drawing.Point(5, 60)
        Me.grppd.Name = "grppd"
        Me.grppd.Size = New System.Drawing.Size(307, 244)
        Me.grppd.TabIndex = 6
        Me.grppd.TabStop = False
        Me.grppd.Text = "Pick days"
        '
        'btnrci
        '
        Me.btnrci.BackMouseDownColor = System.Drawing.SystemColors.Control
        Me.btnrci.BackMouseOverColor = System.Drawing.SystemColors.Control
        Me.btnrci.BorderBottomColor = System.Drawing.SystemColors.ControlDark
        Me.btnrci.BorderLeftColor = System.Drawing.SystemColors.ControlLight
        Me.btnrci.BorderRightColor = System.Drawing.SystemColors.ControlDark
        Me.btnrci.BorderSize = 1
        Me.btnrci.BorderTopColor = System.Drawing.SystemColors.ControlLight
        Me.btnrci.ButtonSettings = System.Drawing.SystemColors.Control
        Me.btnrci.DrawTextShadow = True
        Me.btnrci.Location = New System.Drawing.Point(139, 224)
        Me.btnrci.Name = "btnrci"
        Me.btnrci.Size = New System.Drawing.Size(160, 16)
        Me.btnrci.TabIndex = 10
        Me.btnrci.TextAlign = VSEssentials.VSHotButton.eTextAlign.Center
        Me.btnrci.TextCaption = "Remove checked items"
        Me.btnrci.TextColor = System.Drawing.SystemColors.ControlText
        Me.btnrci.TextFont = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnrci.TextLeft = 0
        Me.btnrci.TextOffsetX = 0
        Me.btnrci.TextOffsetY = 0
        Me.btnrci.TextShadowColor = System.Drawing.SystemColors.ControlDark
        Me.btnrci.TextTop = 0
        '
        'btnsave
        '
        Me.btnsave.BackMouseDownColor = System.Drawing.SystemColors.Control
        Me.btnsave.BackMouseOverColor = System.Drawing.SystemColors.Control
        Me.btnsave.BorderBottomColor = System.Drawing.SystemColors.ControlDark
        Me.btnsave.BorderLeftColor = System.Drawing.SystemColors.ControlLight
        Me.btnsave.BorderRightColor = System.Drawing.SystemColors.ControlDark
        Me.btnsave.BorderSize = 1
        Me.btnsave.BorderTopColor = System.Drawing.SystemColors.ControlLight
        Me.btnsave.ButtonSettings = System.Drawing.SystemColors.Control
        Me.btnsave.DrawTextShadow = True
        Me.btnsave.Location = New System.Drawing.Point(8, 224)
        Me.btnsave.Name = "btnsave"
        Me.btnsave.Size = New System.Drawing.Size(56, 16)
        Me.btnsave.TabIndex = 9
        Me.btnsave.TextAlign = VSEssentials.VSHotButton.eTextAlign.Center
        Me.btnsave.TextCaption = "Save"
        Me.btnsave.TextColor = System.Drawing.SystemColors.ControlText
        Me.btnsave.TextFont = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnsave.TextLeft = 0
        Me.btnsave.TextOffsetX = 0
        Me.btnsave.TextOffsetY = 0
        Me.btnsave.TextShadowColor = System.Drawing.SystemColors.ControlDark
        Me.btnsave.TextTop = 0
        '
        'lstdates
        '
        Me.lstdates.Location = New System.Drawing.Point(8, 41)
        Me.lstdates.Name = "lstdates"
        Me.lstdates.Size = New System.Drawing.Size(288, 169)
        Me.lstdates.TabIndex = 8
        '
        'dtpdays
        '
        Me.dtpdays.Location = New System.Drawing.Point(8, 16)
        Me.dtpdays.Name = "dtpdays"
        Me.dtpdays.Size = New System.Drawing.Size(240, 20)
        Me.dtpdays.TabIndex = 7
        '
        'nud
        '
        Me.nud.BackColor = System.Drawing.SystemColors.Window
        Me.nud.Location = New System.Drawing.Point(264, 16)
        Me.nud.Name = "nud"
        Me.nud.Size = New System.Drawing.Size(48, 20)
        Me.nud.TabIndex = 4
        Me.nud.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'chkpdm
        '
        Me.chkpdm.Location = New System.Drawing.Point(7, 40)
        Me.chkpdm.Name = "chkpdm"
        Me.chkpdm.Size = New System.Drawing.Size(280, 19)
        Me.chkpdm.TabIndex = 5
        Me.chkpdm.Text = "Pick days to modify since database time"
        '
        'chkpdd
        '
        Me.chkpdd.Location = New System.Drawing.Point(8, 18)
        Me.chkpdd.Name = "chkpdd"
        Me.chkpdd.Size = New System.Drawing.Size(248, 20)
        Me.chkpdd.TabIndex = 3
        Me.chkpdd.Text = "Pick no of days since database time"
        '
        'chkulmt
        '
        Me.chkulmt.Location = New System.Drawing.Point(157, 8)
        Me.chkulmt.Name = "chkulmt"
        Me.chkulmt.Size = New System.Drawing.Size(170, 16)
        Me.chkulmt.TabIndex = 1
        Me.chkulmt.Text = "use local machine time"
        '
        'chkudt
        '
        Me.chkudt.Location = New System.Drawing.Point(10, 8)
        Me.chkudt.Name = "chkudt"
        Me.chkudt.Size = New System.Drawing.Size(136, 16)
        Me.chkudt.TabIndex = 0
        Me.chkudt.Text = "Use database time"
        '
        'frmtimesheetmanager
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(328, 338)
        Me.Controls.Add(Me.chkudt)
        Me.Controls.Add(Me.pnltimesheet)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmtimesheetmanager"
        Me.Text = "Timesheet manager"
        Me.pnltimesheet.ResumeLayout(False)
        Me.grpdts.ResumeLayout(False)
        Me.grppd.ResumeLayout(False)
        CType(Me.nud, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
#Region "public members"
    Public strpersonnel As String
#End Region

    Private Sub frmtimesheetmanager_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub btnsave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Try
            MsgBox(strpersonnel)
        Catch cv As Exception

        End Try
    End Sub
    Private Sub ls() 'loadsettings
        Try
            Dim strin As String = Me.strpersonnel
            Dim arr() As String
            arr = strin.Split(",")
            Dim int As Integer = arr.GetUpperBound(0)
            strin = arr(int)

        Catch ex As Exception

        End Try
    End Sub
End Class
