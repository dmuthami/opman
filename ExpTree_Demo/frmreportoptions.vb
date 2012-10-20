Public Class frmreportoptions
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnok As System.Windows.Forms.Button
    Friend WithEvents chkexport As System.Windows.Forms.CheckBox
    Friend WithEvents chkprint As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkprint = New System.Windows.Forms.CheckBox
        Me.chkexport = New System.Windows.Forms.CheckBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnok = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkprint)
        Me.GroupBox1.Controls.Add(Me.chkexport)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(248, 136)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Choose report option"
        '
        'chkprint
        '
        Me.chkprint.Location = New System.Drawing.Point(4, 50)
        Me.chkprint.Name = "chkprint"
        Me.chkprint.Size = New System.Drawing.Size(196, 16)
        Me.chkprint.TabIndex = 2
        Me.chkprint.Text = "Print"
        '
        'chkexport
        '
        Me.chkexport.Location = New System.Drawing.Point(3, 24)
        Me.chkexport.Name = "chkexport"
        Me.chkexport.Size = New System.Drawing.Size(197, 14)
        Me.chkexport.TabIndex = 1
        Me.chkexport.Text = "Export to excel"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnok)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 138)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(248, 32)
        Me.Panel1.TabIndex = 3
        '
        'btnok
        '
        Me.btnok.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnok.Location = New System.Drawing.Point(6, 6)
        Me.btnok.Name = "btnok"
        Me.btnok.Size = New System.Drawing.Size(66, 23)
        Me.btnok.TabIndex = 4
        Me.btnok.Text = "Ok"
        '
        'frmreportoptions
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(248, 170)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmreportoptions"
        Me.Text = "Report options"
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private Sub btnok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnok.Click
        Try
            If Me.chkexport.Checked = True Then
                myForms.Main.reportoption = "0"
            Else
                myForms.Main.reportoption = "1"
            End If
        Catch we As Exception
        End Try
        Try
            Me.Close()
        Catch rt As Exception
        End Try
    End Sub
    Private Sub frmreportoptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.chkexport.Checked = True
        Catch we As Exception

        End Try
    End Sub
    Private Sub chkexport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkexport.Click
        Try
            If chkexport.Checked = True Then
                chkprint.Checked = False
            Else
                chkprint.Checked = True
            End If
        Catch we As Exception
        End Try
    End Sub
    Private Sub chkprint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkprint.Click
        Try
            If chkprint.Checked = True Then
                chkexport.Checked = False
            Else
                chkexport.Checked = True
            End If
        Catch we As Exception
        End Try
    End Sub
End Class
