Public Class frmRamaniequip
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents DataGrid5 As System.Windows.Forms.DataGrid
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents StiButton1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRamaniequip))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.DataGrid5 = New System.Windows.Forms.DataGrid
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.StiButton1 = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        CType(Me.DataGrid5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.DataGrid5)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(384, 194)
        Me.Panel1.TabIndex = 2
        '
        'DataGrid5
        '
        Me.DataGrid5.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid5.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGrid5.CaptionVisible = False
        Me.DataGrid5.DataMember = ""
        Me.DataGrid5.FlatMode = True
        Me.DataGrid5.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid5.Location = New System.Drawing.Point(8, 24)
        Me.DataGrid5.Name = "DataGrid5"
        Me.DataGrid5.PreferredRowHeight = 20
        Me.DataGrid5.ReadOnly = True
        Me.DataGrid5.Size = New System.Drawing.Size(368, 160)
        Me.DataGrid5.TabIndex = 14
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.StiButton1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 194)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(384, 32)
        Me.Panel2.TabIndex = 1
        '
        'StiButton1
        '
        Me.StiButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StiButton1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiButton1.Location = New System.Drawing.Point(304, 6)
        Me.StiButton1.Name = "StiButton1"
        Me.StiButton1.TabIndex = 0
        Me.StiButton1.Text = "Close"
        '
        'frmRamaniequip
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(384, 226)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRamaniequip"
        Me.Text = "Assign Ramani equipment"
        Me.Panel1.ResumeLayout(False)
        CType(Me.DataGrid5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
End Class
