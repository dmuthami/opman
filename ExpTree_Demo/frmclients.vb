Public Class frmclients
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
    Friend WithEvents grpcontactgrid As System.Windows.Forms.GroupBox
    Friend WithEvents btnshowall As System.Windows.Forms.Button
    Friend WithEvents btnaddnew As System.Windows.Forms.Button
    Friend WithEvents dtgClients As System.Windows.Forms.DataGrid
    Friend WithEvents grpContactsearch As System.Windows.Forms.GroupBox
    Friend WithEvents btnsearchname As System.Windows.Forms.Button
    Friend WithEvents cboclientssearchfield As System.Windows.Forms.ComboBox
    Friend WithEvents txtparams As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmclients))
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
        Me.grpcontactgrid.SuspendLayout()
        CType(Me.dtgClients, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpContactsearch.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpcontactgrid
        '
        Me.grpcontactgrid.Controls.Add(Me.btnshowall)
        Me.grpcontactgrid.Controls.Add(Me.btnaddnew)
        Me.grpcontactgrid.Controls.Add(Me.dtgClients)
        Me.grpcontactgrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpcontactgrid.Location = New System.Drawing.Point(0, 83)
        Me.grpcontactgrid.Name = "grpcontactgrid"
        Me.grpcontactgrid.Size = New System.Drawing.Size(584, 375)
        Me.grpcontactgrid.TabIndex = 5
        Me.grpcontactgrid.TabStop = False
        '
        'btnshowall
        '
        Me.btnshowall.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnshowall.Location = New System.Drawing.Point(120, 13)
        Me.btnshowall.Name = "btnshowall"
        Me.btnshowall.Size = New System.Drawing.Size(72, 24)
        Me.btnshowall.TabIndex = 16
        Me.btnshowall.Text = "Show all"
        '
        'btnaddnew
        '
        Me.btnaddnew.FlatStyle = System.Windows.Forms.FlatStyle.System
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
        Me.dtgClients.Size = New System.Drawing.Size(561, 327)
        Me.dtgClients.TabIndex = 5
        '
        'grpContactsearch
        '
        Me.grpContactsearch.Controls.Add(Me.btnsearchname)
        Me.grpContactsearch.Controls.Add(Me.cboclientssearchfield)
        Me.grpContactsearch.Controls.Add(Me.txtparams)
        Me.grpContactsearch.Controls.Add(Me.Label3)
        Me.grpContactsearch.Controls.Add(Me.Label4)
        Me.grpContactsearch.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpContactsearch.Location = New System.Drawing.Point(0, 0)
        Me.grpContactsearch.Name = "grpContactsearch"
        Me.grpContactsearch.Size = New System.Drawing.Size(584, 83)
        Me.grpContactsearch.TabIndex = 4
        Me.grpContactsearch.TabStop = False
        '
        'btnsearchname
        '
        Me.btnsearchname.FlatStyle = System.Windows.Forms.FlatStyle.System
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
        Me.cboclientssearchfield.Size = New System.Drawing.Size(152, 21)
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
        'frmclients
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 458)
        Me.Controls.Add(Me.grpcontactgrid)
        Me.Controls.Add(Me.grpContactsearch)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmclients"
        Me.Text = "frmclients"
        Me.grpcontactgrid.ResumeLayout(False)
        CType(Me.dtgClients, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpContactsearch.ResumeLayout(False)
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
