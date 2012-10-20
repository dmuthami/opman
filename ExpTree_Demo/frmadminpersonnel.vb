Public Class frmadminpersonnel
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
    Friend WithEvents StiGroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents StiCheckBox16 As System.Windows.Forms.CheckBox
    Friend WithEvents StiCheckBox15 As System.Windows.Forms.CheckBox
    Friend WithEvents StiCheckBox14 As System.Windows.Forms.CheckBox
    Friend WithEvents cbreportpersonnel As System.Windows.Forms.CheckBox
    Friend WithEvents cbeditpersonnel As System.Windows.Forms.CheckBox
    Friend WithEvents cbviewonly As System.Windows.Forms.CheckBox
    Friend WithEvents cbmakechanges As System.Windows.Forms.CheckBox
    Friend WithEvents btnok As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents grpchanges As System.Windows.Forms.GroupBox
    Friend WithEvents chkitadminrights As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmadminpersonnel))
        Me.StiGroupBox1 = New System.Windows.Forms.GroupBox
        Me.grpchanges = New System.Windows.Forms.GroupBox
        Me.StiCheckBox16 = New System.Windows.Forms.CheckBox
        Me.StiCheckBox15 = New System.Windows.Forms.CheckBox
        Me.StiCheckBox14 = New System.Windows.Forms.CheckBox
        Me.chkitadminrights = New System.Windows.Forms.CheckBox
        Me.cbreportpersonnel = New System.Windows.Forms.CheckBox
        Me.cbeditpersonnel = New System.Windows.Forms.CheckBox
        Me.cbviewonly = New System.Windows.Forms.CheckBox
        Me.cbmakechanges = New System.Windows.Forms.CheckBox
        Me.btnok = New System.Windows.Forms.Button
        Me.btnclose = New System.Windows.Forms.Button
        Me.StiGroupBox1.SuspendLayout()
        Me.grpchanges.SuspendLayout()
        Me.SuspendLayout()
        '
        'StiGroupBox1
        '
        Me.StiGroupBox1.Controls.Add(Me.grpchanges)
        Me.StiGroupBox1.Controls.Add(Me.cbviewonly)
        Me.StiGroupBox1.Controls.Add(Me.cbmakechanges)
        Me.StiGroupBox1.Controls.Add(Me.btnok)
        Me.StiGroupBox1.Controls.Add(Me.btnclose)
        Me.StiGroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiGroupBox1.Location = New System.Drawing.Point(2, -1)
        Me.StiGroupBox1.Name = "StiGroupBox1"
        Me.StiGroupBox1.Size = New System.Drawing.Size(374, 169)
        Me.StiGroupBox1.TabIndex = 0
        Me.StiGroupBox1.TabStop = False
        Me.StiGroupBox1.Text = "Personnel"
        '
        'grpchanges
        '
        Me.grpchanges.Controls.Add(Me.StiCheckBox16)
        Me.grpchanges.Controls.Add(Me.StiCheckBox15)
        Me.grpchanges.Controls.Add(Me.StiCheckBox14)
        Me.grpchanges.Controls.Add(Me.chkitadminrights)
        Me.grpchanges.Controls.Add(Me.cbreportpersonnel)
        Me.grpchanges.Controls.Add(Me.cbeditpersonnel)
        Me.grpchanges.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpchanges.Location = New System.Drawing.Point(8, 36)
        Me.grpchanges.Name = "grpchanges"
        Me.grpchanges.Size = New System.Drawing.Size(360, 96)
        Me.grpchanges.TabIndex = 3
        Me.grpchanges.TabStop = False
        Me.grpchanges.Text = "Pick changes to apply"
        '
        'StiCheckBox16
        '
        Me.StiCheckBox16.Location = New System.Drawing.Point(192, 64)
        Me.StiCheckBox16.Name = "StiCheckBox16"
        Me.StiCheckBox16.Size = New System.Drawing.Size(144, 16)
        Me.StiCheckBox16.TabIndex = 9
        Me.StiCheckBox16.Text = "View only"
        '
        'StiCheckBox15
        '
        Me.StiCheckBox15.Location = New System.Drawing.Point(192, 40)
        Me.StiCheckBox15.Name = "StiCheckBox15"
        Me.StiCheckBox15.Size = New System.Drawing.Size(144, 16)
        Me.StiCheckBox15.TabIndex = 8
        Me.StiCheckBox15.Text = "View only"
        '
        'StiCheckBox14
        '
        Me.StiCheckBox14.Location = New System.Drawing.Point(192, 16)
        Me.StiCheckBox14.Name = "StiCheckBox14"
        Me.StiCheckBox14.Size = New System.Drawing.Size(144, 16)
        Me.StiCheckBox14.TabIndex = 7
        Me.StiCheckBox14.Text = "View only"
        '
        'chkitadminrights
        '
        Me.chkitadminrights.Location = New System.Drawing.Point(16, 64)
        Me.chkitadminrights.Name = "chkitadminrights"
        Me.chkitadminrights.Size = New System.Drawing.Size(176, 16)
        Me.chkitadminrights.TabIndex = 6
        Me.chkitadminrights.Text = "IT administrative rights"
        '
        'cbreportpersonnel
        '
        Me.cbreportpersonnel.Location = New System.Drawing.Point(16, 40)
        Me.cbreportpersonnel.Name = "cbreportpersonnel"
        Me.cbreportpersonnel.Size = New System.Drawing.Size(168, 16)
        Me.cbreportpersonnel.TabIndex = 5
        Me.cbreportpersonnel.Text = "Print reports on personnel"
        '
        'cbeditpersonnel
        '
        Me.cbeditpersonnel.Location = New System.Drawing.Point(16, 16)
        Me.cbeditpersonnel.Name = "cbeditpersonnel"
        Me.cbeditpersonnel.Size = New System.Drawing.Size(168, 16)
        Me.cbeditpersonnel.TabIndex = 4
        Me.cbeditpersonnel.Text = "Add and edit personnel"
        '
        'cbviewonly
        '
        Me.cbviewonly.Location = New System.Drawing.Point(8, 15)
        Me.cbviewonly.Name = "cbviewonly"
        Me.cbviewonly.Size = New System.Drawing.Size(104, 16)
        Me.cbviewonly.TabIndex = 1
        Me.cbviewonly.Text = "View only"
        '
        'cbmakechanges
        '
        Me.cbmakechanges.Location = New System.Drawing.Point(128, 16)
        Me.cbmakechanges.Name = "cbmakechanges"
        Me.cbmakechanges.Size = New System.Drawing.Size(200, 16)
        Me.cbmakechanges.TabIndex = 2
        Me.cbmakechanges.Text = "Can make changes to personnel"
        '
        'btnok
        '
        Me.btnok.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnok.Location = New System.Drawing.Point(8, 136)
        Me.btnok.Name = "btnok"
        Me.btnok.TabIndex = 10
        Me.btnok.Text = "Ok"
        '
        'btnclose
        '
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(292, 136)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 11
        Me.btnclose.Text = "Close"
        '
        'frmadminpersonnel
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(384, 170)
        Me.Controls.Add(Me.StiGroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmadminpersonnel"
        Me.Text = "administer rights to personnel"
        Me.TopMost = True
        Me.StiGroupBox1.ResumeLayout(False)
        Me.grpchanges.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
#Region "private members"
    Public strin As String
#End Region

    Private Sub btnok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnok.Click
        Try
            Dim strf As String
            If Me.cbviewonly.Checked = True Then
                strf = "1,0,0,0,0,0"
            Else
                strf = "1"
                If Me.cbeditpersonnel.Checked = True Then
                    strf += ",1"
                Else
                    strf += ",0"
                End If
                If Me.cbreportpersonnel.Checked = True Then
                    strf += ",1"
                Else
                    strf += ",0"
                End If
                If Me.chkitadminrights.Checked = True Then
                    strf += ",1"
                Else
                    strf += ",0"
                End If
                strf += ",0,0"
            End If
            myForms.admin.strpersonnel = strf
        Catch xc As Exception

        End Try
    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            Me.Close()
        Catch cv As Exception

        End Try
    End Sub
    Private Sub cbviewonly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbviewonly.Click
        Try
            If cbviewonly.Checked = True Then
                Me.cbmakechanges.Checked = False
                Me.grpchanges.Enabled = False
            Else
                Me.cbmakechanges.Checked = True
                Me.grpchanges.Enabled = True

            End If

        Catch zx As Exception

        End Try
    End Sub
    Private Sub cbmakechanges_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbmakechanges.Click
        Try
            If cbmakechanges.Checked = True Then
                Me.cbviewonly.Checked = False
                Me.grpchanges.Enabled = True

            Else
                Me.cbviewonly.Checked = True
                Me.grpchanges.Enabled = False
            End If

        Catch zx As Exception

        End Try
    End Sub
    Private Sub frmadminpersonnel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim array() As String
            array = strin.Split(",")
            If array(0) = "0" Then
                cbviewonly.Checked = True
                cbmakechanges.Checked = False
                Me.grpchanges.Enabled = False
            Else
                cbviewonly.Checked = False
                cbmakechanges.Checked = True
                Me.grpchanges.Enabled = True
                If array(1) = "1" Then
                    Me.cbeditpersonnel.Checked = True
                Else
                    Me.cbeditpersonnel.Checked = False
                End If
                If array(2) = "1" Then
                    Me.cbreportpersonnel.Checked = True
                Else
                    Me.cbreportpersonnel.Checked = False
                End If
            End If


        Catch ax As Exception

        End Try
    End Sub

End Class
