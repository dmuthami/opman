Public Class frmadminjobs
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
    Friend WithEvents StiCheckBox13 As System.Windows.Forms.CheckBox
    Friend WithEvents btnok As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents cbreportjobs As System.Windows.Forms.CheckBox
    Friend WithEvents cbeditjobs As System.Windows.Forms.CheckBox
    Friend WithEvents cbviewonly As System.Windows.Forms.CheckBox
    Friend WithEvents cbmakechanges As System.Windows.Forms.CheckBox
    Friend WithEvents grpchanges As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmadminjobs))
        Me.StiGroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnok = New System.Windows.Forms.Button
        Me.btnclose = New System.Windows.Forms.Button
        Me.grpchanges = New System.Windows.Forms.GroupBox
        Me.StiCheckBox16 = New System.Windows.Forms.CheckBox
        Me.StiCheckBox15 = New System.Windows.Forms.CheckBox
        Me.StiCheckBox14 = New System.Windows.Forms.CheckBox
        Me.StiCheckBox13 = New System.Windows.Forms.CheckBox
        Me.cbreportjobs = New System.Windows.Forms.CheckBox
        Me.cbeditjobs = New System.Windows.Forms.CheckBox
        Me.cbviewonly = New System.Windows.Forms.CheckBox
        Me.cbmakechanges = New System.Windows.Forms.CheckBox
        Me.StiGroupBox1.SuspendLayout()
        Me.grpchanges.SuspendLayout()
        Me.SuspendLayout()
        '
        'StiGroupBox1
        '
        Me.StiGroupBox1.Controls.Add(Me.btnok)
        Me.StiGroupBox1.Controls.Add(Me.btnclose)
        Me.StiGroupBox1.Controls.Add(Me.grpchanges)
        Me.StiGroupBox1.Controls.Add(Me.cbviewonly)
        Me.StiGroupBox1.Controls.Add(Me.cbmakechanges)
        Me.StiGroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiGroupBox1.Location = New System.Drawing.Point(2, -1)
        Me.StiGroupBox1.Name = "StiGroupBox1"
        Me.StiGroupBox1.Size = New System.Drawing.Size(360, 169)
        Me.StiGroupBox1.TabIndex = 0
        Me.StiGroupBox1.TabStop = False
        Me.StiGroupBox1.Text = "Jobs"
        '
        'btnok
        '
        Me.btnok.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnok.Location = New System.Drawing.Point(8, 135)
        Me.btnok.Name = "btnok"
        Me.btnok.TabIndex = 10
        Me.btnok.Text = "Ok"
        '
        'btnclose
        '
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(272, 135)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 11
        Me.btnclose.Text = "Close"
        '
        'grpchanges
        '
        Me.grpchanges.Controls.Add(Me.StiCheckBox16)
        Me.grpchanges.Controls.Add(Me.StiCheckBox15)
        Me.grpchanges.Controls.Add(Me.StiCheckBox14)
        Me.grpchanges.Controls.Add(Me.StiCheckBox13)
        Me.grpchanges.Controls.Add(Me.cbreportjobs)
        Me.grpchanges.Controls.Add(Me.cbeditjobs)
        Me.grpchanges.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpchanges.Location = New System.Drawing.Point(8, 36)
        Me.grpchanges.Name = "grpchanges"
        Me.grpchanges.Size = New System.Drawing.Size(344, 96)
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
        'StiCheckBox13
        '
        Me.StiCheckBox13.Location = New System.Drawing.Point(16, 64)
        Me.StiCheckBox13.Name = "StiCheckBox13"
        Me.StiCheckBox13.Size = New System.Drawing.Size(176, 16)
        Me.StiCheckBox13.TabIndex = 6
        Me.StiCheckBox13.Text = "View only"
        '
        'cbreportjobs
        '
        Me.cbreportjobs.Location = New System.Drawing.Point(16, 40)
        Me.cbreportjobs.Name = "cbreportjobs"
        Me.cbreportjobs.Size = New System.Drawing.Size(168, 16)
        Me.cbreportjobs.TabIndex = 5
        Me.cbreportjobs.Text = "Print reports on jobs"
        '
        'cbeditjobs
        '
        Me.cbeditjobs.Location = New System.Drawing.Point(16, 16)
        Me.cbeditjobs.Name = "cbeditjobs"
        Me.cbeditjobs.Size = New System.Drawing.Size(168, 16)
        Me.cbeditjobs.TabIndex = 4
        Me.cbeditjobs.Text = "Add and edit jobs"
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
        Me.cbmakechanges.Size = New System.Drawing.Size(160, 16)
        Me.cbmakechanges.TabIndex = 2
        Me.cbmakechanges.Text = "Can make changes to jobs"
        '
        'frmadminjobs
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(368, 170)
        Me.Controls.Add(Me.StiGroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmadminjobs"
        Me.Text = "Administer rights to jobs"
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

    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            Me.Close()
        Catch ez As Exception

        End Try
    End Sub
    Private Sub btnok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnok.Click
        Try
            Dim strf As String
            If Me.cbviewonly.Checked = True Then
                strf = "1,0,0,0,0,0"
            Else
                strf = "1"
                If Me.cbeditjobs.Checked = True Then
                    strf += ",1"
                Else
                    strf += ",0"
                End If
                If Me.cbreportjobs.Checked = True Then
                    strf += ",1"
                Else
                    strf += ",0"
                End If

                strf += ",0,0,0"
            End If
            myForms.admin.strjobs = strf
        Catch xc As Exception

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

    Private Sub frmadminjobs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
                    Me.cbeditjobs.Checked = True
                Else
                    Me.cbeditjobs.Checked = False
                End If
                If array(2) = "1" Then
                    Me.cbreportjobs.Checked = True
                Else
                    Me.cbreportjobs.Checked = False
                End If
            End If


        Catch ax As Exception

        End Try
    End Sub
End Class
