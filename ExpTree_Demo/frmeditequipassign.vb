Public Class frmeditequipassign
    Inherits System.Windows.Forms.Form
    Public htirow
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
            Try
                myForms.iseditassignequip = False
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dtppurchasedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpreleasedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtequipid As System.Windows.Forms.TextBox
    Friend WithEvents txtmodelname As System.Windows.Forms.TextBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtptime As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmeditequipassign))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtmodelname = New System.Windows.Forms.TextBox
        Me.txtequipid = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dtptime = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtpreleasedate = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.dtppurchasedate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtmodelname)
        Me.GroupBox1.Controls.Add(Me.txtequipid)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.btnSave)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(354, 156)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txtmodelname
        '
        Me.txtmodelname.Enabled = False
        Me.txtmodelname.Location = New System.Drawing.Point(152, 37)
        Me.txtmodelname.Name = "txtmodelname"
        Me.txtmodelname.Size = New System.Drawing.Size(192, 20)
        Me.txtmodelname.TabIndex = 2
        Me.txtmodelname.Text = ""
        '
        'txtequipid
        '
        Me.txtequipid.Enabled = False
        Me.txtequipid.Location = New System.Drawing.Point(152, 12)
        Me.txtequipid.Name = "txtequipid"
        Me.txtequipid.Size = New System.Drawing.Size(192, 20)
        Me.txtequipid.TabIndex = 1
        Me.txtequipid.Text = ""
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Location = New System.Drawing.Point(279, 133)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(72, 20)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSave.Location = New System.Drawing.Point(4, 133)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(56, 20)
        Me.btnSave.TabIndex = 7
        Me.btnSave.Text = "Ok"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dtptime)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.dtpreleasedate)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.dtppurchasedate)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(0, 64)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(352, 64)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Please pick the relevant dates"
        '
        'dtptime
        '
        Me.dtptime.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtptime.Location = New System.Drawing.Point(256, 15)
        Me.dtptime.Name = "dtptime"
        Me.dtptime.ShowUpDown = True
        Me.dtptime.Size = New System.Drawing.Size(88, 20)
        Me.dtptime.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(168, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Time Assigned"
        '
        'dtpreleasedate
        '
        Me.dtpreleasedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpreleasedate.Location = New System.Drawing.Point(176, 40)
        Me.dtpreleasedate.Name = "dtpreleasedate"
        Me.dtpreleasedate.Size = New System.Drawing.Size(168, 20)
        Me.dtpreleasedate.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(7, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(161, 16)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Estimated Release Date"
        '
        'dtppurchasedate
        '
        Me.dtppurchasedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtppurchasedate.Location = New System.Drawing.Point(88, 16)
        Me.dtppurchasedate.Name = "dtppurchasedate"
        Me.dtppurchasedate.Size = New System.Drawing.Size(80, 20)
        Me.dtppurchasedate.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(-2, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Date Assigned"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(0, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Model Name"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Equipment Id"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        '
        'frmeditequipassign
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(354, 156)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmeditequipassign"
        Me.Text = "Edit equipment assignment"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            myForms.iseditassignequip = False
            Me.Dispose(True)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xcv As Exception

        End Try
        Try
            Dim ds3 As System.Data.DataSet = New System.Data.DataSet
            ds3 = myForms.tojobs.dtgequip.DataSource
            Dim sdate1, sdate2 As String
            sdate1 = dtppurchasedate.Value.Year & "-" _
                   & dtppurchasedate.Value.Month & "-" _
                   & dtppurchasedate.Value.Day & " " _
                   & dtptime.Value.Hour & ":" _
                   & dtptime.Value.Minute & ":" _
                   & dtptime.Value.Second

            sdate2 = dtpreleasedate.Value.Year & "-" _
                 & dtpreleasedate.Value.Month & "-" _
                 & dtpreleasedate.Value.Day & " " _
                 & dtpreleasedate.Value.Hour & ":" _
                 & dtpreleasedate.Value.Minute & ":" _
                 & dtpreleasedate.Value.Second
            ds3.Tables(0).Rows(htirow).Item("date_assigned") = sdate1
            ds3.Tables(0).Rows(htirow).Item("estimate_release_date") = sdate2
            ds3.Tables(0).AcceptChanges()
            MessageBox.Show("Very successful", "Updates", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub frmeditequipassign_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

#Region "validation"
    Private Sub txtequipid_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtequipid.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtequipid, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtequipid, "")
            End If
        Catch xc As Exception

        End Try
    End Sub

    Private Sub txtmodelname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmodelname.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmodelname, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmodelname, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

End Class
