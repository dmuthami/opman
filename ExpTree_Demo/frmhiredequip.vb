Imports System
Imports ADODB


Public Class frmhiredequip
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtcost As VSEssentials.LabelTextBox
    Friend WithEvents dtpreleasetime As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpssigntime As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpreleasedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpssigndate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txthourlyrate As System.Windows.Forms.TextBox
    Friend WithEvents txtname As System.Windows.Forms.TextBox
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmhiredequip))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtcost = New VSEssentials.LabelTextBox
        Me.dtpreleasetime = New System.Windows.Forms.DateTimePicker
        Me.dtpssigntime = New System.Windows.Forms.DateTimePicker
        Me.dtpreleasedate = New System.Windows.Forms.DateTimePicker
        Me.dtpssigndate = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txthourlyrate = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtname = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnadd = New System.Windows.Forms.Button
        Me.btnclose = New System.Windows.Forms.Button
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(352, 284)
        Me.Panel1.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtcost)
        Me.GroupBox3.Controls.Add(Me.dtpreleasetime)
        Me.GroupBox3.Controls.Add(Me.dtpssigntime)
        Me.GroupBox3.Controls.Add(Me.dtpreleasedate)
        Me.GroupBox3.Controls.Add(Me.dtpssigndate)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.txthourlyrate)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(8, 159)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(336, 121)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        '
        'txtcost
        '
        Me.txtcost.LabelText = "Total  cost"
        Me.txtcost.Location = New System.Drawing.Point(176, 82)
        Me.txtcost.Name = "txtcost"
        Me.txtcost.Size = New System.Drawing.Size(152, 24)
        Me.txtcost.TabIndex = 10
        Me.txtcost.TextBoxText = ""
        '
        'dtpreleasetime
        '
        Me.dtpreleasetime.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtpreleasetime.Location = New System.Drawing.Point(240, 47)
        Me.dtpreleasetime.Name = "dtpreleasetime"
        Me.dtpreleasetime.ShowUpDown = True
        Me.dtpreleasetime.Size = New System.Drawing.Size(88, 20)
        Me.dtpreleasetime.TabIndex = 8
        '
        'dtpssigntime
        '
        Me.dtpssigntime.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtpssigntime.Location = New System.Drawing.Point(240, 10)
        Me.dtpssigntime.Name = "dtpssigntime"
        Me.dtpssigntime.ShowUpDown = True
        Me.dtpssigntime.Size = New System.Drawing.Size(88, 20)
        Me.dtpssigntime.TabIndex = 6
        '
        'dtpreleasedate
        '
        Me.dtpreleasedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpreleasedate.Location = New System.Drawing.Point(80, 45)
        Me.dtpreleasedate.Name = "dtpreleasedate"
        Me.dtpreleasedate.Size = New System.Drawing.Size(88, 20)
        Me.dtpreleasedate.TabIndex = 7
        '
        'dtpssigndate
        '
        Me.dtpssigndate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpssigndate.Location = New System.Drawing.Point(80, 10)
        Me.dtpssigndate.Name = "dtpssigndate"
        Me.dtpssigndate.Size = New System.Drawing.Size(88, 20)
        Me.dtpssigndate.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(176, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 30)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Release time"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(176, 10)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 30)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Assign time"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(9, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 30)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Release date"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(9, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 20)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Assign date"
        '
        'txthourlyrate
        '
        Me.txthourlyrate.Location = New System.Drawing.Point(80, 83)
        Me.txthourlyrate.Name = "txthourlyrate"
        Me.txthourlyrate.Size = New System.Drawing.Size(88, 20)
        Me.txthourlyrate.TabIndex = 9
        Me.txthourlyrate.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(9, 83)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 20)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Hourly rate"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtname)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(8, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(336, 40)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'txtname
        '
        Me.txtname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtname.Location = New System.Drawing.Point(113, 13)
        Me.txtname.Name = "txtname"
        Me.txtname.Size = New System.Drawing.Size(215, 20)
        Me.txtname.TabIndex = 1
        Me.txtname.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(9, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 20)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Equipment name"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtdesc)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 44)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(336, 112)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Description"
        '
        'txtdesc
        '
        Me.txtdesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdesc.Location = New System.Drawing.Point(8, 16)
        Me.txtdesc.Multiline = True
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(320, 88)
        Me.txtdesc.TabIndex = 3
        Me.txtdesc.Text = ""
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnadd)
        Me.Panel2.Controls.Add(Me.btnclose)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 284)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(352, 32)
        Me.Panel2.TabIndex = 11
        '
        'btnadd
        '
        Me.btnadd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnadd.Location = New System.Drawing.Point(5, 6)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.TabIndex = 12
        Me.btnadd.Text = "Add"
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(272, 6)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 13
        Me.btnclose.Text = "Close"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmhiredequip
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(352, 316)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "frmhiredequip"
        Me.Text = "Assign hired equipment"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
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
    Public jobno As String
    Public autono As String
    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        Dim Thread566 As System.Threading.Thread
        Try
            Try
                Dim connectstr As String
                connectstr = "DSN=" & myForms.qconnstr
                'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
                Dim connect As New ADODB.Connection
                connect.Mode = ConnectModeEnum.adModeReadWrite
                connect.CursorLocation = CursorLocationEnum.adUseClient
                connect.ConnectionString = connectstr
                connect.Open()
                Dim asubstring As String
                Dim bsubstring As String
                asubstring = Me.dtpssigndate.Value.Year & "-" _
                                 & dtpssigndate.Value.Month & "-" _
                                 & dtpssigndate.Value.Day & " "
                asubstring += Me.dtpssigntime.Text
                bsubstring = Me.dtpreleasedate.Value.Year & "-" _
                                            & dtpreleasedate.Value.Month & "-" _
                                            & dtpreleasedate.Value.Day & " "
                bsubstring += Me.dtpreleasetime.Text
                Dim strsql As String
                Dim Tasks As New taskclass
                'Tasks.mid = Mid()
                'Dim a() As String
                'a = txtdesc.Lines
                'Dim desc As String = rml(a)
                Dim isadd As Boolean = True
                '-------------------
                Dim arr() As String
                Dim strr As String
                Dim y As Integer
                txtdesc.Text = Me.txtdesc.Text.Trim()
                arr = txtdesc.Lines
                y = arr.GetUpperBound(0)
                Dim alpha As Integer
                For alpha = 0 To y
                    strr += arr(alpha) + vbCrLf
                    Application.DoEvents()
                Next
                '----------------------------------
                If Me.btnadd.Text = "Add" Then
                    strsql = "insert into hiredequip"
                    strsql += "( job_no,equipname,description,assigndate,releasedate,hourly_rate) values "
                    strsql += " ( '" & jobno & "','" & Me.txtname.Text & "',  '" & strr & "','" & asubstring & "','" & bsubstring & "'," _
                    & "'" & Me.txthourlyrate.Text.Trim & "')"
                Else
                    strsql = "update hiredequip set"
                    strsql += " job_no='" & jobno & "', equipname='" & txtname.Text.Trim & "'" _
                    & ",description='" & strr & "',assigndate='" & asubstring & "'" _
                    & ",hourly_rate='" & Me.txthourlyrate.Text & "'," _
                    & "releasedate='" & bsubstring & "'"
                    strsql += " where ano = '" & Integer.Parse(autono) & "';"
                    isadd = False
                End If
                Thread566 = New System.Threading.Thread( _
                AddressOf Tasks.hiredequipinvoke)

                connect.BeginTrans()
                connect.Execute(strsql)
                connect.CommitTrans()
                Try
                    connect.Close()
                Catch es As Exception
                End Try
            Catch ex As Exception
            Finally

            End Try
            Thread566.IsBackground = True
            Thread566.Start()
        Catch qw As Exception
        End Try
        Try
            myForms.CustomerForm2.mygross()
        Catch zx As Exception

        End Try
    End Sub
    Private Sub frmhiredequip_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.hired = Nothing
            Me.Close()

        Catch ex As Exception
        End Try
    End Sub

#Region "validation"
    Private Sub txtcost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcost.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtcost, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtcost, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtname.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtname, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtname, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtdesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdesc.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtdesc, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtdesc, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

End Class
