Imports System
Imports ADODB

Public Class frmtimeoff
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
            myForms.timeoff = Nothing
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents dtptimeoff As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpdayoff As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents txtname As System.Windows.Forms.TextBox
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmtimeoff))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtptimeoff = New System.Windows.Forms.DateTimePicker
        Me.dtpdayoff = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.txtname = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnadd = New System.Windows.Forms.Button
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(400, 274)
        Me.Panel1.TabIndex = 16
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtptimeoff)
        Me.GroupBox1.Controls.Add(Me.dtpdayoff)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.txtname)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(400, 274)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'dtptimeoff
        '
        Me.dtptimeoff.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtptimeoff.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtptimeoff.Location = New System.Drawing.Point(296, 242)
        Me.dtptimeoff.Name = "dtptimeoff"
        Me.dtptimeoff.ShowUpDown = True
        Me.dtptimeoff.Size = New System.Drawing.Size(96, 20)
        Me.dtptimeoff.TabIndex = 5
        Me.dtptimeoff.Value = New Date(2006, 4, 28, 9, 58, 58, 500)
        '
        'dtpdayoff
        '
        Me.dtpdayoff.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpdayoff.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpdayoff.Location = New System.Drawing.Point(88, 242)
        Me.dtpdayoff.Name = "dtpdayoff"
        Me.dtpdayoff.Size = New System.Drawing.Size(112, 20)
        Me.dtpdayoff.TabIndex = 4
        Me.dtpdayoff.Value = New Date(2006, 4, 28, 9, 58, 58, 515)
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(208, 242)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 20)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Time off"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(6, 242)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 20)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Day off"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.txtdesc)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(1, 42)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(391, 200)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Description"
        '
        'txtdesc
        '
        Me.txtdesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdesc.Location = New System.Drawing.Point(8, 16)
        Me.txtdesc.Multiline = True
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(375, 178)
        Me.txtdesc.TabIndex = 3
        Me.txtdesc.Text = ""
        '
        'txtname
        '
        Me.txtname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtname.Location = New System.Drawing.Point(112, 16)
        Me.txtname.Name = "txtname"
        Me.txtname.Size = New System.Drawing.Size(280, 20)
        Me.txtname.TabIndex = 1
        Me.txtname.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 20)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Employees name"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnclose)
        Me.Panel2.Controls.Add(Me.btnadd)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 274)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(400, 32)
        Me.Panel2.TabIndex = 6
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(320, 6)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 8
        Me.btnclose.Text = "Close"
        '
        'btnadd
        '
        Me.btnadd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnadd.Location = New System.Drawing.Point(3, 4)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.TabIndex = 7
        Me.btnadd.Text = "Add"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmtimeoff
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(400, 306)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmtimeoff"
        Me.Text = "Time off"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
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
    Public mid As String
    Public autono As String

    Private Sub frmtimeoff_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.timeoff = Nothing
            Me.Close()
        Catch we As Exception

        End Try
    End Sub
    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        Try
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
                    Dim sdate As String 'date of employment
                    sdate = Me.dtpdayoff.Value.Year & "-" _
                     & dtpdayoff.Value.Month & "-" _
                     & dtpdayoff.Value.Day & " " _
                     & dtpdayoff.Value.Hour & ":" _
                     & dtpdayoff.Value.Minute & ":" _
                     & dtpdayoff.Value.Second

                    Dim strsql As String
                    Dim Tasks As New taskclass
                    Tasks.mid = mid
                    'Dim a() As String
                    'a = txtdesc.Lines
                    'Dim desc As String = rml(a)
                    '-------------------
                    Dim arr() As String
                    Dim strr, strr2 As String
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
                        If Me.Text = "Add Day off" Then
                            strsql = "insert into dayoff"
                            strsql += "( idno,description,dateoff) values "
                            strsql += " ( '" & mid & "', '" & strr.Trim & "','" & sdate & "')"
                            Thread566 = New System.Threading.Thread( _
                            AddressOf Tasks.dayoffinvoke)
                        Else
                            strsql = "insert into timeoff "
                            strsql += "( idno,description,dateoff,timeoff) values "
                            strsql += " ( '" & mid & "', '" & strr.Trim & "','" & sdate & "','" & dtptimeoff.Value & "')"
                            Thread566 = New System.Threading.Thread( _
                             AddressOf Tasks.timeoffinvoke)
                        End If
                        txtdesc.Text = ""
                    Else
                        If Me.Text = "Edit time off" Then
                            strsql = "update timeoff  set"
                            strsql += " idno='" & mid & "', description='" & strr.Trim & "'," _
                                      & " dateoff='" & sdate & "',timeoff='" & dtptimeoff.Value & "',ano='" & Integer.Parse(autono) & "'"
                            strsql += " where ano = '" & Integer.Parse(autono) & "';"
                            Thread566 = New System.Threading.Thread( _
                             AddressOf Tasks.timeoffinvoke)

                        Else
                            strsql = "update dayoff  set"
                            strsql += " idno='" & mid & "', description='" & strr.Trim & "'," _
                                      & " dateoff='" & sdate & "',ano='" & Integer.Parse(autono) & "'"
                            strsql += " where ano = '" & Integer.Parse(autono) & "';"
                            Thread566 = New System.Threading.Thread( _
                             AddressOf Tasks.dayoffinvoke)

                        End If

                    End If

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
        Catch we As Exception

        End Try
    End Sub

#Region "validation"
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
#End Region

End Class
