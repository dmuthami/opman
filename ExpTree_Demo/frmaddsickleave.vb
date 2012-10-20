Imports System
Imports ADODB

Public Class frmaddsickleave
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
            myForms.sickleave = Nothing
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents dtpedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpsdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents txtemployeename As System.Windows.Forms.TextBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmaddsickleave))
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnadd = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpedate = New System.Windows.Forms.DateTimePicker
        Me.dtpsdate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.txtemployeename = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnclose)
        Me.Panel2.Controls.Add(Me.btnadd)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 268)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(474, 32)
        Me.Panel2.TabIndex = 6
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(394, 6)
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
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(474, 300)
        Me.Panel1.TabIndex = 14
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.dtpedate)
        Me.GroupBox1.Controls.Add(Me.dtpsdate)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.txtemployeename)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(474, 300)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'dtpedate
        '
        Me.dtpedate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpedate.Location = New System.Drawing.Point(304, 248)
        Me.dtpedate.Name = "dtpedate"
        Me.dtpedate.Size = New System.Drawing.Size(168, 20)
        Me.dtpedate.TabIndex = 5
        Me.dtpedate.Value = New Date(2006, 4, 28, 9, 58, 58, 500)
        '
        'dtpsdate
        '
        Me.dtpsdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpsdate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpsdate.Location = New System.Drawing.Point(80, 248)
        Me.dtpsdate.Name = "dtpsdate"
        Me.dtpsdate.Size = New System.Drawing.Size(160, 20)
        Me.dtpsdate.TabIndex = 4
        Me.dtpsdate.Value = New Date(2006, 4, 28, 9, 58, 58, 515)
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(248, 250)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 20)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "End Date"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(8, 247)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 17)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Start date"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.txtdesc)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(1, 40)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(471, 200)
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
        Me.txtdesc.Size = New System.Drawing.Size(455, 178)
        Me.txtdesc.TabIndex = 3
        Me.txtdesc.Text = ""
        '
        'txtemployeename
        '
        Me.txtemployeename.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtemployeename.Location = New System.Drawing.Point(112, 16)
        Me.txtemployeename.Name = "txtemployeename"
        Me.txtemployeename.Size = New System.Drawing.Size(354, 20)
        Me.txtemployeename.TabIndex = 1
        Me.txtemployeename.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 20)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Employees name"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        '
        'frmaddsickleave
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(474, 300)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmaddsickleave"
        Me.Text = "Add leave"
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
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
                Dim sdate As String 'date of employment
                sdate = Me.dtpsdate.Value.Year & "-" _
                 & dtpsdate.Value.Month & "-" _
                 & dtpsdate.Value.Day & " " _
                 & dtpsdate.Value.Hour & ":" _
                 & dtpsdate.Value.Minute & ":" _
                 & dtpsdate.Value.Second
                Dim sdate1 As String 'date of termination
                sdate1 = Me.dtpedate.Value.Year & "-" _
                 & dtpedate.Value.Month & "-" _
                 & dtpedate.Value.Day & " " _
                 & dtpedate.Value.Hour & ":" _
                 & dtpedate.Value.Minute & ":" _
                 & dtpedate.Value.Second
                Dim strsql As String
                Dim Tasks As New taskclass
                Tasks.mid = mid
                'Dim a() As String
                'a = txtdesc.Lines
                'Dim desc As String = rml(a)
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
                    If Me.Text = "Add sick off" Then
                        strsql = "insert into sickoff"
                        strsql += "( idno,description,sdate,edate) values "
                        strsql += " ( '" & mid & "', '" & strr.Trim & "','" & sdate & "','" & sdate1 & "')"
                        Thread566 = New System.Threading.Thread( _
                        AddressOf Tasks.sickoffinvoke)
                    Else
                        strsql = "insert into leaves"
                        strsql += "( idno,description,sdate,edate) values "
                        strsql += " ( '" & mid & "', '" & strr.Trim & "','" & sdate & "','" & sdate1 & "')"
                        Thread566 = New System.Threading.Thread( _
                         AddressOf Tasks.leavesinvoke)
                    End If
                    txtdesc.Text = ""
                Else
                    If Me.Text = "Edit sick off" Then
                        strsql = "update sickoff set "
                        strsql += "  idno='" & mid & "', description='" & strr.Trim & "'" _
                        & " ,sdate='" & sdate & "',edate='" & sdate1 & "',ano='" & Integer.Parse(autono) & "'"
                        strsql += " where ano = '" & Integer.Parse(autono) & "';"
                        Thread566 = New System.Threading.Thread( _
                        AddressOf Tasks.sickoffinvoke)
                    Else
                        strsql = "update leaves set "
                        strsql += "  idno='" & mid & "', description='" & strr.Trim & "'" _
                        & " ,sdate='" & sdate & "',edate='" & sdate1 & "',ano='" & Integer.Parse(autono) & "'"
                        strsql += " where ano = '" & Integer.Parse(autono) & "';"
                        Thread566 = New System.Threading.Thread( _
                         AddressOf Tasks.leavesinvoke)
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
    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.sickleave = Nothing
            Me.Close()

        Catch qw As Exception
        End Try
    End Sub

    Private Sub frmaddsickleave_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
    Private Sub txtemployeename_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtemployeename.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtemployeename, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtemployeename, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

End Class
