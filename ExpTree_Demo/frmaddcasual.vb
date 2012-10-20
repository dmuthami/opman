
Imports System
Imports ADODB

Public Class frmaddcasual
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
            myForms.casuals = Nothing
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtname As System.Windows.Forms.TextBox
    Friend WithEvents txttask As System.Windows.Forms.TextBox
    Friend WithEvents txtwage As System.Windows.Forms.TextBox
    Friend WithEvents dtpdatehired As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmaddcasual))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnadd = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dtpdatehired = New System.Windows.Forms.DateTimePicker
        Me.txtname = New System.Windows.Forms.TextBox
        Me.txttask = New System.Windows.Forms.TextBox
        Me.txtwage = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(442, 220)
        Me.Panel1.TabIndex = 10
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnclose)
        Me.Panel2.Controls.Add(Me.btnadd)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 220)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(442, 32)
        Me.Panel2.TabIndex = 5
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(362, 6)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 7
        Me.btnclose.Text = "Close"
        '
        'btnadd
        '
        Me.btnadd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnadd.Location = New System.Drawing.Point(3, 4)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.TabIndex = 6
        Me.btnadd.Text = "Add"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dtpdatehired)
        Me.GroupBox1.Controls.Add(Me.txtname)
        Me.GroupBox1.Controls.Add(Me.txttask)
        Me.GroupBox1.Controls.Add(Me.txtwage)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(442, 218)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'dtpdatehired
        '
        Me.dtpdatehired.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpdatehired.Location = New System.Drawing.Point(88, 188)
        Me.dtpdatehired.Name = "dtpdatehired"
        Me.dtpdatehired.Size = New System.Drawing.Size(160, 20)
        Me.dtpdatehired.TabIndex = 3
        '
        'txtname
        '
        Me.txtname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtname.Location = New System.Drawing.Point(88, 11)
        Me.txtname.Multiline = True
        Me.txtname.Name = "txtname"
        Me.txtname.Size = New System.Drawing.Size(344, 45)
        Me.txtname.TabIndex = 1
        Me.txtname.Text = ""
        '
        'txttask
        '
        Me.txttask.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txttask.Location = New System.Drawing.Point(88, 64)
        Me.txttask.Multiline = True
        Me.txttask.Name = "txttask"
        Me.txttask.Size = New System.Drawing.Size(344, 120)
        Me.txttask.TabIndex = 2
        Me.txttask.Text = ""
        '
        'txtwage
        '
        Me.txtwage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtwage.Location = New System.Drawing.Point(320, 188)
        Me.txtwage.Name = "txtwage"
        Me.txtwage.Size = New System.Drawing.Size(112, 20)
        Me.txtwage.TabIndex = 4
        Me.txtwage.Text = ""
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(254, 188)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 16)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Wage paid"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(8, 188)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 16)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Date hired"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(11, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 33)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Task performed"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(11, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 16)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Name"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmaddcasual
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(442, 252)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmaddcasual"
        Me.Text = "Add casuals"
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
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
                Dim sdate As String 'date of employment
                sdate = Me.dtpdatehired.Value.Year & "-" _
                 & dtpdatehired.Value.Month & "-" _
                 & dtpdatehired.Value.Day & " " _
                 & dtpdatehired.Value.Hour & ":" _
                 & dtpdatehired.Value.Minute & ":" _
                 & dtpdatehired.Value.Second

                Dim strsql As String
                Dim Tasks As New taskclass
                'Tasks.mid = Mid()
                'Dim a() As String
                'a = txtdesc.Lines
                'Dim desc As String = rml(a)
                Dim isadd As Boolean = True
                Dim strin As String = txttask.Text.Trim
                Dim strin1 As String = txtwage.Text.Trim
                Dim strin2 As String = txtname.Text.Trim
                strin = strin.Replace("'", "\'")
                strin1 = strin1.Replace("'", "\'")
                strin2 = strin2.Replace("'", "\'")
                If Me.btnadd.Text = "Add" Then
                    strsql = "insert into casuals"
                    strsql += "( job_no,description,task,datehired,wagespaid,namme) values "
                    strsql += " ( '" & jobno & "','" & "" & "', '" & strin & "','" & sdate & "'," _
                    & "'" & strin1 & "','" & strin2 & "')"
                Else
                    strsql = "update casuals set"
                    strsql += " job_no='" & jobno & "', task='" & strin & "'" _
                    & ",datehired='" & sdate & "',wagespaid='" & strin1 & "'," _
                    & "namme='" & strin2 & "'"
                    strsql += " where ano = '" & Integer.Parse(autono) & "';"
                    isadd = False
                End If
                Thread566 = New System.Threading.Thread( _
                AddressOf Tasks.casualsinvoke)
                If isadd = True Then
                    txttask.Text = ""
                    txtwage.Text = ""
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
        '---------refresh grosss margin
        Try
            myForms.CustomerForm2.mygross()
        Catch zx As Exception

        End Try
        '----------------------
    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.casuals = Nothing
            Me.Close()

        Catch ex As Exception
        End Try
    End Sub

    Private Sub frmaddcasual_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

#Region "validation"
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
    Private Sub txttask_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttask.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txttask, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txttask, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtwage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtwage.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtwage, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtwage, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region


End Class
