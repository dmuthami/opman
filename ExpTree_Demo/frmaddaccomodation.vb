Imports System
Imports ADODB


Public Class frmaddaccomodation
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
            myForms.accomodation = Nothing
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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtaccomodationcost As System.Windows.Forms.TextBox
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtname As System.Windows.Forms.TextBox
    Friend WithEvents cboentry As System.Windows.Forms.ComboBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmaddaccomodation))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtaccomodationcost = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cboentry = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtname = New System.Windows.Forms.TextBox
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnadd = New System.Windows.Forms.Button
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(442, 246)
        Me.Panel1.TabIndex = 12
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.txtaccomodationcost)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(4, 193)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(434, 50)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        '
        'txtaccomodationcost
        '
        Me.txtaccomodationcost.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtaccomodationcost.Location = New System.Drawing.Point(100, 13)
        Me.txtaccomodationcost.Name = "txtaccomodationcost"
        Me.txtaccomodationcost.Size = New System.Drawing.Size(324, 20)
        Me.txtaccomodationcost.TabIndex = 5
        Me.txtaccomodationcost.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(9, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 20)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Cost incurred"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.cboentry)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtname)
        Me.GroupBox1.Controls.Add(Me.txtdesc)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(442, 194)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'cboentry
        '
        Me.cboentry.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboentry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboentry.Items.AddRange(New Object() {"Food", "Accomodation"})
        Me.cboentry.Location = New System.Drawing.Point(104, 170)
        Me.cboentry.Name = "cboentry"
        Me.cboentry.Size = New System.Drawing.Size(328, 22)
        Me.cboentry.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Enter your name"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(8, 170)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Pick an entry"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Name of hotel attended"
        '
        'txtname
        '
        Me.txtname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtname.Location = New System.Drawing.Point(104, 16)
        Me.txtname.Multiline = True
        Me.txtname.Name = "txtname"
        Me.txtname.Size = New System.Drawing.Size(328, 40)
        Me.txtname.TabIndex = 1
        Me.txtname.Text = ""
        '
        'txtdesc
        '
        Me.txtdesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdesc.Location = New System.Drawing.Point(104, 62)
        Me.txtdesc.Multiline = True
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(328, 100)
        Me.txtdesc.TabIndex = 2
        Me.txtdesc.Text = ""
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnclose)
        Me.Panel2.Controls.Add(Me.btnadd)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 246)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(442, 32)
        Me.Panel2.TabIndex = 6
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(362, 6)
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
        '
        'frmaddaccomodation
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(442, 278)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmaddaccomodation"
        Me.Text = "Add accomodation costs"
        Me.Panel1.ResumeLayout(False)
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
    Private Sub frmaddaccomodation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
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

                Dim strsql As String
                Dim Tasks As New taskclass
                'Tasks.mid = Mid()
                'Dim a() As String
                'a = txtdesc.Lines
                'Dim desc As String = rml(a)
                Dim isadd As Boolean = True
                Dim strin As String = txtdesc.Text.Trim
                Dim strin1 As String = txtaccomodationcost.Text.Trim
                Dim strin2 As String = txtname.Text.Trim
                strin = strin.Replace("'", "\'")
                strin1 = strin1.Replace("'", "\'")
                strin2 = strin2.Replace("'", "\'")

                If Me.btnadd.Text = "Add" Then
                    strsql = "insert into accomodation"
                    strsql += "( job_no,description,costincurred, namme,entry) values "
                    strsql += " ( '" & jobno & "','" & strin & "',"
                    strsql += " '" & strin1 & "','" & strin2 & "'," _
                    & "'" & Me.cboentry.Text.Trim & "')"

                Else
                    strsql = "update accomodation set"
                    strsql += " job_no='" & jobno & "', description='" & strin & "',"
                    strsql += " namme='" & strin2 & "', entry='" & Me.cboentry.Text.Trim & "'" _
                    & ",costincurred='" & strin1 & "'"
                    strsql += " where ano = '" & Integer.Parse(autono) & "';"
                    isadd = False
                End If
                Thread566 = New System.Threading.Thread( _
                AddressOf Tasks.accomodationinvoke)
                If isadd = True Then
                    txtdesc.Text = ""
                    txtaccomodationcost.Text = ""
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
            myForms.accomodation = Nothing
            Me.Close()

        Catch ex As Exception
        End Try
    End Sub

    '#Region "validation"
    '    Private Sub txtaccomodationcost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtaccomodationcost.KeyPress
    '        Try
    '            Dim vt As New validation()
    '            If vt._validatetextbox(e) = True Then
    '                Me.errp.SetError(Me.txtaccomodationcost, _
    '                                      "not allowed chars: ''','%','*','\','*','1'")
    '                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
    '            Else
    '                Me.errp.SetError(Me.txtaccomodationcost, "")
    '            End If
    '        Catch xc As Exception

    '        End Try
    '    End Sub
    '    Private Sub txtdesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdesc.KeyPress
    '        Try
    '            Dim vt As New validation()
    '            If vt._validatetextbox(e) = True Then
    '                Me.errp.SetError(Me.txtdesc, _
    '                                      "not allowed chars: ''','%','*','\','*','1'")
    '                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
    '            Else
    '                Me.errp.SetError(Me.txtdesc, "")
    '            End If
    '        Catch xc As Exception

    '        End Try
    '    End Sub
    '    Private Sub txtname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtname.KeyPress
    '        Try
    '            Dim vt As New validation()
    '            If vt._validatetextbox(e) = True Then
    '                Me.errp.SetError(Me.txtname, _
    '                                      "not allowed chars: ''','%','*','\','*','1'")
    '                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
    '            Else
    '                Me.errp.SetError(Me.txtname, "")
    '            End If
    '        Catch xc As Exception

    '        End Try
    '    End Sub
    '#End Region

End Class
