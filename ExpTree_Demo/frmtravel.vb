Imports System
Imports ADODB


Public Class frmtravel
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
            myForms.travel = Nothing
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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txttravelcost As System.Windows.Forms.TextBox
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtother As System.Windows.Forms.TextBox
    Friend WithEvents txtkilometers As System.Windows.Forms.TextBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmtravel))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtkilometers = New System.Windows.Forms.TextBox
        Me.txtother = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txttravelcost = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnadd = New System.Windows.Forms.Button
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(608, 298)
        Me.Panel1.TabIndex = 14
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtkilometers)
        Me.GroupBox1.Controls.Add(Me.txtother)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtdesc)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(608, 296)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Description"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 19)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Other modes"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 19)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Kilometers"
        '
        'txtkilometers
        '
        Me.txtkilometers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtkilometers.Location = New System.Drawing.Point(112, 72)
        Me.txtkilometers.Name = "txtkilometers"
        Me.txtkilometers.Size = New System.Drawing.Size(488, 20)
        Me.txtkilometers.TabIndex = 2
        Me.txtkilometers.Text = ""
        '
        'txtother
        '
        Me.txtother.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtother.Location = New System.Drawing.Point(112, 96)
        Me.txtother.Multiline = True
        Me.txtother.Name = "txtother"
        Me.txtother.Size = New System.Drawing.Size(488, 128)
        Me.txtother.TabIndex = 3
        Me.txtother.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(10, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 19)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Vehicle/driving"
        '
        'txtdesc
        '
        Me.txtdesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdesc.Location = New System.Drawing.Point(112, 16)
        Me.txtdesc.Multiline = True
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(488, 48)
        Me.txtdesc.TabIndex = 1
        Me.txtdesc.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.txttravelcost)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 240)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(608, 48)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        '
        'txttravelcost
        '
        Me.txttravelcost.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txttravelcost.Location = New System.Drawing.Point(113, 13)
        Me.txttravelcost.Name = "txttravelcost"
        Me.txttravelcost.Size = New System.Drawing.Size(487, 20)
        Me.txttravelcost.TabIndex = 5
        Me.txttravelcost.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(9, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Travel costs"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnclose)
        Me.Panel2.Controls.Add(Me.btnadd)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 298)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(608, 32)
        Me.Panel2.TabIndex = 6
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(528, 6)
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
        'frmtravel
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(608, 330)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmtravel"
        Me.Text = "Add travel costs"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
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
    Private Sub frmtravel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.travel = Nothing
            Me.Close()

        Catch ex As Exception
        End Try
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
                txtother.Text = Me.txtother.Text.Trim()
                arr = txtdesc.Lines
                y = arr.GetUpperBound(0)
                For alpha = 0 To y
                    strr2 += arr(alpha) + vbCrLf
                    Application.DoEvents()
                Next
                If Me.btnadd.Text = "Add" Then
                    strsql = "insert into travel"
                    strsql += "( job_no,description,costincurred,kilometers,othermodes) values "
                    strsql += " ( '" & jobno & "','" & strr & "'," _
                    & "'" & Me.txttravelcost.Text.Trim & "','" & Me.txtkilometers.Text.Trim & "','" & strr2 & "')"
                Else
                    strsql = "update travel set"
                    strsql += " job_no='" & jobno & "', description='" & strr & "'" _
                    & ",costincurred='" & txttravelcost.Text.Trim & "',"
                    strsql += " kilometers='" & txtkilometers.Text.Trim & "', othermodes='" & strr2 & "'"
                    strsql += " where ano = '" & Integer.Parse(autono) & "';"
                    isadd = False
                End If
                Thread566 = New System.Threading.Thread( _
                                  AddressOf Tasks.travelinvoke)
                If isadd = True Then
                    txtdesc.Text = ""
                    txttravelcost.Text = ""
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

#Region "validation"
    Private Sub txttravelcost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttravelcost.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txttravelcost, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txttravelcost, "")
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
    Private Sub txtother_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtother.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtother, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtother, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtkilometers_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtkilometers.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtkilometers, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtkilometers, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

End Class
