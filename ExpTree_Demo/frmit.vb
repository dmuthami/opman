Imports System
Imports System.Threading
Imports ADODB


Imports System.ArgumentOutOfRangeException
Public Class frmit
    Inherits System.Windows.Forms.Form
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
    Friend WithEvents cboid As New System.Windows.Forms.ComboBox()

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        cboid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        cboid.Name = "cboid"
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
    Friend WithEvents StiGroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents StiGroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents StiGroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboviews As System.Windows.Forms.ComboBox
    Friend WithEvents dtgissues As System.Windows.Forms.DataGrid
    Friend WithEvents rtbcomments As System.Windows.Forms.RichTextBox
    Friend WithEvents dtpreportdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents rtbissues As System.Windows.Forms.RichTextBox
    Friend WithEvents chksolved As System.Windows.Forms.CheckBox
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents btnrefresh As System.Windows.Forms.Button
    Friend WithEvents btnjobsquery As VSEssentials.VSHotButton
    Friend WithEvents cbofor As System.Windows.Forms.ComboBox
    Friend WithEvents btnsearch As System.Windows.Forms.Button
    Friend WithEvents lblfor As System.Windows.Forms.Label
    Friend WithEvents cbowhere As System.Windows.Forms.ComboBox
    Friend WithEvents txtparam As System.Windows.Forms.TextBox
    Friend WithEvents lblwhere As System.Windows.Forms.Label
    Friend WithEvents lblis As System.Windows.Forms.Label
    Friend WithEvents btndelete As System.Windows.Forms.Button
    Friend WithEvents btnshowall As System.Windows.Forms.Button
    Friend WithEvents stiexporttoexcel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.StiGroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblis = New System.Windows.Forms.Label
        Me.txtparam = New System.Windows.Forms.TextBox
        Me.cbowhere = New System.Windows.Forms.ComboBox
        Me.lblwhere = New System.Windows.Forms.Label
        Me.btnsearch = New System.Windows.Forms.Button
        Me.cbofor = New System.Windows.Forms.ComboBox
        Me.cboviews = New System.Windows.Forms.ComboBox
        Me.StiGroupBox4 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpreportdate = New System.Windows.Forms.DateTimePicker
        Me.rtbissues = New System.Windows.Forms.RichTextBox
        Me.StiGroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnjobsquery = New VSEssentials.VSHotButton
        Me.chksolved = New System.Windows.Forms.CheckBox
        Me.rtbcomments = New System.Windows.Forms.RichTextBox
        Me.lblfor = New System.Windows.Forms.Label
        Me.StiGroupBox2 = New System.Windows.Forms.GroupBox
        Me.stiexporttoexcel = New System.Windows.Forms.Button
        Me.btnshowall = New System.Windows.Forms.Button
        Me.btndelete = New System.Windows.Forms.Button
        Me.btnadd = New System.Windows.Forms.Button
        Me.btnrefresh = New System.Windows.Forms.Button
        Me.dtgissues = New System.Windows.Forms.DataGrid
        Me.StiGroupBox1.SuspendLayout()
        Me.StiGroupBox4.SuspendLayout()
        Me.StiGroupBox3.SuspendLayout()
        Me.StiGroupBox2.SuspendLayout()
        CType(Me.dtgissues, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StiGroupBox1
        '
        Me.StiGroupBox1.Controls.Add(Me.lblis)
        Me.StiGroupBox1.Controls.Add(Me.txtparam)
        Me.StiGroupBox1.Controls.Add(Me.cbowhere)
        Me.StiGroupBox1.Controls.Add(Me.lblwhere)
        Me.StiGroupBox1.Controls.Add(Me.btnsearch)
        Me.StiGroupBox1.Controls.Add(Me.cbofor)
        Me.StiGroupBox1.Controls.Add(Me.cboviews)
        Me.StiGroupBox1.Controls.Add(Me.StiGroupBox4)
        Me.StiGroupBox1.Controls.Add(Me.StiGroupBox3)
        Me.StiGroupBox1.Controls.Add(Me.lblfor)
        Me.StiGroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.StiGroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.StiGroupBox1.Name = "StiGroupBox1"
        Me.StiGroupBox1.Size = New System.Drawing.Size(640, 280)
        Me.StiGroupBox1.TabIndex = 0
        Me.StiGroupBox1.TabStop = False
        '
        'lblis
        '
        Me.lblis.Location = New System.Drawing.Point(432, 253)
        Me.lblis.Name = "lblis"
        Me.lblis.Size = New System.Drawing.Size(16, 16)
        Me.lblis.TabIndex = 20
        Me.lblis.Text = "is"
        Me.lblis.Visible = False
        '
        'txtparam
        '
        Me.txtparam.Location = New System.Drawing.Point(448, 253)
        Me.txtparam.Name = "txtparam"
        Me.txtparam.Size = New System.Drawing.Size(152, 20)
        Me.txtparam.TabIndex = 11
        Me.txtparam.Text = ""
        Me.txtparam.Visible = False
        '
        'cbowhere
        '
        Me.cbowhere.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbowhere.Items.AddRange(New Object() {"issues", "comments"})
        Me.cbowhere.Location = New System.Drawing.Point(313, 253)
        Me.cbowhere.Name = "cbowhere"
        Me.cbowhere.Size = New System.Drawing.Size(112, 22)
        Me.cbowhere.TabIndex = 10
        Me.cbowhere.Visible = False
        '
        'lblwhere
        '
        Me.lblwhere.Location = New System.Drawing.Point(273, 253)
        Me.lblwhere.Name = "lblwhere"
        Me.lblwhere.Size = New System.Drawing.Size(40, 16)
        Me.lblwhere.TabIndex = 17
        Me.lblwhere.Text = "where"
        Me.lblwhere.Visible = False
        '
        'btnsearch
        '
        Me.btnsearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsearch.Location = New System.Drawing.Point(600, 253)
        Me.btnsearch.Name = "btnsearch"
        Me.btnsearch.Size = New System.Drawing.Size(32, 24)
        Me.btnsearch.TabIndex = 12
        Me.btnsearch.Text = "Go"
        Me.btnsearch.Visible = False
        '
        'cbofor
        '
        Me.cbofor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbofor.Location = New System.Drawing.Point(159, 253)
        Me.cbofor.Name = "cbofor"
        Me.cbofor.Size = New System.Drawing.Size(112, 22)
        Me.cbofor.TabIndex = 9
        Me.cbofor.Visible = False
        '
        'cboviews
        '
        Me.cboviews.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboviews.Items.AddRange(New Object() {"View all issues", "View solved issues", "View unsolved issues"})
        Me.cboviews.Location = New System.Drawing.Point(8, 253)
        Me.cboviews.Name = "cboviews"
        Me.cboviews.Size = New System.Drawing.Size(128, 22)
        Me.cboviews.TabIndex = 8
        '
        'StiGroupBox4
        '
        Me.StiGroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StiGroupBox4.Controls.Add(Me.Label3)
        Me.StiGroupBox4.Controls.Add(Me.dtpreportdate)
        Me.StiGroupBox4.Controls.Add(Me.rtbissues)
        Me.StiGroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiGroupBox4.Location = New System.Drawing.Point(0, 8)
        Me.StiGroupBox4.Name = "StiGroupBox4"
        Me.StiGroupBox4.Size = New System.Drawing.Size(632, 112)
        Me.StiGroupBox4.TabIndex = 1
        Me.StiGroupBox4.TabStop = False
        Me.StiGroupBox4.Text = "Type your IT issues here"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(152, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Date issue(s) reported"
        '
        'dtpreportdate
        '
        Me.dtpreportdate.CalendarFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpreportdate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpreportdate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpreportdate.Location = New System.Drawing.Point(168, 86)
        Me.dtpreportdate.Name = "dtpreportdate"
        Me.dtpreportdate.Size = New System.Drawing.Size(80, 20)
        Me.dtpreportdate.TabIndex = 3
        Me.dtpreportdate.Value = New Date(2006, 6, 6, 11, 26, 29, 625)
        '
        'rtbissues
        '
        Me.rtbissues.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbissues.AutoSize = True
        Me.rtbissues.AutoWordSelection = True
        Me.rtbissues.BulletIndent = 2
        Me.rtbissues.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbissues.Location = New System.Drawing.Point(8, 14)
        Me.rtbissues.Name = "rtbissues"
        Me.rtbissues.Size = New System.Drawing.Size(616, 67)
        Me.rtbissues.TabIndex = 2
        Me.rtbissues.Text = ""
        '
        'StiGroupBox3
        '
        Me.StiGroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StiGroupBox3.Controls.Add(Me.btnjobsquery)
        Me.StiGroupBox3.Controls.Add(Me.chksolved)
        Me.StiGroupBox3.Controls.Add(Me.rtbcomments)
        Me.StiGroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiGroupBox3.Location = New System.Drawing.Point(0, 120)
        Me.StiGroupBox3.Name = "StiGroupBox3"
        Me.StiGroupBox3.Size = New System.Drawing.Size(632, 128)
        Me.StiGroupBox3.TabIndex = 4
        Me.StiGroupBox3.TabStop = False
        Me.StiGroupBox3.Text = "Comments from administrator"
        '
        'btnjobsquery
        '
        Me.btnjobsquery.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnjobsquery.BackMouseDownColor = System.Drawing.SystemColors.Control
        Me.btnjobsquery.BackMouseOverColor = System.Drawing.SystemColors.Control
        Me.btnjobsquery.BorderBottomColor = System.Drawing.SystemColors.ControlDark
        Me.btnjobsquery.BorderLeftColor = System.Drawing.SystemColors.ControlLight
        Me.btnjobsquery.BorderRightColor = System.Drawing.SystemColors.ControlDark
        Me.btnjobsquery.BorderSize = 1
        Me.btnjobsquery.BorderTopColor = System.Drawing.SystemColors.ControlLight
        Me.btnjobsquery.ButtonSettings = System.Drawing.SystemColors.Control
        Me.btnjobsquery.DrawTextShadow = True
        Me.btnjobsquery.Location = New System.Drawing.Point(504, 96)
        Me.btnjobsquery.Name = "btnjobsquery"
        Me.btnjobsquery.Size = New System.Drawing.Size(120, 24)
        Me.btnjobsquery.TabIndex = 7
        Me.btnjobsquery.Tag = ""
        Me.btnjobsquery.TextAlign = VSEssentials.VSHotButton.eTextAlign.Center
        Me.btnjobsquery.TextCaption = "Go to Jobs query"
        Me.btnjobsquery.TextColor = System.Drawing.SystemColors.ControlText
        Me.btnjobsquery.TextFont = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnjobsquery.TextLeft = 0
        Me.btnjobsquery.TextOffsetX = 0
        Me.btnjobsquery.TextOffsetY = 0
        Me.btnjobsquery.TextShadowColor = System.Drawing.SystemColors.ControlDark
        Me.btnjobsquery.TextTop = 0
        '
        'chksolved
        '
        Me.chksolved.Enabled = False
        Me.chksolved.Location = New System.Drawing.Point(8, 97)
        Me.chksolved.Name = "chksolved"
        Me.chksolved.Size = New System.Drawing.Size(168, 24)
        Me.chksolved.TabIndex = 6
        Me.chksolved.Text = "IT problem solved"
        '
        'rtbcomments
        '
        Me.rtbcomments.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbcomments.AutoSize = True
        Me.rtbcomments.AutoWordSelection = True
        Me.rtbcomments.BulletIndent = 2
        Me.rtbcomments.Enabled = False
        Me.rtbcomments.Location = New System.Drawing.Point(6, 17)
        Me.rtbcomments.Name = "rtbcomments"
        Me.rtbcomments.Size = New System.Drawing.Size(618, 79)
        Me.rtbcomments.TabIndex = 5
        Me.rtbcomments.Text = ""
        '
        'lblfor
        '
        Me.lblfor.Location = New System.Drawing.Point(138, 253)
        Me.lblfor.Name = "lblfor"
        Me.lblfor.Size = New System.Drawing.Size(24, 16)
        Me.lblfor.TabIndex = 13
        Me.lblfor.Text = "for"
        Me.lblfor.Visible = False
        '
        'StiGroupBox2
        '
        Me.StiGroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StiGroupBox2.Controls.Add(Me.stiexporttoexcel)
        Me.StiGroupBox2.Controls.Add(Me.btnshowall)
        Me.StiGroupBox2.Controls.Add(Me.btndelete)
        Me.StiGroupBox2.Controls.Add(Me.btnadd)
        Me.StiGroupBox2.Controls.Add(Me.btnrefresh)
        Me.StiGroupBox2.Controls.Add(Me.dtgissues)
        Me.StiGroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiGroupBox2.Location = New System.Drawing.Point(0, 280)
        Me.StiGroupBox2.Name = "StiGroupBox2"
        Me.StiGroupBox2.Size = New System.Drawing.Size(640, 296)
        Me.StiGroupBox2.TabIndex = 13
        Me.StiGroupBox2.TabStop = False
        '
        'stiexporttoexcel
        '
        Me.stiexporttoexcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.stiexporttoexcel.Location = New System.Drawing.Point(177, 11)
        Me.stiexporttoexcel.Name = "stiexporttoexcel"
        Me.stiexporttoexcel.Size = New System.Drawing.Size(95, 23)
        Me.stiexporttoexcel.TabIndex = 16
        Me.stiexporttoexcel.Text = "Export to excel"
        '
        'btnshowall
        '
        Me.btnshowall.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnshowall.Location = New System.Drawing.Point(274, 11)
        Me.btnshowall.Name = "btnshowall"
        Me.btnshowall.Size = New System.Drawing.Size(69, 23)
        Me.btnshowall.TabIndex = 17
        Me.btnshowall.Text = "Show all"
        Me.btnshowall.Visible = False
        '
        'btndelete
        '
        Me.btndelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelete.Location = New System.Drawing.Point(344, 11)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(135, 23)
        Me.btndelete.TabIndex = 18
        Me.btndelete.Text = "Delete selected entry"
        Me.btndelete.Visible = False
        '
        'btnadd
        '
        Me.btnadd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnadd.Location = New System.Drawing.Point(8, 11)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.Size = New System.Drawing.Size(88, 23)
        Me.btnadd.TabIndex = 14
        Me.btnadd.Text = "Add IT issue"
        '
        'btnrefresh
        '
        Me.btnrefresh.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnrefresh.Location = New System.Drawing.Point(97, 11)
        Me.btnrefresh.Name = "btnrefresh"
        Me.btnrefresh.Size = New System.Drawing.Size(80, 23)
        Me.btnrefresh.TabIndex = 15
        Me.btnrefresh.Text = "Refresh"
        '
        'dtgissues
        '
        Me.dtgissues.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgissues.DataMember = ""
        Me.dtgissues.FlatMode = True
        Me.dtgissues.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgissues.Location = New System.Drawing.Point(8, 40)
        Me.dtgissues.Name = "dtgissues"
        Me.dtgissues.ReadOnly = True
        Me.dtgissues.Size = New System.Drawing.Size(624, 248)
        Me.dtgissues.TabIndex = 19
        '
        'frmit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(640, 576)
        Me.Controls.Add(Me.StiGroupBox2)
        Me.Controls.Add(Me.StiGroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmit"
        Me.Text = "It issues"
        Me.StiGroupBox1.ResumeLayout(False)
        Me.StiGroupBox4.ResumeLayout(False)
        Me.StiGroupBox3.ResumeLayout(False)
        Me.StiGroupBox2.ResumeLayout(False)
        CType(Me.dtgissues, System.ComponentModel.ISupportInitialize).EndInit()
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
    Private ano As String
    Private user_no As String
    Private Threaditcb As Thread
#End Region

#Region "frmit"
    Private Sub frmit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Try
                Me.cboviews.SelectedIndex = 0
                Me.cbowhere.SelectedIndex = 0
            Catch cv As Exception
            End Try
            '-------------load combos
            Dim task As New taskclass
            Dim Threadit As Thread = New System.Threading.Thread( _
            AddressOf task.itcomboinvoke)
            Threadit.IsBackground = True
            Threadit.Start()
            '--------------
            '---------------config controls
            Try
                configcontrols()
            Catch we As Exception
            End Try
            '--------------
            Try
                Me.user_no = myForms.id_no
            Catch cb As Exception

            End Try
            Me.dtpreportdate.Value = Now
            Me.Invalidate(True)
        Catch xc As Exception

        End Try
    End Sub
    Private Sub loadgrid()
        Try

        Catch xc As Exception

        End Try
    End Sub
    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        Try
            Try
                Dim connectstr As String
                connectstr = "DSN=" & myForms.qconnstr
                Dim connect As New ADODB.Connection
                connect.Mode = ConnectModeEnum.adModeReadWrite
                connect.CursorLocation = CursorLocationEnum.adUseClient
                connect.ConnectionString = connectstr
                connect.Open()
                Dim strsql As String
                Dim isadd As Boolean = False
                Dim strin As String = Me.rtbissues.Text.Trim
                Dim strin1 As String = Me.rtbcomments.Text.Trim
                strin = strin.Replace("'", "\'")
                strin1 = strin1.Replace("'", "\'")
                rtbissues.Text = strin
                rtbcomments.Text = strin1
                '-------------------
                Dim arr() As String
                Dim strr, strr2 As String
                Dim y As Integer
                rtbissues.Text = Me.rtbissues.Text.Trim()
                arr = rtbissues.Lines
                y = arr.GetUpperBound(0)
                Dim alpha As Integer
                For alpha = 0 To y
                    strr += arr(alpha) + vbCrLf
                    Application.DoEvents()
                Next
                '----------------------------------
                rtbcomments.Text = Me.rtbcomments.Text.Trim()
                arr = rtbcomments.Lines
                y = arr.GetUpperBound(0)
                For alpha = 0 To y
                    strr2 += arr(alpha) + vbCrLf
                    Application.DoEvents()
                Next
                Dim sdate As String
                sdate = dtpreportdate.Value.Year & "-" _
                 & dtpreportdate.Value.Month & "-" _
                 & dtpreportdate.Value.Day & " " _
                 & dtpreportdate.Value.Hour & ":" _
                 & dtpreportdate.Value.Minute & ":" _
                 & dtpreportdate.Value.Second

                Dim ischecked As Boolean = Me.chksolved.CheckState
                strsql = "insert into it"
                strsql += "( id_no,issues,comments, solved,report_date) values "
                strsql += " ( '" & myForms.id_no & "','" & strr & "',"
                strsql += " '" & strr2 & "','" & ischecked & "'," _
                & "'" & sdate & "');"
                connect.BeginTrans()
                connect.Execute(strsql)
                connect.CommitTrans()
                Try
                    Dim ds As New DataSet
                    ds = Me.dtgissues.DataSource
                    Dim myrow As System.Data.DataRow = ds.Tables(0).NewRow
                    myrow.Item("issues") = strr
                    myrow.Item("comments") = strr2
                    myrow.Item("solved") = ischecked
                    myrow.Item("report_date") = sdate
                    myrow.Item("id_no") = myForms.id_no
                    Try
                        myrow.Item("namme") = myForms.Main._name
                    Catch ex As Exception

                    End Try
                    ds.Tables(0).Rows.Add(myrow)
                Catch ex As Exception

                End Try
                isadd = True
                If isadd = True Then
                    rtbissues.Text = ""
                    rtbcomments.Text = ""
                End If
                Try
                    connect.Close()
                Catch es As Exception
                End Try
            Catch ex As Exception
            Finally
            End Try
            'Dim Tasks As New taskclass
            'Tasks.itindex = Me.cboviews.SelectedIndex.ToString()
            'Tasks.itno = myForms.id_no
            'Dim Thread566 As System.Threading.Thread
            'Thread566 = New System.Threading.Thread( _
            '               AddressOf Tasks.itinvoke)
            'Thread566.IsBackground = True
            'Thread566.Start()
        Catch qw As Exception
        End Try
    End Sub
    Private Sub configcontrols()
        Try
            Dim validate As Boolean = myForms.Main.canmanipulateit()
            If validate = True Then
                myForms.itissues.lblfor.Visible = True
                myForms.itissues.lblis.Visible = True
                myForms.itissues.lblwhere.Visible = True
                myForms.itissues.cbofor.Visible = True
                myForms.itissues.cbowhere.Visible = True
                myForms.itissues.btnsearch.Visible = True
                myForms.itissues.rtbcomments.Enabled = True
                myForms.itissues.chksolved.Enabled = True
                myForms.itissues.txtparam.Visible = True
                myForms.itissues.btndelete.Visible = True
                myForms.itissues.btnshowall.Visible = True
            Else
                myForms.itissues.lblfor.Visible = False
                myForms.itissues.lblis.Visible = False
                myForms.itissues.lblwhere.Visible = False
                myForms.itissues.cbofor.Visible = False
                myForms.itissues.cbowhere.Visible = False
                myForms.itissues.btnsearch.Visible = False
                myForms.itissues.rtbcomments.Enabled = False
                myForms.itissues.chksolved.Enabled = False
                myForms.itissues.txtparam.Visible = False
                myForms.itissues.btndelete.Visible = False
                myForms.itissues.btnshowall.Visible = False
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnrefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrefresh.Click
        Try
            Try
                Dim connectstr As String
                connectstr = "DSN=" & myForms.qconnstr
                Dim connect As New ADODB.Connection
                connect.Mode = ConnectModeEnum.adModeReadWrite
                connect.CursorLocation = CursorLocationEnum.adUseClient
                connect.ConnectionString = connectstr
                connect.Open()
                Dim strsql As String
                Dim isadd As Boolean = False
                Dim strin As String = Me.rtbissues.Text.Trim
                Dim strin1 As String = Me.rtbcomments.Text.Trim
                Dim sdate As String
                sdate = dtpreportdate.Value.Year & "-" _
                 & dtpreportdate.Value.Month & "-" _
                 & dtpreportdate.Value.Day & " " _
                 & dtpreportdate.Value.Hour & ":" _
                 & dtpreportdate.Value.Minute & ":" _
                 & dtpreportdate.Value.Second
                strin = strin.Replace("'", "\'")
                strin1 = strin1.Replace("'", "\'")
                Dim ischecked As Boolean = Me.chksolved.CheckState
                strsql = "update it set "
                strsql += " id_no='" & user_no & "',issues='" & strin & "',"
                strsql += " comments='" & strin1 & "',solved='" & ischecked & "'," _
                & "report_date='" & sdate & "'"
                strsql += " where ano='" & ano & "';"
                connect.BeginTrans()
                connect.Execute(strsql)
                connect.CommitTrans()
                Try
                    Dim ds As New DataSet
                    ds = Me.dtgissues.DataSource
                    ds.Tables(0).Rows(hti.Row).Item("issues") = strin
                    ds.Tables(0).Rows(hti.Row).Item("comments") = strin1
                    ds.Tables(0).Rows(hti.Row).Item("solved") = ischecked
                    ds.Tables(0).Rows(hti.Row).Item("report_date") = sdate
                    ds.Tables(0).Rows(hti.Row).Item("id_no") = user_no
                    Try
                        ds.Tables(0).Rows(hti.Row).Item("namme") = myForms.Main._name
                    Catch ex As Exception

                    End Try
                Catch ex As Exception

                End Try
                isadd = True
                If isadd = True Then
                    rtbissues.Text = ""
                    rtbcomments.Text = ""
                End If
                Try
                    connect.Close()
                Catch es As Exception
                End Try
            Catch ex As Exception
            Finally
            End Try
            Dim Tasks As New taskclass
            Tasks.itindex = Me.cboviews.SelectedIndex.ToString()
            Tasks.itno = myForms.id_no
            Dim Thread566 As System.Threading.Thread
            Thread566 = New System.Threading.Thread( _
                           AddressOf Tasks.itinvoke)
            Thread566.IsBackground = True
            Thread566.Start()
        Catch qw As Exception
        End Try
    End Sub
    Private Sub dtgissues_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgissues.MouseDown
        Try
            hti = dtgissues.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgissues_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgissues.DoubleClick
        Try
            Dim ds As DataSet = New DataSet
            ds = Me.dtgissues.DataSource
            ano = ds.Tables(0).Rows(hti.Row).Item("ano")
            Try
                dtgissues.Select(Integer.Parse(hti.Row))
            Catch ex As Exception
            End Try
            Try
                Me.rtbissues.Text = ds.Tables(0).Rows(hti.Row).Item("issues")
            Catch es As Exception
            End Try
            Try
                Me.rtbcomments.Text = ds.Tables(0).Rows(hti.Row).Item("comments")
            Catch es As Exception

            End Try

            Try
                Me.dtpreportdate.Value = CDate(ds.Tables(0).Rows(hti.Row).Item("report_date"))
            Catch es As Exception

            End Try
            Try
                Me.chksolved.Checked = Convert.ToBoolean(ds.Tables(0).Rows(hti.Row).Item("mybool"))
            Catch es As Exception

            End Try
            Try
                user_no = ds.Tables(0).Rows(hti.Row).Item("id_no")
            Catch es As Exception

            End Try
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())

        End Try
    End Sub
    Private Sub cbofor_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbofor.SelectedIndexChanged
        Try
            Dim indexx As Integer
            indexx = Me.cbofor.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.cboid.SelectedIndex = indexx
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub cboviews_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboviews.SelectedIndexChanged
        Try
            If Me.btnsearch.Visible = False Then
                '-------------load default view
                If Threaditcb Is Nothing Then
                    Try
                        Threaditcb.Abort()
                    Catch we As Exception
                    End Try
                End If
                Dim task As New taskclass
                task.itindex = cboviews.SelectedIndex.ToString()
                task.itno = myForms.id_no
                Threaditcb = New System.Threading.Thread( _
                     AddressOf task.itinvoke)
                Threaditcb.IsBackground = True
                Threaditcb.Start()
                '--------------
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
        Try
            Dim task As New taskclass
            task.itindex = cboviews.SelectedIndex.ToString()
            task.itno = Me.cboid.Text
            task.itsearch = True
            If Me.cbowhere.SelectedIndex = 0 Then
                task.ffield = "issues"
            Else
                task.ffield = "comments"
            End If
            Dim Threadit As Thread = New System.Threading.Thread( _
                 AddressOf task.itinvoke)
            Threadit.IsBackground = True
            Threadit.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnjobsquery_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnjobsquery.Click
        Try
            If myForms.Main.canmanipulateit = False Then
                MessageBox.Show("Can't view jobs query,contact administrator", "View Jobs Query", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch jh As Exception

        End Try
        Try
            myForms.Main.ToolBar2.Buttons(2).Pushed = True
            myForms.Main.ToolBar2.Buttons(1).Pushed = False
            myForms.Main.ToolBar2.Buttons(0).Pushed = False
            Try
                myForms.itissues.Close()
                myForms.itissues = Nothing
            Catch er As Exception
            End Try
            Dim ad As New frmjobs
            myForms.qjobs = ad
            myForms.qjobs.Size = myForms.Main.pnlpersonnel.Size
            myForms.qjobs.TopLevel = False
            myForms.qjobs.Parent = myForms.Main.pnlpersonnel
            myForms.qjobs.Dock = DockStyle.Fill
            myForms.qjobs.Show()
            myForms.qjobs.BringToFront()
        Catch we As Exception

        End Try
    End Sub
    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
        Try
            dtgissues.Select(hti.Row)
            If MessageBox.Show("Do you wish to delete the current row?", "Deleting", MessageBoxButtons.YesNo, _
            MessageBoxIcon.Information) = DialogResult.No Then
                Exit Try
            End If
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgissues.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hti.Row).Item("ano")
            str = "delete from it where"
            str += "  ano='" & sid & "'"

            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(hti.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try

            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnshowall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnshowall.Click
        Dim task As New taskclass
        task.showall = True
        Dim Threadit As Thread = New System.Threading.Thread( _
       AddressOf task.itinvoke)
        Threadit.IsBackground = True
        Threadit.Start()
    End Sub
#End Region



    Private Sub stiexporttoexcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stiexporttoexcel.Click
        Try
            Dim ds1 As System.Data.DataSet = New System.Data.DataSet
            Dim ds2 As System.Data.DataSet = New System.Data.DataSet
            ds2 = myForms.itissues.dtgissues.DataSource
            ds1 = ds2.Copy
            myForms.Main.dgrid.DataSource = Nothing
            Dim MyTable As New DataTable
            MyTable = ds1.Tables(0)
            Try
                MyTable.Columns.Remove("ano")
                MyTable.Columns.Remove("mybool")
                'MyTable.Columns.Remove("ano")
                'MyTable.Columns.Remove("job_no1")
                'MyTable.Columns.Remove("milliseconds")
                'MyTable.Columns.Remove("Edit")
                'MyTable.Columns.Remove("ddate1")
                'MyTable.Columns.Remove("isadded")
            Catch cv As Exception
            End Try
            Try
                MyTable.Columns("id_no").ColumnName = "Identification number"
                MyTable.Columns("issues").ColumnName = "Issues"
                MyTable.Columns("comments").ColumnName = "Comments"
                MyTable.Columns("solved").ColumnName = "Solved"
                MyTable.Columns("report_date").ColumnName = "Report Date"

            Catch cv As Exception
            End Try
            Try
                Dim sfd As System.Windows.Forms.SaveFileDialog _
                = New System.Windows.Forms.SaveFileDialog
                sfd.Filter = "Excel files (*.xls)|*.xls"
                sfd.CheckFileExists = False
                sfd.CheckPathExists = True
                sfd.ValidateNames = True
                sfd.ShowDialog()
                Dim m As String = sfd.FileName
                If m.Trim.Length > 0 Then
                    exporttoexcel.exportexcel.exportToExcel(ds1, m)
                    MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch we As Exception
            End Try
        Catch ex As Exception
        End Try
        'Me.btnview_Click(Me, e)
    End Sub
    Private Sub btnjobsquery_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles btnjobsquery.Paint

    End Sub
    Private Sub dtgissues_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtgissues.Navigate
    End Sub
    Private Sub StiGroupBox4_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StiGroupBox4.Enter
    End Sub
End Class
