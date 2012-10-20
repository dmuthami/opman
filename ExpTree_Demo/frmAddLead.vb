Imports System
Imports System.Threading
Imports ADODB


Imports System.IO
Public Class frmAddLead
    Inherits System.Windows.Forms.Form
    Dim frm As frmHome
    Dim frm1 As frmMe
    Public clientno, cname, leadno 'client number and name
    Public fromclients As Boolean
    Private cboclientno As New System.Windows.Forms.ComboBox()
    Private clickcombo As Boolean = False
    Private journalpath As String
    Private Delegate Sub mydelegate1()
    Private Delegate Sub mydelegate()

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
            addleads = False
            myForms.CustomerForm = Nothing
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
    Friend WithEvents pnljournal As System.Windows.Forms.Panel
    Friend WithEvents rtbjournal As System.Windows.Forms.RichTextBox
    Friend WithEvents btnSaveJournal As System.Windows.Forms.Button
    Friend WithEvents pnlAddLead As System.Windows.Forms.Panel
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents lblLeadtitle As System.Windows.Forms.Label
    Friend WithEvents txtleadtitle As System.Windows.Forms.TextBox
    Friend WithEvents dtpsniffed As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateSniffed As System.Windows.Forms.Label
    Friend WithEvents pnlSearch As System.Windows.Forms.Panel
    Friend WithEvents BtnAddLead As System.Windows.Forms.Button
    Friend WithEvents txtCboContactnames As System.Windows.Forms.TextBox
    Friend WithEvents cboContactName As System.Windows.Forms.ComboBox
    Friend WithEvents lblContactName As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents grpdesc As System.Windows.Forms.GroupBox
    Friend WithEvents grpnormal As System.Windows.Forms.GroupBox
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents lblOpenjournal As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtdepartment As System.Windows.Forms.ComboBox
    Friend WithEvents cboProspect As System.Windows.Forms.ComboBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents btndepartments As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddLead))
        Me.pnljournal = New System.Windows.Forms.Panel
        Me.btnSaveJournal = New System.Windows.Forms.Button
        Me.rtbjournal = New System.Windows.Forms.RichTextBox
        Me.pnlAddLead = New System.Windows.Forms.Panel
        Me.btndepartments = New System.Windows.Forms.Button
        Me.cboProspect = New System.Windows.Forms.ComboBox
        Me.txtdepartment = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblAmount = New System.Windows.Forms.Label
        Me.txtAmount = New System.Windows.Forms.TextBox
        Me.grpnormal = New System.Windows.Forms.GroupBox
        Me.lblCompany = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.grpdesc = New System.Windows.Forms.GroupBox
        Me.lblOpenjournal = New System.Windows.Forms.LinkLabel
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.lblLeadtitle = New System.Windows.Forms.Label
        Me.txtleadtitle = New System.Windows.Forms.TextBox
        Me.dtpsniffed = New System.Windows.Forms.DateTimePicker
        Me.lblDateSniffed = New System.Windows.Forms.Label
        Me.pnlSearch = New System.Windows.Forms.Panel
        Me.BtnAddLead = New System.Windows.Forms.Button
        Me.txtCboContactnames = New System.Windows.Forms.TextBox
        Me.cboContactName = New System.Windows.Forms.ComboBox
        Me.lblContactName = New System.Windows.Forms.Label
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.lblStatus = New System.Windows.Forms.Label
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.pnljournal.SuspendLayout()
        Me.pnlAddLead.SuspendLayout()
        Me.grpnormal.SuspendLayout()
        Me.grpdesc.SuspendLayout()
        Me.pnlSearch.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnljournal
        '
        Me.pnljournal.Controls.Add(Me.btnSaveJournal)
        Me.pnljournal.Controls.Add(Me.rtbjournal)
        Me.pnljournal.Location = New System.Drawing.Point(0, 0)
        Me.pnljournal.Name = "pnljournal"
        Me.pnljournal.Size = New System.Drawing.Size(360, 496)
        Me.pnljournal.TabIndex = 0
        Me.pnljournal.Visible = False
        '
        'btnSaveJournal
        '
        Me.btnSaveJournal.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSaveJournal.Location = New System.Drawing.Point(8, 5)
        Me.btnSaveJournal.Name = "btnSaveJournal"
        Me.btnSaveJournal.Size = New System.Drawing.Size(128, 20)
        Me.btnSaveJournal.TabIndex = 1
        Me.btnSaveJournal.Text = "Save Journal"
        '
        'rtbjournal
        '
        Me.rtbjournal.Location = New System.Drawing.Point(8, 32)
        Me.rtbjournal.Name = "rtbjournal"
        Me.rtbjournal.Size = New System.Drawing.Size(344, 408)
        Me.rtbjournal.TabIndex = 0
        Me.rtbjournal.Text = ""
        '
        'pnlAddLead
        '
        Me.pnlAddLead.Controls.Add(Me.btndepartments)
        Me.pnlAddLead.Controls.Add(Me.cboProspect)
        Me.pnlAddLead.Controls.Add(Me.txtdepartment)
        Me.pnlAddLead.Controls.Add(Me.Label1)
        Me.pnlAddLead.Controls.Add(Me.lblAmount)
        Me.pnlAddLead.Controls.Add(Me.txtAmount)
        Me.pnlAddLead.Controls.Add(Me.grpnormal)
        Me.pnlAddLead.Controls.Add(Me.grpdesc)
        Me.pnlAddLead.Controls.Add(Me.lblLeadtitle)
        Me.pnlAddLead.Controls.Add(Me.txtleadtitle)
        Me.pnlAddLead.Controls.Add(Me.dtpsniffed)
        Me.pnlAddLead.Controls.Add(Me.lblDateSniffed)
        Me.pnlAddLead.Controls.Add(Me.pnlSearch)
        Me.pnlAddLead.Controls.Add(Me.btnExit)
        Me.pnlAddLead.Controls.Add(Me.btnSave)
        Me.pnlAddLead.Controls.Add(Me.lblStatus)
        Me.pnlAddLead.Location = New System.Drawing.Point(5, 6)
        Me.pnlAddLead.Name = "pnlAddLead"
        Me.pnlAddLead.Size = New System.Drawing.Size(352, 490)
        Me.pnlAddLead.TabIndex = 2
        '
        'btndepartments
        '
        Me.btndepartments.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndepartments.Location = New System.Drawing.Point(120, 385)
        Me.btndepartments.Name = "btndepartments"
        Me.btndepartments.Size = New System.Drawing.Size(32, 20)
        Me.btndepartments.TabIndex = 101
        Me.btndepartments.Tag = "Add or edit departments"
        Me.btndepartments.Text = "A"
        '
        'cboProspect
        '
        Me.cboProspect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProspect.Location = New System.Drawing.Point(120, 360)
        Me.cboProspect.Name = "cboProspect"
        Me.cboProspect.Size = New System.Drawing.Size(224, 22)
        Me.cboProspect.TabIndex = 41
        '
        'txtdepartment
        '
        Me.txtdepartment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.txtdepartment.Items.AddRange(New Object() {"", "Survey", "GI", "RS", "BD"})
        Me.txtdepartment.Location = New System.Drawing.Point(152, 384)
        Me.txtdepartment.Name = "txtdepartment"
        Me.txtdepartment.Size = New System.Drawing.Size(192, 22)
        Me.txtdepartment.TabIndex = 37
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 384)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "Department"
        '
        'lblAmount
        '
        Me.lblAmount.Location = New System.Drawing.Point(16, 431)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(96, 16)
        Me.lblAmount.TabIndex = 34
        Me.lblAmount.Text = "Amount"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(120, 431)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(224, 20)
        Me.txtAmount.TabIndex = 8
        Me.txtAmount.Text = ""
        '
        'grpnormal
        '
        Me.grpnormal.Controls.Add(Me.lblCompany)
        Me.grpnormal.Controls.Add(Me.Label2)
        Me.grpnormal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpnormal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpnormal.Location = New System.Drawing.Point(8, 191)
        Me.grpnormal.Name = "grpnormal"
        Me.grpnormal.Size = New System.Drawing.Size(336, 56)
        Me.grpnormal.TabIndex = 25
        Me.grpnormal.TabStop = False
        '
        'lblCompany
        '
        Me.lblCompany.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lblCompany.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompany.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCompany.Location = New System.Drawing.Point(112, 16)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(216, 32)
        Me.lblCompany.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 24)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Company"
        '
        'grpdesc
        '
        Me.grpdesc.Controls.Add(Me.lblOpenjournal)
        Me.grpdesc.Controls.Add(Me.txtDesc)
        Me.grpdesc.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpdesc.Location = New System.Drawing.Point(8, 88)
        Me.grpdesc.Name = "grpdesc"
        Me.grpdesc.Size = New System.Drawing.Size(336, 264)
        Me.grpdesc.TabIndex = 4
        Me.grpdesc.TabStop = False
        Me.grpdesc.Text = "Description"
        '
        'lblOpenjournal
        '
        Me.lblOpenjournal.Location = New System.Drawing.Point(8, 240)
        Me.lblOpenjournal.Name = "lblOpenjournal"
        Me.lblOpenjournal.Size = New System.Drawing.Size(128, 16)
        Me.lblOpenjournal.TabIndex = 41
        Me.lblOpenjournal.TabStop = True
        Me.lblOpenjournal.Text = "Open New Journal"
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(8, 16)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(320, 216)
        Me.txtDesc.TabIndex = 5
        Me.txtDesc.Text = ""
        '
        'lblLeadtitle
        '
        Me.lblLeadtitle.Location = New System.Drawing.Point(11, 66)
        Me.lblLeadtitle.Name = "lblLeadtitle"
        Me.lblLeadtitle.Size = New System.Drawing.Size(77, 16)
        Me.lblLeadtitle.TabIndex = 23
        Me.lblLeadtitle.Text = "Lead Title"
        '
        'txtleadtitle
        '
        Me.txtleadtitle.Location = New System.Drawing.Point(96, 65)
        Me.txtleadtitle.Multiline = True
        Me.txtleadtitle.Name = "txtleadtitle"
        Me.txtleadtitle.Size = New System.Drawing.Size(240, 20)
        Me.txtleadtitle.TabIndex = 3
        Me.txtleadtitle.Text = ""
        '
        'dtpsniffed
        '
        Me.dtpsniffed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpsniffed.Location = New System.Drawing.Point(120, 407)
        Me.dtpsniffed.Name = "dtpsniffed"
        Me.dtpsniffed.Size = New System.Drawing.Size(224, 20)
        Me.dtpsniffed.TabIndex = 7
        '
        'lblDateSniffed
        '
        Me.lblDateSniffed.Location = New System.Drawing.Point(16, 407)
        Me.lblDateSniffed.Name = "lblDateSniffed"
        Me.lblDateSniffed.Size = New System.Drawing.Size(104, 16)
        Me.lblDateSniffed.TabIndex = 16
        Me.lblDateSniffed.Text = "Date"
        '
        'pnlSearch
        '
        Me.pnlSearch.Controls.Add(Me.BtnAddLead)
        Me.pnlSearch.Controls.Add(Me.txtCboContactnames)
        Me.pnlSearch.Controls.Add(Me.cboContactName)
        Me.pnlSearch.Controls.Add(Me.lblContactName)
        Me.pnlSearch.Location = New System.Drawing.Point(8, 8)
        Me.pnlSearch.Name = "pnlSearch"
        Me.pnlSearch.Size = New System.Drawing.Size(336, 56)
        Me.pnlSearch.TabIndex = 0
        Me.pnlSearch.Visible = False
        '
        'BtnAddLead
        '
        Me.BtnAddLead.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BtnAddLead.ForeColor = System.Drawing.Color.FromArgb(CType(80, Byte), CType(173, Byte), CType(209, Byte))
        Me.BtnAddLead.Location = New System.Drawing.Point(7, 8)
        Me.BtnAddLead.Name = "BtnAddLead"
        Me.BtnAddLead.Size = New System.Drawing.Size(24, 24)
        Me.BtnAddLead.TabIndex = 19
        Me.BtnAddLead.Text = "A"
        '
        'txtCboContactnames
        '
        Me.txtCboContactnames.Location = New System.Drawing.Point(88, 8)
        Me.txtCboContactnames.Name = "txtCboContactnames"
        Me.txtCboContactnames.Size = New System.Drawing.Size(220, 20)
        Me.txtCboContactnames.TabIndex = 1
        Me.txtCboContactnames.Text = ""
        '
        'cboContactName
        '
        Me.cboContactName.Location = New System.Drawing.Point(88, 8)
        Me.cboContactName.MaxLength = 10
        Me.cboContactName.Name = "cboContactName"
        Me.cboContactName.Size = New System.Drawing.Size(240, 22)
        Me.cboContactName.TabIndex = 1
        '
        'lblContactName
        '
        Me.lblContactName.Location = New System.Drawing.Point(32, 8)
        Me.lblContactName.Name = "lblContactName"
        Me.lblContactName.Size = New System.Drawing.Size(56, 32)
        Me.lblContactName.TabIndex = 0
        Me.lblContactName.Text = "Contact Name"
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnExit.Location = New System.Drawing.Point(216, 459)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(120, 20)
        Me.btnExit.TabIndex = 10
        Me.btnExit.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSave.Location = New System.Drawing.Point(8, 459)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(120, 20)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "Save"
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(16, 358)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(104, 24)
        Me.lblStatus.TabIndex = 9
        Me.lblStatus.Text = "Status"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmAddLead
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(362, 500)
        Me.Controls.Add(Me.pnlAddLead)
        Me.Controls.Add(Me.pnljournal)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmAddLead"
        Me.Text = "Add Lead"
        Me.pnljournal.ResumeLayout(False)
        Me.pnlAddLead.ResumeLayout(False)
        Me.grpnormal.ResumeLayout(False)
        Me.grpdesc.ResumeLayout(False)
        Me.pnlSearch.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    <System.STAThread()> _
        Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub

#Region "Add leads"
    Private Sub btndepartments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndepartments.Click
        Try
            Dim n As New frmdepartments
            n.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub frmAddLead_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.grpnormal.Location = New System.Drawing.Point(8, 8)
            loaddt()

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
        Try
            Dim _thread As Thread = New Thread(AddressOf ld)
            _thread.IsBackground = True
            _thread.Start()
        Catch ex As Exception

        End Try

    End Sub
    Private Sub loaddt()
        Try
            Me.lblCompany.Text = Me.cname
            Me.cboProspect.Items.Add("Prospect")
            Me.cboProspect.Items.Add("Proposal")
            Me.cboProspect.Items.Add("PHO")
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            ''''''''''''---------------validation
            'If Me.txtAmount.Text = "" Then
            '    MessageBox.Show("Please input an amount", "Add leads", _
            '    MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Try
            'End If
            If Me.txtdepartment.Text = "" Then
                MessageBox.Show("Please pick a department", "Add leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.txtDesc.Text = "" Then
                MessageBox.Show("Please input description", "Add leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.txtleadtitle.Text = "" Then
                MessageBox.Show("Please input a title for the lead", "Add leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If Me.cboProspect.Text = "" Then
                MessageBox.Show("Please pick a status for this lead", "Add leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            '--------------end

            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

            If Me.grpnormal.Visible = True Then
                Dim lno = newlno(Me.clientno)  'code to automate lead numbers
                leadno = lno
                'code to save journal--------
                Call checkdirectory()
                '--------------------
                Dim sdate As String
                sdate = dtpsniffed.Value.Year & "-" _
                & dtpsniffed.Value.Month & "-" _
                & dtpsniffed.Value.Day
                journalpath = journalpath.Replace("\", "|")
                Dim strsql
                strsql = "insert into leads" _
                                & " (leads_no,client_no,descrip,status,date_sniffed,title,journal,amount,department) values"
                strsql = strsql & "(" & "'" & lno & "',"

                strsql = strsql & "'" & clientno & "',"
                strsql = strsql & "'" & txtDesc.Text & "',"
                strsql = strsql & "'" & cboProspect.Text & "',"
                strsql = strsql & "'" & sdate & "',"
                strsql = strsql & "'" & txtleadtitle.Text & "',"
                strsql = strsql & "'" & journalpath & "',"
                strsql = strsql & "'" & txtAmount.Text & "',"
                strsql = strsql & "'" & txtdepartment.Text & "'"
                strsql = strsql & ");"
                strsql += " update clients set leads_no='" & lno & "'"
                strsql += " where lower(client_no)='" & CStr(Me.clientno).ToLower() & "'"
                strsql += ";"
                connect.BeginTrans()
                connect.IsolationLevel = ADODB.IsolationLevelEnum.adXactSerializable
                connect.Execute(strsql)
                connect.CommitTrans()

                MessageBox.Show(Text:="Lead has been successfully added", _
                buttons:=MessageBoxButtons.OK, caption:="Add Client")

                refreshleads = True
                refreshleadshome = True
            Else
                If Me.cboContactName.Text.Trim = "" Then
                    MsgBox("A valid client name is required" _
                    , MsgBoxStyle.Information)
                    Exit Try
                End If
                If Me.clientno = "" Then
                    MsgBox("Please pick a valid client from the drop down" _
                             , MsgBoxStyle.Information)
                End If
                Dim lno = newlno(Me.clientno)  'code to automate lead numbers
                leadno = lno
                ''code to save journal--------
                'Call checkdirectory()
                ''--------------------
                Dim sdate As String
                sdate = dtpsniffed.Value.Year & "-" _
                & dtpsniffed.Value.Month & "-" _
                & dtpsniffed.Value.Day
                Dim arr() As String
                Dim strr As String
                Dim y As Integer
                arr = txtDesc.Lines
                y = arr.GetUpperBound(0)
                Dim alpha As Integer
                For alpha = 0 To y
                    strr += arr(alpha) + vbCrLf
                    Application.DoEvents()
                Next
                Dim strsql
                strsql = "insert into leads" _
                             & " (leads_no,client_no,descrip,status,date_sniffed,title,journal,amount) values"
                strsql = strsql & "(" & "'" & lno & "',"
                strsql = strsql & "'" & clientno & "',"
                strsql = strsql & "'" & strr & "',"
                strsql = strsql & "'" & cboProspect.Text & "',"
                strsql = strsql & "'" & sdate & "',"
                strsql = strsql & "'" & txtleadtitle.Text & "',"
                strsql = strsql & "'" & journalpath & "',"
                strsql = strsql & "'" & txtAmount.Text & "'"
                strsql = strsql & ");"
                strsql += " update clients set leads_no='" & lno & "'"
                strsql += " where lower(client_no)='" & CStr(Me.clientno).ToLower() & "'"
                strsql += ";"
                connect.BeginTrans()
                connect.IsolationLevel = ADODB.IsolationLevelEnum.adXactSerializable
                connect.Execute(strsql)
                connect.CommitTrans()
                Dim Threadleads As System.Threading.Thread = New System.Threading.Thread( _
                                                              AddressOf loaddirectory)
                Threadleads.IsBackground = True
                Threadleads.Start()
                MessageBox.Show(Text:="Lead has been successfully added", _
              buttons:=MessageBoxButtons.OK, caption:="Add Client", _
              Icon:=MessageBoxIcon.Information)

                refreshleads = True
                refreshleadshome = True
            End If
            updateclientstatus() 'update client status

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Cursor.Current = currentcursor
        End Try
        Try
            Dim tthread As System.Threading.Thread = New System.Threading.Thread(AddressOf threadjobs)
            Try
                If tthread.IsAlive = True Then
                    tthread.Abort()
                End If
            Catch ex As Exception

            End Try

            tthread.IsBackground = True
            tthread.Start()
        Catch ex As Exception

        End Try
        Try
            Dim tthread1 As System.Threading.Thread = New System.Threading.Thread(AddressOf myForms.Main.loadleads)
            Try
                If tthread1.IsAlive = True Then
                    tthread1.Abort()
                End If
            Catch ex As Exception

            End Try

            tthread1.IsBackground = True
            tthread1.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub threadjobs()
        Try
            myForms.CustomerForm3.Invoke(New mydelegate(AddressOf myForms.CustomerForm3.loadleads))

        Catch ex As Exception

        End Try
    End Sub
    Private Sub loaddirectory()
        Try
            Me.Invoke(New mydelegate1(AddressOf checkdirectory))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub updateclientstatus()
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

            Dim cstatus = editclientstatus(clientno)
            If cstatus = "" Then
                cstatus = Me.cboProspect.Text
            End If
            Dim strsql
            connect.BeginTrans()
            connect.IsolationLevel = ADODB.IsolationLevelEnum.adXactSerializable

            strsql += " update clients set least_status='" & cstatus & "'"
            strsql += " where lower(client_no)='" & CStr(Me.clientno).ToLower() & "'"
            strsql += ";"
            connect.Execute(strsql)
            connect.CommitTrans()
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            addleads = False
            myForms.CustomerForm = Nothing
            Me.Dispose(True)

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub txtCboContactnames_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCboContactnames.TextChanged
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()
            Cursor.Current = Cursors.WaitCursor
            If Me.clickcombo = False Then
                Me.cboclientno.Items.Clear()


                Dim rs As New ADODB.Recordset
                With rs
                    .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                    .CursorType = ADODB.CursorTypeEnum.adOpenStatic
                    Dim str
                    str = "select name,client_no from clients where lower(name) like " _
                    & "'%" & Me.txtCboContactnames.Text.Trim.ToLower() & "%'" _
                    & "order by name asc"
                    .Open(str, connect)
                    cboContactName.Items.Clear()
                    If .BOF = False And .EOF = False Then
                        .MoveFirst()
                        While .EOF = False
                            cboContactName.Items.Add(.Fields("name").Value)
                            cboclientno.Items.Add(.Fields("client_no").Value)
                            .MoveNext()
                            Application.DoEvents()
                        End While

                        ' ComboBox1.Sorted = True

                    End If
                End With
                rs.Close()
            End If
            Me.clickcombo = False

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Me.clientno = ""
            Cursor.Current = currentcursor
        End Try
    End Sub
    Private Sub cboContactName_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboContactName.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cboContactName.SelectedIndex
            If indexx = -1 Then
                Exit Try

            End If
            Me.cboclientno.SelectedIndex = indexx
            Dim strp
            strp = Me.cboclientno.Text
            Me.clickcombo = True
            Me.txtCboContactnames.Text = Me.cboContactName.Text
            Me.clientno = strp
            'MsgBox(strp)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub BtnAddLead_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddLead.Click
        Try
            If myfrmAddClientsform = 0 Then
                Dim form As New frmAddClients

                form.Show()
                addleads = True
            Else
                ' Dim form As New frmAddClients()
                addleads = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub lblViewJournal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            pnlAddLead.Visible = False
            pnljournal.Visible = True
            Me.Text = "Write Journal"
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub cboProspect_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Select Case cboProspect.Text.ToLower()
                Case "prospect"
                    txtAmount.Enabled = False
                Case "suspect"
                    txtAmount.Enabled = False
                Case Else
                    txtAmount.Enabled = True


            End Select

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
#End Region

#Region "journal"
    Private Sub checkdirectory()
        Try
            Dim myvar As String = "value=" & myForms.qfolderpath
            'str = Configuration.ConfigurationSettings.AppSettings("folderpath")

            Dim myfile, mypath
            mypath = myvar
            mypath = mypath & "\"
            mypath = mypath & clientno

            myfile = Dir(mypath, FileAttribute.Directory)
            If myfile <> "" Then
                mypath += "\" & leadno
                myfile = Dir(mypath, FileAttribute.Directory)
                If myfile <> "" Then
                    Me.storedata(mypath)

                Else
                    MkDir(mypath)
                    Me.storedata(mypath)
                End If

            Else
                MkDir(mypath)
                mypath += "\" & leadno
                MkDir(mypath)
                Me.storedata(mypath)
            End If

        Catch ex As Exception
            'MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub storedata(ByVal str As String)
        Try
            ''Me.txtMailBody.SaveFile(str & Me.txtMailSubject.Text, System.Windows.Forms.RichTextBoxStreamType.RichText)
            Dim i, j
            Dim str1, myfilename As String
            Dim a()
            'For i = 0 To Me.lstAttachment.Items.Count - 1


            '    lstAttachment.SetSelected(i, True)
            '    str1 = lstAttachment.SelectedItems(i).ToString()
            '    a = str1.Split("\")
            '    j = a.GetUpperBound(0)
            '    myfilename = a(j)
            '    Try
            '        File.Copy(str1, str & "\" & myfilename)
            '    Catch ex As Exception
            '    End Try
            'Next
            Dim dtp As New System.Windows.Forms.DateTimePicker
            journalpath = str & "\" & leadno & "_" & dtp.Value.Year & dtp.Value.Month & dtp.Value.Day _
            & dtp.Value.Hour & dtp.Value.Minute & dtp.Value.Second & dtp.Value.Millisecond & ".txt"
            rtbjournal.SaveFile(journalpath)

        Catch ex As Exception
            'MsgBox(ex.Message.ToString())
        End Try
    End Sub
    Private Sub btnSaveJournal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveJournal.Click
        Try
            Me.pnlAddLead.Visible = True
            Me.pnljournal.Visible = False
            Me.Text = "Add Lead"
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub
#End Region

#Region "validation"
    Private Sub txtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesc.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtDesc, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtDesc, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtleadtitle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtleadtitle.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtleadtitle, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtleadtitle, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtCboContactnames_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCboContactnames.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtCboContactnames, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtCboContactnames, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtAmount, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtAmount, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtdepartment_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdepartment.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtdepartment, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtdepartment, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

#Region "loaddepartments"
    Public Sub loaddepart()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str = "select * from department"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                myForms.CustomerForm.txtdepartment.Items.Clear()
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While rs.EOF = False
                        myForms.CustomerForm.txtdepartment.Items.Add(.Fields("dept").Value)
                        .MoveNext()
                        Application.DoEvents()
                    End While
                End If
            End With
            Try
                rs.Close()

            Catch er34b As Exception

            End Try
            Try

                connect.Close()
            Catch er344 As Exception

            End Try
        Catch ex As Exception

        End Try
        Try
            System.GC.Collect()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ld()
        Try
            myForms.CustomerForm.Invoke(New mydelegate(AddressOf loaddepart))
        Catch ex As Exception

        End Try
    End Sub
#End Region

End Class
