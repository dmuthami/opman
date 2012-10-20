Imports System.String
Imports ADODB
Imports System.Text.StringBuilder
Imports System.Object
Imports System.Data
Imports System



Imports System.Threading
Public Class frmEditClients
    Inherits System.Windows.Forms.Form
    Public Delegate Sub mydelegate()
    Public mynumber

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
            editclients = False
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
    Friend WithEvents tbcchangeclients As System.Windows.Forms.TabControl
    Friend WithEvents tpgedit As System.Windows.Forms.TabPage
    Friend WithEvents tpgDel As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtClientName As System.Windows.Forms.TextBox
    Friend WithEvents txtClientNo As System.Windows.Forms.TextBox
    Friend WithEvents lblClientName As System.Windows.Forms.Label
    Friend WithEvents lblClientNo As System.Windows.Forms.Label
    Friend WithEvents pnlsqlstmt As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnBackUp As System.Windows.Forms.Button
    Friend WithEvents rtbsqlstmt As System.Windows.Forms.RichTextBox
    Friend WithEvents btnEditClient As System.Windows.Forms.Button
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEditClients))
        Me.tbcchangeclients = New System.Windows.Forms.TabControl
        Me.tpgedit = New System.Windows.Forms.TabPage
        Me.btnEditClient = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtClientName = New System.Windows.Forms.TextBox
        Me.txtClientNo = New System.Windows.Forms.TextBox
        Me.lblClientName = New System.Windows.Forms.Label
        Me.lblClientNo = New System.Windows.Forms.Label
        Me.tpgDel = New System.Windows.Forms.TabPage
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnBackUp = New System.Windows.Forms.Button
        Me.pnlsqlstmt = New System.Windows.Forms.Panel
        Me.rtbsqlstmt = New System.Windows.Forms.RichTextBox
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.tbcchangeclients.SuspendLayout()
        Me.tpgedit.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.tpgDel.SuspendLayout()
        Me.pnlsqlstmt.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbcchangeclients
        '
        Me.tbcchangeclients.Controls.Add(Me.tpgedit)
        Me.tbcchangeclients.Controls.Add(Me.tpgDel)
        Me.tbcchangeclients.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbcchangeclients.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbcchangeclients.Location = New System.Drawing.Point(0, 0)
        Me.tbcchangeclients.Multiline = True
        Me.tbcchangeclients.Name = "tbcchangeclients"
        Me.tbcchangeclients.SelectedIndex = 0
        Me.tbcchangeclients.Size = New System.Drawing.Size(378, 424)
        Me.tbcchangeclients.TabIndex = 5
        '
        'tpgedit
        '
        Me.tpgedit.Controls.Add(Me.btnEditClient)
        Me.tpgedit.Controls.Add(Me.GroupBox2)
        Me.tpgedit.Controls.Add(Me.GroupBox1)
        Me.tpgedit.Location = New System.Drawing.Point(4, 24)
        Me.tpgedit.Name = "tpgedit"
        Me.tpgedit.Size = New System.Drawing.Size(370, 396)
        Me.tpgedit.TabIndex = 0
        Me.tpgedit.Text = "Edit"
        '
        'btnEditClient
        '
        Me.btnEditClient.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnEditClient.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnEditClient.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEditClient.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.btnEditClient.Location = New System.Drawing.Point(11, 338)
        Me.btnEditClient.Name = "btnEditClient"
        Me.btnEditClient.Size = New System.Drawing.Size(120, 20)
        Me.btnEditClient.TabIndex = 5
        Me.btnEditClient.Text = "Save Changes"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtDesc)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 128)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(360, 208)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Description"
        '
        'txtDesc
        '
        Me.txtDesc.AcceptsTab = True
        Me.txtDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.Location = New System.Drawing.Point(8, 16)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDesc.Size = New System.Drawing.Size(344, 176)
        Me.txtDesc.TabIndex = 4
        Me.txtDesc.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtClientName)
        Me.GroupBox1.Controls.Add(Me.txtClientNo)
        Me.GroupBox1.Controls.Add(Me.lblClientName)
        Me.GroupBox1.Controls.Add(Me.lblClientNo)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(360, 120)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txtClientName
        '
        Me.txtClientName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClientName.Location = New System.Drawing.Point(104, 40)
        Me.txtClientName.Multiline = True
        Me.txtClientName.Name = "txtClientName"
        Me.txtClientName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtClientName.Size = New System.Drawing.Size(248, 72)
        Me.txtClientName.TabIndex = 2
        Me.txtClientName.Text = ""
        '
        'txtClientNo
        '
        Me.txtClientNo.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.txtClientNo.Enabled = False
        Me.txtClientNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClientNo.Location = New System.Drawing.Point(104, 16)
        Me.txtClientNo.Name = "txtClientNo"
        Me.txtClientNo.Size = New System.Drawing.Size(248, 20)
        Me.txtClientNo.TabIndex = 1
        Me.txtClientNo.Text = ""
        '
        'lblClientName
        '
        Me.lblClientName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientName.Location = New System.Drawing.Point(8, 40)
        Me.lblClientName.Name = "lblClientName"
        Me.lblClientName.Size = New System.Drawing.Size(88, 16)
        Me.lblClientName.TabIndex = 1
        Me.lblClientName.Text = "Client Name"
        '
        'lblClientNo
        '
        Me.lblClientNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientNo.Location = New System.Drawing.Point(8, 16)
        Me.lblClientNo.Name = "lblClientNo"
        Me.lblClientNo.Size = New System.Drawing.Size(88, 16)
        Me.lblClientNo.TabIndex = 0
        Me.lblClientNo.Text = "Client No"
        '
        'tpgDel
        '
        Me.tpgDel.Controls.Add(Me.btnDelete)
        Me.tpgDel.Controls.Add(Me.btnBackUp)
        Me.tpgDel.Controls.Add(Me.pnlsqlstmt)
        Me.tpgDel.Location = New System.Drawing.Point(4, 24)
        Me.tpgDel.Name = "tpgDel"
        Me.tpgDel.Size = New System.Drawing.Size(370, 396)
        Me.tpgDel.TabIndex = 1
        Me.tpgDel.Text = "Delete"
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnDelete.Location = New System.Drawing.Point(245, 339)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(120, 20)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "Delete"
        '
        'btnBackUp
        '
        Me.btnBackUp.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnBackUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnBackUp.Location = New System.Drawing.Point(5, 339)
        Me.btnBackUp.Name = "btnBackUp"
        Me.btnBackUp.Size = New System.Drawing.Size(120, 20)
        Me.btnBackUp.TabIndex = 2
        Me.btnBackUp.Text = "Back Up "
        '
        'pnlsqlstmt
        '
        Me.pnlsqlstmt.Controls.Add(Me.rtbsqlstmt)
        Me.pnlsqlstmt.Location = New System.Drawing.Point(8, 8)
        Me.pnlsqlstmt.Name = "pnlsqlstmt"
        Me.pnlsqlstmt.Size = New System.Drawing.Size(360, 320)
        Me.pnlsqlstmt.TabIndex = 0
        '
        'rtbsqlstmt
        '
        Me.rtbsqlstmt.Location = New System.Drawing.Point(8, 8)
        Me.rtbsqlstmt.Name = "rtbsqlstmt"
        Me.rtbsqlstmt.Size = New System.Drawing.Size(344, 304)
        Me.rtbsqlstmt.TabIndex = 0
        Me.rtbsqlstmt.Text = ""
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnClose)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(0, 384)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(378, 40)
        Me.GroupBox3.TabIndex = 6
        Me.GroupBox3.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(251, 11)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(120, 24)
        Me.btnClose.TabIndex = 7
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        '
        'frmEditClients
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(378, 424)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.tbcchangeclients)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(352, 384)
        Me.Name = "frmEditClients"
        Me.Text = "Change Client Details"
        Me.tbcchangeclients.ResumeLayout(False)
        Me.tpgedit.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.tpgDel.ResumeLayout(False)
        Me.pnlsqlstmt.ResumeLayout(False)
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

#Region "edit client"
   
    Private Sub frmEditClients_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call loaddata()
        ' mythread1 = New System.Threading.Thread(AddressOf loaddata)

    End Sub
    Private Sub editclient()
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Dim strsql As String
            Cursor.Current = Cursors.WaitCursor

            If Me.txtClientName.Text = "" Or txtClientName.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="A client must have a name", _
                buttons:=MessageBoxButtons.OK, caption:="Edit Client")
                Exit Try
            End If
            If Me.txtClientName.Text = "" Then
                MessageBox.Show(Text:="A client must have a name", _
                buttons:=MessageBoxButtons.OK, caption:="Edit Client")
                Exit Try

            End If
            If Me.txtDesc.Text.Trim() = "" Then

                MessageBox.Show(Text:="Please input description", _
                 buttons:=MessageBoxButtons.OK, caption:="Edit Client")
                Exit Try
            End If
            If Me.txtClientNo.Text.Trim.ToUpper = Nothing Then
                MessageBox.Show(Text:="A client must have a number", _
                buttons:=MessageBoxButtons.OK, caption:="Edit Client")
                Exit Try
            End If


            Dim number As String
            number = Me.txtClientNo.Text.Trim.ToUpper
            Dim r As New ADODB.Recordset
            Dim cmd As New ADODB.Command
            With cmd
                .CommandType = CommandTypeEnum.adCmdText
                .ActiveConnection = connect
                Dim mystr As String
                mystr = "select client_no" _
                & " from clients " _
                & " where client_no='" & number & "'"
                .CommandText = mystr
                r = .Execute()

            End With
            'If r.BOF = False And r.EOF = False Then
            '    MessageBox.Show(Text:="A similar client number already exists", _
            '    buttons:=MessageBoxButtons.OK, caption:="Add Client")
            '    Exit Try
            'End If

            'Dim a() As String
            'a = Me.txtDesc.Lines
            'Dim strdesc As String

            'strdesc = rml(a)
            'a.Initialize()

            'Dim cname As String
            'a = txtClientName.Lines
            'cname = rml(a)

            '-------------------
            Dim arr() As String
            Dim strr As String
            Dim y As Integer
            txtDesc.Text = Me.txtDesc.Text.Trim()
            arr = txtDesc.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            '----------------------------------
            txtClientName.Text = Me.txtClientName.Text.Trim()
            arr = txtClientName.Lines
            y = arr.GetUpperBound(0)
            Dim strr2 As String
            For alpha = 0 To y
                strr2 += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            Dim cno As String
            cno = Me.txtClientNo.Text.Trim
            cno = cno.ToUpper
            strsql = "update clients set"
            strsql = strsql & " client_no='" & cno & "',"
            strsql = strsql & " name='" & strr2 & "',"
            strsql = strsql & " description='" & strr & "'"
            strsql = strsql & " where client_no='" & myclientno & "'"
            'contacts table
            strsql = strsql & ";"
            strsql = strsql & " update contact set"
            strsql = strsql & " client_no='" & cno & "'"
            strsql = strsql & " where client_no='" & myclientno & "'"
            'job sheet table
            strsql = strsql & ";"
            strsql = strsql & " update jobsheet set"
            strsql = strsql & " client_no='" & cno & "'"
            strsql = strsql & " where client_no='" & myclientno & "'"
            'rcljobs
            strsql = strsql & ";"
            strsql = strsql & " update rcljobs set"
            strsql = strsql & " client_no='" & cno & "'"
            strsql = strsql & " where client_no='" & myclientno & "'"
            connect.BeginTrans()
            connect.IsolationLevel = ADODB.IsolationLevelEnum.adXactSerializable
            connect.Execute(strsql)
            connect.CommitTrans()
            myclientno = Me.txtClientNo.Text
            Try
                myForms.CustomerForm3.lblClientNo.Text = myclientno
                myForms.CustomerForm3.lblClientName.Text = strr2
            Catch ex As Exception

            End Try

            MessageBox.Show(Text:="Client details have been successfully updated", _
                caption:="", buttons:=MessageBoxButtons.OK, _
                Icon:=MessageBoxIcon.Information)
            refreshclients = True
            myclientno = cno
            myclientname = strr2
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
            Try
                connect.RollbackTrans()
            Catch r As Exception
                'MessageBox.Show(ex.InnerException.ToString, "Error", MessageBoxButtons.OK)
            End Try


        Finally
            txtClientName.Text = ""
            txtDesc.Text = ""
            Cursor.Current = currentcursor
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try

    End Sub
    Private Sub loaddata()
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Try
            Dim cmd As New ADODB.Command
            Dim rs As New ADODB.Recordset
            Dim str As String
            With cmd
                .ActiveConnection = connect
                .CommandType = ADODB.CommandTypeEnum.adCmdText

                str = " select *  from clients"
                str = str & " where client_no='" & Me.txtClientNo.Text & "'"
                .CommandText = str
                rs = cmd.Execute
            End With
            Me.txtClientName.Text = rs.Fields("name").Value
            Me.txtDesc.Text = rs.Fields("description").Value
            rs.Close()
            rs = Nothing
        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Function removetabs(ByVal str As String) As String
        Dim i As Integer
        Dim s As String
        s = str
        Do
            i = s.IndexOf(vbTab)
            If i <> 0 Then
                s = s.Remove(i, 1)
            End If

        Loop Until str.IndexOf(vbTab) = 0
        removetabs = str
    End Function
    Private Sub txtDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Protected Overrides Sub Finalize()
        'editclients = False
        MyBase.Finalize()
    End Sub
    Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
        Try
            ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
            ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
            If keyData = System.Windows.Forms.Keys.Return Then
                'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)
                Dim E As System.EventArgs

                'Call btnEditClient_Click_1(Me, E)

                Return True ' True means we've processed the key
            Else
                Return MyBase.ProcessDialogKey(keyData)
            End If
        Catch ex As Exception
            'Trace.WriteLine(ex.ToString())
            MsgBox(ex.Message.ToString, , Title:="Return key")

        End Try
    End Function
    Private Sub btnEditClient_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditClient.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateclients()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate clients contact administrator", "Clients", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Me.Invoke(New mydelegate(AddressOf editclient))

        Catch ex As Exception
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
    Private Sub txtClientName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtClientName.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtClientName, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtClientName, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        editclients = False
        Me.Dispose(False)
    End Sub
End Class

