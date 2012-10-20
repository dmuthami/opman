Imports System.String
Imports ADODB
Imports System.Text.StringBuilder
Imports System.Object
Imports System.Data
Imports System




Imports System.Exception
Imports System.SystemException
Imports System.IO

Imports System.Threading
Public Class frmAddClients
    Inherits System.Windows.Forms.Form
    Public Delegate Sub mydelegate()
    Public Delegate Sub mydelegate1()
    Private clientno, leadno
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
            myfrmAddClientsform = 0
            refreshclients = True
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
    Friend WithEvents lblClientNo As System.Windows.Forms.Label
    Friend WithEvents lblClientName As System.Windows.Forms.Label
    Friend WithEvents txtClientNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnAddClient As System.Windows.Forms.Button
    Friend WithEvents txtClientName As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddClients))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtClientName = New System.Windows.Forms.TextBox
        Me.txtClientNo = New System.Windows.Forms.TextBox
        Me.lblClientName = New System.Windows.Forms.Label
        Me.lblClientNo = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.btnAddClient = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cboStatus = New System.Windows.Forms.ComboBox
        Me.lblStatus = New System.Windows.Forms.Label
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtClientName)
        Me.GroupBox1.Controls.Add(Me.txtClientNo)
        Me.GroupBox1.Controls.Add(Me.lblClientName)
        Me.GroupBox1.Controls.Add(Me.lblClientNo)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(7, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(317, 120)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txtClientName
        '
        Me.txtClientName.BackColor = System.Drawing.Color.GhostWhite
        Me.txtClientName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClientName.Location = New System.Drawing.Point(120, 40)
        Me.txtClientName.Multiline = True
        Me.txtClientName.Name = "txtClientName"
        Me.txtClientName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtClientName.Size = New System.Drawing.Size(192, 72)
        Me.txtClientName.TabIndex = 2
        Me.txtClientName.Text = ""
        Me.txtClientName.WordWrap = False
        '
        'txtClientNo
        '
        Me.txtClientNo.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.txtClientNo.Enabled = False
        Me.txtClientNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClientNo.Location = New System.Drawing.Point(120, 16)
        Me.txtClientNo.Name = "txtClientNo"
        Me.txtClientNo.Size = New System.Drawing.Size(192, 20)
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
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtDesc)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(5, 120)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(320, 160)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Description"
        '
        'txtDesc
        '
        Me.txtDesc.BackColor = System.Drawing.Color.GhostWhite
        Me.txtDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.Location = New System.Drawing.Point(8, 16)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDesc.Size = New System.Drawing.Size(304, 136)
        Me.txtDesc.TabIndex = 4
        Me.txtDesc.Text = ""
        Me.txtDesc.WordWrap = False
        '
        'btnAddClient
        '
        Me.btnAddClient.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnAddClient.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddClient.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddClient.Location = New System.Drawing.Point(9, 331)
        Me.btnAddClient.Name = "btnAddClient"
        Me.btnAddClient.Size = New System.Drawing.Size(120, 20)
        Me.btnAddClient.TabIndex = 7
        Me.btnAddClient.Text = "Add Client"
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(205, 331)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(120, 20)
        Me.btnClose.TabIndex = 8
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cboStatus)
        Me.GroupBox3.Controls.Add(Me.lblStatus)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(7, 274)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(320, 56)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        '
        'cboStatus
        '
        Me.cboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatus.Items.AddRange(New Object() {"Suspect"})
        Me.cboStatus.Location = New System.Drawing.Point(136, 20)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(176, 22)
        Me.cboStatus.TabIndex = 6
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(8, 24)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(120, 16)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Status"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmAddClients
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(330, 356)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnAddClient)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmAddClients"
        Me.Text = "Add New Client"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
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
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        myfrmAddClientsform = 0
        refreshclients = True
        Me.Dispose(False)
    End Sub
    Private Sub frmAddClients_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.txtClientName.Enabled = True
            Me.cboStatus.SelectedIndex = 0
            myfrmAddClientsform = 1
            'Dim mythread As System.Threading.Thread
            'mythread = New System.Threading.Thread(AddressOf dbconnect)

            Dim n As String
            Me.txtClientNo.Text = clientnumber() 'increments client number
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnAddClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddClient.Click
        Try
            Me.Invoke(New mydelegate(AddressOf dir1))
        Catch ex As Exception

        End Try
    End Sub
    Private Sub dir1()
        Dim isvalid As Boolean = False
        Dim connect As New ADODB.Connection
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor

            Dim msg
            Me.txtClientNo.Enabled = False
            Dim rs As New ADODB.Recordset
            Dim number As String
            Dim clientname, desc As String
            With rs
                number = clientnumber()
                txtClientNo.Text = number
                'perform validations
                Dim strsql, str As String
                str = Me.txtClientName.Text
                str = str.Trim
                Dim j = str.Length
                If j = 0 Then
                    MessageBox.Show(Text:="Please enter a valid name", _
                    buttons:=MessageBoxButtons.OK, _
                    Icon:=MessageBoxIcon.Information, caption:="Add Client")
                    Exit Try
                End If
                If Me.txtClientName.Text = "" Then
                    MessageBox.Show(Text:="A client must have a name", _
                    buttons:=MessageBoxButtons.OK, caption:="Add Client")
                    Exit Try
                Else
                    ' clientname = Me.txtClientName.Text
                    Dim a() As String
                    a = txtClientName.Lines
                    clientname = rml(a)
                    clientname = clientname.Replace("'", "\'")
                End If
                If Me.txtDesc.Text.Trim() = "" Then
                    desc = "null"
                    MessageBox.Show(Text:="Please input description", _
                     buttons:=MessageBoxButtons.OK, caption:="Add Client")
                    Exit Try
                Else
                    'desc = Me.txtDesc.Text
                    Dim k() As String
                    k = txtDesc.Lines
                    desc = rml(k)
                End If
                If Me.txtClientNo.Text.Trim.ToUpper = Nothing Then
                    MessageBox.Show(Text:="A client must have a number", _
                    buttons:=MessageBoxButtons.OK, caption:="Add Client")
                    Exit Try
                End If
                number = txtClientNo.Text.ToUpper
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
                If r.BOF = False And r.EOF = False Then
                    MessageBox.Show(Text:="Client number already exists", _
                    buttons:=MessageBoxButtons.OK, caption:="Add Client")
                    Exit Try
                End If
                Dim lno = newlno(txtClientNo.Text)
                clientno = txtClientNo.Text
                leadno = lno

                desc = desc.Replace("'", "\'")
                strsql = strsql & "insert into clients (client_no,name,description,least_status,leads_no) values"
                strsql = strsql & "('" & number & "','" & clientname & "','" & desc & "'," _
                & "'" & cboStatus.Text & "','" & lno & "');"
                strsql = strsql & "insert into leads (client_no,status,leads_no) values "
                strsql = strsql & "('" & number & "','" & cboStatus.Text & "','" & lno & "');"
                connect.BeginTrans()
                connect.IsolationLevel = IsolationLevelEnum.adXactSerializable
                connect.Execute(strsql)
                connect.CommitTrans()
                isvalid = True
                Dim Threadleads As Thread = New System.Threading.Thread( _
                                                              AddressOf loaddirectory)
                Threadleads.IsBackground = True
                Threadleads.Start()
                MessageBox.Show(Text:="Client has been successfully added", _
                buttons:=MessageBoxButtons.OK, caption:="Add Client")
            End With

        Catch gen As Exception

        Finally
            If isvalid = True Then
                Cursor.Current = currentcursor
                txtClientNo.Text = ""
                txtClientNo.Text = clientnumber()
                Me.txtClientName.Text = ""
                txtDesc.Text = ""
                txtClientNo.Enabled = False
                txtClientNo.Focus()
            End If

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try

        Try
            Dim tthread As System.Threading.Thread = New System.Threading.Thread(AddressOf myForms.Main.loadgrid)
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
    Private Function clientnumber() As String
        Dim connect As New ADODB.Connection
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Function
        End Try
        Try

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                Dim str = "select max(client_no) from clients"
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    If Convert.IsDBNull(.Fields("max").Value) = True Then
                        clientnumber = "1000"
                    Else
                        clientnumber = .Fields("max").Value
                        clientnumber = (CLng(clientnumber) + 1).ToString()
                    End If

                Else
                    clientnumber = "1000"

                End If
            End With
            Try
                connect.Close()
            Catch er As Exception

            End Try
        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Private Function clientnumber1() As String
        Dim connect As New ADODB.Connection
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Function
        End Try
        Try
            Dim rs As New ADODB.Recordset
            Dim number, number1, number2 As String
            With rs
                'add new client number
                .Open(Source:="select max(client_no)  from clients", _
                activeconnection:=connect, cursortype:=CursorTypeEnum.adOpenForwardOnly)

                number = .Fields("max").Value
                Dim a() As String


                Dim i As Integer
                number1 = number.Chars(0)
                For i = 1 To number.Length - 1
                    If IsNumeric(number.Chars(i)) = True Then


                        number2 = number2 & number.Chars(i)
                    End If

                Next
                Dim no3 As String

                For i = 0 To number2.Length - 1
                    If i = 0 Then
                        If number2.Substring(i, 1) = "0" Then
                            no3 = no3 & number2.Substring(i, 1)
                        End If
                    End If
                Next i
                'number2 = (CSng(number.Remove(2, 3)) + 1) '.ToString()
                Dim no As Integer
                no = CInt(number2) + 1
                number = number1 & no3 & no.ToString
                Me.txtClientNo.Text = number
                Return number.ToString
                rs.Close()
                rs = Nothing
            End With
        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Function
    Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
        Try
            ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
            ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
            If keyData = System.Windows.Forms.Keys.Return Then
                'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)
                Dim E As System.EventArgs
                'Me.Invoke(New mydelegate(AddressOf dir1))
                'Call btnAddClient_Click(Me, E)

                Return True ' True means we've processed the key
            Else
                Return MyBase.ProcessDialogKey(keyData)
            End If
        Catch ex As Exception
            'Trace.WriteLine(ex.ToString())
            MsgBox(ex.Message.ToString, , Title:="Return key")

        End Try
    End Function
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

        End Try
    End Sub
    Private Sub storedata(ByVal str As String)
        Dim connect As New ADODB.Connection
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

        Catch jb As Exception

            Exit Sub
        End Try
        Try
            ''Me.txtMailBody.SaveFile(str & Me.txtMailSubject.Text, System.Windows.Forms.RichTextBoxStreamType.RichText)
            Dim i, j
            Dim str1, myfilename As String
            Dim dtp As New System.Windows.Forms.DateTimePicker
            Dim rtbjournal As New System.Windows.Forms.RichTextBox
            Dim journalpath As String
            journalpath = str & "\" & leadno & "_" & dtp.Value.Year & dtp.Value.Month & dtp.Value.Day _
            & dtp.Value.Hour & dtp.Value.Minute & dtp.Value.Second & dtp.Value.Millisecond & ".txt"
            rtbjournal.SaveFile(journalpath)
            journalpath = journalpath.Replace("\", "|")

            connect.BeginTrans()
            Dim strsql As String
            strsql = "update leads set journal='" & journalpath & "'"
            strsql += " where leads_no='" & leadno & "'"
            connect.Execute(strsql)
            connect.CommitTrans()
        Catch ex As Exception
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    '#Region "validation"
    '    Private Sub txtDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDesc.KeyPress
    '        Try
    '            Dim vt As New validation()
    '            If vt._validatetextbox(e) = True Then
    '                Me.errp.SetError(Me.txtDesc, _
    '                                      "not allowed chars: ''','%','*','\','*','1'")
    '                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
    '            Else
    '                Me.errp.SetError(Me.txtDesc, "")
    '            End If
    '        Catch xc As Exception

    '        End Try
    '    End Sub
    '    Private Sub txtClientName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtClientName.KeyPress
    '        Try
    '            Dim vt As New validation()
    '            If vt._validatetextbox(e) = True Then
    '                Me.errp.SetError(Me.txtClientName, _
    '                                      "not allowed chars: ''','%','*','\','*','1'")
    '                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
    '            Else
    '                Me.errp.SetError(Me.txtClientName, "")
    '            End If
    '        Catch xc As Exception

    '        End Try
    '    End Sub
    '#End Region
End Class
