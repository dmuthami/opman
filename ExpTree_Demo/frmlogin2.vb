
Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Threading
Imports System.ArgumentNullException
Imports System.NullReferenceException
Imports System.ArgumentOutOfRangeException
Imports System.IO
Imports System.Text

'Imports System.Collections.Specialized
'Imports System.Configuration
Imports ADODB
Public Class frmlogin2
    Inherits System.Windows.Forms.Form
    Public Shared login As New frmlogin2()
    Public Delegate Sub mydelegate()
    Public mythread As System.Threading.Thread

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtuid As System.Windows.Forms.TextBox
    Friend WithEvents txtpwd As System.Windows.Forms.TextBox
    Friend WithEvents lbluid As System.Windows.Forms.Label
    Friend WithEvents lblpwd As System.Windows.Forms.Label
    Friend WithEvents btnLogin As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmlogin2))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnLogin = New System.Windows.Forms.Button
        Me.lblpwd = New System.Windows.Forms.Label
        Me.lbluid = New System.Windows.Forms.Label
        Me.txtpwd = New System.Windows.Forms.TextBox
        Me.txtuid = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnLogin)
        Me.GroupBox1.Controls.Add(Me.lblpwd)
        Me.GroupBox1.Controls.Add(Me.lbluid)
        Me.GroupBox1.Controls.Add(Me.txtpwd)
        Me.GroupBox1.Controls.Add(Me.txtuid)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(2, -5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(264, 93)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'btnLogin
        '
        Me.btnLogin.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnLogin.Location = New System.Drawing.Point(82, 66)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.TabIndex = 3
        Me.btnLogin.Text = "Login"
        '
        'lblpwd
        '
        Me.lblpwd.Location = New System.Drawing.Point(11, 39)
        Me.lblpwd.Name = "lblpwd"
        Me.lblpwd.Size = New System.Drawing.Size(104, 17)
        Me.lblpwd.TabIndex = 1
        Me.lblpwd.Text = "Password"
        '
        'lbluid
        '
        Me.lbluid.Location = New System.Drawing.Point(11, 13)
        Me.lbluid.Name = "lbluid"
        Me.lbluid.Size = New System.Drawing.Size(104, 19)
        Me.lbluid.TabIndex = 2
        Me.lbluid.Text = "User name"
        '
        'txtpwd
        '
        Me.txtpwd.Location = New System.Drawing.Point(120, 40)
        Me.txtpwd.Name = "txtpwd"
        Me.txtpwd.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtpwd.Size = New System.Drawing.Size(136, 20)
        Me.txtpwd.TabIndex = 2
        Me.txtpwd.Text = ""
        '
        'txtuid
        '
        Me.txtuid.Location = New System.Drawing.Point(120, 16)
        Me.txtuid.Name = "txtuid"
        Me.txtuid.Size = New System.Drawing.Size(136, 20)
        Me.txtuid.TabIndex = 1
        Me.txtuid.Text = ""
        '
        'frmlogin2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(266, 92)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmlogin2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    <System.STAThread()> _
          Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run(New frmlogin2)
    End Sub
#End Region

#Region "private members"
    Private duser As String
    Private arr() As String
    Private strr As String = "jesus"
    Private strconn As String
    'dsn
    'server
    'port
    'password
    'username
    'database
    'path
    'duser
    'smtp
#End Region

#Region "login"
    Private Sub loginn()

    End Sub
    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            isloading = True
            Dim otter As New ThreadStart(AddressOf cvb)
            mythread = New System.Threading.Thread(otter)
            mythread.Start()
            Me.txtuid.Focus()
            btnLogin.Width = 120
        Catch ex As Exception
            MsgBox(ex.Message.ToString() & vbCrLf _
                      & ex.InnerException().ToString() & vbCrLf _
                      & ex.StackTrace.ToString())
        End Try
        
     
    End Sub
    Private Sub cvb()
        Try
            Call retrievesettings()
            Me.txtuid.Text = Me.duser
            closeprogram()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.Dispose(True)
        Catch xc As Exception
            DisplayError(xc)
        End Try

    End Sub
    Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
        Try
            ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
            ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
            If keyData = System.Windows.Forms.Keys.Return Then
                'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)
                Dim E As System.EventArgs
                'Me.Invoke(New mydelegate(AddressOf loginn))

                Call btnLogin_Click(Me, E)

                Return True ' True means we've processed the key
            Else
                Return MyBase.ProcessDialogKey(keyData)
            End If
        Catch ex As Exception
            '  Trace.WriteLine(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())

        End Try
    End Function
    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Try
            Dim cnnstr As String
            cnnstr = "Data source=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()
            Dim currentcursor As Cursor = Cursor.Current
            Try
                Cursor.Current = Cursors.WaitCursor
                Dim strpassword As String
                Dim strusername As String

                Dim str, strseclevel As String
                'Name, password, id_no, seclevel
                str = "select * from seccheck " _
                    & " where lower(name) = '" & LCase(txtuid.Text) & "'" _
                    & " and password = '" & txtpwd.Text & "'"
                Dim rs As New ADODB.Recordset
                With rs
                    .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                    .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                    .Open(str, connect)
                    If .BOF = False And .EOF = False Then
                        If Convert.IsDBNull(.Fields("id_no").Value) = False Then
                            Dim uid As String = .Fields("id_no").Value
                            myForms.id_no = uid

                        Else
                            myForms.id_no = "null"
                        End If

                        If Convert.ToString(.Fields("seclevel").Value).Trim.Length > 0 Then
                            strseclevel = .Fields("seclevel").Value

                        Else
                            'Try
                            '    strseclevel = .Fields("seclevel").Value
                            'Catch ex As Exception

                            'End Try
                            strseclevel = "null"
                        End If
                        rs.Close()
                        rs = Nothing
                        Dim myMain As New frmHome
                        myForms.Main = myMain

                        '-----------authorization
                        myForms.Main.pnlleads.Visible = False
                        myForms.Main.pnlclients.Visible = False
                        myForms.Main.pnljobs.Visible = False
                        myForms.Main.pnlequipmain.Visible = False
                        myForms.Main.pnlpersonnelmain.Visible = False
                        myForms.Main.pnlequipcontrols.Visible = False
                        myForms.Main.ToolBar2.Visible = False
                        myForms.Main.mnufilesettings.Enabled = False
                        '------------------
                        If strseclevel <> "null" Then
                            myForms.Main.seclevel = strseclevel

                        End If
                        Me.Dispose(False)
                        myForms.Main.Show()

                        '--------------------------authorizing users rights
                        'Dim Tasks As New taskclass
                        'Try
                        '    Tasks.adminarray = strseclevel.Split(":")
                        '    Dim Threadh1 As New System.Threading.Thread( _
                        '                            AddressOf taskclass.homeinvoke)
                        '    Threadh1.Start()
                        'Catch we As Exception
                        'End Try
                        '---------------------------
                    Else
                        Dim myMain As New frmHome
                        myForms.Main = myMain
                        '-----------authorization
                        myForms.Main.pnlleads.Visible = False
                        myForms.Main.pnlclients.Visible = False
                        myForms.Main.pnljobs.Visible = False
                        myForms.Main.pnlequipmain.Visible = False
                        myForms.Main.pnlpersonnelmain.Visible = False
                        myForms.Main.pnlequipcontrols.Visible = False
                        myForms.Main.ToolBar2.Visible = False
                        myForms.Main.mnufilesettings.Enabled = False
                        '------------------

                        myForms.Main.seclevel = "0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0"
                        If MessageBox.Show("Invalid UserName or Password" & vbCrLf & "Click Yes button to continue and change the connection settings or" & vbCrLf _
                        & "Click No button to change your login details", _
                         "Login", MessageBoxButtons.YesNo, _
                           MessageBoxIcon.Information) = DialogResult.Yes Then
                            Me.Dispose(False)
                            myForms.Main.Show()
                        End If

                    End If

                End With
            Catch ex As Exception
                DisplayError(ex)
            Finally
                Cursor.Current = currentcursor
            End Try
            Try
                connect.Close()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            DisplayError(ex)
        End Try

    End Sub
    Private Sub DisplayError(ByVal ex As Exception)
        Try
            If MessageBox.Show(ex.GetType().ToString() & _
                    vbCrLf & vbCrLf & _
                    ex.Message & vbCrLf & vbCrLf & _
                    ex.StackTrace, _
                    "Error", _
                    MessageBoxButtons.AbortRetryIgnore, _
                    MessageBoxIcon.Stop) = DialogResult.Ignore Then
                Dim myMain As New frmHome
                myForms.Main = myMain
                '-----------authorization
                myForms.Main.pnlleads.Visible = False
                myForms.Main.pnlclients.Visible = False
                myForms.Main.pnljobs.Visible = False
                myForms.Main.pnlequipmain.Visible = False
                myForms.Main.pnlpersonnelmain.Visible = False
                myForms.Main.pnlequipcontrols.Visible = False
                myForms.Main.ToolBar2.Visible = False
                myForms.Main.mnufilesettings.Enabled = False
                '------------------
                Me.Dispose(False)
                myForms.Main.Show()
            ElseIf DialogResult.Abort Then
                Application.Exit()
            End If
        Catch vb As Exception
        End Try

    End Sub
    Private Sub retrievesettings()
        Try
            readfile()
            con()
            Try
                System.GC.Collect()
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "encrypt/decrypt"
    Private Sub readfile()
        Try
            Dim objStreamReader As StreamReader
            Dim strLine As String
            'Pass the file path and the file name to the StreamReader constructor.
            objStreamReader = New StreamReader(Application.StartupPath & "\output.txt")
            'Read the first line of text.
            strLine = objStreamReader.ReadLine
            Dim rtb As New System.Windows.Forms.RichTextBox
            'Continue to read until you reach the end of the file.
            Do While Not strLine Is Nothing
                Application.DoEvents()
                rtb.AppendText(strLine & vbCrLf)
                'Read the next line.
                strLine = objStreamReader.ReadLine
            Loop

            arr = rtb.Lines
            'Close the file.
            objStreamReader.Close()
            System.GC.Collect()

        Catch ex As Exception

        End Try
    End Sub
    Sub con()
        Try
            Dim int As Integer = arr.GetUpperBound(0)
            Dim y As Integer = 0
            Dim bn As String
            Dim en As New en_de_crypt
            For y = 0 To int
                Application.DoEvents()
                arr.SetValue(en.EnDeCrypt(Convert.ToString(arr.GetValue(y)), strr, False), y)
                bn = arr.GetValue(y)
                bn = bn.Replace("///", "t")
                arr.SetValue(bn, y)
            Next
            strconn = arr.GetValue(0) & ";"
            strconn += "Password=" & arr.GetValue(3) & ";"
            strconn += "User Id=" & arr.GetValue(4) & ";"
            strconn += "Database=" & arr.GetValue(5) & ";"
            strconn += "Server=" & arr.GetValue(1) & ";"
            strconn += "Port=" & arr.GetValue(2) & ";"
            duser = arr.GetValue(7)
            myForms.qconnstr = strconn
            myForms.qfolderpath = arr.GetValue(6)
            myForms.str_r = strr
            myForms.mailserver = arr.GetValue(8)

            'Me.RichTextBox1.Lines = arr
            Dim cv As String = "yuy"
            System.GC.Collect()
        Catch ex As Exception

        End Try

    End Sub
#End Region

End Class
