Imports System
Imports System.IO
Public Class frmconsettings
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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnsave As System.Windows.Forms.Button
    Friend WithEvents txtuid As System.Windows.Forms.TextBox
    Friend WithEvents txtpwd As System.Windows.Forms.TextBox
    Friend WithEvents txtport As System.Windows.Forms.TextBox
    Friend WithEvents txtserver As System.Windows.Forms.TextBox
    Friend WithEvents txtdsn As System.Windows.Forms.TextBox
    Friend WithEvents txtfpath As System.Windows.Forms.TextBox
    Friend WithEvents txtdefuser As System.Windows.Forms.TextBox
    Friend WithEvents txtdb As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtsmtp As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmconsettings))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtdb = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtdefuser = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtfpath = New System.Windows.Forms.TextBox
        Me.txtdsn = New System.Windows.Forms.TextBox
        Me.txtserver = New System.Windows.Forms.TextBox
        Me.txtport = New System.Windows.Forms.TextBox
        Me.txtpwd = New System.Windows.Forms.TextBox
        Me.txtuid = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnsave = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtsmtp = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.txtdb)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtdefuser)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtfpath)
        Me.GroupBox1.Controls.Add(Me.txtdsn)
        Me.GroupBox1.Controls.Add(Me.txtserver)
        Me.GroupBox1.Controls.Add(Me.txtport)
        Me.GroupBox1.Controls.Add(Me.txtpwd)
        Me.GroupBox1.Controls.Add(Me.txtuid)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(6, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(322, 216)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Database settings"
        '
        'txtdb
        '
        Me.txtdb.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdb.Location = New System.Drawing.Point(117, 40)
        Me.txtdb.Name = "txtdb"
        Me.txtdb.Size = New System.Drawing.Size(200, 20)
        Me.txtdb.TabIndex = 15
        Me.txtdb.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(5, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 16)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Database"
        '
        'txtdefuser
        '
        Me.txtdefuser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdefuser.Location = New System.Drawing.Point(117, 182)
        Me.txtdefuser.Name = "txtdefuser"
        Me.txtdefuser.Size = New System.Drawing.Size(200, 20)
        Me.txtdefuser.TabIndex = 7
        Me.txtdefuser.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(5, 182)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(100, 16)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Default user"
        '
        'txtfpath
        '
        Me.txtfpath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtfpath.Location = New System.Drawing.Point(117, 157)
        Me.txtfpath.Name = "txtfpath"
        Me.txtfpath.Size = New System.Drawing.Size(200, 20)
        Me.txtfpath.TabIndex = 6
        Me.txtfpath.Text = ""
        '
        'txtdsn
        '
        Me.txtdsn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtdsn.Location = New System.Drawing.Point(117, 16)
        Me.txtdsn.Name = "txtdsn"
        Me.txtdsn.Size = New System.Drawing.Size(200, 20)
        Me.txtdsn.TabIndex = 1
        Me.txtdsn.Text = ""
        '
        'txtserver
        '
        Me.txtserver.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtserver.Location = New System.Drawing.Point(117, 64)
        Me.txtserver.Name = "txtserver"
        Me.txtserver.Size = New System.Drawing.Size(200, 20)
        Me.txtserver.TabIndex = 2
        Me.txtserver.Text = ""
        '
        'txtport
        '
        Me.txtport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtport.Location = New System.Drawing.Point(117, 88)
        Me.txtport.Name = "txtport"
        Me.txtport.Size = New System.Drawing.Size(200, 20)
        Me.txtport.TabIndex = 3
        Me.txtport.Text = ""
        '
        'txtpwd
        '
        Me.txtpwd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtpwd.Location = New System.Drawing.Point(117, 111)
        Me.txtpwd.Name = "txtpwd"
        Me.txtpwd.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtpwd.Size = New System.Drawing.Size(200, 20)
        Me.txtpwd.TabIndex = 4
        Me.txtpwd.Text = ""
        '
        'txtuid
        '
        Me.txtuid.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtuid.Location = New System.Drawing.Point(117, 133)
        Me.txtuid.Name = "txtuid"
        Me.txtuid.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtuid.Size = New System.Drawing.Size(200, 20)
        Me.txtuid.TabIndex = 5
        Me.txtuid.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(5, 157)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 16)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Folder path"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(5, 133)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 16)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "User name"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(5, 111)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Password"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(7, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Port"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Server"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(5, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Data source"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnclose)
        Me.GroupBox2.Controls.Add(Me.btnsave)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(0, 262)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(330, 40)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        '
        'btnclose
        '
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(245, 11)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 10
        Me.btnclose.Text = "Close"
        '
        'btnsave
        '
        Me.btnsave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsave.Location = New System.Drawing.Point(3, 10)
        Me.btnsave.Name = "btnsave"
        Me.btnsave.Size = New System.Drawing.Size(101, 23)
        Me.btnsave.TabIndex = 9
        Me.btnsave.Text = "Save settings"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.txtsmtp)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(8, 216)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(320, 45)
        Me.GroupBox3.TabIndex = 9
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Mail settings"
        '
        'txtsmtp
        '
        Me.txtsmtp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtsmtp.Location = New System.Drawing.Point(112, 16)
        Me.txtsmtp.Name = "txtsmtp"
        Me.txtsmtp.Size = New System.Drawing.Size(200, 20)
        Me.txtsmtp.TabIndex = 14
        Me.txtsmtp.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(5, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(99, 16)
        Me.Label9.TabIndex = 15
        Me.Label9.Text = "Smtp server"
        '
        'frmconsettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(330, 302)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmconsettings"
        Me.Text = "Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region ""
    Private arrcon() As String
#End Region

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Try
            arrcon(0) = Me.txtdsn.Text
            arrcon(6) = Me.txtfpath.Text
            arrcon(2) = Me.txtport.Text
            arrcon(3) = Me.txtpwd.Text
            arrcon(1) = Me.txtserver.Text
            arrcon(4) = Me.txtuid.Text
            arrcon(7) = Me.txtdefuser.Text
            arrcon(5) = Me.txtdb.Text
            arrcon(8) = Me.txtsmtp.Text
            Dim en As New en_de_crypt
            'arrcon.SetValue(en.EnDeCrypt(txtpwd.Text, myForms.str_r, False), 3)
            'arrcon.SetValue(en.EnDeCrypt(txtuid.Text, myForms.str_r, False), 4)
            writefile()
            MessageBox.Show("Settings applied successfully", "Save settings", _
           MessageBoxButtons.OK, MessageBoxIcon.Information)
            loaddt()
        Catch ex As Exception
            DisplayError(ex)
        End Try
    End Sub
    Private Sub DisplayError(ByVal ex As Exception)
        Try
            MessageBox.Show(ex.GetType().ToString() & _
                    vbCrLf & vbCrLf & _
                    ex.Message & vbCrLf & vbCrLf & _
                    ex.StackTrace, _
                    "Error", _
                    MessageBoxButtons.AbortRetryIgnore, _
                    MessageBoxIcon.Stop)
        Catch vb As Exception
        End Try

    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub writefile()
        Try
            Dim objStreamWriter As StreamWriter
            'Pass the file path and the file name to the StreamWriter constructor.
            objStreamWriter = New StreamWriter(Application.StartupPath & "\output.txt", False)
            Dim int As Integer = arrcon.GetUpperBound(0)
            Dim y As Integer = 0
            Dim bn As String
            Dim en As New en_de_crypt
            For y = 0 To int
                Application.DoEvents()
                'Write a line of text.
                bn = arrcon(y)
                bn = bn.Replace("t", "///")
                arrcon(y) = bn
                arrcon(y) = en.EnDeCrypt(arrcon(y), myForms.str_r, True)
                objStreamWriter.WriteLine(arrcon(y))
            Next

            'Close the file.
            objStreamWriter.Close()

        Catch ex As Exception

        End Try
    End Sub
#Region "encrypt_dencrypty"
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
                'Write the line to a richtext box
                rtb.AppendText(strLine & vbCrLf)

                'Read the next line.
                strLine = objStreamReader.ReadLine
            Loop

            arrcon = rtb.Lines
            'Close the file.
            objStreamReader.Close()
            System.GC.Collect()

        Catch ex As Exception

        End Try
    End Sub
    Sub con()
        Try
            Dim int As Integer = arrcon.GetUpperBound(0)
            Dim y As Integer = 0
            Dim bn As String
            Dim en As New en_de_crypt
            For y = 0 To int
                Try
                    Application.DoEvents()
                    arrcon.SetValue(en.EnDeCrypt(Convert.ToString(arrcon.GetValue(y)), myForms.str_r, False), y)
                    bn = arrcon.GetValue(y)
                    bn = bn.Replace("///", "t")
                    arrcon.SetValue(bn, y)
                Catch ex As Exception

                End Try

            Next
            'arrcon.SetValue(en.EnDeCrypt(Convert.ToString(arrcon.GetValue(3)), myForms.str_r, True), 3)
            'arrcon.SetValue(en.EnDeCrypt(Convert.ToString(arrcon.GetValue(4)), myForms.str_r, True), 4)
            'Me.RichTextBox1.Lines = arr
            Dim cv As String = "yuy"
            System.GC.Collect()
        Catch ex As Exception

        End Try

    End Sub
#End Region
    Sub loaddt()
        Try
            readfile()
            con()
            Me.txtdsn.Text = arrcon(0)
            Me.txtfpath.Text = arrcon(6)
            Me.txtport.Text = arrcon(2)
            Me.txtpwd.Text = arrcon(3)
            Me.txtdb.Text = arrcon(5)
            Me.txtserver.Text = arrcon(1)
            Me.txtuid.Text = arrcon(4)
            Me.txtdefuser.Text = arrcon(7)
            Me.txtsmtp.Text = arrcon(8)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub frmconsettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            loaddt()
            Me.Invalidate(True)
        Catch ex As Exception

        End Try
    End Sub
End Class
