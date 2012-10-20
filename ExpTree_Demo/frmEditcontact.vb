Imports ADODB
Imports System.Text.StringBuilder
Imports System.Object
Imports System.Data
Imports System



Imports System.Threading

Public Class frmEditcontact
    Inherits System.Windows.Forms.Form
    Public desc, fname, sname, salu, pobox, e_mail1, e_mail2, fax, tel As String
    Public cell, phyadd, mobile2 As String
    Public Delegate Sub mydelegate()

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
            editcontacts = False
            myForms.CustomerForm1 = Nothing
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
    Friend WithEvents lblCellPhone As System.Windows.Forms.Label
    Friend WithEvents txtTelephone As System.Windows.Forms.TextBox
    Friend WithEvents txtCellPhone As System.Windows.Forms.TextBox
    Friend WithEvents lblPostalAddress As System.Windows.Forms.Label
    Friend WithEvents lblEMail As System.Windows.Forms.Label
    Friend WithEvents lblPhysicalAddress As System.Windows.Forms.Label
    Friend WithEvents txtPostalAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtPhysicalAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents txtFax As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lbFirstName As System.Windows.Forms.Label
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents lblSecondName As System.Windows.Forms.Label
    Friend WithEvents txtSecondName As System.Windows.Forms.TextBox
    Friend WithEvents txtEMail1 As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail2 As System.Windows.Forms.TextBox
    Friend WithEvents lblEMail2 As System.Windows.Forms.Label
    Friend WithEvents lbldesc As System.Windows.Forms.Label
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblClientNo As System.Windows.Forms.Label
    Friend WithEvents lblClientName As System.Windows.Forms.Label
    Friend WithEvents cboSalutation As System.Windows.Forms.ComboBox
    Friend WithEvents btnEditContacts As System.Windows.Forms.Button
    Friend WithEvents txtcellphone2 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEditcontact))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtcellphone2 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.lbldesc = New System.Windows.Forms.Label
        Me.lblEMail2 = New System.Windows.Forms.Label
        Me.txtEmail2 = New System.Windows.Forms.TextBox
        Me.lbFirstName = New System.Windows.Forms.Label
        Me.txtFirstName = New System.Windows.Forms.TextBox
        Me.txtFax = New System.Windows.Forms.TextBox
        Me.lblFax = New System.Windows.Forms.Label
        Me.lblSecondName = New System.Windows.Forms.Label
        Me.txtSecondName = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnEditContacts = New System.Windows.Forms.Button
        Me.txtPhysicalAddress = New System.Windows.Forms.TextBox
        Me.txtEMail1 = New System.Windows.Forms.TextBox
        Me.txtPostalAddress = New System.Windows.Forms.TextBox
        Me.lblPhysicalAddress = New System.Windows.Forms.Label
        Me.lblEMail = New System.Windows.Forms.Label
        Me.lblPostalAddress = New System.Windows.Forms.Label
        Me.txtCellPhone = New System.Windows.Forms.TextBox
        Me.txtTelephone = New System.Windows.Forms.TextBox
        Me.lblCellPhone = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboSalutation = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblClientNo = New System.Windows.Forms.Label
        Me.lblClientName = New System.Windows.Forms.Label
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtcellphone2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtdesc)
        Me.GroupBox1.Controls.Add(Me.lbldesc)
        Me.GroupBox1.Controls.Add(Me.lblEMail2)
        Me.GroupBox1.Controls.Add(Me.txtEmail2)
        Me.GroupBox1.Controls.Add(Me.lbFirstName)
        Me.GroupBox1.Controls.Add(Me.txtFirstName)
        Me.GroupBox1.Controls.Add(Me.txtFax)
        Me.GroupBox1.Controls.Add(Me.lblFax)
        Me.GroupBox1.Controls.Add(Me.lblSecondName)
        Me.GroupBox1.Controls.Add(Me.txtSecondName)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.btnEditContacts)
        Me.GroupBox1.Controls.Add(Me.txtPhysicalAddress)
        Me.GroupBox1.Controls.Add(Me.txtEMail1)
        Me.GroupBox1.Controls.Add(Me.txtPostalAddress)
        Me.GroupBox1.Controls.Add(Me.lblPhysicalAddress)
        Me.GroupBox1.Controls.Add(Me.lblEMail)
        Me.GroupBox1.Controls.Add(Me.lblPostalAddress)
        Me.GroupBox1.Controls.Add(Me.txtCellPhone)
        Me.GroupBox1.Controls.Add(Me.txtTelephone)
        Me.GroupBox1.Controls.Add(Me.lblCellPhone)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cboSalutation)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(328, 336)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Client Contacts"
        '
        'txtcellphone2
        '
        Me.txtcellphone2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcellphone2.Location = New System.Drawing.Point(128, 168)
        Me.txtcellphone2.Name = "txtcellphone2"
        Me.txtcellphone2.Size = New System.Drawing.Size(192, 20)
        Me.txtcellphone2.TabIndex = 7
        Me.txtcellphone2.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 168)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 16)
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "Mobile No(2)"
        '
        'txtdesc
        '
        Me.txtdesc.Location = New System.Drawing.Point(128, 264)
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(192, 20)
        Me.txtdesc.TabIndex = 11
        Me.txtdesc.Text = ""
        '
        'lbldesc
        '
        Me.lbldesc.Location = New System.Drawing.Point(16, 264)
        Me.lbldesc.Name = "lbldesc"
        Me.lbldesc.Size = New System.Drawing.Size(104, 16)
        Me.lbldesc.TabIndex = 23
        Me.lbldesc.Text = "Description"
        '
        'lblEMail2
        '
        Me.lblEMail2.Location = New System.Drawing.Point(16, 240)
        Me.lblEMail2.Name = "lblEMail2"
        Me.lblEMail2.Size = New System.Drawing.Size(104, 16)
        Me.lblEMail2.TabIndex = 22
        Me.lblEMail2.Text = "E Mail Address(2)"
        '
        'txtEmail2
        '
        Me.txtEmail2.Location = New System.Drawing.Point(128, 240)
        Me.txtEmail2.Name = "txtEmail2"
        Me.txtEmail2.Size = New System.Drawing.Size(192, 20)
        Me.txtEmail2.TabIndex = 10
        Me.txtEmail2.Text = ""
        '
        'lbFirstName
        '
        Me.lbFirstName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbFirstName.Location = New System.Drawing.Point(16, 44)
        Me.lbFirstName.Name = "lbFirstName"
        Me.lbFirstName.Size = New System.Drawing.Size(104, 16)
        Me.lbFirstName.TabIndex = 20
        Me.lbFirstName.Text = "First Name"
        '
        'txtFirstName
        '
        Me.txtFirstName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstName.Location = New System.Drawing.Point(128, 44)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(192, 20)
        Me.txtFirstName.TabIndex = 2
        Me.txtFirstName.Text = ""
        '
        'txtFax
        '
        Me.txtFax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFax.Location = New System.Drawing.Point(128, 116)
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(192, 20)
        Me.txtFax.TabIndex = 5
        Me.txtFax.Text = ""
        '
        'lblFax
        '
        Me.lblFax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFax.Location = New System.Drawing.Point(16, 116)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(104, 16)
        Me.lblFax.TabIndex = 14
        Me.lblFax.Text = "Fax"
        '
        'lblSecondName
        '
        Me.lblSecondName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSecondName.Location = New System.Drawing.Point(16, 68)
        Me.lblSecondName.Name = "lblSecondName"
        Me.lblSecondName.Size = New System.Drawing.Size(104, 16)
        Me.lblSecondName.TabIndex = 13
        Me.lblSecondName.Text = "Second  Name"
        '
        'txtSecondName
        '
        Me.txtSecondName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSecondName.Location = New System.Drawing.Point(128, 68)
        Me.txtSecondName.Name = "txtSecondName"
        Me.txtSecondName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSecondName.Size = New System.Drawing.Size(192, 20)
        Me.txtSecondName.TabIndex = 3
        Me.txtSecondName.Text = ""
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(199, 312)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(120, 20)
        Me.btnClose.TabIndex = 14
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        '
        'btnEditContacts
        '
        Me.btnEditContacts.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnEditContacts.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnEditContacts.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEditContacts.Location = New System.Drawing.Point(8, 312)
        Me.btnEditContacts.Name = "btnEditContacts"
        Me.btnEditContacts.Size = New System.Drawing.Size(120, 20)
        Me.btnEditContacts.TabIndex = 13
        Me.btnEditContacts.Text = "Save Changes"
        '
        'txtPhysicalAddress
        '
        Me.txtPhysicalAddress.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhysicalAddress.Location = New System.Drawing.Point(128, 288)
        Me.txtPhysicalAddress.Name = "txtPhysicalAddress"
        Me.txtPhysicalAddress.Size = New System.Drawing.Size(192, 20)
        Me.txtPhysicalAddress.TabIndex = 12
        Me.txtPhysicalAddress.Text = ""
        '
        'txtEMail1
        '
        Me.txtEMail1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEMail1.Location = New System.Drawing.Point(128, 216)
        Me.txtEMail1.Name = "txtEMail1"
        Me.txtEMail1.Size = New System.Drawing.Size(192, 20)
        Me.txtEMail1.TabIndex = 9
        Me.txtEMail1.Text = ""
        '
        'txtPostalAddress
        '
        Me.txtPostalAddress.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostalAddress.Location = New System.Drawing.Point(128, 192)
        Me.txtPostalAddress.Name = "txtPostalAddress"
        Me.txtPostalAddress.Size = New System.Drawing.Size(192, 20)
        Me.txtPostalAddress.TabIndex = 8
        Me.txtPostalAddress.Text = ""
        '
        'lblPhysicalAddress
        '
        Me.lblPhysicalAddress.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhysicalAddress.Location = New System.Drawing.Point(16, 288)
        Me.lblPhysicalAddress.Name = "lblPhysicalAddress"
        Me.lblPhysicalAddress.Size = New System.Drawing.Size(104, 16)
        Me.lblPhysicalAddress.TabIndex = 6
        Me.lblPhysicalAddress.Text = "Physical Address"
        '
        'lblEMail
        '
        Me.lblEMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEMail.Location = New System.Drawing.Point(16, 216)
        Me.lblEMail.Name = "lblEMail"
        Me.lblEMail.Size = New System.Drawing.Size(104, 16)
        Me.lblEMail.TabIndex = 5
        Me.lblEMail.Text = "E Mail Address(1)"
        '
        'lblPostalAddress
        '
        Me.lblPostalAddress.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPostalAddress.Location = New System.Drawing.Point(16, 192)
        Me.lblPostalAddress.Name = "lblPostalAddress"
        Me.lblPostalAddress.Size = New System.Drawing.Size(104, 16)
        Me.lblPostalAddress.TabIndex = 4
        Me.lblPostalAddress.Text = "P.O Box"
        '
        'txtCellPhone
        '
        Me.txtCellPhone.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCellPhone.Location = New System.Drawing.Point(128, 140)
        Me.txtCellPhone.Name = "txtCellPhone"
        Me.txtCellPhone.Size = New System.Drawing.Size(192, 20)
        Me.txtCellPhone.TabIndex = 6
        Me.txtCellPhone.Text = ""
        '
        'txtTelephone
        '
        Me.txtTelephone.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTelephone.Location = New System.Drawing.Point(128, 92)
        Me.txtTelephone.Name = "txtTelephone"
        Me.txtTelephone.Size = New System.Drawing.Size(192, 20)
        Me.txtTelephone.TabIndex = 4
        Me.txtTelephone.Text = ""
        '
        'lblCellPhone
        '
        Me.lblCellPhone.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCellPhone.Location = New System.Drawing.Point(16, 140)
        Me.lblCellPhone.Name = "lblCellPhone"
        Me.lblCellPhone.Size = New System.Drawing.Size(104, 16)
        Me.lblCellPhone.TabIndex = 1
        Me.lblCellPhone.Text = "Mobile No(1)"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 92)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 16)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Telephone"
        '
        'cboSalutation
        '
        Me.cboSalutation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSalutation.Location = New System.Drawing.Point(128, 16)
        Me.cboSalutation.Name = "cboSalutation"
        Me.cboSalutation.Size = New System.Drawing.Size(192, 22)
        Me.cboSalutation.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Salutation"
        '
        'lblClientNo
        '
        Me.lblClientNo.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lblClientNo.Location = New System.Drawing.Point(8, 8)
        Me.lblClientNo.Name = "lblClientNo"
        Me.lblClientNo.Size = New System.Drawing.Size(104, 24)
        Me.lblClientNo.TabIndex = 1
        '
        'lblClientName
        '
        Me.lblClientName.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lblClientName.Location = New System.Drawing.Point(120, 8)
        Me.lblClientName.Name = "lblClientName"
        Me.lblClientName.Size = New System.Drawing.Size(216, 24)
        Me.lblClientName.TabIndex = 2
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmEditcontact
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(344, 380)
        Me.Controls.Add(Me.lblClientName)
        Me.Controls.Add(Me.lblClientNo)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmEditcontact"
        Me.Text = "Change Contacts"
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
#Region "public members"
    Public autono As String
#End Region

#Region " edit contacts"
    Private Sub btnEditContacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditContacts.Click
        Try
            Me.Invoke(New mydelegate(AddressOf editcontact))
        Catch ex As Exception

        End Try
    End Sub
    Private Sub editcontact()
        Dim currentcursor As Cursor = Cursor.Current
        Try

            Dim strsql As String
            Cursor.Current = Cursors.WaitCursor
            'With rs
            Dim str As String

            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim strin As String = txtFirstName.Text.Trim
            Dim strin1 As String = txtTelephone.Text.Trim
            Dim strin2 As String = txtFax.Text.Trim
            Dim strin3 As String = txtCellPhone.Text.Trim
            Dim strin4 As String = txtEmail2.Text.Trim
            Dim strin5 As String = txtSecondName.Text.Trim
            Dim strin6 As String = cboSalutation.Text.Trim
            Dim strin7 As String = txtdesc.Text.Trim
            Dim strin8 As String = txtPostalAddress.Text.Trim
            Dim strin9 As String = txtPhysicalAddress.Text.Trim
            Dim strin10 As String = txtEMail1.Text.Trim
            Dim strin11 As String = txtcellphone2.Text.Trim

            strin = strin.Replace("'", "\'")
            strin1 = strin1.Replace("'", "\'")
            strin2 = strin2.Replace("'", "\'")
            strin3 = strin3.Replace("'", "\'")
            strin4 = strin4.Replace("'", "\'")
            strin5 = strin5.Replace("'", "\'")
            strin6 = strin6.Replace("'", "\'")
            strin7 = strin7.Replace("'", "\'")
            strin8 = strin8.Replace("'", "\'")
            strin9 = strin9.Replace("'", "\'")
            strin10 = strin10.Replace("'", "\'")
            strin11 = strin11.Replace("'", "\'")

            strsql = "update contact set "
            strsql = strsql & "client_no='" & Me.lblClientNo.Text & "',"
            strsql = strsql & "f_name='" & strin & "',"
            strsql = strsql & "tel='" & strin1 & "',"
            strsql = strsql & "fax='" & strin2 & "',"
            strsql = strsql & "cell='" & strin3 & "',"

            strsql = strsql & "e_mail2='" & strin4 & "',"
            strsql = strsql & "s_name='" & strin5 & "',"
            strsql = strsql & "salutation='" & strin6 & "',"
            strsql = strsql & "description='" & strin7 & "',"

            strsql = strsql & "pobox='" & strin8 & "',"
            strsql = strsql & "physicaladd='" & strin9 & "',"
            strsql = strsql & "mobile2='" & strin11 & "',"
            strsql = strsql & "e_mail1='" & strin10 & "'"

            strsql = strsql & " where ano='" & autono & "' ;"
            connect.BeginTrans()
            connect.IsolationLevel = ADODB.IsolationLevelEnum.adXactSerializable

            connect.Execute(strsql)
            connect.CommitTrans()
            MessageBox.Show(Text:="Changes to Contacts have been made", _
            caption:="Add contact", buttons:=MessageBoxButtons.OK, _
            Icon:=MessageBoxIcon.Information)
            refreshcontacts = True

        Catch ex As Exception

        Finally
            'Me.txtFirstName.Text = ""
            'txtTelephone.Text = ""
            'txtFax.Text = ""
            'txtCellPhone.Text = ""
            'txtPostalAddress.Text = ""
            'txtPhysicalAddress.Text = ""
            'Me.txtEMail1.Text = ""
            'Me.txtEmail2.Text = ""
            'Me.txtSecondName.Text = ""
            'Me.txtdesc.Text = ""
            Me.cboSalutation.Focus()
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
    End Sub
    Private Sub threadjobs()
        Try
            myForms.CustomerForm3.Invoke(New mydelegate(AddressOf myForms.CustomerForm3.loadgridcontact))

        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        editcontacts = False
        myForms.CustomerForm1 = Nothing
        Me.Dispose(True)
    End Sub
    Private Sub frmEditcontact_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        editcontacts = True
        Me.lblClientNo.Text = myclientno
        Me.lblClientName.Text = myclientname
        Me.txtdesc.Text = desc
        Me.txtCellPhone.Text = cell
        Me.txtEMail1.Text = e_mail1
        Me.txtEmail2.Text = e_mail2
        Me.txtFax.Text = fax
        Me.txtFirstName.Text = fname
        Me.txtPhysicalAddress.Text = phyadd
        Me.txtPostalAddress.Text = pobox
        Me.txtSecondName.Text = sname
        Me.txtTelephone.Text = tel

        Me.cboSalutation.Items.Add("Mr")
        Me.cboSalutation.Items.Add("Mrs")
        Me.cboSalutation.Items.Add("Prof")
        Me.cboSalutation.Items.Add("Dr")
        Me.cboSalutation.Items.Add("Miss")
        Me.cboSalutation.Text = Me.salu
        Me.txtcellphone2.Text = mobile2

    End Sub
    Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
        Try
            ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
            ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
            If keyData = System.Windows.Forms.Keys.Return Then
                'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)
                Dim E As System.EventArgs
                'Me.Invoke(New mydelegate(AddressOf editcontact))
                'Call btnEditContacts_Click(Me, E)

                Return True ' True means we've processed the key
            Else
                Return MyBase.ProcessDialogKey(keyData)
            End If
        Catch ex As Exception
            'Trace.WriteLine(ex.ToString())
            MsgBox(ex.Message.ToString, , Title:="Return key")

        End Try
    End Function
#End Region

#Region " validation"
    Private Sub txtTelephone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTelephone.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtTelephone, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtTelephone, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtCellPhone_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCellPhone.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtCellPhone, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtCellPhone, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtPostalAddress_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPostalAddress.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtPostalAddress, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtPostalAddress, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtPhysicalAddress_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPhysicalAddress.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtPhysicalAddress, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtPhysicalAddress, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtFax_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFax.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtFax, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtFax, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtFirstName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFirstName.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtFirstName, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtFirstName, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtSecondName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSecondName.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtSecondName, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtSecondName, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtEMail1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEMail1.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtEMail1, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtEMail1, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtEmail2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEmail2.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtEmail2, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtEmail2, "")
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
    Private Sub txtcellphone2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcellphone2.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtcellphone2, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtcellphone2, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region
End Class



