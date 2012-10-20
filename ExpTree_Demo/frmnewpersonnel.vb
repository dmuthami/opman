
Imports System
Imports ADODB
Imports System.Threading

Public Class frmnewpersonnel
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
    Friend WithEvents pnlemployeecontrols As System.Windows.Forms.Panel
    Friend WithEvents btnpassport As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents pnlinput As System.Windows.Forms.Panel
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents pbimage As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents txtname As System.Windows.Forms.TextBox
    Friend WithEvents txtmobileno As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtphoneno As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtbirthday As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents cbogender As System.Windows.Forms.ComboBox
    Friend WithEvents txtidno As System.Windows.Forms.TextBox
    Friend WithEvents txtemail As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtpin As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents txtpostaladdress As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents txtcontractend As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtnhifno As System.Windows.Forms.TextBox
    Friend WithEvents txtnssfno As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txthourlyrate As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtnextofkin As System.Windows.Forms.TextBox
    Friend WithEvents dtpdot As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpdoe As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtmedicalcover As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtcomments As System.Windows.Forms.TextBox
    Friend WithEvents btndepartments As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmnewpersonnel))
        Me.pnlinput = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtcomments = New System.Windows.Forms.TextBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.btndepartments = New System.Windows.Forms.Button
        Me.txtname = New System.Windows.Forms.TextBox
        Me.txtmobileno = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtphoneno = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtbirthday = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.cbogender = New System.Windows.Forms.ComboBox
        Me.txtidno = New System.Windows.Forms.TextBox
        Me.txtemail = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.txtpin = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.txtpostaladdress = New System.Windows.Forms.TextBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.txtcontractend = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtnhifno = New System.Windows.Forms.TextBox
        Me.txtnssfno = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.txthourlyrate = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtnextofkin = New System.Windows.Forms.TextBox
        Me.dtpdot = New System.Windows.Forms.DateTimePicker
        Me.dtpdoe = New System.Windows.Forms.DateTimePicker
        Me.txtmedicalcover = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.pbimage = New System.Windows.Forms.PictureBox
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.pnlemployeecontrols = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnpassport = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.pnlinput.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.pnlemployeecontrols.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlinput
        '
        Me.pnlinput.Controls.Add(Me.GroupBox3)
        Me.pnlinput.Controls.Add(Me.GroupBox5)
        Me.pnlinput.Controls.Add(Me.GroupBox1)
        Me.pnlinput.Controls.Add(Me.StatusBar1)
        Me.pnlinput.Controls.Add(Me.pnlemployeecontrols)
        Me.pnlinput.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlinput.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlinput.Location = New System.Drawing.Point(0, 0)
        Me.pnlinput.Name = "pnlinput"
        Me.pnlinput.Size = New System.Drawing.Size(794, 500)
        Me.pnlinput.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.txtcomments)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(1, 295)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(791, 185)
        Me.GroupBox3.TabIndex = 24
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Comments"
        '
        'txtcomments
        '
        Me.txtcomments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtcomments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcomments.Location = New System.Drawing.Point(8, 16)
        Me.txtcomments.Multiline = True
        Me.txtcomments.Name = "txtcomments"
        Me.txtcomments.Size = New System.Drawing.Size(775, 161)
        Me.txtcomments.TabIndex = 25
        Me.txtcomments.Text = ""
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.btndepartments)
        Me.GroupBox5.Controls.Add(Me.txtname)
        Me.GroupBox5.Controls.Add(Me.txtmobileno)
        Me.GroupBox5.Controls.Add(Me.Label12)
        Me.GroupBox5.Controls.Add(Me.txtphoneno)
        Me.GroupBox5.Controls.Add(Me.Label13)
        Me.GroupBox5.Controls.Add(Me.txtbirthday)
        Me.GroupBox5.Controls.Add(Me.Label16)
        Me.GroupBox5.Controls.Add(Me.Label23)
        Me.GroupBox5.Controls.Add(Me.cbogender)
        Me.GroupBox5.Controls.Add(Me.txtidno)
        Me.GroupBox5.Controls.Add(Me.txtemail)
        Me.GroupBox5.Controls.Add(Me.Label25)
        Me.GroupBox5.Controls.Add(Me.txtpin)
        Me.GroupBox5.Controls.Add(Me.Label26)
        Me.GroupBox5.Controls.Add(Me.Label27)
        Me.GroupBox5.Controls.Add(Me.Label28)
        Me.GroupBox5.Controls.Add(Me.txtpostaladdress)
        Me.GroupBox5.Controls.Add(Me.Label30)
        Me.GroupBox5.Controls.Add(Me.txtcontractend)
        Me.GroupBox5.Controls.Add(Me.Label7)
        Me.GroupBox5.Controls.Add(Me.txtnhifno)
        Me.GroupBox5.Controls.Add(Me.txtnssfno)
        Me.GroupBox5.Controls.Add(Me.Label18)
        Me.GroupBox5.Controls.Add(Me.Label24)
        Me.GroupBox5.Controls.Add(Me.txthourlyrate)
        Me.GroupBox5.Controls.Add(Me.Label19)
        Me.GroupBox5.Controls.Add(Me.Label21)
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.txtnextofkin)
        Me.GroupBox5.Controls.Add(Me.dtpdot)
        Me.GroupBox5.Controls.Add(Me.dtpdoe)
        Me.GroupBox5.Controls.Add(Me.txtmedicalcover)
        Me.GroupBox5.Controls.Add(Me.Label20)
        Me.GroupBox5.Controls.Add(Me.Label6)
        Me.GroupBox5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox5.Location = New System.Drawing.Point(237, 34)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(552, 257)
        Me.GroupBox5.TabIndex = 5
        Me.GroupBox5.TabStop = False
        '
        'btndepartments
        '
        Me.btndepartments.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndepartments.Location = New System.Drawing.Point(112, 112)
        Me.btndepartments.Name = "btndepartments"
        Me.btndepartments.Size = New System.Drawing.Size(32, 20)
        Me.btndepartments.TabIndex = 14
        Me.btndepartments.Tag = "Add or edit job description"
        Me.btndepartments.Text = "A"
        '
        'txtname
        '
        Me.txtname.Location = New System.Drawing.Point(116, 16)
        Me.txtname.Name = "txtname"
        Me.txtname.Size = New System.Drawing.Size(188, 20)
        Me.txtname.TabIndex = 6
        Me.txtname.Text = ""
        '
        'txtmobileno
        '
        Me.txtmobileno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmobileno.Location = New System.Drawing.Point(116, 232)
        Me.txtmobileno.Name = "txtmobileno"
        Me.txtmobileno.Size = New System.Drawing.Size(188, 20)
        Me.txtmobileno.TabIndex = 23
        Me.txtmobileno.Text = ""
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label12.Location = New System.Drawing.Point(7, 232)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(98, 16)
        Me.Label12.TabIndex = 70
        Me.Label12.Text = "Mobile no"
        '
        'txtphoneno
        '
        Me.txtphoneno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtphoneno.Location = New System.Drawing.Point(116, 210)
        Me.txtphoneno.Name = "txtphoneno"
        Me.txtphoneno.Size = New System.Drawing.Size(188, 20)
        Me.txtphoneno.TabIndex = 21
        Me.txtphoneno.Text = ""
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label13.Location = New System.Drawing.Point(7, 210)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(98, 14)
        Me.Label13.TabIndex = 68
        Me.Label13.Text = "Phone no"
        '
        'txtbirthday
        '
        Me.txtbirthday.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbirthday.Location = New System.Drawing.Point(116, 88)
        Me.txtbirthday.Name = "txtbirthday"
        Me.txtbirthday.Size = New System.Drawing.Size(188, 20)
        Me.txtbirthday.TabIndex = 12
        Me.txtbirthday.Text = ""
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label16.Location = New System.Drawing.Point(4, 88)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(98, 16)
        Me.Label16.TabIndex = 54
        Me.Label16.Text = "Birthday"
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(6, 68)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(98, 12)
        Me.Label23.TabIndex = 4
        Me.Label23.Text = "Id no"
        '
        'cbogender
        '
        Me.cbogender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbogender.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbogender.Items.AddRange(New Object() {"Data developer", "Surveyor", "Processor", "Secretary", "Accountant"})
        Me.cbogender.Location = New System.Drawing.Point(144, 111)
        Me.cbogender.Name = "cbogender"
        Me.cbogender.Size = New System.Drawing.Size(160, 21)
        Me.cbogender.TabIndex = 15
        '
        'txtidno
        '
        Me.txtidno.Location = New System.Drawing.Point(116, 64)
        Me.txtidno.Name = "txtidno"
        Me.txtidno.Size = New System.Drawing.Size(188, 20)
        Me.txtidno.TabIndex = 10
        Me.txtidno.Text = ""
        '
        'txtemail
        '
        Me.txtemail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtemail.Location = New System.Drawing.Point(116, 135)
        Me.txtemail.Multiline = True
        Me.txtemail.Name = "txtemail"
        Me.txtemail.Size = New System.Drawing.Size(188, 20)
        Me.txtemail.TabIndex = 17
        Me.txtemail.Text = ""
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label25.Location = New System.Drawing.Point(4, 115)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(98, 16)
        Me.Label25.TabIndex = 44
        Me.Label25.Text = "Job description"
        '
        'txtpin
        '
        Me.txtpin.Location = New System.Drawing.Point(116, 40)
        Me.txtpin.Name = "txtpin"
        Me.txtpin.Size = New System.Drawing.Size(188, 20)
        Me.txtpin.TabIndex = 8
        Me.txtpin.Text = ""
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(6, 41)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(98, 16)
        Me.Label26.TabIndex = 2
        Me.Label26.Text = "Pin no"
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label27.Location = New System.Drawing.Point(5, 140)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(98, 16)
        Me.Label27.TabIndex = 36
        Me.Label27.Text = "E-mail address"
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label28.Location = New System.Drawing.Point(5, 168)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(98, 16)
        Me.Label28.TabIndex = 38
        Me.Label28.Text = "Postal address"
        '
        'txtpostaladdress
        '
        Me.txtpostaladdress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtpostaladdress.Location = New System.Drawing.Point(116, 160)
        Me.txtpostaladdress.Multiline = True
        Me.txtpostaladdress.Name = "txtpostaladdress"
        Me.txtpostaladdress.Size = New System.Drawing.Size(188, 48)
        Me.txtpostaladdress.TabIndex = 19
        Me.txtpostaladdress.Text = ""
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(8, 19)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(96, 16)
        Me.Label30.TabIndex = 0
        Me.Label30.Text = "Name"
        '
        'txtcontractend
        '
        Me.txtcontractend.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtcontractend.Location = New System.Drawing.Point(400, 19)
        Me.txtcontractend.Name = "txtcontractend"
        Me.txtcontractend.Size = New System.Drawing.Size(144, 20)
        Me.txtcontractend.TabIndex = 7
        Me.txtcontractend.Text = ""
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label7.Location = New System.Drawing.Point(304, 19)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(88, 16)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "Contract End Date"
        '
        'txtnhifno
        '
        Me.txtnhifno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnhifno.Location = New System.Drawing.Point(400, 89)
        Me.txtnhifno.Name = "txtnhifno"
        Me.txtnhifno.Size = New System.Drawing.Size(144, 20)
        Me.txtnhifno.TabIndex = 13
        Me.txtnhifno.Text = ""
        '
        'txtnssfno
        '
        Me.txtnssfno.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnssfno.Location = New System.Drawing.Point(400, 66)
        Me.txtnssfno.Name = "txtnssfno"
        Me.txtnssfno.Size = New System.Drawing.Size(144, 20)
        Me.txtnssfno.TabIndex = 11
        Me.txtnssfno.Text = ""
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label18.Location = New System.Drawing.Point(304, 66)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(88, 16)
        Me.Label18.TabIndex = 56
        Me.Label18.Text = "NSSF No"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(304, 42)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(88, 16)
        Me.Label24.TabIndex = 72
        Me.Label24.Text = "Hourly rate"
        '
        'txthourlyrate
        '
        Me.txthourlyrate.Location = New System.Drawing.Point(400, 42)
        Me.txthourlyrate.Name = "txthourlyrate"
        Me.txthourlyrate.Size = New System.Drawing.Size(144, 20)
        Me.txthourlyrate.TabIndex = 9
        Me.txthourlyrate.Text = ""
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label19.Location = New System.Drawing.Point(304, 91)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(88, 16)
        Me.Label19.TabIndex = 58
        Me.Label19.Text = "NHIF No"
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label21.Location = New System.Drawing.Point(305, 140)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(80, 16)
        Me.Label21.TabIndex = 62
        Me.Label21.Text = "Next of Kin"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label4.Location = New System.Drawing.Point(304, 163)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 32)
        Me.Label4.TabIndex = 68
        Me.Label4.Text = "Date of Employment"
        '
        'txtnextofkin
        '
        Me.txtnextofkin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtnextofkin.Location = New System.Drawing.Point(400, 140)
        Me.txtnextofkin.Name = "txtnextofkin"
        Me.txtnextofkin.Size = New System.Drawing.Size(144, 20)
        Me.txtnextofkin.TabIndex = 18
        Me.txtnextofkin.Text = ""
        '
        'dtpdot
        '
        Me.dtpdot.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpdot.Location = New System.Drawing.Point(400, 193)
        Me.dtpdot.Name = "dtpdot"
        Me.dtpdot.Size = New System.Drawing.Size(144, 20)
        Me.dtpdot.TabIndex = 22
        '
        'dtpdoe
        '
        Me.dtpdoe.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpdoe.Location = New System.Drawing.Point(400, 164)
        Me.dtpdoe.Name = "dtpdoe"
        Me.dtpdoe.Size = New System.Drawing.Size(144, 20)
        Me.dtpdoe.TabIndex = 20
        '
        'txtmedicalcover
        '
        Me.txtmedicalcover.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmedicalcover.Location = New System.Drawing.Point(400, 114)
        Me.txtmedicalcover.Name = "txtmedicalcover"
        Me.txtmedicalcover.Size = New System.Drawing.Size(144, 20)
        Me.txtmedicalcover.TabIndex = 16
        Me.txtmedicalcover.Text = ""
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label20.Location = New System.Drawing.Point(304, 114)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(88, 16)
        Me.Label20.TabIndex = 60
        Me.Label20.Text = "Medical Cover "
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label6.Location = New System.Drawing.Point(305, 195)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 32)
        Me.Label6.TabIndex = 70
        Me.Label6.Text = "Date of Termination"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.pbimage)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(1, 33)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(232, 258)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Passport photograph"
        '
        'pbimage
        '
        Me.pbimage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbimage.BackColor = System.Drawing.Color.Gray
        Me.pbimage.Location = New System.Drawing.Point(8, 16)
        Me.pbimage.Name = "pbimage"
        Me.pbimage.Size = New System.Drawing.Size(216, 227)
        Me.pbimage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbimage.TabIndex = 0
        Me.pbimage.TabStop = False
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 484)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(794, 16)
        Me.StatusBar1.TabIndex = 77
        Me.StatusBar1.Text = "All fields marked with an asterix(*) are compulsory"
        '
        'pnlemployeecontrols
        '
        Me.pnlemployeecontrols.Controls.Add(Me.btnclose)
        Me.pnlemployeecontrols.Controls.Add(Me.btnpassport)
        Me.pnlemployeecontrols.Controls.Add(Me.btnAdd)
        Me.pnlemployeecontrols.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlemployeecontrols.Location = New System.Drawing.Point(0, 0)
        Me.pnlemployeecontrols.Name = "pnlemployeecontrols"
        Me.pnlemployeecontrols.Size = New System.Drawing.Size(794, 32)
        Me.pnlemployeecontrols.TabIndex = 0
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(690, 4)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(96, 23)
        Me.btnclose.TabIndex = 3
        Me.btnclose.Text = "Close"
        '
        'btnpassport
        '
        Me.btnpassport.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnpassport.Location = New System.Drawing.Point(5, 3)
        Me.btnpassport.Name = "btnpassport"
        Me.btnpassport.Size = New System.Drawing.Size(123, 23)
        Me.btnpassport.TabIndex = 1
        Me.btnpassport.Text = "Browse for passport"
        '
        'btnAdd
        '
        Me.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAdd.Location = New System.Drawing.Point(130, 4)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(96, 23)
        Me.btnAdd.TabIndex = 2
        Me.btnAdd.Text = "Add personnel"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        '
        'frmnewpersonnel
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(794, 500)
        Me.Controls.Add(Me.pnlinput)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmnewpersonnel"
        Me.Text = "Add new personnel"
        Me.pnlinput.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.pnlemployeecontrols.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub

#Region "personnel"
    Private imagefilename As String
    Private Sub frmnewpersonnel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim _thread As Thread = New Thread(AddressOf ld)
            _thread.IsBackground = True
            _thread.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnpassport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpassport.Click
        Try
            Dim MyImage As Bitmap
            Dim ofd As New System.Windows.Forms.OpenFileDialog
            ofd.Filter = " All image Files(*.BMP;*.JPG;*.GIF;*.JPEG;*.PNG;*.EMF;*.WMF)" _
            & "|*.BMP;*.JPG;*.GIF;*.JPEG;*.PNG;*.EMF;*.WMF"
            ofd.Multiselect = False
            ofd.ShowDialog()
            Dim hb As String = ofd.FileName
            imagefilename = hb
            ' Stretches the image to fit the pictureBox. 
            pbimage.SizeMode = PictureBoxSizeMode.StretchImage
            MyImage = New Bitmap(imagefilename)
            'pbimage.ClientSize = New Size(xSize, ySize)
            pbimage.Image = CType(MyImage, Image)

            'Me.pbimage.Image.FromFile(imagefilename)
            ' pbimage.Refresh()
        Catch sa As Exception

        End Try
    End Sub
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim isvalid As Boolean = False
        Dim connect As New ADODB.Connection
        Try
            If Me.txtidno.Text.Trim.Length = 0 Or _
            Me.txtname.Text.Trim.Length = 0 Or _
              Me.txthourlyrate.Text.Trim.Length = 0 Then
                MessageBox.Show("Please supply a  valid name, identification number and hourly rate", _
                "Add new personnel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim sdate As String 'date of employment
            sdate = Me.dtpdoe.Value.Year & "-" _
             & dtpdoe.Value.Month & "-" _
             & dtpdoe.Value.Day & " " _
             & dtpdoe.Value.Hour & ":" _
             & dtpdoe.Value.Minute & ":" _
             & dtpdoe.Value.Second
            Dim sdate1 As String 'date of termination
            sdate1 = Me.dtpdot.Value.Year & "-" _
             & dtpdot.Value.Month & "-" _
             & dtpdot.Value.Day & " " _
             & dtpdot.Value.Hour & ":" _
             & dtpdot.Value.Minute & ":" _
             & dtpdot.Value.Second
            Try
                imagefilename.Replace("\", "\\")
            Catch cv As Exception
            End Try
            Dim strin As String = Me.txtcomments.Text.Trim()
            strin = strin.Replace("'", "\'")
            Me.txtcomments.Text = strin
            '-------------------
            Dim arr() As String
            Dim strr, strr2 As String
            Dim y As Integer
            txtcomments.Text = Me.txtcomments.Text.Trim()
            arr = txtcomments.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            '----------------------------------
            '-------------validation
            Dim rs As New ADODB.Recordset
            Dim str = "select id_no from personnel_info where id_no ='" & txtidno.Text.Trim & "';"
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    isvalid = False
                    MessageBox.Show("Personnel with a similar id number already exists", _
                    "Add new personnel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If
            End With

            '----------
            Dim strsql As String = "insert into personnel_info"
            strsql += "(namme,id_no,hourly_rate,gender,phone_no,mobile_no,postal_address,email,pin_no, " _
            & " birthday,contract_end,nssf_no,nhif_no,medical_cover," _
            & "dateofemployment,nextofkin,dateoftermination,imagefile,comments) values "
            strsql += " ( '" & txtname.Text & "', '" & txtidno.Text.Trim & "','" & txthourlyrate.Text & "','" & cbogender.Text & "',"
            strsql += " '" & txtphoneno.Text & "', '" & txtmobileno.Text & "','" & txtpostaladdress.Text & "','" & txtemail.Text & "','" & txtpin.Text & "',"
            strsql += "'" & txtbirthday.Text & "', '" & txtcontractend.Text.Trim & "','" & txtnssfno.Text & "','" & txtnhifno.Text & "',"
            strsql += "'" & txtmedicalcover.Text & "', '" & sdate & "','" & txtnextofkin.Text & "','" & sdate1 & "','" & imagefilename & "',"
            strsql += "'" & strr & "');"

            strsql += " insert into seccheck (name,id_no) values"
            strsql += "('" & txtname.Text & "', '" & txtidno.Text.Trim & "');"
            connect.BeginTrans()
            connect.Execute(strsql)
            connect.CommitTrans()
            isvalid = True
            Dim Tasks As New taskclass
            Dim Thread1cv As New System.Threading.Thread( _
                AddressOf Tasks.personnelinvoke)
            Thread1cv.Start() '
          
        Catch ex As Exception

        Finally
            If isvalid = True Then
                txtname.Text = ""
                txtidno.Text = ""
                txthourlyrate.Text = ""
                cbogender.Text = ""
                txtphoneno.Text = ""
                txtmobileno.Text = ""
                txtpostaladdress.Text = ""
                txtemail.Text = ""
                txtpin.Text = ""
                Me.txtbirthday.Text = ""
                Me.txtmedicalcover.Text = ""
                Me.txtnextofkin.Text = ""
                Me.txtnhifno.Text = ""
                Me.txtcontractend.Text = ""
                Me.txtnssfno.Text = ""
                Me.cbogender.SelectedIndex = -1
            End If
        End Try
        Try
            connect.Close()
        Catch es As Exception

        End Try
        Try
            Dim Tasks As New taskclass
            Dim Thread1cv1 As New System.Threading.Thread( _
                AddressOf Tasks.admininvoke)
            Thread1cv1.Start() '
        Catch xc As Exception

        End Try
    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            Me.Dispose(True)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btndepartments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndepartments.Click
        Try
            Dim n As New frmjobdescription
            n.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "validation"
    Private Sub txtmobileno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmobileno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmobileno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtphoneno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtphoneno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtphoneno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtbirthday_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtbirthday, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtbirthday, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtname_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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
    Private Sub txtidno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtidno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtidno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtemail_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtemail, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtemail, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtpin_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtpin, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtpin, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtpostaladdress_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtpostaladdress, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtpostaladdress, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtnextofkin_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtnextofkin, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtnextofkin, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtnssfno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtnssfno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtnssfno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtcontractend_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtcontractend, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtcontractend, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtmedicalcover_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmedicalcover, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmedicalcover, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtnhifno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtnhifno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtnhifno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub

#End Region

#Region "loaddepartments"
    Private Delegate Sub mydelegate()
    Public Sub loaddepart()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str = "select * from jobdescription"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                myForms.npersonnel.cbogender.Items.Clear()
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While rs.EOF = False
                        myForms.npersonnel.cbogender.Items.Add(.Fields("description").Value)
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
            myForms.npersonnel.Invoke(New mydelegate(AddressOf loaddepart))
        Catch ex As Exception

        End Try
    End Sub
#End Region

End Class
