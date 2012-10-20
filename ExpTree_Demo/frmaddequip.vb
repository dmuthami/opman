Imports System
Imports System.Data
Imports System.Threading
Imports ADODB


Public Class frmaddequip
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtequipid As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtmodelname As System.Windows.Forms.TextBox
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtphone As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtmouse As System.Windows.Forms.TextBox
    Friend WithEvents txtkeyboard As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents txtsupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txthourlyrate As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txttype As System.Windows.Forms.TextBox
    Friend WithEvents txtguarantee As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtmodelyear As System.Windows.Forms.TextBox
    Friend WithEvents txtcondition As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtlicense As System.Windows.Forms.TextBox
    Friend WithEvents txtserialno As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtmanufacturer As System.Windows.Forms.TextBox
    Friend WithEvents txtmodelno As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtppurchasedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtamount As AMS.TextBox.NumericTextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtunit As System.Windows.Forms.TextBox
    Friend WithEvents txtbatteries As System.Windows.Forms.TextBox
    Friend WithEvents txtdownloadcables As System.Windows.Forms.TextBox
    Friend WithEvents txtmonitor2 As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtmonitor As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmaddequip))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.txtunit = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.txtbatteries = New System.Windows.Forms.TextBox
        Me.txtdownloadcables = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.txtsupplier = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.txthourlyrate = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txttype = New System.Windows.Forms.TextBox
        Me.txtguarantee = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtmodelyear = New System.Windows.Forms.TextBox
        Me.txtcondition = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtlicense = New System.Windows.Forms.TextBox
        Me.txtserialno = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtmanufacturer = New System.Windows.Forms.TextBox
        Me.txtmodelno = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtppurchasedate = New System.Windows.Forms.DateTimePicker
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtamount = New AMS.TextBox.NumericTextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txtphone = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtmonitor2 = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtmouse = New System.Windows.Forms.TextBox
        Me.txtkeyboard = New System.Windows.Forms.TextBox
        Me.txtmonitor = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.txtmodelname = New System.Windows.Forms.TextBox
        Me.txtequipid = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.txtmodelname)
        Me.GroupBox1.Controls.Add(Me.txtequipid)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(520, 424)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.txtunit)
        Me.GroupBox5.Controls.Add(Me.Label20)
        Me.GroupBox5.Controls.Add(Me.txtbatteries)
        Me.GroupBox5.Controls.Add(Me.txtdownloadcables)
        Me.GroupBox5.Controls.Add(Me.Label22)
        Me.GroupBox5.Controls.Add(Me.Label24)
        Me.GroupBox5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox5.Location = New System.Drawing.Point(272, 243)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(240, 173)
        Me.GroupBox5.TabIndex = 24
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Other accessories"
        '
        'txtunit
        '
        Me.txtunit.Location = New System.Drawing.Point(72, 72)
        Me.txtunit.Name = "txtunit"
        Me.txtunit.Size = New System.Drawing.Size(160, 20)
        Me.txtunit.TabIndex = 27
        Me.txtunit.Text = ""
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(8, 72)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(48, 16)
        Me.Label20.TabIndex = 42
        Me.Label20.Text = "Unit"
        '
        'txtbatteries
        '
        Me.txtbatteries.Location = New System.Drawing.Point(72, 13)
        Me.txtbatteries.Name = "txtbatteries"
        Me.txtbatteries.Size = New System.Drawing.Size(160, 20)
        Me.txtbatteries.TabIndex = 25
        Me.txtbatteries.Text = ""
        '
        'txtdownloadcables
        '
        Me.txtdownloadcables.Location = New System.Drawing.Point(72, 37)
        Me.txtdownloadcables.Name = "txtdownloadcables"
        Me.txtdownloadcables.Size = New System.Drawing.Size(160, 20)
        Me.txtdownloadcables.TabIndex = 26
        Me.txtdownloadcables.Text = ""
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(8, 17)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(56, 11)
        Me.Label22.TabIndex = 36
        Me.Label22.Text = "Batteries"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(8, 40)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(64, 32)
        Me.Label24.TabIndex = 34
        Me.Label24.Text = "Download cables"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtsupplier)
        Me.GroupBox4.Controls.Add(Me.Label14)
        Me.GroupBox4.Controls.Add(Me.Label13)
        Me.GroupBox4.Controls.Add(Me.txthourlyrate)
        Me.GroupBox4.Controls.Add(Me.Label12)
        Me.GroupBox4.Controls.Add(Me.txttype)
        Me.GroupBox4.Controls.Add(Me.txtguarantee)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Controls.Add(Me.txtmodelyear)
        Me.GroupBox4.Controls.Add(Me.txtcondition)
        Me.GroupBox4.Controls.Add(Me.Label9)
        Me.GroupBox4.Controls.Add(Me.Label10)
        Me.GroupBox4.Controls.Add(Me.txtlicense)
        Me.GroupBox4.Controls.Add(Me.txtserialno)
        Me.GroupBox4.Controls.Add(Me.Label7)
        Me.GroupBox4.Controls.Add(Me.Label8)
        Me.GroupBox4.Controls.Add(Me.txtmanufacturer)
        Me.GroupBox4.Controls.Add(Me.txtmodelno)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.dtppurchasedate)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.txtamount)
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox4.Location = New System.Drawing.Point(8, 104)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(264, 312)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "GroupBox4"
        '
        'txtsupplier
        '
        Me.txtsupplier.Location = New System.Drawing.Point(92, 19)
        Me.txtsupplier.Name = "txtsupplier"
        Me.txtsupplier.Size = New System.Drawing.Size(165, 20)
        Me.txtsupplier.TabIndex = 6
        Me.txtsupplier.Text = ""
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(12, 19)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(80, 16)
        Me.Label14.TabIndex = 57
        Me.Label14.Text = "Supplier"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(12, 278)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 16)
        Me.Label13.TabIndex = 55
        Me.Label13.Text = "Amount"
        '
        'txthourlyrate
        '
        Me.txthourlyrate.Location = New System.Drawing.Point(92, 254)
        Me.txthourlyrate.Name = "txthourlyrate"
        Me.txthourlyrate.Size = New System.Drawing.Size(165, 20)
        Me.txthourlyrate.TabIndex = 16
        Me.txthourlyrate.Text = ""
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(12, 254)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 16)
        Me.Label12.TabIndex = 54
        Me.Label12.Text = "Hourly_rate"
        '
        'txttype
        '
        Me.txttype.Location = New System.Drawing.Point(92, 230)
        Me.txttype.Name = "txttype"
        Me.txttype.Size = New System.Drawing.Size(168, 20)
        Me.txttype.TabIndex = 15
        Me.txttype.Text = ""
        '
        'txtguarantee
        '
        Me.txtguarantee.Location = New System.Drawing.Point(92, 183)
        Me.txtguarantee.Name = "txtguarantee"
        Me.txtguarantee.Size = New System.Drawing.Size(165, 20)
        Me.txtguarantee.TabIndex = 13
        Me.txtguarantee.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(12, 183)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(72, 16)
        Me.Label11.TabIndex = 53
        Me.Label11.Text = "Guarantee"
        '
        'txtmodelyear
        '
        Me.txtmodelyear.Location = New System.Drawing.Point(92, 159)
        Me.txtmodelyear.Name = "txtmodelyear"
        Me.txtmodelyear.Size = New System.Drawing.Size(165, 20)
        Me.txtmodelyear.TabIndex = 12
        Me.txtmodelyear.Text = ""
        '
        'txtcondition
        '
        Me.txtcondition.Location = New System.Drawing.Point(92, 135)
        Me.txtcondition.Name = "txtcondition"
        Me.txtcondition.Size = New System.Drawing.Size(165, 20)
        Me.txtcondition.TabIndex = 11
        Me.txtcondition.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(12, 159)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(72, 16)
        Me.Label9.TabIndex = 52
        Me.Label9.Text = "Model year"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(12, 135)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 16)
        Me.Label10.TabIndex = 51
        Me.Label10.Text = "Condition"
        '
        'txtlicense
        '
        Me.txtlicense.Location = New System.Drawing.Point(92, 111)
        Me.txtlicense.Name = "txtlicense"
        Me.txtlicense.Size = New System.Drawing.Size(165, 20)
        Me.txtlicense.TabIndex = 10
        Me.txtlicense.Text = ""
        '
        'txtserialno
        '
        Me.txtserialno.Location = New System.Drawing.Point(92, 87)
        Me.txtserialno.Name = "txtserialno"
        Me.txtserialno.Size = New System.Drawing.Size(165, 20)
        Me.txtserialno.TabIndex = 9
        Me.txtserialno.Text = ""
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(12, 111)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 16)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "Licence"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(12, 87)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(80, 16)
        Me.Label8.TabIndex = 49
        Me.Label8.Text = "Serial No"
        '
        'txtmanufacturer
        '
        Me.txtmanufacturer.Location = New System.Drawing.Point(92, 63)
        Me.txtmanufacturer.Name = "txtmanufacturer"
        Me.txtmanufacturer.Size = New System.Drawing.Size(165, 20)
        Me.txtmanufacturer.TabIndex = 8
        Me.txtmanufacturer.Text = ""
        '
        'txtmodelno
        '
        Me.txtmodelno.Location = New System.Drawing.Point(92, 42)
        Me.txtmodelno.Name = "txtmodelno"
        Me.txtmodelno.Size = New System.Drawing.Size(165, 20)
        Me.txtmodelno.TabIndex = 7
        Me.txtmodelno.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(12, 63)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 47
        Me.Label4.Text = "Manufacturer"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(12, 42)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "Model No"
        '
        'dtppurchasedate
        '
        Me.dtppurchasedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtppurchasedate.Location = New System.Drawing.Point(92, 206)
        Me.dtppurchasedate.Name = "dtppurchasedate"
        Me.dtppurchasedate.Size = New System.Drawing.Size(168, 20)
        Me.dtppurchasedate.TabIndex = 14
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(12, 230)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 35
        Me.Label6.Text = "Type"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 198)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 24)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Acquisition Date"
        '
        'txtamount
        '
        Me.txtamount.AllowNegative = True
        Me.txtamount.DigitsInGroup = 0
        Me.txtamount.Flags = 0
        Me.txtamount.Location = New System.Drawing.Point(92, 278)
        Me.txtamount.MaxDecimalPlaces = 4
        Me.txtamount.MaxWholeDigits = 9
        Me.txtamount.Name = "txtamount"
        Me.txtamount.Prefix = ""
        Me.txtamount.RangeMax = 1.7976931348623157E+308
        Me.txtamount.RangeMin = -1.7976931348623157E+308
        Me.txtamount.Size = New System.Drawing.Size(168, 20)
        Me.txtamount.TabIndex = 17
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtphone)
        Me.GroupBox3.Controls.Add(Me.Label19)
        Me.GroupBox3.Controls.Add(Me.txtmonitor2)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Controls.Add(Me.txtmouse)
        Me.GroupBox3.Controls.Add(Me.txtkeyboard)
        Me.GroupBox3.Controls.Add(Me.txtmonitor)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(272, 104)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(240, 137)
        Me.GroupBox3.TabIndex = 18
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Computer accessories"
        '
        'txtphone
        '
        Me.txtphone.Location = New System.Drawing.Point(72, 108)
        Me.txtphone.Name = "txtphone"
        Me.txtphone.Size = New System.Drawing.Size(160, 20)
        Me.txtphone.TabIndex = 23
        Me.txtphone.Text = ""
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(5, 111)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(59, 16)
        Me.Label19.TabIndex = 42
        Me.Label19.Text = "Phone"
        '
        'txtmonitor2
        '
        Me.txtmonitor2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmonitor2.Location = New System.Drawing.Point(72, 84)
        Me.txtmonitor2.Name = "txtmonitor2"
        Me.txtmonitor2.Size = New System.Drawing.Size(160, 20)
        Me.txtmonitor2.TabIndex = 22
        Me.txtmonitor2.Text = ""
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(6, 86)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(66, 16)
        Me.Label18.TabIndex = 40
        Me.Label18.Text = "Monitor(2)"
        '
        'txtmouse
        '
        Me.txtmouse.Location = New System.Drawing.Point(72, 13)
        Me.txtmouse.Name = "txtmouse"
        Me.txtmouse.Size = New System.Drawing.Size(160, 20)
        Me.txtmouse.TabIndex = 19
        Me.txtmouse.Text = ""
        '
        'txtkeyboard
        '
        Me.txtkeyboard.Location = New System.Drawing.Point(72, 37)
        Me.txtkeyboard.Name = "txtkeyboard"
        Me.txtkeyboard.Size = New System.Drawing.Size(160, 20)
        Me.txtkeyboard.TabIndex = 20
        Me.txtkeyboard.Text = ""
        '
        'txtmonitor
        '
        Me.txtmonitor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtmonitor.Location = New System.Drawing.Point(72, 61)
        Me.txtmonitor.Name = "txtmonitor"
        Me.txtmonitor.Size = New System.Drawing.Size(160, 20)
        Me.txtmonitor.TabIndex = 21
        Me.txtmonitor.Text = ""
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 17)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(48, 11)
        Me.Label15.TabIndex = 36
        Me.Label15.Text = "Mouse"
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(7, 64)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(57, 16)
        Me.Label16.TabIndex = 35
        Me.Label16.Text = "Monitor"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(8, 40)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 16)
        Me.Label17.TabIndex = 34
        Me.Label17.Text = "Keyboard"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtdesc)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(8, 40)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(504, 64)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Description"
        '
        'txtdesc
        '
        Me.txtdesc.Location = New System.Drawing.Point(8, 16)
        Me.txtdesc.Multiline = True
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(488, 40)
        Me.txtdesc.TabIndex = 4
        Me.txtdesc.Text = ""
        '
        'txtmodelname
        '
        Me.txtmodelname.Location = New System.Drawing.Point(342, 16)
        Me.txtmodelname.Name = "txtmodelname"
        Me.txtmodelname.Size = New System.Drawing.Size(165, 20)
        Me.txtmodelname.TabIndex = 2
        Me.txtmodelname.Text = ""
        '
        'txtequipid
        '
        Me.txtequipid.Location = New System.Drawing.Point(98, 16)
        Me.txtequipid.Name = "txtequipid"
        Me.txtequipid.Size = New System.Drawing.Size(165, 20)
        Me.txtequipid.TabIndex = 1
        Me.txtequipid.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(270, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Model Name"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Equipment Id"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSave.Location = New System.Drawing.Point(8, 429)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(120, 20)
        Me.btnSave.TabIndex = 28
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Location = New System.Drawing.Point(408, 429)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(120, 20)
        Me.btnClose.TabIndex = 29
        Me.btnClose.Text = "Close"
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmaddequip
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(530, 452)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmaddequip"
        Me.Text = "Add Equipment"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
#Region " add equip"
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Dispose(True)
        Catch we As Exception

        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xcv As Exception

        End Try
        Dim isvalid As Boolean = True
        Try
            If Me.txtequipid.Text.Length < 1 Then
                MessageBox.Show("Please supply an id for the equipment", "Save", _
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
                isvalid = False
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
            '-------------check if id number already exists
            Dim rs As New ADODB.Recordset
            Dim Str As String = " select equip_id  from equip_info" _
            & " where lower(equip_id) like '%" & txtequipid.Text.Trim.ToLower() & "%'"

            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(Str, connect)
                If .BOF = False And .EOF = False Then
                    MessageBox.Show(" Equipment id  already exists", "Save", _
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    isvalid = False
                    Exit Try
                End If
                .Close()
            End With
            '--------------end of sanity check
            Dim pdate As String
            pdate = dtppurchasedate.Value.Year & "-" _
                       & dtppurchasedate.Value.Month & "-" _
                       & dtppurchasedate.Value.Day & " " _
                       & dtppurchasedate.Value.Hour & ":" _
                       & dtppurchasedate.Value.Minute & ":" _
                       & dtppurchasedate.Value.Second
            Dim strsql As String
            strsql = "insert into equip_info (equip_id,manufacturer,model_no,serial_no,model_name,purchase_date, " _
            & "description,license,guarantee,condition,type,model_year,amount,supplier,mouse,keyboard,monitor,monitor2,phone,hourly_rate " _
             & " ,batteries,downloadcables,unit) "
            strsql += "values ('" & Me.txtequipid.Text.Trim & "','" & Me.txtmanufacturer.Text.Trim & "','" & Me.txtmodelno.Text.Trim & "',"
            strsql += "'" & Me.txtserialno.Text.Trim & "','" & Me.txtmodelname.Text.Trim & "','" & pdate & "','" & Me.txtsupplier.Text.Trim & "'," _
            & "'" & Me.txtlicense.Text.Trim & "','" & Me.txtguarantee.Text.Trim & "','" & Me.txtcondition.Text.Trim & "'," _
            & "'" & Me.txttype.Text.Trim & "','" & Me.txtmodelyear.Text.Trim & "','" & Me.txtamount.Text.Trim() & "','" & Me.txtdesc.Text.Trim() & "'," _
            & "'" & Me.txtmouse.Text.Trim & "','" & Me.txtkeyboard.Text.Trim & "','" & Me.txtmonitor.Text.Trim() & "'," _
             & "'" & Me.txtmonitor2.Text.Trim & "','" & Me.txtphone.Text.Trim & "','" & Me.txthourlyrate.Text.Trim & "'," _
            & "'" & Me.txtbatteries.Text.Trim & "','" & Me.txtdownloadcables.Text.Trim & "','" & Me.txtunit.Text.Trim & "');"

            strsql += "insert into assigned_info (equip_id,status) values ('" & Me.txtequipid.Text.Trim & "','" & "0" & "');"
            strsql += "insert into equip_finances (equip_id,hourly_rate) values ('" & Me.txtequipid.Text.Trim & "','" & Me.txthourlyrate.Text.Trim & "');"
            strsql += "insert into current_equip (equip_id) values ('" & Me.txtequipid.Text.Trim & "');"
            connect.BeginTrans()
            connect.IsolationLevel = IsolationLevelEnum.adXactSerializable
            connect.Execute(strsql)
            connect.CommitTrans()

            MessageBox.Show("Addition successful", "Adding equipments", _
            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Try
                connect.Close()
            Catch qw As Exception
            End Try
        Catch ex As Exception
        Finally
            If isvalid = True Then
                txtequipid.Text = ""
                txtmanufacturer.Text = ""
                txtmodelno.Text = ""
                txtserialno.Text = ""
                txtmodelname.Text = ""
                txtdesc.Text = ""
                txtlicense.Text = ""
                txtguarantee.Text = ""
                txtcondition.Text = ""
                txttype.Text = ""
                txtmodelyear.Text = ""
                txthourlyrate.Text = ""

                Me.txtmouse.Text = ""
                Me.txtamount.Text = ""
                Me.txtmonitor.Text = ""
                Me.txtmonitor2.Text = ""
                Me.txtphone.Text = ""
                Me.txtkeyboard.Text = ""
                Me.txtsupplier.Text = ""
                Me.txtunit.Text = "'"
                Me.txtbatteries.Text = ""
                Me.txtdownloadcables.Text = ""
            End If

        End Try
        Try
            Dim Tasks As New taskclass
            Dim Threade2 As New System.Threading.Thread( _
                AddressOf taskclass.equipinvoke)
            Threade2.Start()
        Catch xzz As Exception

        End Try
    End Sub
    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
    Private Sub frmaddequip_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

#Region " text box validation"
    Private Sub txtmodelname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmodelname.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmodelname, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmodelname, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtmanufacturer_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmanufacturer, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmanufacturer, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtmodelno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmodelno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmodelno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtlicense_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtlicense, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtlicense, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtserialno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtserialno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtserialno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtmodelyear_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmodelyear, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmodelyear, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtcondition_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtcondition, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtcondition, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtguarantee_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtguarantee, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtguarantee, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txttype_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txttype, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txttype, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txthourlyrate_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txthourlyrate, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txthourlyrate, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtmonitor_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmonitor, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmonitor, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtkeyboard_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtkeyboard, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtkeyboard, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtmouse_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtmouse, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtmouse, "")
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
#End Region

#End Region


End Class
