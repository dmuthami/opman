Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Common

Imports System.IO
Imports Microsoft.VisualBasic
Imports ADODB

Imports System.ArgumentNullException
Imports System.NullReferenceException
Imports System.ArgumentOutOfRangeException


Public Class frmMe
    Inherits System.Windows.Forms.Form
    Public mycaller As frmAddLead
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
    Dim hti1 As System.Windows.Forms.DataGrid.HitTestInfo
    Dim htileads As System.Windows.Forms.DataGrid.HitTestInfo
    Public Delegate Sub mydelegate()
    Public Delegate Sub mydelegate1()

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
            jobsform = False
            myForms.CustomerForm3 = Nothing
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
    Friend WithEvents lblClientName As System.Windows.Forms.Label
    Friend WithEvents lblClientNo As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAddNewJob As System.Windows.Forms.Button
    Public WithEvents dtgJobs As System.Windows.Forms.DataGrid
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCCD As System.Windows.Forms.Button
    Friend WithEvents btnAddContacts As System.Windows.Forms.Button
    Public WithEvents dtgContacts As System.Windows.Forms.DataGrid
    Friend WithEvents grpLeads As System.Windows.Forms.GroupBox
    Friend WithEvents dtgLeads As System.Windows.Forms.DataGrid
    Friend WithEvents btnAddLeads As System.Windows.Forms.Button
    Friend WithEvents btndeletecontact As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMe))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblClientName = New System.Windows.Forms.Label
        Me.lblClientNo = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.btnAddNewJob = New System.Windows.Forms.Button
        Me.dtgJobs = New System.Windows.Forms.DataGrid
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btndeletecontact = New System.Windows.Forms.Button
        Me.btnCCD = New System.Windows.Forms.Button
        Me.btnAddContacts = New System.Windows.Forms.Button
        Me.dtgContacts = New System.Windows.Forms.DataGrid
        Me.grpLeads = New System.Windows.Forms.GroupBox
        Me.btnAddLeads = New System.Windows.Forms.Button
        Me.dtgLeads = New System.Windows.Forms.DataGrid
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dtgJobs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dtgContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpLeads.SuspendLayout()
        CType(Me.dtgLeads, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblClientName)
        Me.GroupBox1.Controls.Add(Me.lblClientNo)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(800, 56)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lblClientName
        '
        Me.lblClientName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblClientName.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lblClientName.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientName.Location = New System.Drawing.Point(216, 16)
        Me.lblClientName.Name = "lblClientName"
        Me.lblClientName.Size = New System.Drawing.Size(568, 32)
        Me.lblClientName.TabIndex = 2
        Me.lblClientName.Text = "Client Name"
        Me.lblClientName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblClientNo
        '
        Me.lblClientNo.BackColor = System.Drawing.Color.FromArgb(CType(206, Byte), CType(237, Byte), CType(247, Byte))
        Me.lblClientNo.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClientNo.Location = New System.Drawing.Point(8, 16)
        Me.lblClientNo.Name = "lblClientNo"
        Me.lblClientNo.Size = New System.Drawing.Size(200, 32)
        Me.lblClientNo.TabIndex = 1
        Me.lblClientNo.Text = "Client No"
        Me.lblClientNo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.btnAddNewJob)
        Me.GroupBox4.Controls.Add(Me.dtgJobs)
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(360, 208)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(440, 320)
        Me.GroupBox4.TabIndex = 11
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Existing Jobs"
        '
        'btnAddNewJob
        '
        Me.btnAddNewJob.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnAddNewJob.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddNewJob.Location = New System.Drawing.Point(8, 14)
        Me.btnAddNewJob.Name = "btnAddNewJob"
        Me.btnAddNewJob.Size = New System.Drawing.Size(120, 20)
        Me.btnAddNewJob.TabIndex = 12
        Me.btnAddNewJob.Text = "Add New Job"
        '
        'dtgJobs
        '
        Me.dtgJobs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgJobs.CaptionText = "Existing jobs"
        Me.dtgJobs.DataMember = ""
        Me.dtgJobs.FlatMode = True
        Me.dtgJobs.GridLineColor = System.Drawing.SystemColors.Window
        Me.dtgJobs.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgJobs.Location = New System.Drawing.Point(8, 37)
        Me.dtgJobs.Name = "dtgJobs"
        Me.dtgJobs.PreferredColumnWidth = 80
        Me.dtgJobs.ReadOnly = True
        Me.dtgJobs.Size = New System.Drawing.Size(424, 276)
        Me.dtgJobs.TabIndex = 13
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.btndeletecontact)
        Me.GroupBox2.Controls.Add(Me.btnCCD)
        Me.GroupBox2.Controls.Add(Me.btnAddContacts)
        Me.GroupBox2.Controls.Add(Me.dtgContacts)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 56)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(792, 152)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Contacts"
        '
        'btndeletecontact
        '
        Me.btndeletecontact.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeletecontact.Location = New System.Drawing.Point(264, 16)
        Me.btndeletecontact.Name = "btndeletecontact"
        Me.btndeletecontact.Size = New System.Drawing.Size(152, 20)
        Me.btndeletecontact.TabIndex = 6
        Me.btndeletecontact.Text = "Delete selected contact"
        '
        'btnCCD
        '
        Me.btnCCD.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnCCD.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCCD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCCD.Location = New System.Drawing.Point(128, 16)
        Me.btnCCD.Name = "btnCCD"
        Me.btnCCD.Size = New System.Drawing.Size(136, 20)
        Me.btnCCD.TabIndex = 5
        Me.btnCCD.Text = "Change Client Details"
        '
        'btnAddContacts
        '
        Me.btnAddContacts.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnAddContacts.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddContacts.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddContacts.Location = New System.Drawing.Point(8, 16)
        Me.btnAddContacts.Name = "btnAddContacts"
        Me.btnAddContacts.Size = New System.Drawing.Size(120, 20)
        Me.btnAddContacts.TabIndex = 4
        Me.btnAddContacts.Text = "Add Contacts"
        '
        'dtgContacts
        '
        Me.dtgContacts.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgContacts.CaptionText = "Contacts"
        Me.dtgContacts.DataMember = ""
        Me.dtgContacts.FlatMode = True
        Me.dtgContacts.GridLineColor = System.Drawing.SystemColors.Window
        Me.dtgContacts.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgContacts.Location = New System.Drawing.Point(8, 37)
        Me.dtgContacts.Name = "dtgContacts"
        Me.dtgContacts.PreferredColumnWidth = 100
        Me.dtgContacts.ReadOnly = True
        Me.dtgContacts.Size = New System.Drawing.Size(776, 107)
        Me.dtgContacts.TabIndex = 7
        '
        'grpLeads
        '
        Me.grpLeads.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpLeads.Controls.Add(Me.btnAddLeads)
        Me.grpLeads.Controls.Add(Me.dtgLeads)
        Me.grpLeads.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpLeads.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpLeads.Location = New System.Drawing.Point(8, 208)
        Me.grpLeads.Name = "grpLeads"
        Me.grpLeads.Size = New System.Drawing.Size(352, 320)
        Me.grpLeads.TabIndex = 8
        Me.grpLeads.TabStop = False
        Me.grpLeads.Text = "Leads"
        '
        'btnAddLeads
        '
        Me.btnAddLeads.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnAddLeads.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnAddLeads.Location = New System.Drawing.Point(11, 14)
        Me.btnAddLeads.Name = "btnAddLeads"
        Me.btnAddLeads.Size = New System.Drawing.Size(120, 20)
        Me.btnAddLeads.TabIndex = 9
        Me.btnAddLeads.Text = "Add New  Lead"
        '
        'dtgLeads
        '
        Me.dtgLeads.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgLeads.CaptionText = "Leads"
        Me.dtgLeads.DataMember = ""
        Me.dtgLeads.FlatMode = True
        Me.dtgLeads.GridLineColor = System.Drawing.SystemColors.Window
        Me.dtgLeads.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgLeads.Location = New System.Drawing.Point(8, 36)
        Me.dtgLeads.Name = "dtgLeads"
        Me.dtgLeads.ReadOnly = True
        Me.dtgLeads.Size = New System.Drawing.Size(336, 276)
        Me.dtgLeads.TabIndex = 10
        '
        'frmMe
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(800, 530)
        Me.Controls.Add(Me.grpLeads)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMe"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Client Details"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.dtgJobs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dtgContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpLeads.ResumeLayout(False)
        CType(Me.dtgLeads, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
              Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
#Region "former code"
    Private Function returnhittest1(ByVal hit As System.Windows.Forms.DataGrid.HitTestType) As String
        Try
            Select Case hit
                Case System.Windows.Forms.DataGrid.HitTestType.Cell

                    Dim mycell As New DataGridCell
                    mycell.RowNumber = Me.dtgContacts.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittest1 = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell
                    mycell.RowNumber = Me.dtgContacts.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittest1 = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case Else
                    returnhittest1 = ""
            End Select
        Catch ex As Exception
            MessageBox.Show(Text:="Error:" & ex.Message.ToString, _
            caption:="Throwing an exception", _
            Icon:=MessageBoxIcon.Information, _
            buttons:=MessageBoxButtons.OK)

        End Try
    End Function
    Private Sub dtgContacts_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgContacts.MouseDown
        Try
            hti1 = Me.dtgContacts.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
            MsgBox(ex.Message.ToString())
        Finally

        End Try
    End Sub
    Private Sub dtgContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgContacts.DoubleClick
        Try
            Dim results As String

            results = returnhittest1(hti1.Type)
            If results <> "" Then
                Dim ds As New System.Data.DataSet
                ds = dtgContacts.DataSource
                Dim mycell As New DataGridCell
                Dim a() As String
                a = results.Split("|")
                mycell.RowNumber = CInt(a(0))
                mycell.ColumnNumber = CInt(a(1))
                Dim desc, fname, sname, salu, pobox, e_mail1, e_mail2, fax, tel As String
                Dim cell, phyadd, mobile2 As String
                Try
                    desc = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 1
                    fname = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 2
                    sname = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 3
                    salu = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 4
                    pobox = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 5
                    e_mail1 = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 6
                    If Me.dtgContacts(mycell) = Nothing Then
                        e_mail2 = ""
                    End If
                    e_mail2 = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 7
                    fax = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 8
                    tel = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 9
                    cell = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 10
                    mobile2 = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    mycell.ColumnNumber = 11
                    phyadd = Me.dtgContacts(mycell)
                Catch cv As Exception
                End Try
                Try
                    myForms.CustomerForm1.txtdesc.Text = desc
                    myForms.CustomerForm1.txtFirstName.Text = fname
                    myForms.CustomerForm1.txtSecondName.Text = sname
                    myForms.CustomerForm1.cboSalutation.Text = salu
                    myForms.CustomerForm1.txtEMail1.Text = e_mail1
                    myForms.CustomerForm1.txtEmail2.Text = e_mail2
                    myForms.CustomerForm1.txtFax.Text = fax
                    myForms.CustomerForm1.txtTelephone.Text = tel
                    myForms.CustomerForm1.txtCellPhone.Text = cell
                    myForms.CustomerForm1.txtPhysicalAddress.Text = phyadd
                    myForms.CustomerForm1.txtPostalAddress.Text = pobox
                    myForms.CustomerForm1.txtcellphone2.Text = mobile2
                    Try
                        myForms.CustomerForm1.autono = ds.Tables(0).Rows(hti1.Row).Item("ano")
                    Catch cv As Exception
                    End Try
                Catch xcv As Exception
                    Try
                        Dim myform As New frmEditcontact
                        myform.desc = desc
                        myform.fname = fname
                        myform.sname = sname
                        myform.salu = salu
                        myform.e_mail1 = e_mail1
                        myform.e_mail2 = e_mail2
                        myform.fax = fax
                        myform.tel = tel
                        myform.cell = cell
                        myform.phyadd = phyadd
                        myform.pobox = pobox
                        myform.mobile2 = mobile2
                        myForms.CustomerForm1 = myform
                        Try
                            myForms.CustomerForm1.autono = ds.Tables(0).Rows(hti1.Row).Item("ano")
                        Catch cv As Exception
                        End Try
                        myForms.CustomerForm1.Show()
                    Catch bnm As Exception
                    End Try
                End Try
            End If
        Catch r As Exception

        End Try
    End Sub
    Private Sub frmJobs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ''''''''''''''''''''------------------

            'Me.GroupBox2.Width = Me.Width - 20
            'Me.GroupBox1.Width = Me.Width - 20

            'Me.dtgContacts.Width = Me.GroupBox2.Width - 20
            'Me.GroupBox4.Width = (Me.Width - Me.grpLeads.Width) - 20
            'Me.dtgJobs.Width = Me.GroupBox4.Width - 20
            'Dim h As Integer
            'h = Me.GroupBox1.Height + Me.GroupBox2.Height + 100
            'Me.GroupBox4.Height = Me.Height - h
            'Me.dtgJobs.Height = Me.GroupBox4.Height - 40 - Me.btnAddNewJob.Height

            'loadleads() 'load leads
            'grpLeads.Height = GroupBox4.Height
            'Me.dtgLeads.Height = Me.grpLeads.Height - 40 - Me.btnAddLeads.Height
            '-------------------------------------------------------------
            Me.lblClientName.Text = myclientname
            Me.lblClientNo.Text = myclientno
            Me.dtgContacts.Invoke(New mydelegate(AddressOf loadgridcontact))
            Me.dtgJobs.Invoke(New mydelegate1(AddressOf loadgridexistingjobs))
            loadleads() 'load leads
            addjobs = False
            addcontacts = False
            editcontacts = False
            editclients = False
            editjob = False
            jobsform = True
        Catch ex As Exception
            MsgBox(ex.Message.ToString(), MsgBoxStyle.Information)
        End Try
    End Sub
    Public Sub loadgridexistingjobs()
        Dim currentcursor As Cursor = Cursor.Current
        Try

            '-----------------try this dave
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter
            Dim custDS As DataSet = New DataSet
            Dim adors As New ADODB.Recordset
            Dim str As String = "select * from rcljobs" _
            & " where client_no='" & myclientno & "'"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "rcljobs")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgJobs.SetDataBinding(custDS, tname)
            Try
                connect.Close()
            Catch xc As Exception

            End Try

            '--------------------this is quite cool---------------------------------------------------------------------

            Call AddCustomDataTableStylejobs()
        Catch ex As Exception

        Finally
            Cursor.Current = currentcursor
        End Try
    End Sub
    Public Sub AddCustomDataTableStylecont()
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = "contact"
            Dim mywidth As Integer
            mywidth = Me.dtgContacts.Width
            mywidth = mywidth / 12
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "description"
            myno.HeaderText = "Description"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "f_name"
            myname.HeaderText = "First Name"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "s_name"
            mydesc.HeaderText = "Second Name"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)


            ' Add a second column style.
            Dim mydesc1 As New DataGridTextBoxColumn
            mydesc1.MappingName = "salutation"
            mydesc1.HeaderText = "Salutation"
            mydesc1.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc1)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn
            mydesc2.MappingName = "pobox"
            mydesc2.HeaderText = "P.O Box"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)


            ' Add a second column style.
            Dim mydesc3 As New DataGridTextBoxColumn
            mydesc3.MappingName = "e_mail1"
            mydesc3.HeaderText = "E-Mail Address(1)"
            mydesc3.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc3)


            ' Add a second column style.
            Dim mydesc4 As New DataGridTextBoxColumn
            mydesc4.MappingName = "e_mail2"
            mydesc4.HeaderText = "E-Mail Address(2)"
            mydesc4.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc4)


            ' Add a second column style.
            Dim mydesc5 As New DataGridTextBoxColumn
            mydesc5.MappingName = "fax"
            mydesc5.HeaderText = "Fax"
            mydesc5.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc5)


            ' Add a second column style.
            Dim mydesc6 As New DataGridTextBoxColumn
            mydesc6.MappingName = "tel"
            mydesc6.HeaderText = "Telephone"
            mydesc6.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc6)


            ' Add a second column style.
            Dim mydesc7 As New DataGridTextBoxColumn
            mydesc7.MappingName = "cell"
            mydesc7.HeaderText = "Mobile No(1)"
            mydesc7.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc7)

            ' Add a second column style.
            Dim mydesc71 As New DataGridTextBoxColumn
            mydesc71.MappingName = "mobile2"
            mydesc71.HeaderText = "Mobile No(2)"
            mydesc71.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc71)

            ' Add a second column style.
            Dim mydesc8 As New DataGridTextBoxColumn
            mydesc8.MappingName = "physicaladd"
            mydesc8.HeaderText = "Physical Address"
            mydesc8.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc8)
            ' Add the DataGridTableStyle objects to the collection.
            dtgContacts.TableStyles.Clear()
            dtgContacts.TableStyles.Add(ts1)
        Catch ex As Exception

        End Try

    End Sub 'AddCustomDataTableStyle
    Public Sub loadgridcontact()
        'Dim currentcursor As Cursor = Cursor.Current
        Try
            '-----------------try this dave
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter
            Dim custDS As DataSet = New DataSet
            Dim adors As New ADODB.Recordset
            Dim str As String = "select * from contact where client_no='" & myclientno & "'"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "contact")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgContacts.SetDataBinding(custDS, tname)
            Try
                connect.Close()
            Catch er As Exception

            End Try

            '--------------------this is quite cool---------------------------------------------------------------------


            Call AddCustomDataTableStylecont()

        Catch t As Exception
            'MessageBox.Show("error" & t.InnerException.ToString, "Error, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            'Cursor.Current = currentcursor
        End Try
    End Sub
    Public Sub AddCustomDataTableStylejobs()
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = "rcljobs"
            Dim mywidth As Integer
            mywidth = Me.dtgContacts.Width
            mywidth = mywidth / 4
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "job_no"
            myno.HeaderText = "Job Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "job_tittle"
            myname.HeaderText = "Job Title"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "job_status"
            mydesc.HeaderText = "Job Status"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)


            '' Add a second column style.
            'Dim mydesc1 As New DataGridTextBoxColumn()
            'mydesc1.MappingName = "dept"
            'mydesc1.HeaderText = "Department"
            'mydesc1.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc1)


            '' Add a second column style.
            'Dim mydesc2 As New DataGridTextBoxColumn()
            'mydesc2.MappingName = "cont"
            'mydesc2.HeaderText = "Contact"
            'mydesc2.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc2)


            ' Add a second column style.
            'Dim mydesc3 As New DataGridTextBoxColumn()
            'mydesc3.MappingName = "sdate"
            'mydesc3.HeaderText = "Start Date"
            'mydesc3.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc3)


            ' Add a second column style.
            'Dim mydesc4 As New DataGridTextBoxColumn()
            'mydesc4.MappingName = "fdate"
            'mydesc4.HeaderText = "End Date"
            'mydesc4.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc4)


            ' Add a second column style.
            'Dim mydesc5 As New DataGridTextBoxColumn()
            'mydesc5.MappingName = "deadline"
            'mydesc5.HeaderText = "Deadline"
            'mydesc5.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc5)


            '' Add a second column style.
            'Dim mydesc6 As New DataGridTextBoxColumn()
            'mydesc6.MappingName = "techres"
            'mydesc6.HeaderText = "Technician Responsible"
            'mydesc6.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc6)


            '' Add a second column style.
            'Dim mydesc7 As New DataGridTextBoxColumn()
            'mydesc7.MappingName = "ojob_no"
            'mydesc7.HeaderText = "Old Job Number"
            'mydesc7.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc7)

            ' Add a second column style.
            Dim mydesc7 As New DataGridTextBoxColumn
            mydesc7.MappingName = "amount"
            mydesc7.HeaderText = "Amount"
            mydesc7.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc7)



            ' Add the DataGridTableStyle objects to the collection.
            dtgJobs.TableStyles.Clear()
            dtgJobs.TableStyles.Add(ts1)
        Catch ex As Exception

        End Try

    End Sub
    'Protected Overrides Function ProcessDialogKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean
    '    Try
    '        ' Note: keyData is a bitmask, so the following IF statement is evaluated to true
    '        ' only if just the ENTER key is pressed, without any modifiers (Shift, Alt, Ctrl...).
    '        If keyData = System.Windows.Forms Then
    '            'RaiseEvent btnLogin_Click(Me, EventArgs.Empty)
    '            Dim E As System.EventArgs
    '            Me.GroupBox2.Width = Me.Width - 50
    '            Me.dtgContacts.Width = Me.GroupBox2.Width - 40
    '            Me.GroupBox4.Width = Me.Width - 50
    '            Me.dtgJobs.Width = Me.GroupBox4.Width - 40
    '            Me.GroupBox4.Height = Me.Height - (Me.GroupBox1.Height + Me.GroupBox1.Height + 70)
    '            Me.dtgJobs.Height = Me.GroupBox4.Height - 70
    '            'Call txtClientName_Click(Me, E)

    '            Return True ' True means we've processed the key
    '        Else
    '            Return MyBase.ProcessDialogKey(keyData)
    '        End If
    '    Catch ex As Exception
    '        'Trace.WriteLine(ex.ToString())
    '        MsgBox(ex.Message.ToString, , Title:="Return key")

    '    End Try
    'End Function
    Private Sub dtgJobs_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtgJobs.Navigate

    End Sub
    Private Sub dtgJobs_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgJobs.DoubleClick
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
            Dim x As Boolean = canviewjobs()
            If x = False Then
                MessageBox.Show("Not allowed to view jobs", "Clients", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xc As Exception
        End Try
        Dim currentcursor As Cursor = Cursor.Current
        Try

            myForms.CustomerForm2.Close()
            myForms.CustomerForm2 = Nothing

        Catch zx As Exception

        End Try
        Try
            myclientno = Me.lblClientNo.Text.Trim
        Catch shj As Exception

        End Try
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim results As String
            'Dim int As Integer = dtgJobs.CurrentRowIndex()

            results = returnhittest(hti.Type)
            If results <> "" Then
                Dim mycell As New DataGridCell
                Dim a() As String
                a = results.Split("|")
                mycell.RowNumber = CInt(a(0))
                mycell.ColumnNumber = CInt(a(1))
                '----------
                Dim ds As System.Data.DataSet = New System.Data.DataSet
                ds = Me.dtgJobs.DataSource
                'sdate()
                'fdate()
                'deadline()
                'costagreed()
                'descrip()
                'date_sniffed()
                'amount()
                'journal()
                'department()
                '------------
                Dim jobno As String
                Dim ojobno As String
                Dim jobtitle As String
                Dim contname, ddate As String
                Dim jobstatus, amount As String
                Dim tecres, desc, department As String
                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("job_no")) = False Then
                    jobno = ds.Tables(0).Rows(mycell.RowNumber).Item("job_no")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("job_tittle")) = False Then
                    jobtitle = ds.Tables(0).Rows(mycell.RowNumber).Item("job_tittle")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("cont")) = False Then
                    contname = ds.Tables(0).Rows(mycell.RowNumber).Item("cont")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("job_status")) = False Then
                    jobstatus = ds.Tables(0).Rows(mycell.RowNumber).Item("job_status")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("techres")) = False Then
                    tecres = ds.Tables(0).Rows(mycell.RowNumber).Item("techres")
                End If
                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("ojob_no")) = False Then
                    ojobno = ds.Tables(0).Rows(mycell.RowNumber).Item("ojob_no")
                    ojobno = "Old job no:" & " " & ojobno
                Else
                    ojobno = "Old job no:" & " " & ojobno
                End If
                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("department")) = False Then
                    department = ds.Tables(0).Rows(mycell.RowNumber).Item("department")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("descrip")) = False Then
                    desc = ds.Tables(0).Rows(mycell.RowNumber).Item("descrip")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("amount")) = False Then
                    amount = ds.Tables(0).Rows(mycell.RowNumber).Item("amount")
                End If
                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("sdate")) = False Then
                    ddate = ds.Tables(0).Rows(mycell.RowNumber).Item("sdate")
                End If
                If editjob = False Then

                    Dim form As New frmEditJob
                    form.jobtitle = jobtitle
                    form.myjobno = jobno
                    form.contname = contname
                    form.tecres = tecres
                    form.jobstatus = jobstatus
                    form.description = desc
                    myForms.CustomerForm2 = form

                    'form.Show()
                    myForms.CustomerForm2.txtJobTitle.Text = jobtitle
                    myForms.CustomerForm2.txtJobNo.Text = jobno
                    myForms.CustomerForm2.txtContactName.Text = contname
                    myForms.CustomerForm2.cboTechnicianresponsible.Text = tecres
                    myForms.CustomerForm2.cboJobstatus.Text = jobstatus
                    myForms.CustomerForm2.txtdesc.Text = desc
                    myForms.CustomerForm2.txtamount.Text = amount
                    myForms.CustomerForm2.lblojobno.Text = ojobno
                    myForms.CustomerForm2.txtdepartment.Text = department
                    Try
                        myclientno = Me.lblClientNo.Text.Trim
                    Catch shj As Exception

                    End Try
                    Try
                        myForms.CustomerForm2.txtbudget.Text = ds.Tables(0).Rows(mycell.RowNumber).Item("budgetarycost")
                    Catch assf As Exception
                    End Try
                    Try
                        myForms.CustomerForm2.lblClientNo.Text = Me.lblClientNo.Text
                    Catch assf As Exception
                    End Try
                    Try
                        myForms.CustomerForm2.lbClientName.Text = Me.lblClientName.Text
                    Catch assf As Exception
                    End Try
                    myForms.CustomerForm2.Text = "Edit jobs"
                    myForms.CustomerForm2.Show()
                    editjob = True
                Else
                    myForms.CustomerForm2.jobtitle = jobtitle
                    myForms.CustomerForm2.myjobno = jobno
                    myForms.CustomerForm2.contname = contname
                    myForms.CustomerForm2.tecres = tecres
                    myForms.CustomerForm2.jobstatus = jobstatus
                    myForms.CustomerForm2.description = desc

                    myForms.CustomerForm2.txtJobTitle.Text = jobtitle
                    myForms.CustomerForm2.txtJobNo.Text = jobno
                    myForms.CustomerForm2.txtContactName.Text = contname
                    myForms.CustomerForm2.cboTechnicianresponsible.Text = tecres
                    myForms.CustomerForm2.cboJobstatus.Text = jobstatus
                    myForms.CustomerForm2.txtdesc.Text = desc
                    myForms.CustomerForm2.txtamount.Text = amount
                    myForms.CustomerForm2.lblojobno.Text = ojobno
                    myForms.CustomerForm2.txtdepartment.Text = department
                    Try
                        myclientno = Me.lblClientNo.Text.Trim
                    Catch shj As Exception

                    End Try
                    Try
                        myForms.CustomerForm2.txtbudget.Text = ds.Tables(0).Rows(mycell.RowNumber).Item("budgetarycost")
                    Catch assf As Exception
                    End Try
                    Try
                        myForms.CustomerForm2.lblClientNo.Text = Me.lblClientNo.Text
                    Catch assf As Exception
                    End Try
                    Try
                        myForms.CustomerForm2.lbClientName.Text = Me.lblClientName.Text
                    Catch assf As Exception
                    End Try
                    editjob = True
                End If
            End If
            Try
                connect.Close()
            Catch ex500 As Exception
            End Try
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Cursor.Current = currentcursor
        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Function returnhittest(ByVal hit As System.Windows.Forms.DataGrid.HitTestType) As String
        Try
            Select Case hit
                Case System.Windows.Forms.DataGrid.HitTestType.Cell

                    Dim mycell As New DataGridCell
                    mycell.RowNumber = Me.dtgJobs.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittest = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell
                    mycell.RowNumber = Me.dtgJobs.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittest = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case Else
                    returnhittest = ""
            End Select
        Catch ex As Exception
            MessageBox.Show(Text:="Error:" & ex.Message.ToString, _
            caption:="Throwing an exception", _
            Icon:=MessageBoxIcon.Information, _
            buttons:=MessageBoxButtons.OK)

        End Try
    End Function
    Protected Overrides Sub Finalize()
        'jobsform = False
        MyBase.Finalize()
    End Sub
    Private Sub btnAddNewJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNewJob.Click
        Try
            Dim x As Boolean = canviewjobs()
            If x = False Then
                MessageBox.Show("Not allowed to view jobs", "Clients", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
            myForms.CustomerForm2.Close()
            myForms.CustomerForm2 = Nothing

        Catch zx As Exception

        End Try
        Try
            myclientno = Me.lblClientNo.Text.Trim
        Catch shj As Exception

        End Try
        Try
            myclientno = Me.lblClientNo.Text
            Dim myform As New frmEditJob

            addjobs = True
            myForms.CustomerForm2 = myform
            myForms.CustomerForm2.btnEditJobs.Text = "Add jobs"
            myForms.CustomerForm2.Text = "Add jobs"
            Try
                myForms.CustomerForm2.lblClientNo.Text = Me.lblClientNo.Text
            Catch assf As Exception
            End Try
            Try
                myForms.CustomerForm2.lbClientName.Text = Me.lblClientName.Text
            Catch assf As Exception
            End Try
            myForms.CustomerForm2.Show()
            If addjobs = False Then

            End If

            'Me.Dispose(False)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub btnAddContacts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddContacts.Click
        Try
            If addcontacts = False Then
                'Call GetCellValue(Me.dtgContacts)
                myclientno = lblClientNo.Text
                Dim myform As New frmAddContacts
                addcontacts = True
                myform.Show()
            End If

            'Me.Dispose(False)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnCCD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCCD.Click
        Try
            If editclients = False Then
                Dim myform As New frmEditClients
                myform.txtClientNo.Text = Me.lblClientNo.Text
                myform.mynumber = Me.lblClientNo.Text
                myform.txtClientName.Text = Me.lblClientName.Text
                editclients = True
                myform.Show()
            End If

            'Me.Dispose(False)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtgJobs_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgJobs.MouseDown
        Try
            hti = Me.dtgJobs.HitTest(New Point(e.X, e.Y))
        Catch ex As Exception
            Try

            Catch er As Exception

            End Try
        Finally

        End Try
    End Sub
    Private Sub dtgContacts_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtgContacts.Navigate

    End Sub
    'Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
    '    Try
    '        'Me.GroupBox2.Width = Me.Width - 20
    '        'Me.GroupBox1.Width = Me.Width - 20

    '        'Me.dtgContacts.Width = Me.GroupBox2.Width - 20
    '        'Me.GroupBox4.Width = (Me.Width - Me.grpLeads.Width) - 20
    '        'Me.dtgJobs.Width = Me.GroupBox4.Width - 20
    '        'Dim h As Integer
    '        'h = Me.GroupBox1.Height + Me.GroupBox2.Height + 100
    '        'Me.GroupBox4.Height = Me.Height - h
    '        'Me.dtgJobs.Height = Me.GroupBox4.Height - 40 - Me.btnAddNewJob.Height
    '        'grpLeads.Height = GroupBox4.Height
    '        'Me.dtgLeads.Height = Me.grpLeads.Height - 40 - Me.btnAddLeads.Height
    '        ''Call loadgridcontact()
    '        ''Call loadgridexistingjobs()

    '        'Me.dtgContacts.Invoke(New mydelegate(AddressOf AddCustomDataTableStylecont))
    '        'Me.dtgJobs.Invoke(New mydelegate(AddressOf AddCustomDataTableStylejobs))

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Private Sub initializecontacts()
    '    Try
    '        CType(Me.dtgContacts, System.ComponentModel.ISupportInitialize).BeginInit()
    '        Me.dtgContacts.DataMember = ""
    '        Me.dtgContacts.HeaderForeColor = System.Drawing.SystemColors.ControlText
    '        Me.dtgContacts.Location = New System.Drawing.Point(8, 38)
    '        Me.dtgContacts.Name = "dtgContacts"
    '        Me.dtgContacts.PreferredColumnWidth = 100
    '        Me.dtgContacts.ReadOnly = True
    '        Me.dtgContacts.Size = New System.Drawing.Size(760, 112)
    '        Me.dtgContacts.TabIndex = 0
    '        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.dtgContacts})
    '        CType(Me.dtgContacts, System.ComponentModel.ISupportInitialize).EndInit()
    '        Me.GroupBox2.Width = Me.Width - 20
    '        Me.dtgContacts.Width = Me.GroupBox2.Width - 20
    '    Catch er As Exception

    '    End Try
    'End Sub
    'Private Sub initializejobs()
    '    Try
    '        CType(Me.dtgJobs, System.ComponentModel.ISupportInitialize).BeginInit()
    '        Me.dtgJobs.DataMember = ""
    '        Me.dtgJobs.HeaderForeColor = System.Drawing.SystemColors.ControlText
    '        Me.dtgJobs.Location = New System.Drawing.Point(8, 40)
    '        Me.dtgJobs.Name = "dtgJobs"
    '        Me.dtgJobs.PreferredColumnWidth = 80
    '        Me.dtgJobs.ReadOnly = True
    '        Me.dtgJobs.Size = New System.Drawing.Size(760, 216)
    '        Me.dtgJobs.TabIndex = 0
    '        Me.GroupBox4.Controls.AddRange(New System.Windows.Forms.Control() {Me.dtgJobs})
    '        CType(Me.dtgJobs, System.ComponentModel.ISupportInitialize).EndInit()
    '        'Me.GroupBox2.Width = Me.Width - 20
    '        Me.GroupBox1.Width = Me.Width - 20

    '        'Me.dtgContacts.Width = Me.GroupBox2.Width - 20
    '        Me.GroupBox4.Width = (Me.Width - Me.grpLeads.Width) - 20
    '        Me.dtgJobs.Width = Me.GroupBox4.Width - 20
    '        Dim h As Integer
    '        h = Me.GroupBox1.Height + Me.GroupBox2.Height + 100
    '        Me.GroupBox4.Height = Me.Height - h
    '        Me.dtgJobs.Height = Me.GroupBox4.Height - 40 - Me.btnAddNewJob.Height
    '        grpLeads.Height = GroupBox4.Height
    '        Me.dtgLeads.Height = Me.grpLeads.Height - 40 - Me.btnAddLeads.Height
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Private Function canviewjobs() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strjobss.Split(",")
            If arr(0) = "1" Then
                canviewjobs = True
            Else
                canviewjobs = False
            End If
        Catch ex As Exception
            Try
                canviewjobs = False
            Catch exc As Exception

            End Try
        End Try
    End Function
#End Region

#Region "leads"
    Private Function canmanipulateleads() As Boolean
        Try
            Dim arr() As String
            arr = myForms.Main.strleadss.Split(",")
            If arr(1) = "1" Then
                canmanipulateleads = True
            Else
                canmanipulateleads = False
            End If
        Catch ex As Exception
            Try
                canmanipulateleads = False
            Catch exc As Exception

            End Try
        End Try
    End Function
    Private Sub btnAddLeads_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddLeads.Click
        Try
            Dim x As Boolean = canmanipulateleads()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate leads contact administrator", "Leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If addleads = False Then
                Dim form As New frmAddLead
                form.grpnormal.Visible = True
                form.pnlSearch.Visible = False
                form.clientno = Me.lblClientNo.Text.ToString()
                form.cname = Me.lblClientName.Text.ToString()
                'form.Show()
                myForms.CustomerForm = form
                myForms.CustomerForm.Show()
                addleads = True
            Else

                myForms.CustomerForm.grpnormal.Visible = True
                myForms.CustomerForm.pnlSearch.Visible = False
                myForms.CustomerForm.clientno = Me.lblClientNo.Text.ToString()
                myForms.CustomerForm.cname = Me.lblClientName.Text.ToString()

            End If

        Catch ex As Exception

        End Try
    End Sub
    Public Sub loadleads()
        Dim cnnstr As String
        cnnstr = "DSN=" & myForms.qconnstr
        'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = cnnstr
        connect.Open()

        Try
            '-----------------try this dave

            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter
            Dim custDS As DataSet = New DataSet
            Dim adors As New ADODB.Recordset
            Dim str As String = "select " _
                  & "*" _
                  & "from" _
                  & " leads" _
                  & " where lower(status) <> '" & "rrrrrr" & "'" _
                  & "and client_no='" & lblClientNo.Text & "'"
            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leads")
            Dim tname As String = custDS.Tables(0).TableName()
            Me.dtgLeads.SetDataBinding(custDS, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------
            dtgLeads.SetDataBinding(custDS, tname)

            addcustleadstablestyle(tname)

        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub
    Public Sub addcustleadstablestyle(ByVal tname As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = dtgLeads.Width - 20
            mywidth = mywidth / 3

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "leads_no"
            myno.HeaderText = "Lead Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            '' Add a second column style.
            'Dim myname As New DataGridTextBoxColumn()
            'myname.MappingName = "client_no"
            'myname.HeaderText = "Client Number"
            'myname.Width = mywidth
            'ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "title"
            mydesc.HeaderText = "Title"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn
            mydesc2.MappingName = "status"
            mydesc2.HeaderText = "Status"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)


            '' Add a second column style.
            'Dim mydesc200 As New DataGridTextBoxColumn()
            'mydesc200.MappingName = "date_sniffed"
            'mydesc200.HeaderText = "Date"
            'mydesc200.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc200)

            'Dim mydesc20 As New DataGridTextBoxColumn()
            'mydesc20.MappingName = "amount"
            'mydesc20.HeaderText = "Amount"
            'mydesc20.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc20)
            ' Add the DataGridTableStyle objects to the collection.
            dtgLeads.TableStyles.Clear()
            dtgLeads.TableStyles.Add(ts1)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtgLeads_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgLeads.DoubleClick
        Try
            Dim x As Boolean = canmanipulateleads()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate leads contact administrator", "Leads", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch ex As Exception

        End Try

        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim leadno, cno, status, descrip, title, ddate
            Dim amount, department
            Dim results As String
            results = returnhittestleads(htileads.Type)
            If results <> "" Then
                Dim mycell As New DataGridCell
                Dim a() As String
                a = results.Split("|")
                mycell.RowNumber = CInt(a(0))
                mycell.ColumnNumber = CInt(a(1))
                '----------
                Dim ds As System.Data.DataSet = New System.Data.DataSet
                ds = Me.dtgLeads.DataSource
                '------------
                ' journal()
                '---------
                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("leads_no")) = False Then
                    leadno = ds.Tables(0).Rows(mycell.RowNumber).Item("leads_no")

                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("client_no")) = False Then
                    cno = ds.Tables(0).Rows(mycell.RowNumber).Item("client_no")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("title")) = False Then
                    title = ds.Tables(0).Rows(mycell.RowNumber).Item("title")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("status")) = False Then
                    status = ds.Tables(0).Rows(mycell.RowNumber).Item("status")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("date_sniffed")) = False Then
                    ddate = ds.Tables(0).Rows(mycell.RowNumber).Item("date_sniffed")
                End If
                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("amount")) = False Then
                    amount = ds.Tables(0).Rows(mycell.RowNumber).Item("amount")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("descrip")) = False Then
                    descrip = ds.Tables(0).Rows(mycell.RowNumber).Item("descrip")
                End If

                If Convert.IsDBNull(ds.Tables(0).Rows(mycell.RowNumber).Item("department")) = False Then
                    department = ds.Tables(0).Rows(mycell.RowNumber).Item("department")
                End If

                If editleads = False Then

                    Dim form As New frmEditLead
                    form.clientno = Me.lblClientNo.Text
                    form.cstatus = status
                    form.cname = lblClientName.Text
                    form.leadno = leadno
                    form.desription = descrip
                    form.title = title
                    form.amount = amount

                    myForms.CustomerForm4 = form
                    Try
                        myForms.CustomerForm4.txtdepartment.Text = department
                    Catch qw As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.dtpsniffed.Value = CDate(ddate)
                    Catch exs5 As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.cboProspect.Text = status
                    Catch exs4 As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.txtDesc.Text = descrip
                    Catch exs3 As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.txtAmount.Text = amount
                    Catch exs2 As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.txttitle.Text = title
                    Catch exs1 As Exception

                    End Try
                    myForms.CustomerForm4.txtAmount.Enabled = False
                    myForms.CustomerForm4.Show()
                    'form.Show()
                    'myForms.CustomerForm = form
                    editleads = True
                Else

                    myForms.CustomerForm4.clientno = cno
                    myForms.CustomerForm4.cstatus = status
                    myForms.CustomerForm4.cname = lblClientName.Text
                    myForms.CustomerForm4.leadno = leadno
                    myForms.CustomerForm4.desription = descrip
                    myForms.CustomerForm4.amount = amount
                    myForms.CustomerForm4.title = title
                    myForms.CustomerForm4.txtAmount.Enabled = False
                    Try
                        myForms.CustomerForm4.txtdepartment.Text = department
                    Catch qw As Exception

                    End Try
                    Try
                        myForms.CustomerForm4.txtAmount.Text = amount
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.txtAmount.Text = ""
                        Catch ev As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.dtpsniffed.Value = CDate(ddate)
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.dtpsniffed.Text = ""
                        Catch ev As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.txtDesc.Text = descrip
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.txtDesc.Text = ""
                        Catch ev As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.txttitle.Text = title
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.txttitle.Text = ""
                        Catch ev As Exception
                        End Try
                    End Try
                    Try
                        myForms.CustomerForm4.cboProspect.Text = status
                    Catch er As Exception
                        Try
                            myForms.CustomerForm4.cboProspect.Text = ""
                        Catch ev As Exception
                        End Try
                    End Try
                    editleads = True
                    'form.descriptions = descrip
                    'form.status = status
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally
            Cursor.Current = currentcursor
        End Try
    End Sub
    Private Function returnhittestleads(ByVal hit As System.Windows.Forms.DataGrid.HitTestType) As String
        Try
            Select Case hit
                Case System.Windows.Forms.DataGrid.HitTestType.Cell

                    Dim mycell As New DataGridCell
                    mycell.RowNumber = Me.dtgLeads.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittestleads = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell
                    mycell.RowNumber = Me.dtgLeads.CurrentRowIndex
                    mycell.ColumnNumber = 0
                    returnhittestleads = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case Else
                    returnhittestleads = ""
            End Select
        Catch ex As Exception
            MessageBox.Show(Text:="Error:" & ex.Message.ToString, _
            caption:="Throwing an exception", _
            Icon:=MessageBoxIcon.Information, _
            buttons:=MessageBoxButtons.OK)

        End Try
    End Function
    Private Sub dtgLeads_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgLeads.MouseDown
        Try

        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
            htileads = Me.dtgLeads.HitTest(New Point(e.X, e.Y))
        End Try
    End Sub
#End Region

#Region "all"
    'Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
    '    Try
    '        If refreshcontacts = True Then
    '            'Me.Invoke(New mydelegate(AddressOf initializecontacts))
    '            loadgridcontact()
    '            refreshcontacts = False
    '        End If
    '        If refreshjobs = True Then
    '            'Me.Invoke(New mydelegate(AddressOf initializejobs))
    '            Me.dtgContacts.Invoke(New mydelegate(AddressOf loadgridexistingjobs))
    '            refreshjobs = False
    '        End If
    '        If refreshclients = True Then
    '            Me.lblClientNo.Text = myclientno
    '            Me.lblClientName.Text = myclientname

    '        End If
    '        If refreshleads = True Then
    '            loadleads()
    '            refreshleads = False
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub
#End Region

    Private Sub btndeletecontact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeletecontact.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateclients
            If x = False Then
                MessageBox.Show("Not allowed to manipulate clients contact administrator", "Clients", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Me.dtgContacts.Select(hti1.Row)
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
            ds = Me.dtgContacts.DataSource
            Dim sid, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hti1.Row).Item("ano")
            str = "delete from contact where ano='" & sid & "';"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(hti1.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
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

End Class
