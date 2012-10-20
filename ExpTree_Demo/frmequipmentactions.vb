
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Threading
Imports ADODB


Public Class frmequipmentactions
    Inherits System.Windows.Forms.Form

    Private comboControl As New System.Windows.Forms.ComboBox()
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
            myForms.isequipactions = False
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnservice_info As System.Windows.Forms.Button
    Friend WithEvents btnmainte_info As System.Windows.Forms.Button
    Friend WithEvents btndecommission As System.Windows.Forms.Button
    Friend WithEvents btnrelease As System.Windows.Forms.Button
    Friend WithEvents btnassign As System.Windows.Forms.Button
    Friend WithEvents txtassignedby As System.Windows.Forms.TextBox
    Friend WithEvents cbojob As System.Windows.Forms.ComboBox
    Friend WithEvents txtreleasedesc As System.Windows.Forms.RichTextBox
    Friend WithEvents txtassignedto As System.Windows.Forms.TextBox
    Friend WithEvents dtpprd As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpdas As System.Windows.Forms.DateTimePicker
    Friend WithEvents grpassign As System.Windows.Forms.GroupBox
    Friend WithEvents grprelease As System.Windows.Forms.GroupBox
    Friend WithEvents grpdecommission As System.Windows.Forms.GroupBox
    Friend WithEvents txtdecommissionrelease As System.Windows.Forms.RichTextBox
    Friend WithEvents txtequipname As System.Windows.Forms.TextBox
    Friend WithEvents txtequipid As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmequipmentactions))
        Me.grpassign = New System.Windows.Forms.GroupBox
        Me.dtpdas = New System.Windows.Forms.DateTimePicker
        Me.dtpprd = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtassignedby = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbojob = New System.Windows.Forms.ComboBox
        Me.btnassign = New System.Windows.Forms.Button
        Me.grprelease = New System.Windows.Forms.GroupBox
        Me.txtreleasedesc = New System.Windows.Forms.RichTextBox
        Me.txtassignedto = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnrelease = New System.Windows.Forms.Button
        Me.grpdecommission = New System.Windows.Forms.GroupBox
        Me.txtdecommissionrelease = New System.Windows.Forms.RichTextBox
        Me.txtequipname = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtequipid = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btndecommission = New System.Windows.Forms.Button
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnservice_info = New System.Windows.Forms.Button
        Me.btnmainte_info = New System.Windows.Forms.Button
        Me.grpassign.SuspendLayout()
        Me.grprelease.SuspendLayout()
        Me.grpdecommission.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpassign
        '
        Me.grpassign.Controls.Add(Me.dtpdas)
        Me.grpassign.Controls.Add(Me.dtpprd)
        Me.grpassign.Controls.Add(Me.Label8)
        Me.grpassign.Controls.Add(Me.Label7)
        Me.grpassign.Controls.Add(Me.txtassignedby)
        Me.grpassign.Controls.Add(Me.Label6)
        Me.grpassign.Controls.Add(Me.Label1)
        Me.grpassign.Controls.Add(Me.cbojob)
        Me.grpassign.Controls.Add(Me.btnassign)
        Me.grpassign.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpassign.Location = New System.Drawing.Point(8, 4)
        Me.grpassign.Name = "grpassign"
        Me.grpassign.Size = New System.Drawing.Size(376, 140)
        Me.grpassign.TabIndex = 0
        Me.grpassign.TabStop = False
        Me.grpassign.Text = "Assign to job"
        '
        'dtpdas
        '
        Me.dtpdas.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpdas.Location = New System.Drawing.Point(144, 64)
        Me.dtpdas.Name = "dtpdas"
        Me.dtpdas.Size = New System.Drawing.Size(224, 20)
        Me.dtpdas.TabIndex = 3
        '
        'dtpprd
        '
        Me.dtpprd.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpprd.Location = New System.Drawing.Point(144, 88)
        Me.dtpprd.Name = "dtpprd"
        Me.dtpprd.Size = New System.Drawing.Size(224, 20)
        Me.dtpprd.TabIndex = 4
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 85)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(128, 16)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Proposed release date"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 16)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Date assigned"
        '
        'txtassignedby
        '
        Me.txtassignedby.Location = New System.Drawing.Point(144, 39)
        Me.txtassignedby.Name = "txtassignedby"
        Me.txtassignedby.Size = New System.Drawing.Size(224, 20)
        Me.txtassignedby.TabIndex = 2
        Me.txtassignedby.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 41)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 16)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "Assigned by"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Pick a job"
        '
        'cbojob
        '
        Me.cbojob.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbojob.Location = New System.Drawing.Point(144, 16)
        Me.cbojob.Name = "cbojob"
        Me.cbojob.Size = New System.Drawing.Size(224, 22)
        Me.cbojob.TabIndex = 1
        '
        'btnassign
        '
        Me.btnassign.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnassign.Location = New System.Drawing.Point(280, 112)
        Me.btnassign.Name = "btnassign"
        Me.btnassign.Size = New System.Drawing.Size(88, 23)
        Me.btnassign.TabIndex = 5
        Me.btnassign.Text = "Assign to job"
        '
        'grprelease
        '
        Me.grprelease.Controls.Add(Me.txtreleasedesc)
        Me.grprelease.Controls.Add(Me.txtassignedto)
        Me.grprelease.Controls.Add(Me.Label2)
        Me.grprelease.Controls.Add(Me.btnrelease)
        Me.grprelease.Enabled = False
        Me.grprelease.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grprelease.Location = New System.Drawing.Point(8, 144)
        Me.grprelease.Name = "grprelease"
        Me.grprelease.Size = New System.Drawing.Size(376, 136)
        Me.grprelease.TabIndex = 6
        Me.grprelease.TabStop = False
        Me.grprelease.Text = "Release equipment"
        '
        'txtreleasedesc
        '
        Me.txtreleasedesc.Location = New System.Drawing.Point(8, 40)
        Me.txtreleasedesc.Name = "txtreleasedesc"
        Me.txtreleasedesc.Size = New System.Drawing.Size(360, 64)
        Me.txtreleasedesc.TabIndex = 8
        Me.txtreleasedesc.Text = "Description"
        '
        'txtassignedto
        '
        Me.txtassignedto.Location = New System.Drawing.Point(136, 16)
        Me.txtassignedto.Name = "txtassignedto"
        Me.txtassignedto.Size = New System.Drawing.Size(232, 20)
        Me.txtassignedto.TabIndex = 7
        Me.txtassignedto.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 24)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Currently assigned to:"
        '
        'btnrelease
        '
        Me.btnrelease.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnrelease.Location = New System.Drawing.Point(250, 109)
        Me.btnrelease.Name = "btnrelease"
        Me.btnrelease.Size = New System.Drawing.Size(120, 23)
        Me.btnrelease.TabIndex = 9
        Me.btnrelease.Text = "Release equipment"
        '
        'grpdecommission
        '
        Me.grpdecommission.Controls.Add(Me.txtdecommissionrelease)
        Me.grpdecommission.Controls.Add(Me.txtequipname)
        Me.grpdecommission.Controls.Add(Me.Label4)
        Me.grpdecommission.Controls.Add(Me.txtequipid)
        Me.grpdecommission.Controls.Add(Me.Label3)
        Me.grpdecommission.Controls.Add(Me.btndecommission)
        Me.grpdecommission.Enabled = False
        Me.grpdecommission.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpdecommission.Location = New System.Drawing.Point(8, 280)
        Me.grpdecommission.Name = "grpdecommission"
        Me.grpdecommission.Size = New System.Drawing.Size(376, 152)
        Me.grpdecommission.TabIndex = 10
        Me.grpdecommission.TabStop = False
        Me.grpdecommission.Text = "Decommision/reinstate"
        '
        'txtdecommissionrelease
        '
        Me.txtdecommissionrelease.Location = New System.Drawing.Point(8, 56)
        Me.txtdecommissionrelease.Name = "txtdecommissionrelease"
        Me.txtdecommissionrelease.Size = New System.Drawing.Size(360, 64)
        Me.txtdecommissionrelease.TabIndex = 13
        Me.txtdecommissionrelease.Text = "Description"
        '
        'txtequipname
        '
        Me.txtequipname.Location = New System.Drawing.Point(136, 35)
        Me.txtequipname.Name = "txtequipname"
        Me.txtequipname.Size = New System.Drawing.Size(232, 20)
        Me.txtequipname.TabIndex = 12
        Me.txtequipname.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Equipment name"
        '
        'txtequipid
        '
        Me.txtequipid.Location = New System.Drawing.Point(136, 15)
        Me.txtequipid.Name = "txtequipid"
        Me.txtequipid.Size = New System.Drawing.Size(232, 20)
        Me.txtequipid.TabIndex = 11
        Me.txtequipid.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Equipment id"
        '
        'btndecommission
        '
        Me.btndecommission.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndecommission.Location = New System.Drawing.Point(264, 123)
        Me.btndecommission.Name = "btndecommission"
        Me.btndecommission.Size = New System.Drawing.Size(104, 23)
        Me.btndecommission.TabIndex = 14
        Me.btndecommission.Text = "Decommission"
        '
        'btnclose
        '
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(304, 440)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 17
        Me.btnclose.Text = "Close"
        '
        'btnservice_info
        '
        Me.btnservice_info.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnservice_info.Location = New System.Drawing.Point(160, 440)
        Me.btnservice_info.Name = "btnservice_info"
        Me.btnservice_info.Size = New System.Drawing.Size(128, 23)
        Me.btnservice_info.TabIndex = 16
        Me.btnservice_info.Text = "Service information"
        Me.btnservice_info.Visible = False
        '
        'btnmainte_info
        '
        Me.btnmainte_info.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnmainte_info.Location = New System.Drawing.Point(8, 440)
        Me.btnmainte_info.Name = "btnmainte_info"
        Me.btnmainte_info.Size = New System.Drawing.Size(144, 23)
        Me.btnmainte_info.TabIndex = 15
        Me.btnmainte_info.Text = "Maintenace information"
        Me.btnmainte_info.Visible = False
        '
        'frmequipmentactions
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(384, 466)
        Me.Controls.Add(Me.btnclose)
        Me.Controls.Add(Me.grpdecommission)
        Me.Controls.Add(Me.grprelease)
        Me.Controls.Add(Me.grpassign)
        Me.Controls.Add(Me.btnservice_info)
        Me.Controls.Add(Me.btnmainte_info)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmequipmentactions"
        Me.Text = "Equipment Actions"
        Me.grpassign.ResumeLayout(False)
        Me.grprelease.ResumeLayout(False)
        Me.grpdecommission.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
            Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private Delegate Sub mydelegatet1()
    Private Delegate Sub mydelegatet2()
    Public eid As String
    Private Sub frmequipmentactions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call loadthree()
        Catch sd As Exception

        End Try
    End Sub
    Public Sub loadthree()
        Try
            Dim threadsc As New System.Threading.Thread(AddressOf _
             populatecbojob)
            threadsc.Start()

            Dim threadxz As New System.Threading.Thread(AddressOf _
                    release)
            threadxz.Start()

        Catch we As Exception

        End Try
    End Sub
#Region "assign"
    Private Sub assigninvoke()

    End Sub
    Private Sub assign()
        Try
            Me.Invoke(New mydelegatet1(AddressOf populatecbojob))
        Catch cv As Exception

        End Try
    End Sub
    Public Sub populatecbojob()
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = " SELECT rcljobs.*,  clients.name" _
            & " FROM clients INNER JOIN" _
            & "  rcljobs ON clients.client_no = rcljobs.client_no"

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            cbojob.Items.Add(Convert.ToString(.Fields("job_no").Value) & " : " & _
                            Convert.ToString(.Fields("job_tittle").Value) & " : " & _
                            Convert.ToString(.Fields("name").Value))
                            comboControl.Items.Add(.Fields("job_no").Value)
                        Catch es300 As Exception
                        End Try
                        Application.DoEvents()
                        .MoveNext()
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
        Catch we As Exception

        End Try
        Try
            Dim tasks As taskclass
            Me.txtassignedby.Text = tasks.globalnamme
        Catch asc As Exception

        End Try
    End Sub
    Private Sub btnassign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnassign.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xcv As Exception

        End Try
        Try
            If cbojob.Text.Length < 1 Then
                MessageBox.Show("Please select a job ", "Assign equipment to job", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            Else
                Call assignequip3()
            End If
        Catch we As Exception

        End Try
    End Sub
    Private Sub assignequip3()
        Dim isnew As Boolean = False
        Dim isassign As Boolean = False
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim strsql As String = "" 'IN SHARE MODE
            Dim ano As String
            connect.BeginTrans()
            strsql = " BEGIN WORK;" _
                    & " LOCK TABLE assigned_info,tblno,current_equip ;" _
               & ";"
            connect.Execute(strsql)
            Dim clientnumber1 As String
            Try
                Dim rsv As New ADODB.Recordset
                With rsv
                    .CursorLocation = CursorLocationEnum.adUseClient
                    .CursorType = CursorTypeEnum.adOpenForwardOnly
                    Dim str = "select max(auto_no) from tblno"
                    .Open(str, connect)
                    If .BOF = False And .EOF = False Then
                        If Convert.IsDBNull(.Fields("max").Value) = True Then
                            clientnumber1 = "1"
                            isnew = True
                        Else
                            clientnumber1 = .Fields("max").Value
                            clientnumber1 = (CLng(clientnumber1) + 1).ToString()
                        End If

                    Else
                        clientnumber1 = "1"
                        isnew = True

                    End If
                End With
                Try

                Catch er As Exception

                End Try
            Catch ex As Exception
            Finally
                ano = clientnumber1
            End Try
            If isnew = False Then
                connect.Execute(" update tblno set auto_no='" & ano & "';")
            Else
                connect.Execute(" insert into tblno (auto_no) values ('" & ano & "');")
            End If

            strsql = "SELECT *  FROM assigned_info " _
                  & "WHERE equip_id ='" & eid & "'" _
                  & "     "
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(strsql, connect)
                If .BOF = False And .EOF = False Then
                    Dim snow, edate As String
                    Dim d As New System.Windows.Forms.DateTimePicker
                    d.Value = Now
                    snow = d.Value.Year & "-" _
                         & d.Value.Month & "-" _
                         & d.Value.Day & " " _
                         & "00" & ":" _
                         & "00" & ":" _
                         & "00"

                    edate = dtpprd.Value.Year & "-" _
                   & dtpprd.Value.Month & "-" _
                   & dtpprd.Value.Day & " " _
                   & "00" & ":" _
                   & "00" & ":" _
                   & "00"
                    Dim jjobno As String
                    Dim a() As String = Me.cbojob.Text.Split(":")
                    jjobno = a(0)
                    strsql = " update assigned_info set status='" & "1" & "' where equip_id='" & eid & "';"
                    strsql += " insert into  current_equip (equip_id,job_no,task,other,description,assigned_by,date_assigned," _
                  & " estimate_release_date,autonumber) values  " _
                  & " ("
                    strsql += "  '" & eid & "'," _
                    & " '" & jjobno & "','" & "" & "',"
                    strsql += " '" & "" & "','" & "" & "','" & Me.txtassignedby.Text.Trim & "'," _
                    & "'" & snow & "','" & edate & "','" & ano & "');"
                    strsql += " commit work;"
                    connect.Execute(strsql)
                    isassign = True
                End If

            End With
            connect.CommitTrans()
            Try
                connect.Close()
            Catch gh As Exception
            End Try
        Catch ex As Exception
        End Try
        Try
            If isassign = True Then
                Dim threadxz As New System.Threading.Thread(AddressOf _
                               release)
                threadxz.Start()
                isassign = False
            End If

        Catch sx As Exception

        End Try

    End Sub
#End Region
#Region "release"
    Public Sub releaseinvoke()

    End Sub
    Public Sub release()
        Me.Invoke(New mydelegatet2(AddressOf assignequip))
    End Sub
    Public Sub assignequip()
        Dim loadesc As Boolean = False
        Try

            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " SELECT      assigned_info.status,  rcljobs.job_tittle,  current_equip.equip_id, current_equip.job_no,  current_equip.other, " _
            '             & "  current_equip.task,  current_equip.description,  current_equip.assigned_by,  current_equip.date_assigned, " _
            '             & "  current_equip.estimate_release_date,  current_equip.date_released, current_equip.autonumber" _
            '            & " FROM         assigned_info INNER JOIN" _
            '           & "  current_equip ON  assigned_info.equip_id = current_equip.equip_id INNER JOIN" _
            '          & "   rcljobs ON  current_equip.job_no =  rcljobs.job_no " _
            '            & " WHERE      (current_equip.equip_id = '" & eid & "') AND  (assigned_info.status = '" & "1" & "')"
            'str += " where assigned_info.status='" & "1" & "'"
            'str += " and  assigned_info.equip_id='" & eid & "';"
            Dim str As String = " SELECT     assigned_info.status  , current_equip.equip_id,  current_equip.job_no, current_equip.other,  current_equip.task, " _
                                            & " current_equip.description,  current_equip.assigned_by,  current_equip.date_assigned,  current_equip.estimate_release_date, " _
                                            & " current_equip.date_released, current_equip.autonumber" _
                                            & " FROM assigned_info INNER JOIN" _
                                            & " current_equip ON  assigned_info.equip_id =  current_equip.equip_id" _
                                            & " WHERE       current_equip.equip_id = '" & eid & "'  and assigned_info.status='" & "1" & "';"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Me.txtassignedto.Text = .Fields("job_no").Value
                    Me.grprelease.Enabled = True
                    Me.grpassign.Enabled = False
                    Me.grpdecommission.Enabled = False
                    loadesc = True
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
            If loadesc = True Then
                Call loaddesc()
            End If
        Catch ex As Exception


        End Try
    End Sub
    Public Sub loaddesc()
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select " _
                      & "rcljobs.client_no,rcljobs.job_no,rcljobs.job_tittle,rcljobs.job_status, " _
                      & "clients.client_no ,clients.name" _
                      & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
                      & " lower(rcljobs.job_no) like" _
                      & "'%" & Me.txtassignedto.Text.Trim & "%' "
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    txtreleasedesc.Multiline = True
                    Me.txtreleasedesc.Text = "Job Tittle: " & .Fields("job_tittle").Value & vbCrLf
                    Me.txtreleasedesc.AppendText("Contact Name: " & .Fields("name").Value & vbCrLf)
                    Me.txtreleasedesc.AppendText("Job Status: " & .Fields("job_status").Value & vbCrLf)
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
        Catch we As Exception
        End Try
    End Sub
    Private Sub btnrelease_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrelease.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xcv As Exception

        End Try
        Try
            Call assignequip2()
        Catch qw As Exception

        End Try
    End Sub
    Private Sub assignequip2()
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim strsql As String = "" 'IN SHARE MODE
            Dim ano As String
            connect.BeginTrans()
            strsql = " BEGIN WORK;" _
                    & " LOCK TABLE assigned_info,current_equip;"
            connect.Execute(strsql)
            strsql = "SELECT *  FROM assigned_info " _
       & "WHERE equip_id ='" & eid & "'" _
       & "     "
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(strsql, connect)
                If .BOF = False And .EOF = False Then
                    If .Fields("status").Value = "1" Then
                        strsql = " update assigned_info set status='" & "0" & "' where equip_id='" & eid & "';"
                        strsql += " INSERT INTO history_equip " _
                        & " select *  from current_equip where equip_id='" & eid & "'; "
                        strsql += " delete from current_equip where equip_id='" & eid & "';"
                        strsql += " commit work;"
                        connect.Execute(strsql)
                        txtdecommissionrelease.Text = ""
                        txtequipid.Text = ""
                        txtequipname.Text = ""
                        txtreleasedesc.Text = ""
                        txtassignedto.Text = ""

                        grpassign.Enabled = True
                        grpdecommission.Enabled = False
                        grprelease.Enabled = False
                    End If
                End If
            End With
            connect.CommitTrans()
            Try
                connect.Close()
            Catch gh As Exception
            End Try
        Catch ex As Exception
        End Try

    End Sub
#End Region
    Public Sub commissioninvoke()

    End Sub
    Public Sub commission()

    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.isequipactions = False
            myForms.equipactions.Dispose(True)
        Catch sd As Exception

        End Try
    End Sub

    Private Sub btnmainte_info_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnmainte_info.Click

    End Sub

    Private Sub btndecommission_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndecommission.Click

    End Sub
End Class
