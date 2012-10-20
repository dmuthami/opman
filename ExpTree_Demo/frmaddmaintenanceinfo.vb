
Imports System
Imports System.Data
Imports System.Threading
Imports ADODB

Public Class frmaddmaintenanceinfo
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
            myForms.isadddmaintenace = False
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboequipid As System.Windows.Forms.ComboBox
    Friend WithEvents txtinvoiceno As System.Windows.Forms.TextBox
    Friend WithEvents txtcost As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtdesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpservicedate As System.Windows.Forms.DateTimePicker
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmaddmaintenanceinfo))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnadd = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cboequipid = New System.Windows.Forms.ComboBox
        Me.txtinvoiceno = New System.Windows.Forms.TextBox
        Me.txtcost = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtdesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpservicedate = New System.Windows.Forms.DateTimePicker
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnclose)
        Me.Panel1.Controls.Add(Me.btnadd)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 258)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(384, 32)
        Me.Panel1.TabIndex = 7
        '
        'btnclose
        '
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(277, 8)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(104, 23)
        Me.btnclose.TabIndex = 9
        Me.btnclose.Text = "Close"
        '
        'btnadd
        '
        Me.btnadd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnadd.Location = New System.Drawing.Point(2, 8)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.Size = New System.Drawing.Size(104, 23)
        Me.btnadd.TabIndex = 8
        Me.btnadd.Text = "Add"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboequipid)
        Me.GroupBox1.Controls.Add(Me.txtinvoiceno)
        Me.GroupBox1.Controls.Add(Me.txtcost)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.dtpservicedate)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(384, 258)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'cboequipid
        '
        Me.cboequipid.Location = New System.Drawing.Point(104, 15)
        Me.cboequipid.Name = "cboequipid"
        Me.cboequipid.Size = New System.Drawing.Size(272, 22)
        Me.cboequipid.TabIndex = 1
        '
        'txtinvoiceno
        '
        Me.txtinvoiceno.Location = New System.Drawing.Point(104, 227)
        Me.txtinvoiceno.Name = "txtinvoiceno"
        Me.txtinvoiceno.Size = New System.Drawing.Size(272, 20)
        Me.txtinvoiceno.TabIndex = 6
        Me.txtinvoiceno.Text = ""
        '
        'txtcost
        '
        Me.txtcost.Location = New System.Drawing.Point(104, 203)
        Me.txtcost.Name = "txtcost"
        Me.txtcost.Size = New System.Drawing.Size(272, 20)
        Me.txtcost.TabIndex = 5
        Me.txtcost.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(4, 227)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 23)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Invoice Number"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(1, 203)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Cost Incurred"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtdesc)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(2, 63)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(376, 136)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Description"
        '
        'txtdesc
        '
        Me.txtdesc.Location = New System.Drawing.Point(8, 16)
        Me.txtdesc.Multiline = True
        Me.txtdesc.Name = "txtdesc"
        Me.txtdesc.Size = New System.Drawing.Size(360, 112)
        Me.txtdesc.TabIndex = 4
        Me.txtdesc.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Service Date"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 18)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Equipment Id"
        '
        'dtpservicedate
        '
        Me.dtpservicedate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpservicedate.Location = New System.Drawing.Point(104, 40)
        Me.dtpservicedate.Name = "dtpservicedate"
        Me.dtpservicedate.Size = New System.Drawing.Size(272, 20)
        Me.dtpservicedate.TabIndex = 2
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        '
        'frmaddmaintenanceinfo
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(384, 290)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmaddmaintenanceinfo"
        Me.Text = "Add maintenance information"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
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
    Private autono As String
    Private threadmaintenace As System.Threading.Thread
    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
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
            If Me.cboequipid.Text.Length < 1 Then
                MessageBox.Show("Please pick an equipment id", "Save", _
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
            ''-------------check if id number already exists
            'Dim rs As New ADODB.Recordset()
            'Dim Str As String = " select equip_id    from equip_info" _
            '& " where lower(equip_id) like '" & txtequipid.Text.Trim.ToLower() & "'"

            'With rs
            '    .CursorLocation = CursorLocationEnum.adUseClient
            '    .CursorType = CursorTypeEnum.adOpenStatic
            '    .Open(Str, connect)
            '    If .BOF = False And .EOF = False Then
            '        MessageBox.Show(" Equipment id  already exists", "Save", _
            '        MessageBoxButtons.OK, MessageBoxIcon.Information)
            '        isvalid = False
            '        Exit Try
            '    End If
            '    .Close()
            'End With
            ''--------------end of sanity check
            Dim pdate As String
            pdate = dtpservicedate.Value.Year & "-" _
                       & dtpservicedate.Value.Month & "-" _
                       & dtpservicedate.Value.Day & " " _
                       & dtpservicedate.Value.Hour & ":" _
                       & dtpservicedate.Value.Minute & ":" _
                       & dtpservicedate.Value.Second
            Dim a() As String
            a = cboequipid.Text.Split(":")
            Dim ff As String = a(0)
            '-------------------
            Dim arr() As String
            Dim strr As String
            Dim y As Integer
            txtdesc.Text = Me.txtdesc.Text.Trim()
            arr = txtdesc.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            '----------------------------------
            Dim strsql As String
            strsql = "insert into maintenance_info (equip_id,service_date,description,cost_incurred,invoice_no,autonumber) "

            strsql += " values ('" & ff.Trim & "','" & pdate & "','" & strr & "',"
            strsql += "'" & Me.txtcost.Text.Trim & "','" & Me.txtinvoiceno.Text.Trim & "','" & autono & "');"

            connect.BeginTrans()
            connect.IsolationLevel = IsolationLevelEnum.adXactSerializable
            connect.Execute(strsql)
            connect.CommitTrans()
            Dim Tasks As taskclass
            threadmaintenace = New System.Threading.Thread( _
                AddressOf Tasks.servequipdetailsinvoke)
            threadmaintenace.Start()
            MessageBox.Show("Addition successful", "Adding equipments", _
            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Try
                connect.Close()
            Catch wq As Exception

            End Try
        Catch ex As Exception
        Finally
            If isvalid = True Then
                Me.cboequipid.Text = ""
                Me.txtcost.Text = ""
                txtdesc.Text = ""

            End If

        End Try

    End Sub
    Private Sub frmaddmaintenanceinfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim threadsd As System.Threading.Thread

            threadsd = New System.Threading.Thread( _
                       AddressOf loadcbo)
            threadsd.Start()
            Dim threadsd1 As System.Threading.Thread

            threadsd1 = New System.Threading.Thread( _
                       AddressOf clientnumber)
            threadsd1.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.isadddmaintenace = False
            myForms.addmaintenace.Dispose(True)
        Catch we As Exception
        End Try
    End Sub
    Private Sub clientnumber()
        Dim clientnumber1 As String
        Try
            Dim cnnstr As String
            cnnstr = "DSN=" & myForms.qconnstr
            'cnnstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = cnnstr
            connect.Open()

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenForwardOnly
                Dim str = "select max(autonumber) from maintenance_info"
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    If Convert.IsDBNull(.Fields("max").Value) = True Then
                        clientnumber1 = "1"
                    Else
                        clientnumber1 = .Fields("max").Value
                        clientnumber1 = (CLng(clientnumber1) + 1).ToString()
                    End If

                Else
                    clientnumber1 = "1"

                End If
            End With
            Try
                connect.Close()
            Catch er As Exception

            End Try
        Catch ex As Exception
        Finally
            autono = clientnumber1
        End Try
    End Sub
    Private Sub loadcbo()

        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect1 As New ADODB.Connection
            connect1.Mode = ConnectModeEnum.adModeReadWrite
            connect1.CursorLocation = CursorLocationEnum.adUseClient
            connect1.ConnectionString = connectstr
            connect1.Open()
            'Dim str As String = " SELECT  equip_info.*,  assigned_info.equip_id AS id2,  assigned_info.status " _
            '                    & " FROM  assigned_info RIGHT OUTER JOIN " _
            '                    & " equip_info ON  assigned_info.status =  equip_info.equip_id"
            'str += " and assigned_info.status='" & "0" & "'"
            Dim str As String = "SELECT *  from equip_info order by equip_id asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect1)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            Me.cboequipid.Items.Add(Convert.ToString(.Fields("equip_id").Value) & " : " & _
                            Convert.ToString(.Fields("model_name").Value) & " : " & _
                            Convert.ToString(.Fields("model_no").Value))
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

                connect1.Close()
            Catch er344 As Exception

            End Try
        Catch ex As Exception

        End Try

    End Sub
    Private Sub cno()
        Try

        Catch vb As Exception
        End Try
    End Sub

#Region " validation"
    Private Sub txtinvoiceno_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtinvoiceno.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtinvoiceno, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtinvoiceno, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
    Private Sub txtcost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcost.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtcost, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtcost, "")
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

End Class
