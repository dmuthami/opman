
Imports System
Imports System.Data
Imports System.Threading
Imports ADODB

Public Class frmadmin
    Inherits System.Windows.Forms.Form
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
    Private previousrow As New DataGridCell()
    '---------------other variables
    Private utask As Boolean = False
    Private ujobtitle As Boolean = False
    '---------------controls
    Public WithEvents comboControl As System.Windows.Forms.ComboBox
    Public WithEvents comboid As System.Windows.Forms.ComboBox
    Public WithEvents txttask As System.Windows.Forms.TextBox
    Public WithEvents datagridtextBox As DataGridTextBoxColumn
    Public WithEvents datagridtextBox1 As DataGridTextBoxColumn
    '--------------end of controls

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
    Friend WithEvents pnluserrights As System.Windows.Forms.Panel
    Friend WithEvents pnlgrid As System.Windows.Forms.Panel
    Friend WithEvents dtgusers As System.Windows.Forms.DataGrid
    Friend WithEvents btnaddnewuser As System.Windows.Forms.Button
    Friend WithEvents btnedituserdetails As System.Windows.Forms.Button
    Friend WithEvents btndeleteuser As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnshowall As System.Windows.Forms.Button
    Friend WithEvents pnlcontrols As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmadmin))
        Me.pnluserrights = New System.Windows.Forms.Panel
        Me.pnlcontrols = New System.Windows.Forms.Panel
        Me.btnshowall = New System.Windows.Forms.Button
        Me.btnaddnewuser = New System.Windows.Forms.Button
        Me.btnedituserdetails = New System.Windows.Forms.Button
        Me.btndeleteuser = New System.Windows.Forms.Button
        Me.btnclose = New System.Windows.Forms.Button
        Me.pnlgrid = New System.Windows.Forms.Panel
        Me.dtgusers = New System.Windows.Forms.DataGrid
        Me.pnluserrights.SuspendLayout()
        Me.pnlcontrols.SuspendLayout()
        Me.pnlgrid.SuspendLayout()
        CType(Me.dtgusers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnluserrights
        '
        Me.pnluserrights.Controls.Add(Me.pnlcontrols)
        Me.pnluserrights.Controls.Add(Me.pnlgrid)
        Me.pnluserrights.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnluserrights.Location = New System.Drawing.Point(0, 0)
        Me.pnluserrights.Name = "pnluserrights"
        Me.pnluserrights.Size = New System.Drawing.Size(480, 522)
        Me.pnluserrights.TabIndex = 0
        '
        'pnlcontrols
        '
        Me.pnlcontrols.Controls.Add(Me.btnshowall)
        Me.pnlcontrols.Controls.Add(Me.btnaddnewuser)
        Me.pnlcontrols.Controls.Add(Me.btnedituserdetails)
        Me.pnlcontrols.Controls.Add(Me.btndeleteuser)
        Me.pnlcontrols.Controls.Add(Me.btnclose)
        Me.pnlcontrols.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlcontrols.Location = New System.Drawing.Point(0, 0)
        Me.pnlcontrols.Name = "pnlcontrols"
        Me.pnlcontrols.Size = New System.Drawing.Size(480, 56)
        Me.pnlcontrols.TabIndex = 10
        '
        'btnshowall
        '
        Me.btnshowall.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnshowall.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnshowall.ForeColor = System.Drawing.SystemColors.WindowText
        Me.btnshowall.Location = New System.Drawing.Point(3, 32)
        Me.btnshowall.Name = "btnshowall"
        Me.btnshowall.Size = New System.Drawing.Size(192, 20)
        Me.btnshowall.TabIndex = 2
        Me.btnshowall.Text = "Show all users"
        '
        'btnaddnewuser
        '
        Me.btnaddnewuser.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnaddnewuser.Location = New System.Drawing.Point(3, 8)
        Me.btnaddnewuser.Name = "btnaddnewuser"
        Me.btnaddnewuser.Size = New System.Drawing.Size(88, 20)
        Me.btnaddnewuser.TabIndex = 0
        Me.btnaddnewuser.Text = "Add new user"
        '
        'btnedituserdetails
        '
        Me.btnedituserdetails.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnedituserdetails.Location = New System.Drawing.Point(91, 8)
        Me.btnedituserdetails.Name = "btnedituserdetails"
        Me.btnedituserdetails.Size = New System.Drawing.Size(104, 20)
        Me.btnedituserdetails.TabIndex = 1
        Me.btnedituserdetails.Text = "Save changes"
        '
        'btndeleteuser
        '
        Me.btndeleteuser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btndeleteuser.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeleteuser.Location = New System.Drawing.Point(299, 8)
        Me.btndeleteuser.Name = "btndeleteuser"
        Me.btndeleteuser.Size = New System.Drawing.Size(94, 40)
        Me.btndeleteuser.TabIndex = 3
        Me.btndeleteuser.Text = "Delete user(s)"
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(395, 8)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(75, 40)
        Me.btnclose.TabIndex = 4
        Me.btnclose.Text = "Close"
        '
        'pnlgrid
        '
        Me.pnlgrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlgrid.Controls.Add(Me.dtgusers)
        Me.pnlgrid.Location = New System.Drawing.Point(0, 56)
        Me.pnlgrid.Name = "pnlgrid"
        Me.pnlgrid.Size = New System.Drawing.Size(472, 456)
        Me.pnlgrid.TabIndex = 1
        '
        'dtgusers
        '
        Me.dtgusers.DataMember = ""
        Me.dtgusers.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtgusers.FlatMode = True
        Me.dtgusers.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgusers.Location = New System.Drawing.Point(0, 0)
        Me.dtgusers.Name = "dtgusers"
        Me.dtgusers.ReadOnly = True
        Me.dtgusers.Size = New System.Drawing.Size(472, 456)
        Me.dtgusers.TabIndex = 5
        '
        'frmadmin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(480, 522)
        Me.Controls.Add(Me.pnluserrights)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmadmin"
        Me.Text = "Confer user priviledges"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnluserrights.ResumeLayout(False)
        Me.pnlcontrols.ResumeLayout(False)
        Me.pnlgrid.ResumeLayout(False)
        CType(Me.dtgusers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Public strclients, strleads, strjobs, strequip, strpersonnel, strval As String
    Private Sub frmadmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim Tasks As New taskclass
            Tasks.loadadmincontrols = False
            Dim Threada1 As New System.Threading.Thread( _
                AddressOf taskclass.admininvoke)
            Threada1.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.admin.Dispose(True)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtgusers_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgusers.MouseDown
        Try
            hti = dtgusers.HitTest(New Point(e.X, e.Y))
            If hti.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
                       AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
                       AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
                       AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
                       AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.None Then
                Try
                    Call populategridcontrols(hti)
                Catch er456 As Exception

                End Try

            End If
        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub populategridcontrols(ByVal hti As System.Windows.Forms.DataGrid.HitTestInfo)
        Try
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgusers.DataSource
            Try
                'myForms.admin.datagridtextBox.TextBox.Controls.Remove(myForms.admin.comboControl)
            Catch j As Exception

            End Try
            If ds.Tables(0).Rows(hti.Row).Item("mybool") = True Then
                Try
                    'myForms.admin.datagridtextBox.TextBox.Controls.Remove(myForms.admin.comboControl)
                Catch d As Exception
                End Try

            Else
                Try
                    myForms.admin.comboControl.SendToBack()
                    myForms.admin.datagridtextBox.TextBox.Controls.Add(myForms.admin.comboControl)
                    myForms.admin.comboControl.BringToFront()
                    myForms.admin.datagridtextBox.TextBox.BackColor = Color.White
                Catch d As Exception
                End Try

            End If
            If previousrow.ColumnNumber = 2 Then
                If ujobtitle = True Then
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("name") = Me.comboControl.Text
                    If Me.comboControl.SelectedIndex > -1 _
                    AndAlso ds.Tables(0).Rows(previousrow.RowNumber).Item("mybool") = False Then

                        'ds.Tables(0).Rows(previousrow.RowNumber).Item("isadded") = Me.comboid.Text
                    End If
                    ujobtitle = False
                End If
            End If
            If previousrow.ColumnNumber = 3 Then
                If utask = True Then
                    ds.Tables(0).Rows(previousrow.RowNumber).Item("password") = Me.txttask.Text
                    utask = False
                End If
            End If
            Try
                Dim myarray() As String
                strval = ds.Tables(0).Rows(hti.Row).Item("seclevel")
                myarray = strval.Split(":")
                strclients = myarray(0)
                strjobs = myarray(1)
                strleads = myarray(2)
                strequip = myarray(3)
                strpersonnel = myarray(4)
            Catch xc As Exception

            End Try
            Try
                Dim discontinuedColumn As Integer = 0
                Dim pt As Point = Me.dtgusers.PointToClient( _
                    Control.MousePosition)
                'Dim hti As DataGrid.HitTestInfo = _
                '    Me.dtgusers.HitTest(pt)
                Dim bmb As BindingManagerBase = _
                    Me.BindingContext(Me.dtgusers.DataSource, _
                    Me.dtgusers.DataMember)
                If hti.Row < bmb.Count _
                   AndAlso hti.Type = DataGrid.HitTestType.Cell _
                   AndAlso hti.Column = discontinuedColumn Then
                    Me.dtgusers(hti.Row, discontinuedColumn) = _
                       Not CBool(Me.dtgusers(hti.Row, _
                             discontinuedColumn))
                    'Dim ds As System.Data.DataSet = New System.Data.DataSet()
                    'ds = Me.dtgusers.DataSource
                    'MsgBox(ds.Tables(0).Rows(hti.Row).Item("Delete"))
                ElseIf hti.Row < bmb.Count _
                   AndAlso hti.Type = DataGrid.HitTestType.Cell Then
                    'strleads = "0,0,0,0,0,0"
                    'strclients = "0,0,0,0,0,0"
                    'strjobs = "0,0,0,0,0,0"
                    'strequip = "0,0,0,0,0,0"
                    'strpersonnel = "0,0,0,0,0,0"
                    Try
                        Dim myarray() As String
                        strval = ds.Tables(0).Rows(hti.Row).Item("seclevel")
                        myarray = strval.Split(":")
                        strclients = myarray(0)
                        strjobs = myarray(1)
                        strleads = myarray(2)
                        strequip = myarray(3)
                        strpersonnel = myarray(4)
                    Catch xc As Exception

                    End Try
                    If hti.Column = 4 Then
                        Me.dtgusers(hti.Row, 4) = _
                                          Not CBool(Me.dtgusers(hti.Row, _
                                               4))
                        If Me.dtgusers(hti.Row, 4) = True Then
                            strleads = "1,0,0,0,0,0"
                            Dim cbn As New frmadministrator
                            cbn.strin = strleads
                            cbn.ShowDialog()
                        Else
                            strleads = "0,0,0,0,0,0"
                        End If
                        ds.Tables(0).Rows(hti.Row).Item("uupdate") = True
                        strval = strclients & ":" & strjobs & ":" & strleads & ":" & strequip & ":" & strpersonnel
                        ds.Tables(0).Rows(hti.Row).Item("seclevel") = strval
                    ElseIf hti.Column = 5 Then
                        Me.dtgusers(hti.Row, 5) = _
                                          Not CBool(Me.dtgusers(hti.Row, _
                                                5))
                        If Me.dtgusers(hti.Row, 5) = True Then
                            strclients = "1,0,0,0,0,0"
                            Dim cbn As New frmadminclients
                            cbn.strin = strclients
                            cbn.ShowDialog()
                        Else
                            strclients = "0,0,0,0,0,0"
                        End If
                        ds.Tables(0).Rows(hti.Row).Item("uupdate") = True
                        strval = strclients & ":" & strjobs & ":" & strleads & ":" & strequip & ":" & strpersonnel
                        ds.Tables(0).Rows(hti.Row).Item("seclevel") = strval
                    ElseIf hti.Column = 6 Then
                        Me.dtgusers(hti.Row, 6) = _
                                          Not CBool(Me.dtgusers(hti.Row, _
                                                6))
                        If Me.dtgusers(hti.Row, 6) = True Then
                            strjobs = "1,0,0,0,0,0"
                            Dim cbn As New frmadminjobs
                            cbn.strin = strjobs
                            cbn.ShowDialog()
                        Else
                            strjobs = "0,0,0,0,0,0"
                        End If
                        ds.Tables(0).Rows(hti.Row).Item("uupdate") = True '----------for update
                        strval = strclients & ":" & strjobs & ":" & strleads & ":" & strequip & ":" & strpersonnel
                        ds.Tables(0).Rows(hti.Row).Item("seclevel") = strval
                    ElseIf hti.Column = 7 Then
                        Me.dtgusers(hti.Row, 7) = _
                                          Not CBool(Me.dtgusers(hti.Row, _
                                               7))
                        If Me.dtgusers(hti.Row, 7) = True Then
                            strequip = "1,0,0,0,0,0"
                            Dim cbn As New frmadminjobs
                            cbn.strin = strequip
                            cbn.ShowDialog()

                        Else
                            strequip = "0,0,0,0,0,0"
                        End If
                        ds.Tables(0).Rows(hti.Row).Item("uupdate") = True
                        strval = strclients & ":" & strjobs & ":" & strleads & ":" & strequip & ":" & strpersonnel
                        ds.Tables(0).Rows(hti.Row).Item("seclevel") = strval
                    ElseIf hti.Column = 8 Then
                        Me.dtgusers(hti.Row, 8) = _
                                                          Not CBool(Me.dtgusers(hti.Row, _
                                                               8))
                        If Me.dtgusers(hti.Row, 8) = True Then
                            strpersonnel = "1,0,0,0,0,0"
                            Dim cbn As New frmadminpersonnel
                            cbn.strin = strpersonnel
                            cbn.ShowDialog()
                        Else
                            strpersonnel = "0,0,0,0,0,0"
                        End If
                        ds.Tables(0).Rows(hti.Row).Item("uupdate") = True
                        strval = strclients & ":" & strjobs & ":" & strleads & ":" & strequip & ":" & strpersonnel
                        ds.Tables(0).Rows(hti.Row).Item("seclevel") = strval
                    End If
                End If

            Catch ex As Exception

            End Try
            If hti.Column = 2 Then
                Try
                    Me.comboControl.Text = ds.Tables(0).Rows(hti.Row).Item("name")
                    ds.Tables(0).Rows(hti.Row).Item("uupdate") = True
                Catch cf As Exception
                End Try
                ujobtitle = True
            End If
            If hti.Column = 3 Then
                Try
                    Me.txttask.Text = ds.Tables(0).Rows(hti.Row).Item("password")
                    ds.Tables(0).Rows(hti.Row).Item("uupdate") = True
                Catch cf As Exception
                End Try
                utask = True
            End If

        Catch ex As Exception
        Finally
            Me.previousrow = dtgusers.CurrentCell
        End Try
    End Sub
    Private Sub btndeleteuser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btndeleteuser.Click
        Dim currentCursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgusers.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, str As String
            Dim myrow As System.Data.DataRow
            While y < i

                If ds.Tables(0).Rows(y).Item("Delete") = True Then
                    sid = ds.Tables(0).Rows(y).Item("ano")
                    str = "delete from seccheck where id_no='" & sid & "'"
                    Try
                        connect.BeginTrans()
                        connect.Execute(str)
                        connect.CommitTrans()
                    Catch cv As Exception
                    End Try
                    Try
                        myrow = ds.Tables(0).Rows(y)
                        ds.Tables(0).Rows.Remove(myrow)
                        y = y - 1
                        i = ds.Tables(0).Rows.Count
                    Catch g As Exception
                    End Try
                End If
                y += 1
            End While
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception

        Finally
            Cursor.Current = currentCursor
        End Try
        Try
            Dim Tasks As New taskclass
            Dim Threadr1 As New System.Threading.Thread( _
                AddressOf taskclass.comboinvoke)
            Threadr1.Start()
        Catch er As Exception

        End Try
    End Sub
    Private Sub btnedituserdetails_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnedituserdetails.Click
        Dim currentCursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgusers.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, strsec, str As String
            Dim isexist As Boolean = False
            For y = 0 To i - 1
                If Convert.IsDBNull(ds.Tables(0).Rows(y).Item("seclevel")) = True Then
                    strsec = "0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0"
                Else
                    strsec = ds.Tables(0).Rows(y).Item("seclevel")
                End If

                If Convert.ToString(ds.Tables(0).Rows(y).Item("ano")).Length > 0 Then
                    If ds.Tables(0).Rows(y).Item("uupdate") = True Then
                        str = "update seccheck set name='" & ds.Tables(0).Rows(y).Item("name") & "'" _
                                           & " ,password='" & ds.Tables(0).Rows(y).Item("password") & "',seclevel='" & strsec & "',"
                        str += "id_no='" & ds.Tables(0).Rows(y).Item("id2") & "' where id_no='" & ds.Tables(0).Rows(y).Item("id2") & "'"

                        Try
                            connect.BeginTrans()
                            connect.Execute(str)
                            connect.CommitTrans()
                            ds.Tables(0).Rows(y).Item("mybool") = True
                            ds.Tables(0).Rows(y).Item("uupdate") = False
                        Catch cv As Exception
                            Try
                                connect.RollbackTrans()
                            Catch n As Exception

                            End Try
                        End Try
                    End If

                Else
                    'And(Convert.ToString(ds.Tables(0).Rows(y).Item("password")).Trim.Length > 0)
                    If Convert.ToString(ds.Tables(0).Rows(y).Item("name")).Trim.Length > 0 _
                       Then
                        str = "insert into seccheck ( name,password,seclevel, id_no) values (  '" & ds.Tables(0).Rows(y).Item("name") & "'" _
                        & ",  '" & ds.Tables(0).Rows(y).Item("password") & "','" & strsec & "'"
                        str += ",'" & ds.Tables(0).Rows(y).Item("isadded") & "')"
                        '-------------check if id number already exists
                        Dim rs As New ADODB.Recordset
                        Dim Str2 As String = " select id_no from seccheck" _
                        & " where lower(id_no) like '" & Convert.ToString(ds.Tables(0).Rows(y).Item("id2")).ToLower & "'"
                        With rs
                            .CursorLocation = CursorLocationEnum.adUseClient
                            .CursorType = CursorTypeEnum.adOpenStatic
                            .Open(Str2, connect)
                            If .BOF = False And .EOF = False Then
                                isexist = True
                            Else
                                isexist = False
                            End If
                            .Close()
                        End With
                        '--------------end of sanity check
                        Try
                            If isexist = False Then
                                connect.BeginTrans()
                                connect.Execute(str)
                                connect.CommitTrans()
                                ds.Tables(0).Rows(y).Item("mybool") = True
                            End If

                        Catch cv As Exception
                            Try
                                connect.RollbackTrans()
                            Catch n As Exception

                            End Try
                        End Try
                    End If
                End If
                Application.DoEvents()
            Next
            Dim Tasks As New taskclass
            Dim Threaday As New System.Threading.Thread( _
                AddressOf taskclass.comboinvoke)
            Threaday.Start()
            MessageBox.Show("User rights have been conferred successfully", _
            "User rights", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception

        Finally
            Cursor.Current = currentCursor
        End Try
    End Sub
    Private Sub btnaddnewuser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnaddnewuser.Click
        Try
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgusers.DataSource
            Dim f As Integer = ds.Tables(0).Rows.Count
            Dim myrow As System.Data.DataRow = ds.Tables(0).NewRow
            ds.Tables(0).Rows.Add(myrow)
            ds.Tables(0).Rows(f).Item("Delete") = False
            ds.Tables(0).Rows(f).Item("Clients") = False
            ds.Tables(0).Rows(f).Item("Leads") = False
            ds.Tables(0).Rows(f).Item("Jobs") = False
            ds.Tables(0).Rows(f).Item("Equipment") = False
            ds.Tables(0).Rows(f).Item("Personnel") = False
            ds.Tables(0).Rows(f).Item("mybool") = False
            ds.Tables(0).Rows(f).Item("seclevel") = "0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0:0,0,0,0,0,0"
            ds.AcceptChanges()
        Catch es As Exception
        End Try
    End Sub
    Private Sub dtgusers_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgusers.CurrentCellChanged
        Try

        Catch ex As Exception

        Finally
            previousrow = Me.dtgusers.CurrentCell
        End Try
    End Sub
    Private Sub comboControl_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles comboControl.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.comboControl.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.comboid.SelectedIndex = indexx
            Dim strp
            strp = comboid.Text

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub btnshowall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnshowall.Click
        Try
            Dim Tasks As New taskclass
            Tasks.loadadmincontrols = True
            Dim Threada2 As New System.Threading.Thread( _
                AddressOf taskclass.admininvoke)
            Threada2.Start()
        Catch ew As Exception

        End Try
    End Sub
End Class
