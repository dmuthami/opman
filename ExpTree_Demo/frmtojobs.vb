
Imports System
Imports System.Data
Imports System.Threading
Imports ADODB


Public Class frmtojobs
    Inherits System.Windows.Forms.Form
    Public WithEvents comboControl As System.Windows.Forms.ComboBox
    Dim Threadjobs As System.Threading.Thread
    Private hti As DataGrid.HitTestInfo
    Public namme, jjobno, ano As String
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        comboControl = New System.Windows.Forms.ComboBox()
        comboControl.Cursor = System.Windows.Forms.Cursors.Arrow
        comboControl.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        comboControl.Dock = DockStyle.Fill
        comboControl.Visible = True
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            Try
                Call shut()
            Catch we As Exception
            End Try
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
    Friend WithEvents pnltojobs As System.Windows.Forms.Panel
    Friend WithEvents pnljob As System.Windows.Forms.Panel
    Friend WithEvents dtgjobs As System.Windows.Forms.DataGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlbottomcontrols As System.Windows.Forms.Panel
    Friend WithEvents dtgequip As System.Windows.Forms.DataGrid
    Friend WithEvents cbojob As System.Windows.Forms.ComboBox
    Friend WithEvents btnavailable As System.Windows.Forms.Button
    Friend WithEvents btnassigned As System.Windows.Forms.Button
    Friend WithEvents btnsave As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmtojobs))
        Me.pnltojobs = New System.Windows.Forms.Panel
        Me.btnassigned = New System.Windows.Forms.Button
        Me.btnavailable = New System.Windows.Forms.Button
        Me.dtgequip = New System.Windows.Forms.DataGrid
        Me.pnljob = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbojob = New System.Windows.Forms.ComboBox
        Me.dtgjobs = New System.Windows.Forms.DataGrid
        Me.pnlbottomcontrols = New System.Windows.Forms.Panel
        Me.btnclose = New System.Windows.Forms.Button
        Me.btnsave = New System.Windows.Forms.Button
        Me.pnltojobs.SuspendLayout()
        CType(Me.dtgequip, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnljob.SuspendLayout()
        CType(Me.dtgjobs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlbottomcontrols.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnltojobs
        '
        Me.pnltojobs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnltojobs.AutoScroll = True
        Me.pnltojobs.Controls.Add(Me.btnassigned)
        Me.pnltojobs.Controls.Add(Me.btnavailable)
        Me.pnltojobs.Controls.Add(Me.dtgequip)
        Me.pnltojobs.Location = New System.Drawing.Point(0, 144)
        Me.pnltojobs.Name = "pnltojobs"
        Me.pnltojobs.Size = New System.Drawing.Size(536, 344)
        Me.pnltojobs.TabIndex = 0
        '
        'btnassigned
        '
        Me.btnassigned.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnassigned.Location = New System.Drawing.Point(168, 3)
        Me.btnassigned.Name = "btnassigned"
        Me.btnassigned.Size = New System.Drawing.Size(168, 23)
        Me.btnassigned.TabIndex = 3
        Me.btnassigned.Text = "Show assigned equipment"
        '
        'btnavailable
        '
        Me.btnavailable.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnavailable.Location = New System.Drawing.Point(8, 2)
        Me.btnavailable.Name = "btnavailable"
        Me.btnavailable.Size = New System.Drawing.Size(160, 23)
        Me.btnavailable.TabIndex = 2
        Me.btnavailable.Text = "Show available equipment"
        '
        'dtgequip
        '
        Me.dtgequip.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgequip.CaptionText = "Available equipment"
        Me.dtgequip.DataMember = ""
        Me.dtgequip.FlatMode = True
        Me.dtgequip.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgequip.Location = New System.Drawing.Point(8, 24)
        Me.dtgequip.Name = "dtgequip"
        Me.dtgequip.ReadOnly = True
        Me.dtgequip.Size = New System.Drawing.Size(520, 320)
        Me.dtgequip.TabIndex = 4
        '
        'pnljob
        '
        Me.pnljob.Controls.Add(Me.Label1)
        Me.pnljob.Controls.Add(Me.cbojob)
        Me.pnljob.Controls.Add(Me.dtgjobs)
        Me.pnljob.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnljob.Location = New System.Drawing.Point(0, 0)
        Me.pnljob.Name = "pnljob"
        Me.pnljob.Size = New System.Drawing.Size(536, 144)
        Me.pnljob.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Pick a job"
        '
        'cbojob
        '
        Me.cbojob.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbojob.Location = New System.Drawing.Point(72, 8)
        Me.cbojob.Name = "cbojob"
        Me.cbojob.Size = New System.Drawing.Size(304, 22)
        Me.cbojob.TabIndex = 0
        '
        'dtgjobs
        '
        Me.dtgjobs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgjobs.CaptionText = "Job details"
        Me.dtgjobs.DataMember = ""
        Me.dtgjobs.FlatMode = True
        Me.dtgjobs.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgjobs.Location = New System.Drawing.Point(8, 32)
        Me.dtgjobs.Name = "dtgjobs"
        Me.dtgjobs.ReadOnly = True
        Me.dtgjobs.Size = New System.Drawing.Size(520, 112)
        Me.dtgjobs.TabIndex = 1
        '
        'pnlbottomcontrols
        '
        Me.pnlbottomcontrols.Controls.Add(Me.btnclose)
        Me.pnlbottomcontrols.Controls.Add(Me.btnsave)
        Me.pnlbottomcontrols.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlbottomcontrols.Location = New System.Drawing.Point(0, 490)
        Me.pnlbottomcontrols.Name = "pnlbottomcontrols"
        Me.pnlbottomcontrols.Size = New System.Drawing.Size(536, 24)
        Me.pnlbottomcontrols.TabIndex = 5
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(456, 0)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(75, 24)
        Me.btnclose.TabIndex = 7
        Me.btnclose.Text = "Close"
        '
        'btnsave
        '
        Me.btnsave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsave.Location = New System.Drawing.Point(8, 1)
        Me.btnsave.Name = "btnsave"
        Me.btnsave.Size = New System.Drawing.Size(96, 23)
        Me.btnsave.TabIndex = 6
        Me.btnsave.Text = "Save changes"
        '
        'frmtojobs
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(536, 514)
        Me.Controls.Add(Me.pnlbottomcontrols)
        Me.Controls.Add(Me.pnltojobs)
        Me.Controls.Add(Me.pnljob)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmtojobs"
        Me.Text = "To jobs"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnltojobs.ResumeLayout(False)
        CType(Me.dtgequip, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnljob.ResumeLayout(False)
        CType(Me.dtgjobs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlbottomcontrols.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    <System.STAThread()> _
            Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private Sub frmtojobs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Try
                Dim Tasks As New taskclass
                Dim Threadeb As New System.Threading.Thread( _
                    AddressOf Tasks.loadname)
                Threadeb.Start()
            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString() & vbCrLf _
                        & ex.InnerException().ToString() & vbCrLf _
                        & ex.StackTrace.ToString())
            End Try
            Call availableequip()
            Me.Invalidate(True)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub availableequip()
        Try
            Dim Tasks As New taskclass
            Dim Threade2 As New System.Threading.Thread( _
                AddressOf Tasks.equipjobinvoke)
            Threade2.Start()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
                    & ex.InnerException().ToString() & vbCrLf _
                    & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub shut()
        Try
            Dim task As taskclass
            task.hastojobloaded = False
            task.showavailableequip = True
        Catch we As Exception

        End Try
    End Sub
    Private Sub cbojob_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbojob.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cbojob.SelectedIndex
            If indexx = -1 Then
                Exit Try
            End If
            Me.comboControl.SelectedIndex = indexx
            Dim strp
            strp = comboControl.Text
            jjobno = strp
            Dim Tasks As New taskclass
            Tasks.jobno = strp
            Try
                If Threadjobs Is Nothing = False Then
                    Try
                        Threadjobs.Abort()
                    Catch we As Exception
                    End Try
                End If
                Threadjobs = New System.Threading.Thread( _
                AddressOf Tasks.jobinvoke)
                Threadjobs.Start()
            Catch ev As Exception

            End Try


        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub dtgequip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgequip.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to delete equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xcv As Exception

        End Try
        Try
            Dim discontinuedColumn As Integer = 0

            Dim bmb As BindingManagerBase = _
                Me.BindingContext(Me.dtgequip.DataSource, _
                Me.dtgequip.DataMember)
            If hti.Row < bmb.Count _
               AndAlso hti.Type = DataGrid.HitTestType.Cell _
               AndAlso hti.Column = discontinuedColumn Then
                '--------code to assign in the database
                Dim tasks As taskclass
                If tasks.showavailableequip = False Then
                    Dim myval As String
                    myval = Convert.ToBoolean(Me.dtgequip(hti.Row, discontinuedColumn))
                    If Convert.ToBoolean(myval) = True Then
                        If MessageBox.Show("Do you wish to de assign the equipments?", "De assign equipment", _
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            Exit Try
                        End If
                    End If
                    assignequip(myval)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet
                    ds = Me.dtgequip.DataSource
                    ds.Tables(0).Rows(hti.Row).Item("assigned_by") = namme
                    ds.Tables(0).Rows(hti.Row).Item("has_changed") = True
                    ds.Tables(0).AcceptChanges()
                    Call Me.btnsave_Click(Me, e)
                    MsgBox("Successfully  de-assigned")
                Else
                    Dim a() As String
                    a = Me.cbojob.Text.Split(":")
                    jjobno = a(0)
                    If jjobno.Length < 1 Then
                        MessageBox.Show("Please select a job", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Try
                    End If
                    Dim myval As String
                    myval = Convert.ToBoolean(Me.dtgequip(hti.Row, discontinuedColumn))
                    assignequip(myval)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet
                    ds = Me.dtgequip.DataSource
                    ds.Tables(0).Rows(hti.Row).Item("assigned_by") = namme
                    ds.Tables(0).Rows(hti.Row).Item("has_changed") = True
                    ds.Tables(0).AcceptChanges()

                    Me.dtgequip(hti.Row, discontinuedColumn) = _
                                      Not CBool(Me.dtgequip(hti.Row, _
                                            discontinuedColumn))
                    '---------refresh grosss margin
                    Try
                        myForms.CustomerForm2.mygross()
                    Catch zx As Exception

                    End Try
                    '----------------------
                End If

                '------------------end of that code


            ElseIf hti.Row < bmb.Count _
               AndAlso hti.Type = DataGrid.HitTestType.Cell Then

            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtgequip_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgequip.MouseDown
        Try
            'hti = dtginventories.HitTest(New Point(e.X, e.Y))
            Dim pt As Point = Me.dtgequip.PointToClient( _
               Control.MousePosition)
            hti = _
                Me.dtgequip.HitTest(pt)
        Catch ex As Exception
            Try
            Catch er As Exception
            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgequip_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgequip.DoubleClick
        Try

            If hti.Type = DataGrid.HitTestType.Cell _
            Or hti.Type = DataGrid.HitTestType.RowHeader Then
                'MsgBox("cell")
                Dim ds As System.Data.DataSet = New System.Data.DataSet
                ds = Me.dtgequip.DataSource
                If myForms.iseditassignequip = False Then
                    Dim vc As New frmeditequipassign
                    myForms.editassignequip = vc
                    myForms.editassignequip.StartPosition = FormStartPosition.CenterParent

                    '----------set properties and methods     
                    myForms.editassignequip.Show()
                    myForms.editassignequip.txtequipid.Text = Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("equip_id"))
                    myForms.editassignequip.txtmodelname.Text = Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("model_name"))
                    myForms.editassignequip.htirow = hti.Row
                    myForms.iseditassignequip = True
                    Try
                        myForms.editassignequip.dtppurchasedate.Value = _
                        CDate(Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("date_assigned")))
                    Catch we As Exception
                    End Try
                    Try
                        myForms.editassignequip.dtpreleasedate.Value = _
                        CDate(Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("estimate_release_date")))
                    Catch we As Exception
                    End Try



                Else

                    myForms.iseditequip = True
                    myForms.editassignequip.txtequipid.Text = Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("equip_id"))
                    myForms.editassignequip.txtmodelname.Text = Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("model_name"))
                    myForms.editassignequip.htirow = hti.Row
                    Try
                        myForms.editassignequip.dtppurchasedate.Value = _
                        CDate(Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("date_assigned")))
                    Catch we As Exception
                    End Try
                    Try
                        myForms.editassignequip.dtpreleasedate.Value = _
                        CDate(Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("estimate_release_date")))
                    Catch we As Exception
                    End Try
                End If

            End If

        Catch we As Exception

        End Try
    End Sub
    Private Sub assignequip(ByVal status As String)
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgequip.DataSource
            Dim f As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim strsql As String = "" 'IN SHARE MODE
            connect.BeginTrans()
            strsql = " BEGIN WORK;" _
                    & " LOCK TABLE assigned_info,tblno ;" _
               & ";"
            connect.Execute(strsql)
            Dim clientnumber1 As String
            Dim isnew As Boolean = False
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
                  & "WHERE equip_id ='" & ds.Tables(0).Rows(hti.Row).Item("equip_id") & "'" _
                  & "     "
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(strsql, connect)
                If .BOF = False And .EOF = False Then
                    If .Fields("status").Value = "1" Then
                        'MsgBox("Already assigned")
                        Dim sdate As String
                        Dim d As New System.Windows.Forms.DateTimePicker
                        d.Value = Now
                        sdate = d.Value.Year & "-" _
                         & d.Value.Month & "-" _
                         & d.Value.Day & " " _
                         & d.Value.Hour & ":" _
                         & d.Value.Minute & ":" _
                         & d.Value.Second
                        ds.Tables(0).Rows(hti.Row).Item("date_released") = sdate
                        strsql = " update assigned_info set status='" & "0" & "' where equip_id='" & ds.Tables(0).Rows(hti.Row).Item("equip_id") & "';"
                        strsql += " update current_equip set date_released='" & sdate & "' where equip_id='" & ds.Tables(0).Rows(hti.Row).Item("equip_id") & "';"
                        'strsql += "INSERT INTO history_equip SELECT * FROM current_equip WHERE equip_id='" & ds.Tables(0).Rows(hti.Row).Item("equip_id") & "';"
                        strsql += " commit work;"
                        connect.Execute(strsql)

                    Else
                        Dim sdate As String
                        Dim d As New System.Windows.Forms.DateTimePicker
                        d.Value = Now
                        sdate = d.Value.Year & "-" _
                         & d.Value.Month & "-" _
                         & d.Value.Day & " " _
                         & d.Value.Hour & ":" _
                         & d.Value.Minute & ":" _
                         & d.Value.Second
                        ds.Tables(0).Rows(hti.Row).Item("date_assigned") = sdate
                        strsql = " update assigned_info set status='" & "1" & "' where equip_id='" & ds.Tables(0).Rows(hti.Row).Item("equip_id") & "';"
                        strsql += " commit work;"
                        connect.Execute(strsql)
                        'ds.Tables(0).Rows(hti.Row).Item("Assign") = True
                        Dim Tasks As taskclass
                        If Tasks.showavailableequip = True Then
                            MsgBox("Successfully  assigned")
                        End If

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
    Private Sub btnavailable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnavailable.Click
        Try
            Me.btnsave.Enabled = True
            Dim Tasks As New taskclass
            Tasks.showavailableequip = True
            Dim Threadhh As New System.Threading.Thread( _
                AddressOf Tasks.equipjobinvoke)
            Threadhh.Start()
        Catch qw As Exception

        End Try
    End Sub
    Private Sub btnassigned_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnassigned.Click
        Try
            Me.btnsave.Enabled = False
            Dim Tasks As New taskclass
            Tasks.showavailableequip = False
            Dim Threadhh As New System.Threading.Thread( _
                AddressOf Tasks.equipjobinvoke)
            Threadhh.Start()
        Catch qw As Exception

        End Try
    End Sub
    Private Sub btnsave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Dim ccursor As Cursor = Cursor.Current
        Try
            Try
                Dim x As Boolean = myForms.Main.canmanipulateequip()
                If x = False Then
                    MessageBox.Show("Not allowed to delete equipment contact administrator", "Equipment", _
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
            Catch xcv As Exception

            End Try

            Cursor.Current = Cursors.IBeam
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim a() As String
            a = Me.cbojob.Text.Split(":")
            jjobno = a(0)
            'If jjobno.Length < 1 Then
            '    MessageBox.Show("Please select a job", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Try
            'End If

            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgequip.DataSource
            Dim f As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim strsql, equipid As String
            Dim dateassigned, datereleased, erd, fff As String
            For kappa = 0 To f - 1
                'Convert.ToString(ds.Tables(0).Rows(kappa).Item("job_tittle")).Length < 1
                If ds.Tables(0).Rows(kappa).Item("has_changed") = True Then
                    Dim dtpnow As New System.Windows.Forms.DateTimePicker
                    dtpnow.Value = Now

                    Dim strnow As String
                    strnow = dtpnow.Value.Year & "-" _
                                         & dtpnow.Value.Month & "-" _
                                         & dtpnow.Value.Day & " " _
                                         & dtpnow.Value.Hour & ":" _
                                         & dtpnow.Value.Minute & ":" _
                                         & dtpnow.Value.Second
                    If Convert.ToString(ds.Tables(0).Rows(kappa).Item("date_assigned")).Length < 1 Then
                        dateassigned = strnow
                    Else
                        dateassigned = Convert.ToString(ds.Tables(0).Rows(kappa).Item("date_assigned"))
                    End If
                    If Convert.ToString(ds.Tables(0).Rows(kappa).Item("estimate_release_date")).Length < 1 Then
                        erd = strnow
                    Else
                        erd = Convert.ToString(ds.Tables(0).Rows(kappa).Item("estimate_release_date"))
                    End If
                    If Convert.ToString(ds.Tables(0).Rows(kappa).Item("date_released")).Length < 1 Then
                        datereleased = strnow
                    Else
                        datereleased = Convert.ToString(ds.Tables(0).Rows(kappa).Item("date_released"))
                    End If
                    connect.BeginTrans()
                    strsql = " BEGIN WORK;" _
                      & " LOCK TABLE assigned_info,history_equip,current_equip ;" _
                 & ";"
                    connect.Execute(strsql)
                    equipid = ds.Tables(0).Rows(kappa).Item("equip_id")
                    fff = ds.Tables(0).Rows(kappa).Item("equip_id")
                    strsql = " INSERT INTO history_equip " _
                    & " select *  from current_equip where equip_id='" & fff & "'; "
                    strsql += " delete from current_equip where equip_id='" & equipid & "';"
                    strsql += " insert into  current_equip (equip_id,job_no,task,other,description,assigned_by,date_released,date_assigned," _
                    & " estimate_release_date,autonumber) values  " _
                    & " ("
                    strsql += "  '" & ds.Tables(0).Rows(kappa).Item("equip_id") & "'," _
                    & " '" & jjobno & "','" & "" & "',"
                    strsql += " '" & "" & "','" & "" & "','" & ds.Tables(0).Rows(kappa).Item("assigned_by") & "'," _
                    & "'" & datereleased & "','" & dateassigned & "','" & erd & "','" & ano & "');"
                    'strsql += " where equip_id='" & ds.Tables(0).Rows(kappa).Item("equip_id") & "';"
                    Try

                        connect.Execute(strsql)
                        connect.Execute("commit work;")
                        connect.CommitTrans()
                        ds.Tables(0).Rows(kappa).Item("has_changed") = False
                    Catch xc As Exception
                        Try
                            connect.RollbackTrans()
                        Catch er As Exception
                        End Try
                    End Try
                End If


                System.Windows.Forms.Application.DoEvents()
            Next

        Catch ex As Exception

        Finally

        End Try
        Try
            Dim Tasks As New taskclass
            Dim Threade2 As New System.Threading.Thread( _
                AddressOf Tasks.equipjobinvoke)
            Threade2.Start()
        Catch qw As Exception

        End Try
        Try
            Dim Tasks As New taskclass
            Tasks.erjobno = myForms.CustomerForm2.txtJobNo.Text
            Dim Threadec As New System.Threading.Thread( _
                AddressOf Tasks.ramaniequipinvoke)
            Threadec.Start()
        Catch az As Exception

        End Try
        Cursor.Current = ccursor
    End Sub
    Private Sub btnclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            Call shut()
            myForms.tojobs.Dispose(True)
        Catch we As Exception

        End Try
    End Sub
    Private Sub dtgequip_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtgequip.Navigate

    End Sub
End Class
