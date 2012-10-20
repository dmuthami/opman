
Imports System
Imports System.Data
Imports System.Threading
Imports ADODB


Public Class frmhistory
    Inherits System.Windows.Forms.Form
    '--------------end of local and global variables
    Public WithEvents comboassignedby As System.Windows.Forms.ComboBox 'assigned by
    Public WithEvents combojobno As System.Windows.Forms.ComboBox 'job no
    '--------------
    Public WithEvents txttask As System.Windows.Forms.TextBox 'task
    Public WithEvents rtbdesc As System.Windows.Forms.RichTextBox 'description
    Public WithEvents dtperd As System.Windows.Forms.DateTimePicker 'estimated release date
    Public WithEvents dtpdas As System.Windows.Forms.DateTimePicker 'date assigned
    Public WithEvents dtpdre As System.Windows.Forms.DateTimePicker 'date released

    Public WithEvents datagridtextBox As DataGridTextBoxColumn
    Public WithEvents datagridtextBox1 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox2 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox3 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox4 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox5 As DataGridTextBoxColumn
    Public WithEvents datagridtextBox6 As DataGridTextBoxColumn
    '------------------other controls okay

    Private previouscell As New DataGridCell()
    Private currcell As New DataGridCell()
    Private hti As DataGrid.HitTestInfo


    Private udesc As Boolean = False
    Private utask As Boolean = False
    Private ujno As Boolean = False
    Private udas As Boolean = False
    Private udres As Boolean = False
    Private uderd As Boolean = False
    Private uaby As Boolean = False 'assigned by

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
            myForms.ishistory = False
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
    Friend WithEvents grpcontrols As System.Windows.Forms.GroupBox
    Friend WithEvents grpequiphistory As System.Windows.Forms.GroupBox
    Friend WithEvents dtgequiphistory As System.Windows.Forms.DataGrid
    Friend WithEvents grpequipdetails As System.Windows.Forms.GroupBox
    Friend WithEvents cboequip As System.Windows.Forms.ComboBox
    Friend WithEvents dtgequipdetails As System.Windows.Forms.DataGrid
    Friend WithEvents btnsave As System.Windows.Forms.Button
    Friend WithEvents btndelete As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmhistory))
        Me.pnltojobs = New System.Windows.Forms.Panel
        Me.grpcontrols = New System.Windows.Forms.GroupBox
        Me.btnsave = New System.Windows.Forms.Button
        Me.btndelete = New System.Windows.Forms.Button
        Me.btnclose = New System.Windows.Forms.Button
        Me.grpequiphistory = New System.Windows.Forms.GroupBox
        Me.dtgequiphistory = New System.Windows.Forms.DataGrid
        Me.grpequipdetails = New System.Windows.Forms.GroupBox
        Me.cboequip = New System.Windows.Forms.ComboBox
        Me.dtgequipdetails = New System.Windows.Forms.DataGrid
        Me.pnltojobs.SuspendLayout()
        Me.grpcontrols.SuspendLayout()
        Me.grpequiphistory.SuspendLayout()
        CType(Me.dtgequiphistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpequipdetails.SuspendLayout()
        CType(Me.dtgequipdetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnltojobs
        '
        Me.pnltojobs.AutoScroll = True
        Me.pnltojobs.Controls.Add(Me.grpcontrols)
        Me.pnltojobs.Controls.Add(Me.grpequiphistory)
        Me.pnltojobs.Controls.Add(Me.grpequipdetails)
        Me.pnltojobs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnltojobs.Location = New System.Drawing.Point(0, 0)
        Me.pnltojobs.Name = "pnltojobs"
        Me.pnltojobs.Size = New System.Drawing.Size(424, 490)
        Me.pnltojobs.TabIndex = 0
        '
        'grpcontrols
        '
        Me.grpcontrols.Controls.Add(Me.btnsave)
        Me.grpcontrols.Controls.Add(Me.btndelete)
        Me.grpcontrols.Controls.Add(Me.btnclose)
        Me.grpcontrols.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grpcontrols.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpcontrols.Location = New System.Drawing.Point(0, 450)
        Me.grpcontrols.Name = "grpcontrols"
        Me.grpcontrols.Size = New System.Drawing.Size(424, 40)
        Me.grpcontrols.TabIndex = 5
        Me.grpcontrols.TabStop = False
        '
        'btnsave
        '
        Me.btnsave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsave.Location = New System.Drawing.Point(8, 11)
        Me.btnsave.Name = "btnsave"
        Me.btnsave.Size = New System.Drawing.Size(104, 23)
        Me.btnsave.TabIndex = 5
        Me.btnsave.Text = "Save Changes"
        '
        'btndelete
        '
        Me.btndelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelete.Location = New System.Drawing.Point(116, 11)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(129, 23)
        Me.btndelete.TabIndex = 6
        Me.btndelete.Text = "Delete Current Row"
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(346, 20)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 7
        Me.btnclose.Text = "Close"
        '
        'grpequiphistory
        '
        Me.grpequiphistory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpequiphistory.Controls.Add(Me.dtgequiphistory)
        Me.grpequiphistory.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpequiphistory.Location = New System.Drawing.Point(0, 136)
        Me.grpequiphistory.Name = "grpequiphistory"
        Me.grpequiphistory.Size = New System.Drawing.Size(424, 312)
        Me.grpequiphistory.TabIndex = 3
        Me.grpequiphistory.TabStop = False
        '
        'dtgequiphistory
        '
        Me.dtgequiphistory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgequiphistory.CaptionText = "Equipment history"
        Me.dtgequiphistory.DataMember = ""
        Me.dtgequiphistory.FlatMode = True
        Me.dtgequiphistory.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgequiphistory.Location = New System.Drawing.Point(8, 34)
        Me.dtgequiphistory.Name = "dtgequiphistory"
        Me.dtgequiphistory.ReadOnly = True
        Me.dtgequiphistory.Size = New System.Drawing.Size(408, 270)
        Me.dtgequiphistory.TabIndex = 4
        '
        'grpequipdetails
        '
        Me.grpequipdetails.Controls.Add(Me.cboequip)
        Me.grpequipdetails.Controls.Add(Me.dtgequipdetails)
        Me.grpequipdetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpequipdetails.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpequipdetails.Location = New System.Drawing.Point(0, 0)
        Me.grpequipdetails.Name = "grpequipdetails"
        Me.grpequipdetails.Size = New System.Drawing.Size(424, 144)
        Me.grpequipdetails.TabIndex = 0
        Me.grpequipdetails.TabStop = False
        '
        'cboequip
        '
        Me.cboequip.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboequip.Location = New System.Drawing.Point(8, 16)
        Me.cboequip.Name = "cboequip"
        Me.cboequip.Size = New System.Drawing.Size(224, 23)
        Me.cboequip.TabIndex = 1
        '
        'dtgequipdetails
        '
        Me.dtgequipdetails.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgequipdetails.CaptionText = "Equipment details"
        Me.dtgequipdetails.DataMember = ""
        Me.dtgequipdetails.FlatMode = True
        Me.dtgequipdetails.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgequipdetails.Location = New System.Drawing.Point(8, 53)
        Me.dtgequipdetails.Name = "dtgequipdetails"
        Me.dtgequipdetails.ReadOnly = True
        Me.dtgequipdetails.Size = New System.Drawing.Size(408, 83)
        Me.dtgequipdetails.TabIndex = 2
        '
        'frmhistory
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(424, 490)
        Me.Controls.Add(Me.pnltojobs)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmhistory"
        Me.Text = "History"
        Me.pnltojobs.ResumeLayout(False)
        Me.grpcontrols.ResumeLayout(False)
        Me.grpequiphistory.ResumeLayout(False)
        CType(Me.dtgequiphistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpequipdetails.ResumeLayout(False)
        CType(Me.dtgequipdetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private threadjobsequip As System.Threading.Thread

#Region " history"
    Private Sub frmhistory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim Tasks As New taskclass
            Dim Threadhisto As New System.Threading.Thread( _
                AddressOf Tasks.histcboinvoke)
            Threadhisto.Start()
        Catch we As Exception

        End Try
    End Sub
    Private Sub cboequip_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboequip.SelectedValueChanged
        Try
            Dim a() As String
            a = cboequip.Text.Split(":")

            Dim Tasks As New taskclass
            Tasks.historyjobno = a(0)
            Try
                If threadjobsequip Is Nothing = False Then
                    Try
                        threadjobsequip.Abort()
                    Catch we As Exception
                    End Try
                End If
                threadjobsequip = New System.Threading.Thread( _
                AddressOf Tasks.histequipdetailsinvoke)
                threadjobsequip.Start()
            Catch ev As Exception

            End Try


        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Try
            myForms.ishistory = False
            myForms.historry.Dispose(True)
        Catch es As Exception

        End Try
    End Sub
    Private Sub btnaddsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub
    Private Sub btnsave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xcv As Exception

        End Try
        Dim connectstr As String
        connectstr = "DSN=" & myForms.qconnstr
        'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = connectstr
        connect.Open()
        Dim ds As System.Data.DataSet = New System.Data.DataSet
        ds = Me.dtgequiphistory.DataSource
        Dim i As Integer = ds.Tables(0).Rows.Count
        Dim y As Integer
        Dim sid, myseconds, str, strsql As String
        Dim myrow As System.Data.DataRow
        For y = 0 To i - 1

            sid = ds.Tables(0).Rows(y).Item("ano")
            strsql += " update  history_equip set  "
            strsql += " equip_id ='" & ds.Tables(0).Rows(y).Item("equip_id") & "'," _
            & " job_no='" & ds.Tables(0).Rows(y).Item("job_no") & "',task='" & ds.Tables(0).Rows(y).Item("task") & "',"
            strsql += " other='" & "" & "',description='" & ds.Tables(0).Rows(y).Item("description") & "',assigned_by='" & ds.Tables(0).Rows(y).Item("assigned_by") & "'," _
            & "date_assigned='" & ds.Tables(0).Rows(y).Item("date_assigned") & "',date_released='" & ds.Tables(0).Rows(y).Item("date_released") & "'," _
            & " estimate_release_date='" & ds.Tables(0).Rows(y).Item("estimate_release_date") & "'"
            strsql += " where ano='" & sid & "';"


            Try
                connect.BeginTrans()
                connect.Execute(strsql)
                connect.Execute("commit work;")
                connect.CommitTrans()

                'ds.Tables(0).Rows(y).Item("Edit") = False
            Catch xc As Exception
                Try
                    connect.RollbackTrans()
                Catch er As Exception
                End Try
            End Try

        Next
        MessageBox.Show("Changes have been saved successfully", " Equipment history", _
        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Private Sub btndelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btndelete.Click
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
            dtgequiphistory.Select(hti.Row)
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
            ds = Me.dtgequiphistory.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hti.Row).Item("autonumber")
            str = "delete from history_equip where autonumber='" & sid & "'"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(hti.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try
            myrow = ds.Tables(0).Rows(hti.Row)
            ds.Tables(0).Rows.Remove(myrow)
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
    Private Sub dtgequiphistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgequiphistory.Click
    End Sub
    Private Sub dtgequiphistory_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgequiphistory.MouseDown
        Try
            'hti = dtginventories.HitTest(New Point(e.X, e.Y))
            Dim pt As Point = Me.dtgequiphistory.PointToClient( _
               Control.MousePosition)
            hti = _
                Me.dtgequiphistory.HitTest(pt)
            currcell.RowNumber = hti.Row
            currcell.ColumnNumber = hti.Column
            If hti.Column = 0 Then
                Me.dtgequiphistory(hti.Row, 0) = _
                                       Not CBool(Me.dtgequiphistory(hti.Row, _
                                             0))
            End If
            If hti.Type <> Windows.Forms.DataGrid.HitTestType.Caption _
                                 AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.ColumnHeader _
                                 AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.RowResize _
                                 AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.ColumnResize _
                                 AndAlso hti.Type <> Windows.Forms.DataGrid.HitTestType.None Then
                Try

                Catch er456 As Exception

                End Try

            End If
            Call vxvxvx()
        Catch ex As Exception
            Try
            Catch er As Exception
            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgequiphistory_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgequiphistory.CurrentCellChanged

    End Sub
    Private Sub vxvxvx()
        Try
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgequiphistory.DataSource
            Dim discontinuedColumn As Integer = 0
            Dim bmb As BindingManagerBase = _
                Me.BindingContext(Me.dtgequiphistory.DataSource, _
                Me.dtgequiphistory.DataMember)


            ' -----------updatecells
            'If previouscell.ColumnNumber = 1 Then
            '    If ujno = True Then
            '        Dim a() As String = Me.combojobno.Text.Split(";")
            '        ds.Tables(0).Rows(previouscell.RowNumber).Item("job_no") = a(0)
            '        ujno = False
            '    End If
            'End If
            If previouscell.ColumnNumber = 3 Then
                If utask = True Then
                    ds.Tables(0).Rows(previouscell.RowNumber).Item("task") = Me.txttask.Text
                    ds.Tables(0).AcceptChanges()
                    utask = False
                End If
            End If
            If previouscell.ColumnNumber = 4 Then
                If udesc = True Then
                    ds.Tables(0).Rows(previouscell.RowNumber).Item("description") = Me.rtbdesc.Text
                    ds.Tables(0).AcceptChanges()
                    udesc = False
                End If
            End If
            'If previouscell.ColumnNumber = 4 Then
            '    If udas = True Then
            '        ds.Tables(0).Rows(previouscell.RowNumber).Item("assigned_by") = Me.comboassignedby.Text
            '        udas = False
            '    End If
            'End If

            If previouscell.ColumnNumber = 6 Then
                If uderd = True Then
                    Dim sdate As String
                    sdate = dtperd.Value.Year & "-" _
                     & dtperd.Value.Month & "-" _
                     & dtperd.Value.Day & " " _
                     & "00" & ":" _
                     & "00" & ":" _
                     & "00"
                    ds.Tables(0).Rows(previouscell.RowNumber).Item("estimate_release_date") = sdate
                    ds.Tables(0).AcceptChanges()
                    uderd = False
                End If
            End If
            If previouscell.ColumnNumber = 7 Then
                If udas = True Then
                    Dim sdate As String
                    sdate = dtpdas.Value.Year & "-" _
                     & dtpdas.Value.Month & "-" _
                     & dtpdas.Value.Day & " " _
                     & "00" & ":" _
                     & "00" & ":" _
                     & "00"
                    ds.Tables(0).Rows(previouscell.RowNumber).Item("date_assigned") = sdate
                    ds.Tables(0).AcceptChanges()
                    udas = False
                End If
            End If
            If previouscell.ColumnNumber = 8 Then
                If udres = True Then
                    Dim sdate As String
                    sdate = dtpdre.Value.Year & "-" _
                     & dtpdre.Value.Month & "-" _
                     & dtpdre.Value.Day & " " _
                     & "00" & ":" _
                     & "00" & ":" _
                     & "00"
                    ds.Tables(0).Rows(previouscell.RowNumber).Item("date_released") = sdate
                    ds.Tables(0).AcceptChanges()
                    udres = False
                End If
            End If
            '-----------end of update




            '-------------load content of cell into control
            'If hti.Column = 1 Then
            '    Try
            '        Me.combojobno.Text = ds.Tables(0).Rows(hti.Row).Item("job_no")
            '    Catch cf As Exception
            '    End Try
            '    ujno = True
            'End If
            If hti.Column = 3 Then
                Try
                    Me.txttask.Text = ds.Tables(0).Rows(hti.Row).Item("task")
                Catch cf As Exception
                End Try
                utask = True
            End If
            If hti.Column = 4 Then
                Try
                    Me.rtbdesc.Text = ds.Tables(0).Rows(hti.Row).Item("description")
                Catch cf As Exception
                End Try
                udesc = True
            End If
            'If hti.Column = 4 Then
            '    Try
            '        Me.comboassignedby.Text = ds.Tables(0).Rows(hti.Row).Item("assigned_by")
            '    Catch cf As Exception
            '    End Try
            '    uaby = True
            'End If
            If hti.Column = 6 Then
                Try
                    Me.dtperd.Text = ds.Tables(0).Rows(hti.Row).Item("estimate_release_date")
                Catch cf As Exception
                End Try
                uderd = True
            End If
            If hti.Column = 7 Then
                Try
                    Me.dtpdas.Value = CDate(ds.Tables(0).Rows(hti.Row).Item("date_assigned"))
                Catch cf As Exception
                End Try
                udas = True
            End If
            If hti.Column = 8 Then
                Try
                    Me.dtpdre.Value = CDate(ds.Tables(0).Rows(hti.Row).Item("date_released"))
                Catch cf As Exception
                End Try
                udres = True
            End If



        Catch ex As Exception

        Finally
            previouscell = currcell
        End Try
    End Sub
#End Region


End Class
