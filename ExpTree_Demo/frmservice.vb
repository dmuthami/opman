
Imports System
Imports System.Data
Imports System.Threading
Imports ADODB


Public Class frmservice
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
            myForms.ismaintenace = False
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
    Friend WithEvents pnlmaintenance As System.Windows.Forms.Panel
    Friend WithEvents grpcontrols As System.Windows.Forms.GroupBox
    Friend WithEvents btnaddsave As System.Windows.Forms.Button
    Friend WithEvents btndelete As System.Windows.Forms.Button
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents grpequiphistory As System.Windows.Forms.GroupBox
    Friend WithEvents grpequipdetails As System.Windows.Forms.GroupBox
    Friend WithEvents cboequip As System.Windows.Forms.ComboBox
    Friend WithEvents dtgequipdetails As System.Windows.Forms.DataGrid
    Friend WithEvents dtgequipmaintenace As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmservice))
        Me.pnlmaintenance = New System.Windows.Forms.Panel
        Me.grpcontrols = New System.Windows.Forms.GroupBox
        Me.btnaddsave = New System.Windows.Forms.Button
        Me.btndelete = New System.Windows.Forms.Button
        Me.btnclose = New System.Windows.Forms.Button
        Me.grpequipdetails = New System.Windows.Forms.GroupBox
        Me.cboequip = New System.Windows.Forms.ComboBox
        Me.dtgequipdetails = New System.Windows.Forms.DataGrid
        Me.grpequiphistory = New System.Windows.Forms.GroupBox
        Me.dtgequipmaintenace = New System.Windows.Forms.DataGrid
        Me.pnlmaintenance.SuspendLayout()
        Me.grpcontrols.SuspendLayout()
        Me.grpequipdetails.SuspendLayout()
        CType(Me.dtgequipdetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpequiphistory.SuspendLayout()
        CType(Me.dtgequipmaintenace, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlmaintenance
        '
        Me.pnlmaintenance.Controls.Add(Me.grpcontrols)
        Me.pnlmaintenance.Controls.Add(Me.grpequipdetails)
        Me.pnlmaintenance.Controls.Add(Me.grpequiphistory)
        Me.pnlmaintenance.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlmaintenance.Location = New System.Drawing.Point(0, 0)
        Me.pnlmaintenance.Name = "pnlmaintenance"
        Me.pnlmaintenance.Size = New System.Drawing.Size(424, 490)
        Me.pnlmaintenance.TabIndex = 0
        '
        'grpcontrols
        '
        Me.grpcontrols.Controls.Add(Me.btnaddsave)
        Me.grpcontrols.Controls.Add(Me.btndelete)
        Me.grpcontrols.Controls.Add(Me.btnclose)
        Me.grpcontrols.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grpcontrols.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpcontrols.Location = New System.Drawing.Point(0, 450)
        Me.grpcontrols.Name = "grpcontrols"
        Me.grpcontrols.Size = New System.Drawing.Size(424, 40)
        Me.grpcontrols.TabIndex = 8
        Me.grpcontrols.TabStop = False
        '
        'btnaddsave
        '
        Me.btnaddsave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnaddsave.Location = New System.Drawing.Point(8, 14)
        Me.btnaddsave.Name = "btnaddsave"
        Me.btnaddsave.Size = New System.Drawing.Size(192, 23)
        Me.btnaddsave.TabIndex = 5
        Me.btnaddsave.Text = "Add  Maintenance information"
        '
        'btndelete
        '
        Me.btndelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelete.Location = New System.Drawing.Point(200, 14)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(128, 23)
        Me.btndelete.TabIndex = 6
        Me.btndelete.Text = "Delete Current Row"
        '
        'btnclose
        '
        Me.btnclose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnclose.Location = New System.Drawing.Point(344, 22)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.TabIndex = 7
        Me.btnclose.Text = "Close"
        '
        'grpequipdetails
        '
        Me.grpequipdetails.Controls.Add(Me.cboequip)
        Me.grpequipdetails.Controls.Add(Me.dtgequipdetails)
        Me.grpequipdetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpequipdetails.Location = New System.Drawing.Point(0, 0)
        Me.grpequipdetails.Name = "grpequipdetails"
        Me.grpequipdetails.Size = New System.Drawing.Size(424, 176)
        Me.grpequipdetails.TabIndex = 0
        Me.grpequipdetails.TabStop = False
        '
        'cboequip
        '
        Me.cboequip.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboequip.Location = New System.Drawing.Point(8, 12)
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
        Me.dtgequipdetails.Location = New System.Drawing.Point(8, 50)
        Me.dtgequipdetails.Name = "dtgequipdetails"
        Me.dtgequipdetails.ReadOnly = True
        Me.dtgequipdetails.Size = New System.Drawing.Size(408, 118)
        Me.dtgequipdetails.TabIndex = 2
        '
        'grpequiphistory
        '
        Me.grpequiphistory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpequiphistory.Controls.Add(Me.dtgequipmaintenace)
        Me.grpequiphistory.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpequiphistory.Location = New System.Drawing.Point(0, 176)
        Me.grpequiphistory.Name = "grpequiphistory"
        Me.grpequiphistory.Size = New System.Drawing.Size(424, 280)
        Me.grpequiphistory.TabIndex = 3
        Me.grpequiphistory.TabStop = False
        '
        'dtgequipmaintenace
        '
        Me.dtgequipmaintenace.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgequipmaintenace.CaptionText = "Maintenace information"
        Me.dtgequipmaintenace.DataMember = ""
        Me.dtgequipmaintenace.FlatMode = True
        Me.dtgequipmaintenace.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgequipmaintenace.Location = New System.Drawing.Point(8, 26)
        Me.dtgequipmaintenace.Name = "dtgequipmaintenace"
        Me.dtgequipmaintenace.ReadOnly = True
        Me.dtgequipmaintenace.Size = New System.Drawing.Size(408, 246)
        Me.dtgequipmaintenace.TabIndex = 4
        '
        'frmservice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(424, 490)
        Me.Controls.Add(Me.pnlmaintenance)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmservice"
        Me.Text = "Maintenance information"
        Me.pnlmaintenance.ResumeLayout(False)
        Me.grpcontrols.ResumeLayout(False)
        Me.grpequipdetails.ResumeLayout(False)
        CType(Me.dtgequipdetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpequiphistory.ResumeLayout(False)
        CType(Me.dtgequipmaintenace, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
            Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private threadmaintenace As System.Threading.Thread
    Private hti As DataGrid.HitTestInfo
#Region "maintenace"
    Private Sub frmservice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim Tasks As New taskclass
            Dim Threadserv As New System.Threading.Thread( _
                AddressOf Tasks.servcboinvoke)
            Threadserv.Start()
        Catch we As Exception
        End Try
    End Sub
    Private Sub cboequip_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboequip.SelectedValueChanged
        Try
            Dim a() As String
            a = cboequip.Text.Split(":")

            Dim Tasks As New taskclass
            Tasks.servno = a(0)
            Try
                If threadmaintenace Is Nothing = False Then
                    Try
                        threadmaintenace.Abort()
                    Catch we As Exception
                    End Try
                End If
                threadmaintenace = New System.Threading.Thread( _
                AddressOf Tasks.servequipdetailsinvoke)
                threadmaintenace.Start()
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
            myForms.ismaintenace = False
            myForms.maintenace.Dispose(True)
        Catch es As Exception

        End Try
    End Sub
    Private Sub btnaddsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddsave.Click
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
            myForms.isadddmaintenace = False
            If myForms.isadddmaintenace = False Then
                Dim nb As New frmaddmaintenanceinfo
                myForms.addmaintenace = nb
                myForms.addmaintenace.cboequipid.Text = myForms.maintenace.cboequip.Text
                myForms.addmaintenace.ShowDialog()
                myForms.isadddmaintenace = True
            End If
        Catch fd As Exception

        End Try
    End Sub
    Private Sub dtgequipmaintenace_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgequipmaintenace.DoubleClick
        Try

            If hti.Type = DataGrid.HitTestType.Cell _
            Or hti.Type = DataGrid.HitTestType.RowHeader Then
                'MsgBox("cell")
                Dim ds As System.Data.DataSet = New System.Data.DataSet
                ds = Me.dtgequipmaintenace.DataSource
                If myForms.iseditmaintenace = False Then
                    Dim vc2 As New frmeditmaintenanceinfo
                    myForms.editmaintenace = vc2
                    '----------set properties and methods
                    myForms.editmaintenace.txtequipid.Text = ds.Tables(0).Rows(hti.Row).Item("equip_id")
                    myForms.editmaintenace.txtdesc.Text = ds.Tables(0).Rows(hti.Row).Item("description")
                    myForms.editmaintenace.txtcost.Text = ds.Tables(0).Rows(hti.Row).Item("cost_incurred")
                    myForms.editmaintenace.txtinvoiceno.Text = ds.Tables(0).Rows(hti.Row).Item("invoice_no")
                    myForms.editmaintenace.ano = ds.Tables(0).Rows(hti.Row).Item("autonumber")
                    Try
                        myForms.editmaintenace.dtpservicedate.Value = _
                        CDate(Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("service_date")))
                    Catch we As Exception
                    End Try


                    myForms.editmaintenace.Show()
                    myForms.iseditequip = True
                Else
                    myForms.editmaintenace.txtequipid.Text = ds.Tables(0).Rows(hti.Row).Item("equip_id")
                    myForms.editmaintenace.txtdesc.Text = ds.Tables(0).Rows(hti.Row).Item("description")
                    myForms.editmaintenace.txtcost.Text = ds.Tables(0).Rows(hti.Row).Item("cost_incurred")
                    myForms.editmaintenace.txtinvoiceno.Text = ds.Tables(0).Rows(hti.Row).Item("invoice_no")
                    myForms.editmaintenace.ano = ds.Tables(0).Rows(hti.Row).Item("autonumber")
                    myForms.iseditequip = True
                    Try
                        myForms.editmaintenace.dtpservicedate.Value = _
                     CDate(Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("service_date")))
                    Catch we As Exception
                    End Try
                End If

            End If

        Catch we As Exception

        End Try
    End Sub
    Private Sub dtgequipmaintenace_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgequipmaintenace.Click

    End Sub
    Private Sub dtgequipmaintenace_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgequipmaintenace.MouseDown
        Try
            'hti = dtginventories.HitTest(New Point(e.X, e.Y))
            Dim pt As Point = Me.dtgequipmaintenace.PointToClient( _
               Control.MousePosition)
            hti = _
                Me.dtgequipmaintenace.HitTest(pt)

        Catch ex As Exception
            Try
            Catch er As Exception
            End Try
        Finally
        End Try
    End Sub
#End Region

    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
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
            dtgequipmaintenace.Select(hti.Row)
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
            ds = Me.dtgequipmaintenace.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hti.Row).Item("autonumber")
            str = "delete from maintenance_info where autonumber='" & sid & "'"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(hti.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
            End Try

            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
End Class
