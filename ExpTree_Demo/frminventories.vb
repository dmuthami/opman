Imports System
Imports System.Data
Imports System.Threading
Imports ADODB


Public Class frminventories
    Inherits System.Windows.Forms.Form
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
    Private htid As DataGrid.HitTestInfo


    Public WithEvents btnequipactions As System.Windows.Forms.Button
    Public WithEvents datagridtextBox As DataGridTextBoxColumn
    Private Delegate Sub mydelegatee1()

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
    Friend WithEvents pnlinventories As System.Windows.Forms.Panel
    Friend WithEvents dtginventories As System.Windows.Forms.DataGrid
    Friend WithEvents pnlgrid As System.Windows.Forms.Panel
    Friend WithEvents pnltopcontrols As System.Windows.Forms.Panel
    Friend WithEvents btnshowall As System.Windows.Forms.Button
    Friend WithEvents btntojob As System.Windows.Forms.Button
    Friend WithEvents btnmaintenace As System.Windows.Forms.Button
    Friend WithEvents btnhistory As System.Windows.Forms.Button
    Friend WithEvents btndeleteequipment As System.Windows.Forms.Button
    Friend WithEvents btnaddequipments As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frminventories))
        Me.pnlinventories = New System.Windows.Forms.Panel
        Me.pnlgrid = New System.Windows.Forms.Panel
        Me.dtginventories = New System.Windows.Forms.DataGrid
        Me.pnltopcontrols = New System.Windows.Forms.Panel
        Me.btndeleteequipment = New System.Windows.Forms.Button
        Me.btnaddequipments = New System.Windows.Forms.Button
        Me.btnhistory = New System.Windows.Forms.Button
        Me.btnmaintenace = New System.Windows.Forms.Button
        Me.btntojob = New System.Windows.Forms.Button
        Me.btnshowall = New System.Windows.Forms.Button
        Me.pnlinventories.SuspendLayout()
        Me.pnlgrid.SuspendLayout()
        CType(Me.dtginventories, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnltopcontrols.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlinventories
        '
        Me.pnlinventories.Controls.Add(Me.pnlgrid)
        Me.pnlinventories.Controls.Add(Me.pnltopcontrols)
        Me.pnlinventories.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlinventories.Location = New System.Drawing.Point(0, 0)
        Me.pnlinventories.Name = "pnlinventories"
        Me.pnlinventories.Size = New System.Drawing.Size(496, 428)
        Me.pnlinventories.TabIndex = 0
        '
        'pnlgrid
        '
        Me.pnlgrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlgrid.AutoScroll = True
        Me.pnlgrid.Controls.Add(Me.dtginventories)
        Me.pnlgrid.Location = New System.Drawing.Point(0, 56)
        Me.pnlgrid.Name = "pnlgrid"
        Me.pnlgrid.Size = New System.Drawing.Size(496, 368)
        Me.pnlgrid.TabIndex = 7
        '
        'dtginventories
        '
        Me.dtginventories.AllowSorting = False
        Me.dtginventories.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtginventories.CaptionText = "Equipment"
        Me.dtginventories.DataMember = ""
        Me.dtginventories.FlatMode = True
        Me.dtginventories.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtginventories.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtginventories.Location = New System.Drawing.Point(8, 8)
        Me.dtginventories.Name = "dtginventories"
        Me.dtginventories.PreferredRowHeight = 35
        Me.dtginventories.ReadOnly = True
        Me.dtginventories.Size = New System.Drawing.Size(480, 352)
        Me.dtginventories.TabIndex = 8
        '
        'pnltopcontrols
        '
        Me.pnltopcontrols.Controls.Add(Me.btndeleteequipment)
        Me.pnltopcontrols.Controls.Add(Me.btnaddequipments)
        Me.pnltopcontrols.Controls.Add(Me.btnhistory)
        Me.pnltopcontrols.Controls.Add(Me.btnmaintenace)
        Me.pnltopcontrols.Controls.Add(Me.btntojob)
        Me.pnltopcontrols.Controls.Add(Me.btnshowall)
        Me.pnltopcontrols.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnltopcontrols.Location = New System.Drawing.Point(0, 0)
        Me.pnltopcontrols.Name = "pnltopcontrols"
        Me.pnltopcontrols.Size = New System.Drawing.Size(496, 56)
        Me.pnltopcontrols.TabIndex = 0
        '
        'btndeleteequipment
        '
        Me.btndeleteequipment.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeleteequipment.Location = New System.Drawing.Point(106, 32)
        Me.btndeleteequipment.Name = "btndeleteequipment"
        Me.btndeleteequipment.Size = New System.Drawing.Size(112, 23)
        Me.btndeleteequipment.TabIndex = 6
        Me.btndeleteequipment.Text = "Delete equipment"
        '
        'btnaddequipments
        '
        Me.btnaddequipments.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnaddequipments.Location = New System.Drawing.Point(8, 32)
        Me.btnaddequipments.Name = "btnaddequipments"
        Me.btnaddequipments.Size = New System.Drawing.Size(96, 23)
        Me.btnaddequipments.TabIndex = 5
        Me.btnaddequipments.Text = "Add equipment"
        '
        'btnhistory
        '
        Me.btnhistory.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnhistory.Location = New System.Drawing.Point(282, 4)
        Me.btnhistory.Name = "btnhistory"
        Me.btnhistory.Size = New System.Drawing.Size(104, 23)
        Me.btnhistory.TabIndex = 4
        Me.btnhistory.Text = "History report"
        '
        'btnmaintenace
        '
        Me.btnmaintenace.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnmaintenace.Location = New System.Drawing.Point(160, 4)
        Me.btnmaintenace.Name = "btnmaintenace"
        Me.btnmaintenace.Size = New System.Drawing.Size(120, 23)
        Me.btnmaintenace.TabIndex = 3
        Me.btnmaintenace.Text = "Maintenance Report"
        '
        'btntojob
        '
        Me.btntojob.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btntojob.Location = New System.Drawing.Point(83, 4)
        Me.btntojob.Name = "btntojob"
        Me.btntojob.TabIndex = 2
        Me.btntojob.Text = "To jobs"
        '
        'btnshowall
        '
        Me.btnshowall.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnshowall.Location = New System.Drawing.Point(8, 4)
        Me.btnshowall.Name = "btnshowall"
        Me.btnshowall.TabIndex = 1
        Me.btnshowall.Text = "Show all"
        '
        'frminventories
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(496, 428)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlinventories)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frminventories"
        Me.Text = "Equipments"
        Me.pnlinventories.ResumeLayout(False)
        Me.pnlgrid.ResumeLayout(False)
        CType(Me.dtginventories, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnltopcontrols.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
#Region "private members'"
    Private ccursor As Cursor = Cursor.Current
#End Region

#Region " inventories"
    Private Sub frminventories_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim Tasks As New taskclass
            Dim Threade1 As New System.Threading.Thread( _
                AddressOf Tasks.equipinvoke)
            Threade1.Start()

            '--------load combo boxe
            Dim Threadbv As New System.Threading.Thread( _
                AddressOf nini)
            Threadbv.Start()
        Catch we As Exception

        End Try
    End Sub
    Private Sub dtginventories_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtginventories.MouseDown
        Try
            'hti = dtginventories.HitTest(New Point(e.X, e.Y))
            Dim pt As Point = Me.dtginventories.PointToClient( _
                       Control.MousePosition)
            htid = _
                Me.dtginventories.HitTest(pt)
            If htid.Type = Windows.Forms.DataGrid.HitTestType.ColumnHeader Then
                Cursor.Current = Cursors.Hand
            End If
        Catch ex As Exception
            Try
            Catch er As Exception
            End Try
        Finally
        End Try
    End Sub
    Private Sub dtginventories_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtginventories.Click

    End Sub
    Private Sub dtginventories_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtginventories.DoubleClick
        Try

            If htid.Type = DataGrid.HitTestType.Cell _
            Or htid.Type = DataGrid.HitTestType.RowHeader Then
                'MsgBox("cell")
                Dim ds As System.Data.DataSet = New System.Data.DataSet
                ds = Me.dtginventories.DataSource
                If myForms.iseditequip = False Then
                    Dim vc As New frmeditequip
                    myForms.editequipments = vc
                    '----------set properties and methods
                    myForms.editequipments.myid = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("equip_id"))
                    myForms.editequipments.txtequipid.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("equip_id"))
                    myForms.editequipments.txtmanufacturer.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("manufacturer"))
                    myForms.editequipments.txtmodelno.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("model_no"))
                    myForms.editequipments.txtserialno.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("serial_no"))
                    myForms.editequipments.txtmodelname.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("model_name"))
                    myForms.editequipments.txtdesc.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("supplier"))
                    myForms.editequipments.txtlicense.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("license"))
                    myForms.editequipments.txtguarantee.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("guarantee"))
                    myForms.editequipments.txtcondition.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("condition"))
                    myForms.editequipments.txttype.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("type"))
                    myForms.editequipments.txtmodelyear.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("model_year"))
                    myForms.editequipments.txthourlyrate.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("hourly_rate"))

                    myForms.editequipments.txtsupplier.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("description"))
                    myForms.editequipments.txtmouse.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("mouse"))
                    myForms.editequipments.txtkeyboard.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("keyboard"))
                    myForms.editequipments.txtmonitor.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("monitor"))
                    myForms.editequipments.txtmonitor2.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("monitor2"))
                    myForms.editequipments.txtphone.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("phone"))
                    myForms.editequipments.txtamount.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("amount"))
                    myForms.editequipments.txtbatteries.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("batteries"))
                    myForms.editequipments.txtdownloadcables.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("downloadcables"))
                    myForms.editequipments.txtunit.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("unit"))
                    myForms.iseditequip = True
                    Try
                        myForms.editequipments.dtppurchasedate.Value = _
                        CDate(Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("purchase_date")))
                    Catch we As Exception
                    End Try


                    myForms.editequipments.Show()

                Else
                    myForms.editequipments.myid = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("equip_id"))
                    myForms.editequipments.txtequipid.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("equip_id"))
                    myForms.editequipments.txtmanufacturer.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("manufacturer"))
                    myForms.editequipments.txtmodelno.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("model_no"))
                    myForms.editequipments.txtserialno.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("serial_no"))
                    myForms.editequipments.txtmodelname.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("model_name"))
                    myForms.editequipments.txtdesc.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("supplier"))
                    myForms.editequipments.txtlicense.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("license"))
                    myForms.editequipments.txtguarantee.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("condition"))
                    myForms.editequipments.txtcondition.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("guarantee"))
                    myForms.editequipments.txttype.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("type"))
                    myForms.editequipments.txtmodelyear.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("model_year"))
                    myForms.editequipments.txthourlyrate.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("hourly_rate"))

                    myForms.editequipments.txtsupplier.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("description"))
                    myForms.editequipments.txtmouse.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("mouse"))
                    myForms.editequipments.txtkeyboard.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("keyboard"))
                    myForms.editequipments.txtmonitor.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("monitor"))
                    myForms.editequipments.txtmonitor2.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("monitor2"))
                    myForms.editequipments.txtphone.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("phone"))
                    myForms.editequipments.txtamount.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("amount"))

                    myForms.editequipments.txtbatteries.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("batteries"))
                    myForms.editequipments.txtdownloadcables.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("downloadcables"))
                    myForms.editequipments.txtunit.Text = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("unit"))
                    myForms.iseditequip = True
                    myForms.iseditequip = True
                    Try
                        myForms.editequipments.dtppurchasedate.Value = _
                        CDate(Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("purchase_date")))
                    Catch we As Exception
                    End Try
                End If

            End If

        Catch we As Exception

        End Try
    End Sub
    Private Sub btnequipactions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnequipactions.Click
        Try
            If myForms.isequipactions = False Then
                Dim bv As New frmequipmentactions
                myForms.equipactions = bv
                Dim ds As System.Data.DataSet = New System.Data.DataSet
                ds = Me.dtginventories.DataSource
                Dim tasks As taskclass
                tasks.loadname()
                myForms.equipactions.txtassignedby.Text = tasks.globalnamme
                myForms.equipactions.eid = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("equip_id"))
                myForms.equipactions.Show()
                myForms.isequipactions = True
            Else
                Dim ds As System.Data.DataSet = New System.Data.DataSet
                ds = Me.dtginventories.DataSource

                myForms.equipactions.txtdecommissionrelease.Text = ""
                myForms.equipactions.txtequipid.Text = ""
                myForms.equipactions.txtequipname.Text = ""
                myForms.equipactions.txtreleasedesc.Text = ""
                myForms.equipactions.txtassignedto.Text = ""




                myForms.equipactions.grpassign.Enabled = True
                myForms.equipactions.grpdecommission.Enabled = False
                myForms.equipactions.grprelease.Enabled = False
                myForms.equipactions.eid = Convert.ToString(ds.Tables(0).Rows(htid.Row).Item("equip_id"))
                myForms.equipactions.loadthree()
            End If
        Catch qw As Exception

        End Try
    End Sub
    Private Sub nini()
        Try
            myForms.Main.Invoke(New mydelegatee1(AddressOf loadserchpara))
        Catch ex As Exception

        End Try
    End Sub
    Private Sub loadserchpara()
        Try
            myForms.Main.cboequipsearch.Items.Add("Equipment Id")
            myForms.Main.cboequipsearch.Items.Add("Manufacturer")
            myForms.Main.cboequipsearch.Items.Add("Model No")
            myForms.Main.cboequipsearch.Items.Add("Serial No")
            myForms.Main.cboequipsearch.Items.Add("Model Name")
            myForms.Main.cboequipsearch.Items.Add("Purchase Date")
            myForms.Main.cboequipsearch.Items.Add("Type")
            myForms.Main.cboequipsearch.Items.Add("Model Year")

            '---------------
            myForms.Main.cboequipsearchtrue.Items.Add(" equip_id")
            myForms.Main.cboequipsearchtrue.Items.Add("manufacturer")
            myForms.Main.cboequipsearchtrue.Items.Add(" model_no")
            myForms.Main.cboequipsearchtrue.Items.Add(" serial_no")
            myForms.Main.cboequipsearchtrue.Items.Add("  model_name")
            myForms.Main.cboequipsearchtrue.Items.Add("purchase_date")
            myForms.Main.cboequipsearchtrue.Items.Add("Type")
            myForms.Main.cboequipsearchtrue.Items.Add(" model_year")
        Catch qw As Exception

        End Try
    End Sub
    Private Sub btndeleteequipment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeleteequipment.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to delete equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            dtginventories.Select(htid.Row)
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
            ds = Me.dtginventories.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(htid.Row).Item("equip_id")
            str = "delete from equip_info where equip_id='" & sid & "';"
            str += "delete from assigned_info where equip_id='" & sid & "';"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(htid.Row)
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
    Private Sub btntojob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btntojob.Click
        Try
            Dim task As taskclass
            If task.hastojobloaded = False Then
            End If
            Try
                myForms.tojobs.Close()
                myForms.tojobs = Nothing
            Catch zx As Exception
            End Try
            Dim form As New frmtojobs
            form.StartPosition = FormStartPosition.CenterParent
            myForms.tojobs = form
            myForms.tojobs.Show()
            task.hastojobloaded = True
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnshowall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnshowall.Click
        Try
            Dim Tasks As New taskclass
            Tasks.issearch = False
            Dim Threadshowall As New System.Threading.Thread( _
                AddressOf Tasks.equipinvoke)
            Threadshowall.Start()
        Catch sd As Exception
        End Try
    End Sub
    Private Sub btnaddequipments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddequipments.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulateequip()
            If x = False Then
                MessageBox.Show("Not allowed to manipulate equipment contact administrator", "Equipment", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            Dim form As New frmaddequip
            form.ShowDialog()
        Catch we As Exception

        End Try
    End Sub
    Private Sub btnmaintenace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnmaintenace.Click
        Try
            If myForms.ismaintenace = False Then
                Dim bds As New frmservice
                myForms.maintenace = bds
                myForms.maintenace.StartPosition = FormStartPosition.CenterScreen
                myForms.maintenace.Show()
            End If
        Catch sd As Exception
        End Try
    End Sub
    Private Sub btnhistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnhistory.Click
        Try
            If myForms.ishistory = False Then
                Dim bds As New frmhistory
                myForms.historry = bds
                myForms.historry.StartPosition = FormStartPosition.CenterScreen
                myForms.historry.Show()
            End If
        Catch sd As Exception
        End Try
    End Sub
    Private Sub dtginventories_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtginventories.Navigate

    End Sub
    Private Sub dtginventories_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtginventories.MouseUp
        Try
            If Cursor.Current Is Cursors.Hand Then
                Cursor.Current = ccursor
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region


End Class
