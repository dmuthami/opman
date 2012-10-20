Imports System
Imports System.Data
Imports System.Threading
Imports ADODB


Public Class frmtime
    Inherits System.Windows.Forms.Form
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
    Private jobno As String
    Private afterCurrentCellChanged As Boolean = False
    Private bd As New System.Windows.Forms.DateTimePicker()

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        System.Windows.Forms.Application.EnableVisualStyles()
        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        Catch ex As Exception

        End Try

    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbojob As System.Windows.Forms.ComboBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents pnlpersonnels As System.Windows.Forms.Panel
    Friend WithEvents dtgtimesheet As System.Windows.Forms.DataGrid
    Friend WithEvents pnlgrid As System.Windows.Forms.Panel
    Friend WithEvents lbltimesheet As System.Windows.Forms.Label
    Friend WithEvents grpname As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents StiGroupLine1 As Stimulsoft.Controls.StiGroupLine
    Friend WithEvents cboinvisible As System.Windows.Forms.ComboBox
    Friend WithEvents dtpnow As System.Windows.Forms.DateTimePicker
    Friend WithEvents btndeletetimeentries As System.Windows.Forms.Button
    Friend WithEvents btnaddtime As System.Windows.Forms.Button
    Friend WithEvents btnedit As System.Windows.Forms.Button
    Friend WithEvents msktime As AxMSMask.AxMaskEdBox
    Friend WithEvents rtbdesc As System.Windows.Forms.RichTextBox
    Friend WithEvents etime As System.Windows.Forms.ComboBox
    Friend WithEvents stime As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dtptimesheet As System.Windows.Forms.DateTimePicker
    Friend WithEvents lbldbdate As System.Windows.Forms.Label
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txttask As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmtime))
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.pnlpersonnels = New System.Windows.Forms.Panel
        Me.pnlgrid = New System.Windows.Forms.Panel
        Me.dtgtimesheet = New System.Windows.Forms.DataGrid
        Me.grpname = New System.Windows.Forms.GroupBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.txttask = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.dtptimesheet = New System.Windows.Forms.DateTimePicker
        Me.stime = New System.Windows.Forms.ComboBox
        Me.etime = New System.Windows.Forms.ComboBox
        Me.btnedit = New System.Windows.Forms.Button
        Me.btnaddtime = New System.Windows.Forms.Button
        Me.btndeletetimeentries = New System.Windows.Forms.Button
        Me.cboinvisible = New System.Windows.Forms.ComboBox
        Me.dtpnow = New System.Windows.Forms.DateTimePicker
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lbldbdate = New System.Windows.Forms.Label
        Me.lbltimesheet = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rtbdesc = New System.Windows.Forms.RichTextBox
        Me.msktime = New AxMSMask.AxMaskEdBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbojob = New System.Windows.Forms.ComboBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.errp = New System.Windows.Forms.ErrorProvider
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlpersonnels.SuspendLayout()
        Me.pnlgrid.SuspendLayout()
        CType(Me.dtgtimesheet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpname.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.msktime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.Text = "StatusBarPanel1"
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.Text = "StatusBarPanel2"
        '
        'pnlpersonnels
        '
        Me.pnlpersonnels.AutoScroll = True
        Me.pnlpersonnels.Controls.Add(Me.pnlgrid)
        Me.pnlpersonnels.Controls.Add(Me.grpname)
        Me.pnlpersonnels.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlpersonnels.Location = New System.Drawing.Point(0, 0)
        Me.pnlpersonnels.Name = "pnlpersonnels"
        Me.pnlpersonnels.Size = New System.Drawing.Size(610, 640)
        Me.pnlpersonnels.TabIndex = 2
        '
        'pnlgrid
        '
        Me.pnlgrid.Controls.Add(Me.dtgtimesheet)
        Me.pnlgrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlgrid.Location = New System.Drawing.Point(0, 328)
        Me.pnlgrid.Name = "pnlgrid"
        Me.pnlgrid.Size = New System.Drawing.Size(610, 312)
        Me.pnlgrid.TabIndex = 20
        '
        'dtgtimesheet
        '
        Me.dtgtimesheet.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgtimesheet.CaptionText = "Time sheet details"
        Me.dtgtimesheet.DataMember = ""
        Me.dtgtimesheet.FlatMode = True
        Me.dtgtimesheet.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgtimesheet.Location = New System.Drawing.Point(8, 0)
        Me.dtgtimesheet.Name = "dtgtimesheet"
        Me.dtgtimesheet.ReadOnly = True
        Me.dtgtimesheet.Size = New System.Drawing.Size(594, 300)
        Me.dtgtimesheet.TabIndex = 13
        '
        'grpname
        '
        Me.grpname.Controls.Add(Me.GroupBox3)
        Me.grpname.Controls.Add(Me.Label5)
        Me.grpname.Controls.Add(Me.dtptimesheet)
        Me.grpname.Controls.Add(Me.stime)
        Me.grpname.Controls.Add(Me.etime)
        Me.grpname.Controls.Add(Me.btnedit)
        Me.grpname.Controls.Add(Me.btnaddtime)
        Me.grpname.Controls.Add(Me.btndeletetimeentries)
        Me.grpname.Controls.Add(Me.cboinvisible)
        Me.grpname.Controls.Add(Me.dtpnow)
        Me.grpname.Controls.Add(Me.GroupBox2)
        Me.grpname.Controls.Add(Me.GroupBox1)
        Me.grpname.Controls.Add(Me.Label4)
        Me.grpname.Controls.Add(Me.Label3)
        Me.grpname.Controls.Add(Me.Label1)
        Me.grpname.Controls.Add(Me.cbojob)
        Me.grpname.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpname.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grpname.Location = New System.Drawing.Point(0, 0)
        Me.grpname.Name = "grpname"
        Me.grpname.Size = New System.Drawing.Size(610, 328)
        Me.grpname.TabIndex = 0
        Me.grpname.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.txttask)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(8, 90)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(592, 112)
        Me.GroupBox3.TabIndex = 6
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Enter tasks done"
        '
        'txttask
        '
        Me.txttask.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txttask.Location = New System.Drawing.Point(8, 16)
        Me.txttask.Multiline = True
        Me.txttask.Name = "txttask"
        Me.txttask.Size = New System.Drawing.Size(576, 88)
        Me.txttask.TabIndex = 7
        Me.txttask.Text = ""
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 43)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 16)
        Me.Label5.TabIndex = 32
        Me.Label5.Text = "Pick a date "
        '
        'dtptimesheet
        '
        Me.dtptimesheet.Location = New System.Drawing.Point(112, 41)
        Me.dtptimesheet.Name = "dtptimesheet"
        Me.dtptimesheet.Size = New System.Drawing.Size(328, 20)
        Me.dtptimesheet.TabIndex = 2
        '
        'stime
        '
        Me.stime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.stime.Items.AddRange(New Object() {"0:00  AM", "0:30  AM", "1:00  AM", "1:30  AM", "2:00  AM", "2:30  AM", "3:00  AM", "3:30  AM", "4:00  AM", "4:30  AM", "5:00  AM", "5:30  AM", "6:00  AM", "6:30  AM", "7:00  AM", "7:30  AM", "8:00  AM", "8:30  AM", "9:00  AM", "9:30  AM", "10:00  AM", "10:30  AM", "11:00  AM", "11:30  AM", "12:00  PM", "12:30  PM", "1:00  PM", "1:30  PM", "2:00  PM", "2:30  PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", "6:00 PM", "6:30 PM", "7:00 PM", "7:30 PM", "8:00 PM", "8:30 PM", "9:00 PM", "9:30 PM", "10:00 PM", "10:30 PM", "11:00 PM", "11:30 PM", "0:00 AM"})
        Me.stime.Location = New System.Drawing.Point(112, 65)
        Me.stime.Name = "stime"
        Me.stime.Size = New System.Drawing.Size(120, 22)
        Me.stime.TabIndex = 3
        '
        'etime
        '
        Me.etime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.etime.Items.AddRange(New Object() {"0:00  AM", "0:30  AM", "1:00  AM", "1:30  AM", "2:00  AM", "2:30  AM", "3:00  AM", "3:30  AM", "4:00  AM", "4:30  AM", "5:00  AM", "5:30  AM", "6:00  AM", "6:30  AM", "7:00  AM", "7:30  AM", "8:00  AM", "8:30  AM", "9:00  AM", "9:30  AM", "10:00  AM", "10:30  AM", "11:00  AM", "11:30  AM", "12:00  PM", "12:30  PM", "1:00  PM", "1:30  PM", "2:00  PM", "2:30  PM", "3:00 PM", "3:30 PM", "4:00 PM", "4:30 PM", "5:00 PM", "5:30 PM", "6:00 PM", "6:30 PM", "7:00 PM", "7:30 PM", "8:00 PM", "8:30 PM", "9:00 PM", "9:30 PM", "10:00 PM", "10:30 PM", "11:00 PM", "11:30 PM", "0:00 AM"})
        Me.etime.Location = New System.Drawing.Point(301, 65)
        Me.etime.Name = "etime"
        Me.etime.Size = New System.Drawing.Size(136, 22)
        Me.etime.TabIndex = 4
        '
        'btnedit
        '
        Me.btnedit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnedit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnedit.Location = New System.Drawing.Point(104, 300)
        Me.btnedit.Name = "btnedit"
        Me.btnedit.Size = New System.Drawing.Size(94, 24)
        Me.btnedit.TabIndex = 11
        Me.btnedit.Text = "Refresh"
        '
        'btnaddtime
        '
        Me.btnaddtime.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnaddtime.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnaddtime.Location = New System.Drawing.Point(8, 300)
        Me.btnaddtime.Name = "btnaddtime"
        Me.btnaddtime.Size = New System.Drawing.Size(96, 24)
        Me.btnaddtime.TabIndex = 10
        Me.btnaddtime.Text = "Add time entry"
        '
        'btndeletetimeentries
        '
        Me.btndeletetimeentries.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btndeletetimeentries.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeletetimeentries.Location = New System.Drawing.Point(200, 300)
        Me.btndeletetimeentries.Name = "btndeletetimeentries"
        Me.btndeletetimeentries.Size = New System.Drawing.Size(120, 24)
        Me.btndeletetimeentries.TabIndex = 12
        Me.btndeletetimeentries.Text = "Delete time entry(s)"
        '
        'cboinvisible
        '
        Me.cboinvisible.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboinvisible.Location = New System.Drawing.Point(568, 304)
        Me.cboinvisible.Name = "cboinvisible"
        Me.cboinvisible.Size = New System.Drawing.Size(8, 22)
        Me.cboinvisible.TabIndex = 25
        Me.cboinvisible.Visible = False
        '
        'dtpnow
        '
        Me.dtpnow.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpnow.CustomFormat = ""
        Me.dtpnow.Location = New System.Drawing.Point(584, 304)
        Me.dtpnow.Name = "dtpnow"
        Me.dtpnow.Size = New System.Drawing.Size(8, 20)
        Me.dtpnow.TabIndex = 24
        Me.dtpnow.Value = New Date(2006, 5, 1, 0, 0, 0, 0)
        Me.dtpnow.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.lbldbdate)
        Me.GroupBox2.Controls.Add(Me.lbltimesheet)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(448, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(152, 80)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        '
        'lbldbdate
        '
        Me.lbldbdate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbldbdate.BackColor = System.Drawing.Color.FromArgb(CType(209, Byte), CType(241, Byte), CType(254, Byte))
        Me.lbldbdate.Location = New System.Drawing.Point(8, 36)
        Me.lbldbdate.Name = "lbldbdate"
        Me.lbldbdate.Size = New System.Drawing.Size(128, 16)
        Me.lbldbdate.TabIndex = 21
        Me.lbldbdate.Text = "Database date :"
        Me.lbldbdate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbltimesheet
        '
        Me.lbltimesheet.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbltimesheet.BackColor = System.Drawing.SystemColors.Window
        Me.lbltimesheet.Location = New System.Drawing.Point(8, 16)
        Me.lbltimesheet.Name = "lbltimesheet"
        Me.lbltimesheet.Size = New System.Drawing.Size(128, 16)
        Me.lbltimesheet.TabIndex = 18
        Me.lbltimesheet.Text = "Date here"
        Me.lbltimesheet.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.rtbdesc)
        Me.GroupBox1.Controls.Add(Me.msktime)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 199)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(592, 97)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Additional notes"
        '
        'rtbdesc
        '
        Me.rtbdesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbdesc.Location = New System.Drawing.Point(8, 16)
        Me.rtbdesc.Name = "rtbdesc"
        Me.rtbdesc.Size = New System.Drawing.Size(576, 72)
        Me.rtbdesc.TabIndex = 9
        Me.rtbdesc.Text = ""
        '
        'msktime
        '
        Me.msktime.ContainingControl = Me
        Me.msktime.Location = New System.Drawing.Point(272, 32)
        Me.msktime.Name = "msktime"
        Me.msktime.OcxState = CType(resources.GetObject("msktime.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msktime.Size = New System.Drawing.Size(75, 16)
        Me.msktime.TabIndex = 29
        Me.msktime.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(236, 68)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 16)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "End time"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Start time"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Pick a job"
        '
        'cbojob
        '
        Me.cbojob.Location = New System.Drawing.Point(112, 17)
        Me.cbojob.Name = "cbojob"
        Me.cbojob.Size = New System.Drawing.Size(328, 22)
        Me.cbojob.TabIndex = 1
        '
        'Timer1
        '
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        Me.errp.DataMember = ""
        '
        'frmtime
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(610, 640)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlpersonnels)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmtime"
        Me.Text = "Personnels name here"
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlpersonnels.ResumeLayout(False)
        Me.pnlgrid.ResumeLayout(False)
        CType(Me.dtgtimesheet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpname.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.msktime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    <System.STAThread()> _
            Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub

#Region "timesheet"
    Private Sub frmtime_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            'bd.Value = Now
            'bd.Value.AddDays(1)
            Me.dtpnow.Value = Now
            Timer1.Interval = 1000
            Timer1.Enabled = True
            Timer1.Start()
            stime.SelectedIndex = 16
            etime.SelectedIndex = 16
            Try
                'stime.Value = Now
                'etime.Value = Now
            Catch xc As Exception

            End Try
            Call DoWork()
            Call getjobs()
            Call loadgrid()
            Try
                Me.lbldbdate.Text += myForms.Main.currentdate
            Catch xc As Exception

            End Try
            Me.Invalidate(True)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub DoWork()
        Try
            Dim Tasks As New taskclass
            Dim Thread1 As New System.Threading.Thread( _
                AddressOf taskclass.SomeTask)
            Tasks.Strid_no = myForms.id_no  ' Set a field that is used as an argument
            Thread1.IsBackground = True
            Thread1.Start() ' Start the new thread.
            'Thread1.Join() ' Wait for thread 1 to finish.
            '' Display the return value.
            'MsgBox("Thread 1 returned the value " & Tasks.RetVal)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
                    & ex.InnerException().ToString() & vbCrLf _
                    & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub getjobs()
        Try
            Dim Tasks As New taskclass
            Dim Thread2 As New System.Threading.Thread( _
                AddressOf taskclass.somejob)

            Thread2.Start() ' Start the new thread.
            'Thread1.Join() ' Wait for thread 1 to finish.
            '' Display the return value.
            'MsgBox("Thread 1 returned the value " & Tasks.RetVal)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
                    & ex.InnerException().ToString() & vbCrLf _
                    & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            dtpnow.Value = Now
            lbltimesheet.Text = dtpnow.Value.ToLongDateString & " " & dtpnow.Value.ToLongTimeString
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
             & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub loadgrid()
        Try
            Dim Tasks As New taskclass
            Dim Thread3 As New System.Threading.Thread( _
                AddressOf taskclass.myinvoke)

            Thread3.Start() ' Start the new thread.
            'Thread1.Join() ' Wait for thread 1 to finish.
            '' Display the return value.
            'MsgBox("Thread 1 returned the value " & Tasks.RetVal)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub cbojob_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbojob.SelectedValueChanged
        Try
            Dim indexx As Integer
            indexx = Me.cbojob.SelectedIndex
            If indexx = -1 Then
                Exit Try

            End If
            Me.cboinvisible.SelectedIndex = indexx
            Dim strp
            strp = cboinvisible.Text
            jobno = strp

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try

    End Sub
    Private Function returnhittest(ByVal hit As System.Windows.Forms.DataGrid.HitTestType) As String
        Try
            Select Case hit
                Case System.Windows.Forms.DataGrid.HitTestType.Cell

                    Dim mycell As New DataGridCell

                    mycell.RowNumber = Me.dtgtimesheet.CurrentRowIndex
                    mycell.ColumnNumber = 0

                    returnhittest = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell

                    mycell.RowNumber = Me.dtgtimesheet.CurrentRowIndex
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
    Private Sub tmredit_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'btnaddtime.Text = "Edit"
            'tmredit.Stop()
        Catch ex As Exception

        End Try
    End Sub
    Private kappa As String
    Private Sub edit()
        Try

            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim mask As String
            Dim sdate, sdate2 As String
            Dim mymilliseconds = dtpnow.Value.Millisecond.ToString()
            sdate = dtpnow.Value.Year & "-" _
            & dtpnow.Value.Month & "-" _
            & dtpnow.Value.Day & " " _
            & dtpnow.Value.Hour & ":" _
            & dtpnow.Value.Minute & ":" _
            & dtpnow.Value.Second

            '---------------------------
            Dim datediff As System.TimeSpan
            Dim dtpetime As New System.Windows.Forms.DateTimePicker
            Dim dtpstime As New System.Windows.Forms.DateTimePicker
            dtpstime.Value = CDate("5/12/2006 " & stime.Text)
            dtpetime.Value = CDate("5/12/2006 " & etime.Text)
            datediff = dtpetime.Value.Subtract(dtpstime.Value)
            Dim d As Double = 0
            If datediff.TotalMinutes < d _
              Or datediff.TotalMinutes = 0 Then
                MessageBox.Show("Invalid time", " Timesheet", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            Else
                'mask = datediff.Hours & ":" & datediff.Minutes & ":" & datediff.Seconds
                mask = datediff.TotalHours
            End If
            sdate2 = Me.dtptimesheet.Value.Year & "-" _
                & dtptimesheet.Value.Month & "-" _
                & dtptimesheet.Value.Day & " " _
                & "01" & ":" _
                & "00" & ":" _
                & "00"
            '--------------
            Dim ds As System.Data.DataSet = New System.Data.DataSet
            ds = Me.dtgtimesheet.DataSource
            Dim mymilieseconds As String = ds.Tables(0).Rows(kappa).Item("ano")
            Dim myday As String = ds.Tables(0).Rows(kappa).Item("ddate")
            Dim strsql As String
            Dim strin As String = txttask.Text.Trim
            Dim strin1 As String = rtbdesc.Text.Trim
            strin = strin.Replace("'", "\'")
            strin1 = strin1.Replace("'", "\'")
            txttask.Text = strin
            rtbdesc.Text = strin1
            '-------------------
            Dim arr() As String
            Dim strr, strr2 As String
            Dim y As Integer
            txttask.Text = Me.txttask.Text.Trim()
            arr = txttask.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            '----------------------------------
            rtbdesc.Text = Me.rtbdesc.Text.Trim()
            arr = rtbdesc.Lines
            y = arr.GetUpperBound(0)
            For alpha = 0 To y
                strr2 += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            strsql = "update  daily_time set   "
            strsql += " id_no='" & myForms.id_no & "',job_no='" & jobno & "',task='" & strr & "',"
            strsql += "description='" & Me.cbojob.Text.Trim() & "',ddate='" & sdate2 & "',timespent='" & mask & "',milliseconds='" & mymilliseconds & "'"
            strsql += ",stime='" & stime.Text.Trim() & "',etime='" & etime.Text.Trim() & "',notes='" & strr2 & "'"
            strsql += " where   ano='" & mymilieseconds & "'"
            connect.BeginTrans()
            connect.Execute(strsql)
            connect.CommitTrans()

            Dim nb As String = Me.cbojob.Text
            Try
                Dim xc() As String = nb.Split(":")
                nb = xc(1)
            Catch cvbn As Exception
            End Try
            ds.Tables(0).Rows(kappa).Item("ddate") = sdate
            ds.Tables(0).Rows(kappa).Item("milliseconds") = mymilieseconds
            ds.Tables(0).Rows(kappa).Item("description") = cbojob.Text.Trim()
            ds.Tables(0).Rows(kappa).Item("job_no") = jobno
            ds.Tables(0).Rows(kappa).Item("id_no") = myForms.id_no
            ds.Tables(0).Rows(kappa).Item("task") = txttask.Text.Trim()
            ds.Tables(0).Rows(kappa).Item("timespent") = mask
            ds.Tables(0).Rows(kappa).Item("stime") = stime.Text.Trim()
            ds.Tables(0).Rows(kappa).Item("etime") = etime.Text.Trim()
            ds.Tables(0).Rows(kappa).Item("notes") = rtbdesc.Text.Trim()
            kappa = ""
            Dim task As New taskclass
            task.addtablestyle(ds.Tables(0).TableName)
            Try
                connect.Close()
            Catch ex345 As Exception
            End Try
        Catch ex As Exception
        Finally
            ' btnaddtime.Text = "Add"
        End Try
        '---------refresh grosss margin
        Try
            myForms.CustomerForm2.mygross()
        Catch zx As Exception

        End Try
        '----------------------
    End Sub
    Private Sub dtgtimesheet_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgtimesheet.CurrentCellChanged
        ' focus shift to discontinued column if it is checked if
        ' user clicks any where else
        Try
            Dim discontinuedColumn As Integer = 0
            Dim val As Object = Me.dtgtimesheet( _
                Me.dtgtimesheet.CurrentRowIndex, _
                discontinuedColumn)
            Dim productDiscontinued As Boolean = CBool(val)
            If productDiscontinued Then
                Me.dtgtimesheet.CurrentCell = _
                   New DataGridCell( _
                       Me.dtgtimesheet.CurrentRowIndex, _
                       discontinuedColumn)
            End If
            afterCurrentCellChanged = True

        Catch er As System.Exception
            MsgBox(er.ToString())
        End Try
    End Sub
    Private Sub dtgtimesheet_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgtimesheet.Click
        'Try
        '    Dim discontinuedColumn As Integer = 0
        '    Dim pt As Point = Me.dtgtimesheet.PointToClient( _
        '        Control.MousePosition)
        '    Dim htid As DataGrid.HitTestInfo = _
        '        Me.dtgtimesheet.HitTest(pt)
        '    Dim bmb As BindingManagerBase = _
        '        Me.BindingContext(Me.dtgtimesheet.DataSource, _
        '        Me.dtgtimesheet.DataMember)

        '    '-------------------
        '    'htid.Type = DataGrid.HitTestType.Cell _
        '    '              AndAlso hti.Column = discontinuedColumn
        '    If htid.Type = DataGrid.HitTestType.Cell Then
        '        Dim mycell As New DataGridCell()
        '        mycell = dtgtimesheet.CurrentCell
        '        If mycell.ColumnNumber = discontinuedColumn Then
        '            Dim ds As System.Data.DataSet = New System.Data.DataSet()
        '            ds = Me.dtgtimesheet.DataSource
        '            If CBool(ds.Tables(0).Rows(mycell.RowNumber).Item("Delete")) = False Then
        '                ds.Tables(0).Rows(mycell.RowNumber).Item("Delete") = True
        '            Else
        '                ds.Tables(0).Rows(mycell.RowNumber).Item("Delete") = False
        '            End If

        '            Dim task As New taskclass()
        '            task.addtablestyle(ds.Tables(0).TableName)

        '        End If
        '    End If

        'Catch ex As Exception

        'End Try

        Try
            Dim discontinuedColumn As Integer = 0
            Dim pt As Point = Me.dtgtimesheet.PointToClient( _
                Control.MousePosition)
            'Dim hti As DataGrid.HitTestInfo = _
            '    Me.dtgtimesheet.HitTest(pt)
            Dim bmb As BindingManagerBase = _
                Me.BindingContext(Me.dtgtimesheet.DataSource, _
                Me.dtgtimesheet.DataMember)

            If hti.Row < bmb.Count _
               AndAlso hti.Type = DataGrid.HitTestType.Cell _
               AndAlso hti.Column = discontinuedColumn Then
                Me.dtgtimesheet(hti.Row, discontinuedColumn) = _
                   Not CBool(Me.dtgtimesheet(hti.Row, _
                         discontinuedColumn))
                'Dim ds As System.Data.DataSet = New System.Data.DataSet()
                'ds = Me.dtgtimesheet.DataSource
                'MsgBox(Convert.ToString(ds.Tables(0).Rows(hti.Row).Item("Delete")))


            End If
            afterCurrentCellChanged = False

        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtgtimesheet_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgtimesheet.MouseDown
        Try
            hti = dtgtimesheet.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgtimesheet_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgtimesheet.DoubleClick
        Try
            Dim results As String
            results = returnhittest(hti.Type)
            If results <> "" Then
                Dim ds As DataSet = New DataSet
                ds = Me.dtgtimesheet.DataSource

                'Dim mycell As New DataGridCell()
                'Dim a() As String
                'a = results.Split("|")
                'mycell.RowNumber = CInt(a(0))
                kappa = hti.Row
                Try
                    dtgtimesheet.Select(Integer.Parse(kappa))
                Catch ex As Exception
                End Try
                'mycell.ColumnNumber = 0 + 1
                Try
                    Me.cbojob.Text = ds.Tables(0).Rows(hti.Row).Item("description")
                Catch es As Exception

                End Try
                Try
                    Me.txttask.Text = ds.Tables(0).Rows(hti.Row).Item("task")
                Catch es As Exception

                End Try

                Try
                    Me.stime.Text = ds.Tables(0).Rows(hti.Row).Item("stime")
                Catch es As Exception

                End Try
                Try
                    Me.etime.Text = ds.Tables(0).Rows(hti.Row).Item("etime")
                Catch es As Exception

                End Try
                Try
                    Me.rtbdesc.Text = ""
                    Me.rtbdesc.Text = ds.Tables(0).Rows(hti.Row).Item("notes")
                Catch es As Exception

                End Try
                'mycell.ColumnNumber = 3
                'Me.rtbdesc.Text = dtgtimesheet(mycell)




            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())

        End Try
    End Sub
    Private Sub btnaddtime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddtime.Click
        ''----------validate time entry
        'Dim isbool As Boolean = validatetime()
        'If isbool = False Then
        '    MessageBox.Show("Synchronize system date with database date", " Time sheets", _
        '       MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    myForms.timesheet.dtgtimesheet.DataSource = Nothing
        '    Exit Sub
        'End If
        ''-----------end of validation

        Dim currentCursor As Cursor = Cursor.Current
        Dim isadded As Boolean = False
        Try
            Dim connectstr As String
            connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Cursor.Current = Cursors.WaitCursor
            If Me.cbojob.Text.Trim.Length < 1 Then
                isadded = False
                Exit Try
            End If
            Dim mask As String
            Dim sdate, sdate2 As String
            Dim mymilliseconds = dtpnow.Value.Millisecond.ToString()
            sdate = dtpnow.Value.Year & "-" _
            & dtpnow.Value.Month & "-" _
            & dtpnow.Value.Day & " " _
            & dtpnow.Value.Hour & ":" _
            & dtpnow.Value.Minute & ":" _
            & dtpnow.Value.Second
            mymilliseconds = sdate.Trim() & "|" & dtpnow.Value.Millisecond.ToString()

            '---------------------------
            Dim datediff As System.TimeSpan
            Dim dtpetime As New System.Windows.Forms.DateTimePicker
            Dim dtpstime As New System.Windows.Forms.DateTimePicker
            dtpstime.Value = CDate("5/12/2006 " & stime.Text)
            dtpetime.Value = CDate("5/12/2006 " & etime.Text)
            datediff = dtpetime.Value.Subtract(dtpstime.Value)
            Dim d As Double = 0
            If datediff.TotalMinutes < d _
            Or datediff.TotalMinutes = 0 Then
                MessageBox.Show("Invalid time", " Timesheet", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            Else
                'mask = datediff.Hours & ":" & datediff.Minutes & ":" & datediff.Seconds
                mask = datediff.TotalHours
            End If
            sdate2 = Me.dtptimesheet.Value.Year & "-" _
         & dtptimesheet.Value.Month & "-" _
         & dtptimesheet.Value.Day & " " _
         & "00" & ":" _
         & "00" & ":" _
         & "00"
            '--------------

            Dim strin As String = txttask.Text.Trim
            Dim strin1 As String = rtbdesc.Text.Trim
            strin = strin.Replace("'", "\'")
            strin1 = strin1.Replace("'", "\'")
            txttask.Text = strin
            rtbdesc.Text = strin1
            '-------------------
            Dim arr() As String
            Dim strr, strr2 As String
            Dim y As Integer
            txttask.Text = Me.txttask.Text.Trim()
            arr = txttask.Lines
            y = arr.GetUpperBound(0)
            Dim alpha As Integer
            For alpha = 0 To y
                strr += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            '----------------------------------
            rtbdesc.Text = Me.rtbdesc.Text.Trim()
            arr = rtbdesc.Lines
            y = arr.GetUpperBound(0)
            For alpha = 0 To y
                strr2 += arr(alpha) + vbCrLf
                Application.DoEvents()
            Next
            Dim strsql As String
            strsql = "insert into daily_time (id_no,job_no,task,description,ddate,timespent,milliseconds,stime,etime,notes) "
            strsql += "values ('" & myForms.id_no & "','" & jobno & "','" & strr & "',"
            strsql += "'" & Me.cbojob.Text.Trim() & "','" & sdate2 & "','" & mask & "','" & mymilliseconds & "',"
            strsql += "'" & stime.Text.Trim() & "','" & etime.Text.Trim() & "','" & strr2 & "')"
            If strsql <> "" Then
                connect.BeginTrans()
                connect.IsolationLevel = IsolationLevelEnum.adXactSerializable
                connect.Execute(strsql)
                connect.CommitTrans()
                isadded = True
            End If
            strsql = ""
            Try
                connect.Close()
            Catch ex345 As Exception
            End Try
            ''---------------
            Dim Tasks As New taskclass
            Dim Thread4 As New System.Threading.Thread( _
                AddressOf taskclass.myinvoke)
            Thread4.IsBackground = True
            Thread4.Start() ' Start the new thread.
            '-------------------end of this


        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())

        Finally
            If isadded = True Then
                txttask.Text = ""
                msktime.Mask = ""
                msktime.Text = ""
                msktime.Mask = "##:##:##"
                rtbdesc.Text = ""
                stime.SelectedIndex = 16
                etime.SelectedIndex = 16
                Me.cbojob.SelectedIndex = -1
            End If
            Cursor.Current = currentCursor
        End Try
        '---------refresh grosss margin
        Try
            myForms.CustomerForm2.mygross()
        Catch zx As Exception

        End Try
        '----------------------
    End Sub
    Private Sub btnedit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnedit.Click
        ''----------validate time entry
        'Dim isbool As Boolean = validatetime()
        'If isbool = False Then
        '    MessageBox.Show("Synchronize system date with database date", " Time sheets", _
        '       MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    myForms.timesheet.dtgtimesheet.DataSource = Nothing
        '    Exit Sub
        'End If
        ''-----------end of validation
        Try
            If kappa = "" Then
                MessageBox.Show("Please double click on a row header so as to populate the controls with the data to be edited", " Refresh", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try
            End If
            If MessageBox.Show("Are you sure you want to edit", "Editing", _
            MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                Call edit()
            End If
        Catch er As Exception
        End Try
    End Sub
    Private Sub btndeletetimeentries_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeletetimeentries.Click
        ''----------validate time entry
        'Dim isbool As Boolean = validatetime()
        'If isbool = False Then
        '    MessageBox.Show("Synchronize system date with database date", " Time sheets", _
        '       MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    myForms.timesheet.dtgtimesheet.DataSource = Nothing
        '    Exit Sub
        'End If

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
            ds = Me.dtgtimesheet.DataSource

            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sdate, myseconds, str As String
            Dim myrow As System.Data.DataRow
            For y = 0 To i - 1
                If ds.Tables(0).Rows(y).Item("Delete") = True Then
                    sdate = ds.Tables(0).Rows(y).Item("ddate")
                    myseconds = ds.Tables(0).Rows(y).Item("ano")
                    str = "delete from daily_time where ano='" & myseconds & "'"
                    'str += " and ddate='" & sdate & "' and id_no='" & ds.Tables(0).Rows(y).Item("id_no") & "'"
                    Try
                        connect.BeginTrans()
                        connect.Execute(str)
                        connect.CommitTrans()
                        myrow = ds.Tables(0).Rows(y)
                        ds.Tables(0).Rows.Remove(myrow)
                        y = y - 1
                    Catch cv As Exception
                    End Try
                End If


            Next
            Try
                connect.Close()
            Catch er As Exception

            End Try
        Catch ex As Exception

        End Try
        '---------refresh grosss margin
        Try
            myForms.CustomerForm2.mygross()
        Catch zx As Exception

        End Try
        '----------------------
    End Sub
    Private Function validatetime() As Boolean
        validatetime = False
        Dim dtptext As String
        dtptext = Me.dtptimesheet.Text
        Try
            '-------------------validate time sheet
            If myForms.Main.currentdate.Trim.Length < 1 Then
                validatetime = False
            End If
            Dim dm As New datemanipulation
            Dim dtp As New System.Windows.Forms.DateTimePicker
            dtp.Value = CDate(myForms.Main.currentdate)
            Dim sd As String
            sd = dtp.Value.Year & "-" _
            & dtp.Value.Month & "-" _
            & dtp.Value.Day & " " _
            & "00" & ":" _
            & "00" & ":" _
            & "00"
            dm.dbtime = sd
            dtp.Value = Me.dtptimesheet.Value
            sd = dtp.Value.Year & "-" _
           & dtp.Value.Month & "-" _
           & dtp.Value.Day & " " _
           & dtp.Value.Hour & ":" _
           & dtp.Value.Minute & ":" _
           & dtp.Value.Second
            dm.ctime = sd
            Dim fd As String = dm.datediff
            If fd.Trim.Length < 1 Then
                validatetime = False
            End If
            If Convert.ToDouble(fd) <= 0 _
            And Convert.ToDouble(fd) >= -1 Then
                validatetime = True
            End If

            '-------------------------
        Catch jb As Exception
            validatetime = False
        End Try
    End Function
    Private Sub dtptimesheet_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtptimesheet.TextChanged
        ''----------validate time entry
        'Dim isbool As Boolean = validatetime()
        'If isbool = False Then
        '    MessageBox.Show("Synchronize system date with database date", " Time sheets", _
        '       MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    myForms.timesheet.dtgtimesheet.DataSource = Nothing
        '    Exit Sub
        'End If
        ''-----------end of validation
        Try
            Dim sd As String
            sd = dtptimesheet.Value.Year & "-" _
            & dtptimesheet.Value.Month & "-" _
            & dtptimesheet.Value.Day & " " _
            & dtptimesheet.Value.Hour & ":" _
            & dtptimesheet.Value.Minute & ":" _
            & dtptimesheet.Value.Second

            Dim task As taskclass
            task.strtime = sd
            Dim Thread1 As New System.Threading.Thread( _
                AddressOf task.myinvoke)
            Thread1.IsBackground = True
            Thread1.Start()
        Catch ex As Exception

        End Try
        Try
            stime.SelectedIndex = 16
            etime.SelectedIndex = 16
            txttask.Text = ""
            rtbdesc.Text = ""
        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtgtimesheet_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtgtimesheet.Navigate

    End Sub
    Private Sub dtptimesheet_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtptimesheet.ValueChanged

    End Sub
#End Region

#Region "dump"
    'mask = msktime.FormattedText()
    Dim hh, mm, ss As String
    'Dim a() As String
    'a = mask.Split(":")
    'hh = a(0)
    'mm = a(1)
    'ss = a(2)
    'Dim hme As Boolean = False
    'Try
    '    Dim hhint As Integer = Integer.Parse(hh)
    '    If hh > 23 Then
    '        MessageBox.Show("Please ensure that hours do not exceed 23", _
    '        "Invalid time entry", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        Exit Try
    '    End If
    'Catch ex As Exception
    '    hme = True
    'End Try
    'Try
    '    Dim hhint As Integer = Integer.Parse(hh)
    '    If mm > 59 Then
    '        MessageBox.Show(" Please ensure that minutes do not exceed 59 ", _
    '        "Invalid time entry", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        Exit Try
    '    End If
    'Catch ex As Exception
    '    hme = True
    'End Try
    'Try
    '    Dim hhint As Integer = Integer.Parse(hh)
    '    If ss > 59 Then
    '        MessageBox.Show("Please ensure that seconds do not exceed 59", _
    '         "Invalid time entry", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '        Exit Try
    '    End If
    'Catch ex As Exception
    '    hme = True
    'End Try
    'If hme = True Then
    '    Exit Try
    'End If
    'msktime.Mask = ""
    'msktime.Text = ""
    'msktime.Mask = "##:##:##"
    '----------
#End Region

    '#Region "validation"
    '    Private Sub txttask_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txttask.KeyPress
    '        Try
    '            If _validatetextbox(e) = True Then
    '                Me.errp.SetError(Me.txttask, _
    '                                      "not allowed chars: ''','%','*','\','*','1'")
    '                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
    '            Else
    '                Me.errp.SetError(Me.txttask, "")
    '            End If
    '        Catch xc As Exception

    '        End Try
    '    End Sub
    '    Private Function _validatetextbox(ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
    '        _validatetextbox = True
    '        Try
    '            Select Case e.KeyChar
    '                Case "'"
    '                    e.Handled = True 'it indicates the event is handled.

    '                Case "%"
    '                    e.Handled = True 'it indicates the event is handled.

    '                Case "\"
    '                    e.Handled = True 'it indicates the event is handled.

    '                Case """"
    '                    e.Handled = True 'it indicates the event is handled.

    '                Case Else
    '                    _validatetextbox = False

    '            End Select
    '        Catch we As Exception

    '        End Try
    '    End Function
    '    Private Sub rtbdesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles rtbdesc.KeyPress
    '        Try
    '            If _validatetextbox(e) = True Then
    '                Me.errp.SetError(Me.rtbdesc, _
    '                                      "not allowed chars: ''','%','*','\','*','1'")
    '                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
    '            Else
    '                Me.errp.SetError(Me.txttask, "")
    '            End If
    '        Catch xc As Exception

    '        End Try
    '    End Sub
    '#End Region

    Private Sub grpname_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grpname.Enter

    End Sub
End Class
