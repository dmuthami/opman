Imports System

Imports ADODB
Imports System.Data.OleDb
Public Class frmjobs
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
    Friend WithEvents pnljobs As System.Windows.Forms.Panel
    Friend WithEvents StiGroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents StiGroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents MaskedTextBox2 As AMS.TextBox.MaskedTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnjobsquery As VSEssentials.VSHotButton
    Friend WithEvents btnsearch As System.Windows.Forms.Button
    Friend WithEvents cbojobsearchfield As System.Windows.Forms.ComboBox
    Friend WithEvents btnshowall As System.Windows.Forms.Button
    Friend WithEvents dtpto As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpfrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtgjjobs As System.Windows.Forms.DataGrid
    Friend WithEvents chkusedates As System.Windows.Forms.CheckBox
    Friend WithEvents stiexporttoexcel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnljobs = New System.Windows.Forms.Panel
        Me.StiGroupBox1 = New System.Windows.Forms.GroupBox
        Me.stiexporttoexcel = New System.Windows.Forms.Button
        Me.btnjobsquery = New VSEssentials.VSHotButton
        Me.btnshowall = New System.Windows.Forms.Button
        Me.dtgjjobs = New System.Windows.Forms.DataGrid
        Me.StiGroupBox2 = New System.Windows.Forms.GroupBox
        Me.chkusedates = New System.Windows.Forms.CheckBox
        Me.dtpto = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpfrom = New System.Windows.Forms.DateTimePicker
        Me.btnsearch = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.MaskedTextBox2 = New AMS.TextBox.MaskedTextBox
        Me.cbojobsearchfield = New System.Windows.Forms.ComboBox
        Me.pnljobs.SuspendLayout()
        Me.StiGroupBox1.SuspendLayout()
        CType(Me.dtgjjobs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StiGroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnljobs
        '
        Me.pnljobs.Controls.Add(Me.StiGroupBox1)
        Me.pnljobs.Controls.Add(Me.StiGroupBox2)
        Me.pnljobs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnljobs.Location = New System.Drawing.Point(0, 0)
        Me.pnljobs.Name = "pnljobs"
        Me.pnljobs.Size = New System.Drawing.Size(500, 354)
        Me.pnljobs.TabIndex = 0
        '
        'StiGroupBox1
        '
        Me.StiGroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.StiGroupBox1.Controls.Add(Me.stiexporttoexcel)
        Me.StiGroupBox1.Controls.Add(Me.btnjobsquery)
        Me.StiGroupBox1.Controls.Add(Me.btnshowall)
        Me.StiGroupBox1.Controls.Add(Me.dtgjjobs)
        Me.StiGroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiGroupBox1.Location = New System.Drawing.Point(1, 80)
        Me.StiGroupBox1.Name = "StiGroupBox1"
        Me.StiGroupBox1.Size = New System.Drawing.Size(499, 272)
        Me.StiGroupBox1.TabIndex = 6
        Me.StiGroupBox1.TabStop = False
        '
        'stiexporttoexcel
        '
        Me.stiexporttoexcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.stiexporttoexcel.Location = New System.Drawing.Point(106, 9)
        Me.stiexporttoexcel.Name = "stiexporttoexcel"
        Me.stiexporttoexcel.Size = New System.Drawing.Size(95, 23)
        Me.stiexporttoexcel.TabIndex = 8
        Me.stiexporttoexcel.Text = "Export to excel"
        '
        'btnjobsquery
        '
        Me.btnjobsquery.BackMouseDownColor = System.Drawing.SystemColors.Control
        Me.btnjobsquery.BackMouseOverColor = System.Drawing.SystemColors.Control
        Me.btnjobsquery.BorderBottomColor = System.Drawing.SystemColors.ControlDark
        Me.btnjobsquery.BorderLeftColor = System.Drawing.SystemColors.ControlLight
        Me.btnjobsquery.BorderRightColor = System.Drawing.SystemColors.ControlDark
        Me.btnjobsquery.BorderSize = 1
        Me.btnjobsquery.BorderTopColor = System.Drawing.SystemColors.ControlLight
        Me.btnjobsquery.ButtonSettings = System.Drawing.SystemColors.Control
        Me.btnjobsquery.DrawTextShadow = True
        Me.btnjobsquery.Location = New System.Drawing.Point(205, 9)
        Me.btnjobsquery.Name = "btnjobsquery"
        Me.btnjobsquery.Size = New System.Drawing.Size(130, 23)
        Me.btnjobsquery.TabIndex = 9
        Me.btnjobsquery.Tag = "Go to IT issues"
        Me.btnjobsquery.TextAlign = VSEssentials.VSHotButton.eTextAlign.Center
        Me.btnjobsquery.TextCaption = "View IT issues"
        Me.btnjobsquery.TextColor = System.Drawing.SystemColors.ControlText
        Me.btnjobsquery.TextFont = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnjobsquery.TextLeft = 0
        Me.btnjobsquery.TextOffsetX = 0
        Me.btnjobsquery.TextOffsetY = 0
        Me.btnjobsquery.TextShadowColor = System.Drawing.SystemColors.ControlDark
        Me.btnjobsquery.TextTop = 0
        '
        'btnshowall
        '
        Me.btnshowall.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnshowall.Location = New System.Drawing.Point(4, 9)
        Me.btnshowall.Name = "btnshowall"
        Me.btnshowall.Size = New System.Drawing.Size(98, 23)
        Me.btnshowall.TabIndex = 7
        Me.btnshowall.Text = "Show all jobs"
        '
        'dtgjjobs
        '
        Me.dtgjjobs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtgjjobs.DataMember = ""
        Me.dtgjjobs.FlatMode = True
        Me.dtgjjobs.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgjjobs.Location = New System.Drawing.Point(0, 40)
        Me.dtgjjobs.Name = "dtgjjobs"
        Me.dtgjjobs.ReadOnly = True
        Me.dtgjjobs.Size = New System.Drawing.Size(492, 230)
        Me.dtgjjobs.TabIndex = 10
        '
        'StiGroupBox2
        '
        Me.StiGroupBox2.Controls.Add(Me.chkusedates)
        Me.StiGroupBox2.Controls.Add(Me.dtpto)
        Me.StiGroupBox2.Controls.Add(Me.Label3)
        Me.StiGroupBox2.Controls.Add(Me.dtpfrom)
        Me.StiGroupBox2.Controls.Add(Me.btnsearch)
        Me.StiGroupBox2.Controls.Add(Me.Label2)
        Me.StiGroupBox2.Controls.Add(Me.Label1)
        Me.StiGroupBox2.Controls.Add(Me.MaskedTextBox2)
        Me.StiGroupBox2.Dock = System.Windows.Forms.DockStyle.Top
        Me.StiGroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.StiGroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.StiGroupBox2.Name = "StiGroupBox2"
        Me.StiGroupBox2.Size = New System.Drawing.Size(500, 80)
        Me.StiGroupBox2.TabIndex = 0
        Me.StiGroupBox2.TabStop = False
        '
        'chkusedates
        '
        Me.chkusedates.Checked = True
        Me.chkusedates.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkusedates.Location = New System.Drawing.Point(10, 50)
        Me.chkusedates.Name = "chkusedates"
        Me.chkusedates.Size = New System.Drawing.Size(120, 24)
        Me.chkusedates.TabIndex = 4
        Me.chkusedates.Text = "Use dates"
        '
        'dtpto
        '
        Me.dtpto.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpto.Location = New System.Drawing.Point(250, 24)
        Me.dtpto.Name = "dtpto"
        Me.dtpto.Size = New System.Drawing.Size(120, 20)
        Me.dtpto.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(250, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 10)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "To"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtpfrom
        '
        Me.dtpfrom.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpfrom.Location = New System.Drawing.Point(130, 25)
        Me.dtpfrom.Name = "dtpfrom"
        Me.dtpfrom.Size = New System.Drawing.Size(120, 20)
        Me.dtpfrom.TabIndex = 2
        '
        'btnsearch
        '
        Me.btnsearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsearch.Location = New System.Drawing.Point(171, 48)
        Me.btnsearch.Name = "btnsearch"
        Me.btnsearch.TabIndex = 5
        Me.btnsearch.Text = "Search"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(136, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "From"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Choose search field"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MaskedTextBox2
        '
        Me.MaskedTextBox2.Flags = 0
        Me.MaskedTextBox2.Location = New System.Drawing.Point(629, 328)
        Me.MaskedTextBox2.Mask = ""
        Me.MaskedTextBox2.Name = "MaskedTextBox2"
        Me.MaskedTextBox2.Size = New System.Drawing.Size(160, 20)
        Me.MaskedTextBox2.TabIndex = 0
        '
        'cbojobsearchfield
        '
        Me.cbojobsearchfield.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbojobsearchfield.Items.AddRange(New Object() {"All jobs", "Current", "Completed", "Delivered"})
        Me.cbojobsearchfield.Location = New System.Drawing.Point(7, 26)
        Me.cbojobsearchfield.Name = "cbojobsearchfield"
        Me.cbojobsearchfield.Size = New System.Drawing.Size(120, 22)
        Me.cbojobsearchfield.TabIndex = 1
        '
        'frmjobs
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(500, 354)
        Me.ControlBox = False
        Me.Controls.Add(Me.cbojobsearchfield)
        Me.Controls.Add(Me.pnljobs)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmjobs"
        Me.Text = "Jobs"
        Me.pnljobs.ResumeLayout(False)
        Me.StiGroupBox1.ResumeLayout(False)
        CType(Me.dtgjjobs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StiGroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub
    Private Sub btnjobsquery_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnjobsquery.Click
        Try
            myForms.Main.ToolBar2.Buttons(2).Pushed = True
            myForms.Main.ToolBar2.Buttons(1).Pushed = False
            myForms.Main.ToolBar2.Buttons(0).Pushed = False
            Try
                myForms.qjobs.Close()
                myForms.qjobs = Nothing
            Catch er As Exception
            End Try
            Dim ad As New frmit
            myForms.itissues = ad
            myForms.itissues.Size = myForms.Main.pnlpersonnel.Size
            myForms.itissues.TopLevel = False
            myForms.itissues.Parent = myForms.Main.pnlpersonnel
            myForms.itissues.Dock = DockStyle.Fill
            myForms.itissues.Show()
            myForms.itissues.BringToFront()
        Catch we As Exception

        End Try
    End Sub
    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
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
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            '-----------------try this dave
            Dim dfv, dfv2 As String
            Select Case Me.cbojobsearchfield.Text.Trim()
                Case "All jobs"
                    dfv = ""
                Case "Current"
                    dfv = "Current"
                Case "Completed"
                    dfv = "Complete"
                Case "Delivered"
                    dfv = "Delivered"
                Case Else
                    dfv = ""
            End Select
            '-----------------------dates
            Dim sdate, edate As String
            sdate = Me.dtpfrom.Value.Year & "-" _
            & dtpfrom.Value.Month & "-" _
            & dtpfrom.Value.Day & " " _
            & dtpfrom.Value.Hour & ":" _
            & dtpfrom.Value.Minute & ":" _
            & dtpfrom.Value.Second
            edate = Me.dtpto.Value.Year & "-" _
          & dtpto.Value.Month & "-" _
          & dtpto.Value.Day & " " _
          & dtpto.Value.Hour & ":" _
          & dtpto.Value.Minute & ":" _
          & dtpto.Value.Second
            '-----------------------
            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter
            Dim custDS As DataSet = New DataSet
            Dim adors As New ADODB.Recordset
            Dim str As String
            If Me.chkusedates.Checked = True Then
                str = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status," _
                        & "      rcljobs.date_sniffed " _
                                   & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
                                   & " lower(rcljobs.job_status) like " _
                                   & "'%" & dfv.ToLower & "%'" _
                                   & " where  date_sniffed >= '" & sdate & "'" _
                                   & " and date_sniffed <= '" & edate & "'" _
                                   & " order by rcljobs.client_no"
            Else
                str = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status," _
                                      & "      rcljobs.date_sniffed " _
                                                 & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
                                                 & " lower(rcljobs.job_status) like " _
                                                 & "'%" & dfv.ToLower & "%'" _
                                                 & " where  date_sniffed isnull" _
                                                 & " order by rcljobs.client_no"
            End If

            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leadscv")
            Dim tname As String = custDS.Tables(0).TableName()
            Dim dv As DataViewManager = New DataViewManager

            dv.DataSet = custDS
            Me.dtgjjobs.SetDataBinding(dv, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addcurrjobtablestyle(tname)
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
    Private Sub addcurrjobtablestyle(ByVal tname As String)
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = Me.dtgjjobs.Width
            mywidth = mywidth / 4
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Try

                ' Add a second column style.
                Dim mydesc1 As New DataGridTextBoxColumn
                mydesc1.MappingName = "job_no"
                mydesc1.HeaderText = "Job Number"
                mydesc1.Width = mywidth
                ts1.GridColumnStyles.Add(mydesc1)

            Catch bcv As Exception
            End Try

            ' Add a second column style.
            Dim mydesc4 As New DataGridTextBoxColumn
            mydesc4.MappingName = "job_tittle"
            mydesc4.HeaderText = "Job title"
            mydesc4.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc4)

            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "name"
            myno.HeaderText = "Name"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            Dim mynoc As New DataGridTextBoxColumn
            mynoc.MappingName = "date_sniffed"
            mynoc.HeaderText = "Date"
            mynoc.Width = mywidth
            ts1.GridColumnStyles.Add(mynoc)

            ' Add the DataGridTableStyle objects to the collection.
            dtgjjobs.TableStyles.Clear()
            ts1.AllowSorting = False
            dtgjjobs.TableStyles.Add(ts1)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
    End Sub
    Private Sub btnshowall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnshowall.Click
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
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim custDA As OleDbDataAdapter = New OleDbDataAdapter
            Dim custDS As DataSet = New DataSet
            Dim adors As New ADODB.Recordset
            Dim str As String
            str = "select  rcljobs.client_no,clients.name, rcljobs.job_no, rcljobs.job_tittle,rcljobs.job_status," _
                        & "      rcljobs.date_sniffed " _
                                   & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no  " _
                                  & " order by rcljobs.client_no"


            '--------------oledbdataadapter.fill---------------
            'adoConn.Open(Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            'adoConn.Open("Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;", "", "", -1)
            adors.Open(str, connect, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 1)

            custDA.Fill(custDS, adors, "leadscv")
            Dim tname As String = custDS.Tables(0).TableName()
            Dim dv As DataViewManager = New DataViewManager

            dv.DataSet = custDS
            Me.dtgjjobs.SetDataBinding(dv, tname)
            connect.Close()
            '--------------------this is quite cool---------------------------------------------------------------------

            addcurrjobtablestyle(tname)
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

    Private Sub stiexporttoexcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stiexporttoexcel.Click
        Try
            Dim ds1 As System.Data.DataSet = New System.Data.DataSet
            Dim dv2 As New System.Data.DataViewManager
            dv2 = myForms.qjobs.dtgjjobs.DataSource
            ds1 = dv2.DataSet.Copy
            myForms.Main.dgrid.DataSource = Nothing
            Dim MyTable As New DataTable
            MyTable = ds1.Tables(0)
            Try
                MyTable.Columns.Remove("client_no")
                MyTable.Columns.Remove("job_status")
                'MyTable.Columns.Remove("ano")
                'MyTable.Columns.Remove("job_no1")
                'MyTable.Columns.Remove("milliseconds")
                'MyTable.Columns.Remove("Edit")
                'MyTable.Columns.Remove("ddate1")
                'MyTable.Columns.Remove("isadded")


            Catch cv As Exception
            End Try
            Try
                MyTable.Columns("name").ColumnName = "Name"
                MyTable.Columns("job_no").ColumnName = "Job Number"
                MyTable.Columns("job_tittle").ColumnName = "Job Title"
                MyTable.Columns("date_sniffed").ColumnName = "Date"
                'MyTable.Columns("report_date").ColumnName = "Report Date"

            Catch cv As Exception
            End Try
            Try
                Dim sfd As System.Windows.Forms.SaveFileDialog _
                = New System.Windows.Forms.SaveFileDialog
                sfd.Filter = "Excel files (*.xls)|*.xls"
                sfd.CheckFileExists = False
                sfd.CheckPathExists = True
                sfd.ValidateNames = True
                sfd.ShowDialog()
                Dim m As String = sfd.FileName
                If m.Trim.Length > 0 Then
                    exporttoexcel.exportexcel.exportToExcel(ds1, m)
                    MessageBox.Show("Export successful", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Invalid filename", "Export to excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch we As Exception
            End Try
        Catch ex As Exception
        End Try
    End Sub
End Class
