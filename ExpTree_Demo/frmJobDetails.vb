Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Imports System.Data
Imports System.Data.Common
Imports Microsoft.Data.Odbc

Imports System.IO

Imports Microsoft.VisualBasic

Imports System.Data.SqlClient



Imports System.Threading
Imports ADODB
Imports System.Exception
Imports System.SystemException

Public Class frmJobDetails
    Inherits System.Windows.Forms.Form

    Private isdtguploadstablestleexist As Boolean
    Private tsl As New DataGridTableStyle()


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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblJobnumber As System.Windows.Forms.Label
    Friend WithEvents lblJobTittle As System.Windows.Forms.Label
    Friend WithEvents txtGeneralDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnUpload As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dtgUploads As System.Windows.Forms.DataGrid
    Friend WithEvents btnChangeDetails As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lstMail As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSAve As System.Windows.Forms.Button
    Friend WithEvents txtComents As System.Windows.Forms.TextBox
    Friend WithEvents BtnAddNew As System.Windows.Forms.Button
    Friend WithEvents dtgJobSheet As System.Windows.Forms.DataGrid
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnChangeDetails = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.dtgUploads = New System.Windows.Forms.DataGrid()
        Me.btnUpload = New System.Windows.Forms.Button()
        Me.txtGeneralDesc = New System.Windows.Forms.TextBox()
        Me.lblJobTittle = New System.Windows.Forms.Label()
        Me.lblJobnumber = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lstMail = New System.Windows.Forms.CheckedListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.BtnAddNew = New System.Windows.Forms.Button()
        Me.dtgJobSheet = New System.Windows.Forms.DataGrid()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.txtComents = New System.Windows.Forms.TextBox()
        Me.btnSAve = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dtgUploads, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.dtgJobSheet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnChangeDetails, Me.GroupBox2, Me.btnUpload, Me.txtGeneralDesc, Me.lblJobTittle, Me.lblJobnumber})
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(592, 184)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'btnChangeDetails
        '
        Me.btnChangeDetails.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnChangeDetails.Location = New System.Drawing.Point(8, 152)
        Me.btnChangeDetails.Name = "btnChangeDetails"
        Me.btnChangeDetails.Size = New System.Drawing.Size(184, 24)
        Me.btnChangeDetails.TabIndex = 4
        Me.btnChangeDetails.Text = "Change Job Details"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.AddRange(New System.Windows.Forms.Control() {Me.dtgUploads})
        Me.GroupBox2.Location = New System.Drawing.Point(400, 72)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(184, 104)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        '
        'dtgUploads
        '
        Me.dtgUploads.CaptionBackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.dtgUploads.DataMember = ""
        Me.dtgUploads.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtgUploads.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.dtgUploads.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgUploads.Location = New System.Drawing.Point(8, 17)
        Me.dtgUploads.Name = "dtgUploads"
        Me.dtgUploads.ReadOnly = True
        Me.dtgUploads.RowHeadersVisible = False
        Me.dtgUploads.Size = New System.Drawing.Size(168, 79)
        Me.dtgUploads.TabIndex = 3
        '
        'btnUpload
        '
        Me.btnUpload.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpload.Location = New System.Drawing.Point(400, 48)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(144, 24)
        Me.btnUpload.TabIndex = 2
        Me.btnUpload.Text = "Upload File"
        '
        'txtGeneralDesc
        '
        Me.txtGeneralDesc.AcceptsReturn = True
        Me.txtGeneralDesc.AcceptsTab = True
        Me.txtGeneralDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGeneralDesc.Location = New System.Drawing.Point(8, 56)
        Me.txtGeneralDesc.Multiline = True
        Me.txtGeneralDesc.Name = "txtGeneralDesc"
        Me.txtGeneralDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtGeneralDesc.Size = New System.Drawing.Size(368, 88)
        Me.txtGeneralDesc.TabIndex = 1
        Me.txtGeneralDesc.Text = ""
        '
        'lblJobTittle
        '
        Me.lblJobTittle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobTittle.Location = New System.Drawing.Point(288, 17)
        Me.lblJobTittle.Name = "lblJobTittle"
        Me.lblJobTittle.Size = New System.Drawing.Size(296, 31)
        Me.lblJobTittle.TabIndex = 1
        Me.lblJobTittle.Text = "Job Tittle"
        '
        'lblJobnumber
        '
        Me.lblJobnumber.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobnumber.Location = New System.Drawing.Point(8, 17)
        Me.lblJobnumber.Name = "lblJobnumber"
        Me.lblJobnumber.Size = New System.Drawing.Size(264, 31)
        Me.lblJobnumber.TabIndex = 0
        Me.lblJobnumber.Text = "Job Number"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.AddRange(New System.Windows.Forms.Control() {Me.lstMail, Me.Button1})
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(608, 8)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(160, 184)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        '
        'lstMail
        '
        Me.lstMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstMail.Location = New System.Drawing.Point(8, 48)
        Me.lstMail.Name = "lstMail"
        Me.lstMail.Size = New System.Drawing.Size(144, 109)
        Me.lstMail.TabIndex = 8
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(8, 17)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(144, 24)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Send Email"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox4.Controls.AddRange(New System.Windows.Forms.Control() {Me.BtnAddNew, Me.dtgJobSheet})
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(8, 200)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(544, 288)
        Me.GroupBox4.TabIndex = 2
        Me.GroupBox4.TabStop = False
        '
        'BtnAddNew
        '
        Me.BtnAddNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddNew.Location = New System.Drawing.Point(8, 8)
        Me.BtnAddNew.Name = "BtnAddNew"
        Me.BtnAddNew.Size = New System.Drawing.Size(136, 24)
        Me.BtnAddNew.TabIndex = 5
        Me.BtnAddNew.Text = "Add New"
        '
        'dtgJobSheet
        '
        Me.dtgJobSheet.DataMember = ""
        Me.dtgJobSheet.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtgJobSheet.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgJobSheet.Location = New System.Drawing.Point(8, 32)
        Me.dtgJobSheet.Name = "dtgJobSheet"
        Me.dtgJobSheet.ReadOnly = True
        Me.dtgJobSheet.Size = New System.Drawing.Size(520, 248)
        Me.dtgJobSheet.TabIndex = 6
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtComents, Me.btnSAve})
        Me.GroupBox5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(560, 200)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(216, 288)
        Me.GroupBox5.TabIndex = 3
        Me.GroupBox5.TabStop = False
        '
        'txtComents
        '
        Me.txtComents.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtComents.Location = New System.Drawing.Point(8, 56)
        Me.txtComents.Multiline = True
        Me.txtComents.Name = "txtComents"
        Me.txtComents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtComents.Size = New System.Drawing.Size(200, 224)
        Me.txtComents.TabIndex = 10
        Me.txtComents.Text = ""
        '
        'btnSAve
        '
        Me.btnSAve.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSAve.Location = New System.Drawing.Point(16, 17)
        Me.btnSAve.Name = "btnSAve"
        Me.btnSAve.Size = New System.Drawing.Size(176, 24)
        Me.btnSAve.TabIndex = 9
        Me.btnSAve.Text = "Save Comments"
        '
        'Timer1
        '
        '
        'frmJobDetails
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(784, 510)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.GroupBox5, Me.GroupBox4, Me.GroupBox3, Me.GroupBox1})
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaximizeBox = False
        Me.Name = "frmJobDetails"
        Me.Text = "Job Details"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dtgUploads, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.dtgJobSheet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim mailFrm As New frmMail()
        mailFrm.Show()
    End Sub

    Private Sub frmJobDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call dbconnect()
        Me.lblJobnumber.Text = myjobno
        Me.lblJobTittle.Text = myjobtitle
        Call loaddtgjobsheet()

        Call addemails()
        Call loadgeneraldesc()
        Call loaddtguploads()
        Call readtextfile()
        Timer1.Interval = 1000
        Timer1.Start()
    End Sub
    Private Sub loadgeneraldesc()
        Try
            Dim cmd As New ADODB.Command()
            Dim rs As New ADODB.Recordset()

            With cmd
                .ActiveConnection = connect
                .CommandType = ADODB.CommandTypeEnum.adCmdText
                Dim str As String
                str = "select * from rcljobs" _
                & " where job_no='" & lblJobnumber.Text & "'"
                .CommandText = str
                rs = .Execute

            End With
            With rs
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    Dim a() As String
                    Dim str As String
                    str = (" Job Title: ") & .Fields("job_tittle").Value & "|" & _
                    " Job Status: " & .Fields("job_status").Value & "|" & _
                    " Department: " & .Fields("dept").Value & "|" & _
                    " Start Date: " & .Fields("sdate").Value & "|" & _
                    " Deadline: " & .Fields("deadline").Value & "|" & _
                    " End Date: " & .Fields("fdate").Value & "|" & _
                    " Cost Agreed: " & .Fields("costagreed").Value & "|"
                    'txtGeneralDesc.Text() 
                    a = str.Split("|")
                    Me.txtGeneralDesc.Lines = a



                End If
            End With
            rs.Close()
            rs = Nothing
            cmd.ActiveConnection = Nothing
        Catch ex As Exception

        End Try
    End Sub
    Private Sub addemails()
        Try
            Dim rs As New ADODB.Recordset()
            Dim cmd As New ADODB.Command()
            ' Shutdown the painting of the ListBox as items are added.
            lstMail.BeginUpdate()
            With cmd
                .ActiveConnection = connect
                .CommandType = ADODB.CommandTypeEnum.adCmdText
                Dim str As String
                str = "select email from contacts" _
                & " where client_no='" & myclientno & "'"
                .CommandText = str
                rs = .Execute(str)

            End With
            With rs
                If .BOF = False And .EOF = False Then
                    .MoveFirst()

                    While .EOF = False

                        lstMail.Items.Add(.Fields("email").Value).ToString()
                        .MoveNext()

                    End While

                End If
                .Close()
                rs = Nothing
                cmd.ActiveConnection = Nothing

            End With
            rs = New ADODB.Recordset()
            cmd = New ADODB.Command()
            With cmd
                .ActiveConnection = connect
                .CommandType = ADODB.CommandTypeEnum.adCmdText
                Dim str As String
                str = "select name from seccheck"

                .CommandText = str
                rs = .Execute(str)
            End With
            'lstMail.Items.Add("Ramani Staff E-Mail addresses").ToString()
            With rs
                If .BOF = False And .EOF = False Then
                    .MoveFirst()

                    While .EOF = False

                        lstMail.Items.Add(.Fields("name").Value).ToString()
                        .MoveNext()

                    End While

                End If
                .Close()
                rs = Nothing
                cmd.ActiveConnection = Nothing

            End With
            lstMail.EndUpdate()
            ' Allow the ListBox to repaint and display the new items.
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub loaddtguploads()
        Dim currentCursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            'Fill the DataSet

            Dim dAdap As New OdbcDataAdapter("select" _
            & " filename" _
            & " from fileatts" _
            & " where job_no ='" & lblJobnumber.Text.Trim & "'", "dsn=RCL_DB")  '"dsn=RCL_DB"
            Dim dSet As New DataSet()

            dAdap.Fill(dSet, "fileatts")

            dtgUploads.BeginInit()
            dtgUploads.EndInit()


            dtgUploads.SetDataBinding(dSet, "fileatts")
            Call addtablestyledtgUploads(dtgUploads)
        Catch ex As Exception
            MessageBox.Show(Text:=ex.ToString)
        Finally
            Cursor.Current = currentCursor
        End Try
    End Sub
    Private Sub loaddtgjobsheet()
        Dim currentCursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            'Fill the DataSet
            Dim str As String
            str = "select" _
            & " techname,description,dept,date,duration" _
            & " from jobsheet" _
            & " where job_no ='" & lblJobnumber.Text.Trim & "'"
            Dim dAdap As New OdbcDataAdapter("select" _
            & " techname,description,dept,date,duration" _
            & " from jobsheet" _
            & " where job_no ='" & lblJobnumber.Text.Trim & "'", "dsn=RCL_DB")  '"dsn=RCL_DB"
            Dim dSet As New DataSet()

            dAdap.Fill(dSet, "jobsheet")


            dtgJobSheet.SetDataBinding(dSet, "jobsheet")
            Call addtablestyle(dtgJobSheet)
        Catch ex As Exception
            MessageBox.Show(Text:=ex.ToString)
        Finally
            Cursor.Current = currentCursor
        End Try

    End Sub
    Private Sub addtablestyle(ByVal mygrid As DataGrid)
        ' Create an empty DataGridTableStyle & set mapping name to table.
        Dim tableStyle As New DataGridTableStyle()
        tableStyle.MappingName = "jobsheet"

        Dim mywidth As Integer
        mywidth = mygrid.Width
        mywidth = mywidth / 6



        'Technician Name
        Dim column As New FormattableTextBoxColumn()
        column.MappingName = "techname"
        column.HeaderText = "TECHNICIAN RESPONSIBLE"
        column.Width = mywidth

        tableStyle.GridColumnStyles.Add(column)

        'DESCRIPTION
        column = New FormattableTextBoxColumn()
        column.MappingName = "description"
        column.HeaderText = "DESCRIPTION"
        column.Width = mywidth

        tableStyle.GridColumnStyles.Add(column)

        'Department
        column = New FormattableTextBoxColumn()
        column.MappingName = "dept"
        column.HeaderText = "DEPARTMENT"
        column.Width = mywidth

        tableStyle.GridColumnStyles.Add(column)
        'DATE
        column = New FormattableTextBoxColumn()
        column.MappingName = "date"
        column.HeaderText = "DATE"
        column.Width = mywidth

        tableStyle.GridColumnStyles.Add(column)
        'DURATION
        column = New FormattableTextBoxColumn()
        column.MappingName = "duration"
        column.HeaderText = "DURATION"
        column.Width = mywidth

        tableStyle.GridColumnStyles.Add(column)



        mygrid.TableStyles.Add(tableStyle)
    End Sub
    Private Sub addtablestyledtgUploads(ByVal mygrid As DataGrid)
        ' Create an empty DataGridTableStyle & set mapping name to table.
        Dim tableStyle As New DataGridTableStyle()
        tableStyle.MappingName = "fileatts"
        tsl.MappingName = "invoices"

        Dim mywidth As Integer
        mywidth = mygrid.Width
        ' mywidth = mywidth / 6



        'Technician Name
        Dim column As New FormattableTextBoxColumn()
        column.MappingName = "filename"
        column.HeaderText = " "

        column.Width = mywidth
        tableStyle.GridColumnStyles.Add(column)


        'tsl.GridColumnStyles.Add(column)
        mygrid.TableStyles.Add(tableStyle)
    End Sub
    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Dim currentcursor As Cursor = Cursor.Current
        Dim filename2 As String
        Dim a() As String
        Dim i As Integer
        Try
            Cursor.Current = Cursors.WaitCursor
            If Me.lblJobnumber.Text = Nothing Or lblJobnumber.Text.Trim.Length = 0 Then
                MessageBox.Show(Text:="The job number is not valid", _
                                      caption:="File Uploading", buttons:=MessageBoxButtons.OK, _
                                      Icon:=MessageBoxIcon.Information)
                Exit Try
            End If
            Dim filename As String
            OpenFileDialog1.FileName = ""
            Me.OpenFileDialog1.ShowDialog()

            filename = OpenFileDialog1.FileName

            If filename = "" Or filename.Trim.Length = 0 Then
                MessageBox.Show(Text:="File upload was unsuccessful" _
                & "since you did not select a file", _
                caption:="File Uploading", buttons:=MessageBoxButtons.OK, _
            Icon:=MessageBoxIcon.Information)
                Exit Try
            End If



            a = filename.Split("\")
            i = a.GetUpperBound(0)

            filename2 = ConfigurationSettings.AppSettings("invoicespath") & "\" & a(i)
            File.Copy(sourcefilename:=filename, destfilename:=filename2, overwrite:=False)
            Dim str As String

            connect.BeginTrans()
            str = "insert into fileatts"
            str = str & " values ( "
            str = str & "'" & lblJobnumber.Text & "',"
            str = str & "'" & a(i) & "'"
            str = str & ")"
            connect.Execute(str)
            connect.CommitTrans()

            MessageBox.Show(Text:="File successfully uploaded", _
                      caption:="File Uploading", buttons:=MessageBoxButtons.OK, _
                      Icon:=MessageBoxIcon.Information)
            'dtgUploads.TableStyles.Remove()
            'isdtguploadstablestleexist = True
            'Call loaddtguploads()
        Catch ex As IOException
            MessageBox.Show(Text:="File Name already exists please select another one", _
            caption:="File Uploading", buttons:=MessageBoxButtons.OK, _
            Icon:=MessageBoxIcon.Information)
        Catch t As NotSupportedException
            MessageBox.Show(Text:="Dont use this ':'", _
                      caption:="Error", buttons:=MessageBoxButtons.OK, _
                      Icon:=MessageBoxIcon.Information)
        Catch d As PathTooLongException
            MessageBox.Show(Text:="Transfer file to 'Local Disk C:\' and then try again", _
                      caption:="Error", buttons:=MessageBoxButtons.OK, _
                      Icon:=MessageBoxIcon.Information)
        Catch es As DirectoryNotFoundException
            MessageBox.Show(Text:="Directory not found", _
                      caption:="Error", buttons:=MessageBoxButtons.OK, _
                      Icon:=MessageBoxIcon.Information)
        Catch em As ArgumentException
            MessageBox.Show(Text:="Use a(A) to z(Z) in file names only ", _
                      caption:="Error", buttons:=MessageBoxButtons.OK, _
                      Icon:=MessageBoxIcon.Information)


        Catch ev As Exception
            Dim msg As Boolean
            Try
                connect.RollbackTrans()
                MessageBox.Show(Text:="Database could not be updated", _
                    caption:="File Uploading", buttons:=MessageBoxButtons.OK, _
                    Icon:=MessageBoxIcon.Information)
                msg = True
                File.Delete(filename2)
            Catch f As Exception
                If msg = True Then
                    MessageBox.Show(Text:="Cannot delete the file named: " & a(i), _
                    caption:="Error", buttons:=MessageBoxButtons.OK, _
                    Icon:=MessageBoxIcon.Information)
                End If
            End Try



        Finally
            Cursor.Current = currentcursor
        End Try

    End Sub

    Private Sub btnSAve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSAve.Click
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim a() As String
            a = Me.txtComents.Lines
            Dim filename As String
            Dim nv As NameValueCollection
            Dim i, j As Integer
            i = a.GetUpperBound(0)
            nv = ConfigurationSettings.AppSettings
            filename = nv("invoicespath")
            filename = filename & "\" & lblJobnumber.Text.Trim & ".txt"

            If File.Exists(filename) = False Then
                Dim sr As StreamWriter = File.CreateText(filename)
                For j = 0 To i
                    sr.WriteLine(a(j))
                Next
                sr.Close()
            Else
                File.Delete(filename)
                Dim sr As StreamWriter = File.CreateText(filename)
                For j = 0 To i
                    sr.WriteLine(a(j))
                Next
                sr.Close()

            End If
        Catch ex As Exception

        Finally
            Cursor.Current = currentcursor

        End Try
    End Sub
    Private Sub readtextfile()
        Try
            Dim filename As String
            Dim nv As NameValueCollection


            nv = ConfigurationSettings.AppSettings
            filename = nv("invoicespath")
            filename = filename & "\" & lblJobnumber.Text.Trim & ".txt"
            If File.Exists(filename) = True Then

                '    Dim sr As StreamReader
                '    sr = File.OpenText(filename)
                '    Dim input As String
                '    Dim a() As String
                '    Dim str As String

                '    input = sr.ReadLine()
                '    str = input

                '    While Not sr.Peek <> -1


                '        input = sr.ReadLine()
                '        str = str & "|" & input

                '    End While
                '    sr.Close()
                '    a = str.Split("|")

                '    Me.txtComents.Lines = a

                FileOpen(1, filename, OpenMode.Input)
                Dim x As String
                Do While Not EOF(1)
                    x = LineInput(1)
                    Me.txtComents.AppendText(x + vbCrLf)
                Loop
                FileClose(1)
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtGeneralDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGeneralDesc.TextChanged

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        '' Visual Basic .NET 
        'Dim datobj As New System.Windows.Forms.DataObject()

        'datobj.SetData(System.Windows.Forms.DataFormats.Text, "")

        ''Timer1.Stop()
    End Sub

    Private Sub dtgUploads_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtgUploads.Navigate

    End Sub
End Class

