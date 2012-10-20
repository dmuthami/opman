
Imports System
Imports System.Data
Imports ADODB


Imports System.IO
Public Class frmpersonneladmin
    Inherits System.Windows.Forms.Form
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo

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
    Friend WithEvents pnlgrid As System.Windows.Forms.Panel
    Friend WithEvents dtgpersonnel As System.Windows.Forms.DataGrid
    Friend WithEvents pnlsearch As System.Windows.Forms.Panel
    Friend WithEvents txtparam As System.Windows.Forms.TextBox
    Friend WithEvents lblsearchparameters As System.Windows.Forms.Label
    Friend WithEvents StiGroupLine1 As Stimulsoft.Controls.StiGroupLine
    Friend WithEvents btndeletepersonnel As System.Windows.Forms.Button
    Friend WithEvents btnaddnew As System.Windows.Forms.Button
    Friend WithEvents btnsearch As System.Windows.Forms.Button
    Friend WithEvents btnshowall As System.Windows.Forms.Button
    Friend WithEvents errp As System.Windows.Forms.ErrorProvider
    Friend WithEvents btnexporttoexcel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlgrid = New System.Windows.Forms.Panel
        Me.dtgpersonnel = New System.Windows.Forms.DataGrid
        Me.pnlsearch = New System.Windows.Forms.Panel
        Me.btnexporttoexcel = New System.Windows.Forms.Button
        Me.btnshowall = New System.Windows.Forms.Button
        Me.btnsearch = New System.Windows.Forms.Button
        Me.btnaddnew = New System.Windows.Forms.Button
        Me.btndeletepersonnel = New System.Windows.Forms.Button
        Me.txtparam = New System.Windows.Forms.TextBox
        Me.lblsearchparameters = New System.Windows.Forms.Label
        Me.errp = New System.Windows.Forms.ErrorProvider
        Me.pnlgrid.SuspendLayout()
        CType(Me.dtgpersonnel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlsearch.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlgrid
        '
        Me.pnlgrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlgrid.Controls.Add(Me.dtgpersonnel)
        Me.pnlgrid.Location = New System.Drawing.Point(8, 88)
        Me.pnlgrid.Name = "pnlgrid"
        Me.pnlgrid.Size = New System.Drawing.Size(472, 272)
        Me.pnlgrid.TabIndex = 0
        '
        'dtgpersonnel
        '
        Me.dtgpersonnel.CaptionText = "Personnel"
        Me.dtgpersonnel.DataMember = ""
        Me.dtgpersonnel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dtgpersonnel.FlatMode = True
        Me.dtgpersonnel.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgpersonnel.Location = New System.Drawing.Point(0, 0)
        Me.dtgpersonnel.Name = "dtgpersonnel"
        Me.dtgpersonnel.ReadOnly = True
        Me.dtgpersonnel.Size = New System.Drawing.Size(472, 272)
        Me.dtgpersonnel.TabIndex = 7
        '
        'pnlsearch
        '
        Me.pnlsearch.Controls.Add(Me.btnexporttoexcel)
        Me.pnlsearch.Controls.Add(Me.btnshowall)
        Me.pnlsearch.Controls.Add(Me.btnsearch)
        Me.pnlsearch.Controls.Add(Me.btnaddnew)
        Me.pnlsearch.Controls.Add(Me.btndeletepersonnel)
        Me.pnlsearch.Controls.Add(Me.txtparam)
        Me.pnlsearch.Controls.Add(Me.lblsearchparameters)
        Me.pnlsearch.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlsearch.Location = New System.Drawing.Point(0, 0)
        Me.pnlsearch.Name = "pnlsearch"
        Me.pnlsearch.Size = New System.Drawing.Size(488, 80)
        Me.pnlsearch.TabIndex = 0
        '
        'btnexporttoexcel
        '
        Me.btnexporttoexcel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnexporttoexcel.Location = New System.Drawing.Point(314, 56)
        Me.btnexporttoexcel.Name = "btnexporttoexcel"
        Me.btnexporttoexcel.Size = New System.Drawing.Size(128, 23)
        Me.btnexporttoexcel.TabIndex = 6
        Me.btnexporttoexcel.Text = "Export to excel"
        '
        'btnshowall
        '
        Me.btnshowall.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnshowall.Location = New System.Drawing.Point(78, 56)
        Me.btnshowall.Name = "btnshowall"
        Me.btnshowall.Size = New System.Drawing.Size(128, 23)
        Me.btnshowall.TabIndex = 4
        Me.btnshowall.Text = "Show all personnel"
        '
        'btnsearch
        '
        Me.btnsearch.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnsearch.Location = New System.Drawing.Point(227, 12)
        Me.btnsearch.Name = "btnsearch"
        Me.btnsearch.Size = New System.Drawing.Size(40, 32)
        Me.btnsearch.TabIndex = 2
        Me.btnsearch.Text = "Go"
        '
        'btnaddnew
        '
        Me.btnaddnew.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnaddnew.Location = New System.Drawing.Point(2, 56)
        Me.btnaddnew.Name = "btnaddnew"
        Me.btnaddnew.TabIndex = 3
        Me.btnaddnew.Text = "Add new"
        '
        'btndeletepersonnel
        '
        Me.btndeletepersonnel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndeletepersonnel.Location = New System.Drawing.Point(210, 56)
        Me.btndeletepersonnel.Name = "btndeletepersonnel"
        Me.btndeletepersonnel.Size = New System.Drawing.Size(104, 23)
        Me.btndeletepersonnel.TabIndex = 5
        Me.btndeletepersonnel.Text = "Delete personnel"
        '
        'txtparam
        '
        Me.txtparam.Location = New System.Drawing.Point(8, 24)
        Me.txtparam.Name = "txtparam"
        Me.txtparam.Size = New System.Drawing.Size(216, 20)
        Me.txtparam.TabIndex = 1
        Me.txtparam.Text = ""
        '
        'lblsearchparameters
        '
        Me.lblsearchparameters.Location = New System.Drawing.Point(32, 8)
        Me.lblsearchparameters.Name = "lblsearchparameters"
        Me.lblsearchparameters.Size = New System.Drawing.Size(176, 16)
        Me.lblsearchparameters.TabIndex = 0
        Me.lblsearchparameters.Text = "Type search parameter"
        Me.lblsearchparameters.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'errp
        '
        Me.errp.ContainerControl = Me
        '
        'frmpersonneladmin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(488, 370)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlsearch)
        Me.Controls.Add(Me.pnlgrid)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmpersonneladmin"
        Me.Text = "Personnel Administrator"
        Me.pnlgrid.ResumeLayout(False)
        CType(Me.dtgpersonnel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlsearch.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    <System.STAThread()> _
                Public Shared Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Run()
    End Sub

#Region "Private Members"
    Private wwhich As String = "0"
    '---0 is show all ,1 is Go
#End Region

#Region "personnel admin"
    Private Sub frmpersonneladmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call loadgrid()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub loadgrid()
        Try
            Dim Tasks As New taskclass
            Dim Thread4 As New System.Threading.Thread( _
                AddressOf taskclass.personnelinvoke)

            Thread4.Start() ' Start the new thread.
            'Thread1.Join() ' Wait for thread 1 to finish.
            '' Display the return value.
            'MsgBox("Thread 1 returned the value " & Tasks.RetVal)
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

                    mycell.RowNumber = Me.dtgpersonnel.CurrentRowIndex
                    mycell.ColumnNumber = 0

                    returnhittest = mycell.RowNumber & "|" & mycell.ColumnNumber
                Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                    Dim mycell As New DataGridCell

                    mycell.RowNumber = Me.dtgpersonnel.CurrentRowIndex
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
    Private Sub dtgpersonnel_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgpersonnel.MouseDown
        Try
            hti = dtgpersonnel.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
            Try

            Catch er As Exception

            End Try
        Finally
        End Try
    End Sub
    Private Sub dtgpersonnel_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtgpersonnel.DoubleClick
        Try
            Dim results As String
            results = returnhittest(hti.Type)
            If results <> "" Then

                Dim mycell As New DataGridCell
                Dim a() As String
                a = results.Split("|")
                mycell.RowNumber = CInt(a(0))
                mycell.ColumnNumber = 1

                If hasloadedjobsheet = False Then
                    If Convert.IsDBNull(dtgpersonnel(mycell)) = False Then
                        Dim form As New frmjobsheet
                        myForms.jobsheet = form
                        myForms.jobsheet.myid = dtgpersonnel(mycell)
                        myForms.jobsheet.txtidno.Text = myForms.jobsheet.myid
                        myForms.jobsheet.StartPosition = FormStartPosition.CenterParent
                        Dim task34 As New taskclass
                        task34.hisid_no = myForms.jobsheet.myid
                        myForms.jobsheet.Show()
                    End If
                    hasloadedjobsheet = True
                Else
                    myForms.jobsheet.myid = dtgpersonnel(mycell)
                    Dim task35 As New taskclass
                    task35.hisid_no = myForms.jobsheet.myid
                    task35.mid = myForms.jobsheet.myid
                    Dim Tasks As New taskclass
                    Dim Thread5 As New System.Threading.Thread( _
                       AddressOf taskclass.jobsheetinvoke)
                    Thread5.Start()
                    '-------------reload cbos
                    Dim Thread5bn As New System.Threading.Thread( _
                            AddressOf taskclass.cbosinvoke)
                    Thread5bn.Start()
                    '------------
                End If
                Dim ds As New System.Data.DataSet
                ds = Me.dtgpersonnel.DataSource
                myForms.jobsheet.namme = ds.Tables(0).Rows(hti.Row).Item("namme")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())

        End Try
    End Sub
    Private Sub btndeletepersonnel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndeletepersonnel.Click
        Try
            dtgpersonnel.Select(hti.Row)
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
            ds = Me.dtgpersonnel.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hti.Row).Item("id_no")
            str = " delete from personnel_info where "
            str += " id_no='" & sid & "';"

            str += " delete from seccheck where "
            str += " id_no='" & sid & "';"

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
        Try
            Dim Tasks As New taskclass
            Tasks.loadadmincontrols = False
            Dim Threada1 As New System.Threading.Thread( _
                AddressOf taskclass.admininvoke)
            Threada1.Start()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btnaddnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaddnew.Click
        Try
            Dim form As New frmnewpersonnel
            myForms.npersonnel = form
            myForms.npersonnel.StartPosition = FormStartPosition.CenterParent
            myForms.npersonnel.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Sub btnsearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsearch.Click
        wwhich = "1"
        Try
            Dim Tasks As New taskclass
            Tasks.searchstr = Me.txtparam.Text
            Dim Thread99 As New System.Threading.Thread( _
                AddressOf taskclass.searchinvoke)

            Thread99.Start()
        Catch ert As Exception
            MessageBox.Show(ert.Message.ToString() & vbCrLf _
                     & ert.InnerException().ToString() & vbCrLf _
                     & ert.StackTrace.ToString())
        End Try
    End Sub
    Private Sub dtgpersonnel_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs) Handles dtgpersonnel.Navigate

    End Sub
    Private Sub btnshowall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnshowall.Click
        wwhich = "0"
        Try
            Call loadgrid()
        Catch xc As Exception

        End Try
    End Sub
    Private Sub btnexporttoexcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexporttoexcel.Click
        Try
            Dim ds1 As System.Data.DataSet = New System.Data.DataSet
            Dim ds2 As System.Data.DataSet = New System.Data.DataSet
            ds2 = myForms.adminform.dtgpersonnel.DataSource
            myForms.Main.dgrid.DataSource = Nothing
            ds1 = ds2.Copy
            Dim MyTable As New DataTable
            MyTable = ds1.Tables(0)

            Try
                MyTable.Columns("namme").ColumnName = "Name"
                MyTable.Columns("id_no").ColumnName = "Id No"
                MyTable.Columns("hourly_rate").ColumnName = "Hourly_rate"
                MyTable.Columns("gender").ColumnName = "Job Description"
                MyTable.Columns("phone_no").ColumnName = "Phone Number"
                MyTable.Columns("mobile_no").ColumnName = "Mobile Number"
                MyTable.Columns("email").ColumnName = "E mail"
                MyTable.Columns("pin_no").ColumnName = "Pin Number"
                MyTable.Columns("birthday").ColumnName = "Birthday"
                MyTable.Columns("contract_end").ColumnName = "Contract End"

                MyTable.Columns("nssf_no").ColumnName = "Nssf No"
                MyTable.Columns("nhif_no").ColumnName = "Nhif No"
                MyTable.Columns("medical_cover").ColumnName = "Medical Cover"
                MyTable.Columns("dateofemployment").ColumnName = "Date of employment"
                MyTable.Columns("nextofkin").ColumnName = "Next Of Kin"
                MyTable.Columns("dateoftermination").ColumnName = "Date Of Termination"
                MyTable.Columns("comments").ColumnName = "Comments"

                MyTable.Columns.Remove("imagefile")
                MyTable.Columns.Remove("ano")
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
        Try
            Select Case wwhich
                Case "0"
                    Call loadgrid()
                Case Else
                    btnsearch_Click(Me, e)
            End Select
        Catch cv As Exception
        End Try
        Try
            System.GC.Collect()
        Catch qw As Exception

        End Try
        'Me.btnview_Click(Me, e)
    End Sub
#End Region

#Region "validation"
    Private Sub txtparam_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtparam.KeyPress
        Try
            Dim vt As New validation
            If vt._validatetextbox(e) = True Then
                Me.errp.SetError(Me.txtparam, _
                                      "not allowed chars: ''','%','*','\','*','1'")
                'this.statusBar1.Text="not allowed char..."+e.KeyChar;
            Else
                Me.errp.SetError(Me.txtparam, "")
            End If
        Catch xc As Exception

        End Try
    End Sub
#End Region


End Class
