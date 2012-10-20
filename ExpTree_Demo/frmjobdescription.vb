Imports System
Imports System.Threading
Imports ADODB
Imports System.Data.OleDb
Public Class frmjobdescription
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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btndelete As System.Windows.Forms.Button
    Friend WithEvents dtgdepartments As System.Windows.Forms.DataGrid
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmjobdescription))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btndelete = New System.Windows.Forms.Button
        Me.dtgdepartments = New System.Windows.Forms.DataGrid
        Me.btnadd = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox2.SuspendLayout()
        CType(Me.dtgdepartments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btndelete)
        Me.GroupBox2.Controls.Add(Me.dtgdepartments)
        Me.GroupBox2.Controls.Add(Me.btnadd)
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(328, 262)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Add and edit job description"
        '
        'btndelete
        '
        Me.btndelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btndelete.Location = New System.Drawing.Point(96, 13)
        Me.btndelete.Name = "btndelete"
        Me.btndelete.Size = New System.Drawing.Size(128, 23)
        Me.btndelete.TabIndex = 2
        Me.btndelete.Text = "Delete selected row"
        '
        'dtgdepartments
        '
        Me.dtgdepartments.DataMember = ""
        Me.dtgdepartments.FlatMode = True
        Me.dtgdepartments.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dtgdepartments.Location = New System.Drawing.Point(3, 40)
        Me.dtgdepartments.Name = "dtgdepartments"
        Me.dtgdepartments.Size = New System.Drawing.Size(317, 216)
        Me.dtgdepartments.TabIndex = 4
        '
        'btnadd
        '
        Me.btnadd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnadd.Location = New System.Drawing.Point(5, 12)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.Size = New System.Drawing.Size(88, 23)
        Me.btnadd.TabIndex = 1
        Me.btnadd.Text = "Save changes"
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button2.Location = New System.Drawing.Point(244, 13)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Close"
        '
        'frmjobdescription
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(328, 262)
        Me.Controls.Add(Me.GroupBox2)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmjobdescription"
        Me.Text = "Job description"
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.dtgdepartments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "private members"
    Private Delegate Sub mydelegate()
    Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
#End Region

#Region "job description"
    Private Sub frmjobdescription_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim _thread As Thread = New Thread(AddressOf ld)
            _thread.IsBackground = True
            _thread.Start()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub loaddepart()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str = "select * from jobdescription"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                dtgdepartments.DataSource = Nothing
                If .BOF = False And .EOF = False Then
                    Dim equipDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim equipDS As DataSet = New DataSet

                    equipDA.Fill(equipDS, rs, "decr")
                    Dim tname As String = equipDS.Tables(0).TableName()

                    dtgdepartments.SetDataBinding(equipDS, tname)
                    addequiptablestyle(tname)
                    equipDS.Dispose()
                Else
                    Dim ds As New DataSet
                    Dim dt As New DataTable
                    dt.TableName = "decr"

                    dt.Columns.Add("description")
                    dt.Columns.Add("ano")
                    ds.Tables.Add(dt)
                    dtgdepartments.SetDataBinding(ds, ds.Tables(0).TableName)
                    addequiptablestyle(ds.Tables(0).TableName)

                End If

            End With
            Try
                rs.Close()

            Catch er34b As Exception

            End Try
            Try

                connect.Close()
            Catch er344 As Exception

            End Try
        Catch ex As Exception

        End Try
        Try
            System.GC.Collect()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ld()
        Try
            Me.Invoke(New mydelegate(AddressOf loaddepart))
        Catch ex As Exception

        End Try
    End Sub
    Public Sub addequiptablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = namme
            Dim mywidth As Integer

            mywidth = dtgdepartments.Width - 20
            dtgdepartments.PreferredRowHeight = 33
            mywidth = mywidth

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myn As New DataGridTextBoxColumn
            myn.MappingName = "description"
            myn.HeaderText = "Description"
            myn.Width = mywidth

            ts1.GridColumnStyles.Add(myn)
            ' Add the DataGridTableStyle objects to the collection.
            dtgdepartments.TableStyles.Clear()
            dtgdepartments.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        Dim connect As New ADODB.Connection
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"

            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim ds As New DataSet
            ds = Me.dtgdepartments.DataSource
            Dim int As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim strsql As String
            For y = 0 To int - 1
                If Convert.ToString(ds.Tables(0).Rows(y).Item("ano")).Length < 1 Then
                    strsql += " insert into jobdescription (description) values"
                    strsql += "  ('" & ds.Tables(0).Rows(y).Item("description") & "') ;"
                Else
                    strsql += " update jobdescription set description"
                    strsql += "='" & ds.Tables(0).Rows(y).Item("description") & "' "
                    strsql += " where ano='" & ds.Tables(0).Rows(y).Item("ano") & "'; "
                End If
                Application.DoEvents()
            Next
            Try
                connect.BeginTrans()
                connect.IsolationLevel = ADODB.IsolationLevelEnum.adXactSerializable
                connect.Execute(strsql)
                connect.CommitTrans()
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
        Try
            connect.Close()
        Catch ex As Exception

        End Try
        Try
            Dim _thread As Thread = New Thread(AddressOf ld)
            Try
                If _thread.IsAlive = True Then
                    _thread.Abort()
                End If
            Catch ex As Exception

            End Try

            _thread.IsBackground = True
            _thread.Start()

        Catch ex As Exception

        End Try
        Try
            Dim _thread1 As Thread = New Thread(AddressOf myForms.jobsheet.ld)
            Try
                If _thread1.IsAlive = True Then
                    _thread1.Abort()
                End If
            Catch ex As Exception

            End Try

            _thread1.IsBackground = True
            _thread1.Start()

        Catch ex As Exception

        End Try
        Try
            Dim _thread2 As Thread = New Thread(AddressOf myForms.npersonnel.ld)
            Try
                If _thread2.IsAlive = True Then
                    _thread2.Abort()
                End If
            Catch ex As Exception

            End Try

            _thread2.IsBackground = True
            _thread2.Start()

        Catch ex As Exception

        End Try
     
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            System.GC.Collect()
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
        Try
            Dim x As Boolean = myForms.Main.canmanipulatepersonnel
            If x = False Then
                MessageBox.Show("Not allowed to delete personnel details contact administrator", "Job description", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Catch xcv As Exception

        End Try
        Try
            Me.dtgdepartments.Select(hti.Row)
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
            ds = Me.dtgdepartments.DataSource
            Dim i As Integer = ds.Tables(0).Rows.Count
            Dim y As Integer
            Dim sid, myseconds, str As String
            Dim myrow As System.Data.DataRow
            sid = ds.Tables(0).Rows(hti.Row).Item("ano")
            str = "delete from jobdescription where ano='" & sid & "';"
            Try
                connect.BeginTrans()
                connect.Execute(str)
                connect.CommitTrans()
                myrow = ds.Tables(0).Rows(hti.Row)
                ds.Tables(0).Rows.Remove(myrow)
            Catch cv As Exception
                MessageBox.Show("Department deleted successfully", "Job description", _
               MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try
            Try
                connect.Close()
            Catch er As Exception
            End Try
        Catch ex As Exception
        End Try
        Try
            Dim _thread As Thread = New Thread(AddressOf ld)
            Try
                If _thread.IsAlive = True Then
                    _thread.Abort()
                End If
            Catch ex As Exception

            End Try

            _thread.IsBackground = True
            _thread.Start()

        Catch ex As Exception

        End Try
        Try
            Dim _thread1 As Thread = New Thread(AddressOf myForms.jobsheet.ld)
            Try
                If _thread1.IsAlive = True Then
                    _thread1.Abort()
                End If
            Catch ex As Exception

            End Try

            _thread1.IsBackground = True
            _thread1.Start()

        Catch ex As Exception

        End Try
        Try
            Dim _thread2 As Thread = New Thread(AddressOf myForms.npersonnel.ld)
            Try
                If _thread2.IsAlive = True Then
                    _thread2.Abort()
                End If
            Catch ex As Exception

            End Try

            _thread2.IsBackground = True
            _thread2.Start()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub dtgdepartments_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtgdepartments.MouseDown
        Try
            hti = dtgdepartments.HitTest(New Point(e.X, e.Y))
        Catch ex As System.ArgumentOutOfRangeException
            Try
            Catch er As Exception
            End Try
        Finally
        End Try
    End Sub
#End Region

End Class
