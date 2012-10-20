
Imports ADODB
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Threading.Thread

Public Class taskclass

#Region "personnel"

#Region "timesheet"
    Public Shared Strid_no As String
    Public Shared RetVal As Boolean
    Public Delegate Sub mydelegate()
    Public Delegate Sub mydelegate1()
    Public Shared Sub SomeTask()
        Try
            ' Use the StrArg field as an argument.
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

            Dim rs As New ADODB.Recordset()
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.CursorType = CursorTypeEnum.adOpenStatic
            Dim str As String
            str = "select namme from personnel_info where id_no='" & Strid_no & "'"
            rs.Open(str, connect)
            If rs.BOF = False And rs.EOF = False Then
                Dim sname As String
                If Convert.IsDBNull(rs.Fields("namme").Value) = False Then
                    sname = rs.Fields("namme").Value
                    myForms.timesheet.grpname.Text = sname
                    myForms.Main._name = sname
                Else
                    sname = "null"
                    myForms.timesheet.grpname.Text = sname
                End If
            End If
            Try
                rs.Close()
                connect.Close()
            Catch ex444 As Exception
                MessageBox.Show(ex444.Message.ToString() & vbCrLf _
                       & ex444.InnerException().ToString() & vbCrLf _
                       & ex444.StackTrace.ToString())
            End Try


            'MsgBox("The StrArg contains the string " & Strid_no)
            'RetVal = True ' Set a return value in the return argument.
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
          & ex.InnerException().ToString() & vbCrLf _
          & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub somejob()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim rsjob As New ADODB.Recordset()
            With rsjob
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenKeyset
                Dim str As String = " SELECT rcljobs.*,  clients.name" _
                                                & " FROM clients INNER JOIN" _
                                                & "  rcljobs ON clients.client_no = rcljobs.client_no"
                str += " and lower(rcljobs.job_status) like '%" & "curren" & "%'"
                str += " and rcljobs.job_tittle <> '" & "" & "'"
                str += " order by rcljobs.job_tittle asc"
                Dim strj As String
                strj = "select job_tittle,job_no from rcljobs where lower(job_status)='" & "current" & "'"
                strj += " and job_tittle<>'" & "" & "'"
                strj += " order by job_tittle asc "
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            myForms.timesheet.cbojob.Items.Add((Convert.ToString(.Fields("job_no").Value) & " : " & _
                             Convert.ToString(.Fields("job_tittle").Value) & " : " & _
                             Convert.ToString(.Fields("name").Value)))
                            myForms.timesheet.cboinvisible.Items.Add(.Fields("job_no").Value)

                        Catch es300 As Exception
                        End Try

                        Application.DoEvents()
                        .MoveNext()
                    End While
                Else
                End If
                Try
                    .Close()
                    connect.Close()
                Catch ex245 As Exception

                End Try
            End With
        Catch exsj As Exception
            MessageBox.Show(exsj.Message.ToString() & vbCrLf _
            & exsj.InnerException().ToString() & vbCrLf _
            & exsj.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub myinvoke()
        Try
            myForms.timesheet.dtgtimesheet.Invoke(New mydelegate(AddressOf somedata))
        Catch ex As Exception
            '-----------crashes
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub somedata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim sdate, sdate2 As String
            Dim dtp As New System.Windows.Forms.DateTimePicker()
            If strtime = "" Then
                dtp.Value = Now
                sdate = dtp.Value.Year & "-" _
                & dtp.Value.Month & "-" _
                & dtp.Value.Day & " " _
                & "00" & ":" _
                & "00" & ":" _
                & "00"
                sdate2 = dtp.Value.Year & "-" _
                 & dtp.Value.Month & "-" _
                 & dtp.Value.Day & " " _
                 & "23" & ":" _
                 & "59" & ":" _
                 & "59"
            Else
                dtp.Value = CDate(strtime)
                sdate = dtp.Value.Year & "-" _
                & dtp.Value.Month & "-" _
                & dtp.Value.Day & " " _
                & "00" & ":" _
                & "00" & ":" _
                & "00"

                sdate2 = dtp.Value.Year & "-" _
                & dtp.Value.Month & "-" _
                & dtp.Value.Day & " " _
                & "23" & ":" _
                & "59" & ":" _
                & "59"
            End If

            Dim str As String = "select " _
                     & "rcljobs.job_no,rcljobs.job_tittle,daily_time.* " _
                     & " " _
                     & " from rcljobs inner join daily_time on rcljobs.job_no = daily_time.job_no and " _
                     & " lower(daily_time.id_no) like" _
                     & "'%" & myForms.id_no & "%' " _
                     & " and daily_time.ddate>='" & sdate & "'" _
                    & " and daily_time.ddate<='" & sdate2 & "'" _
                    & " order by daily_time.ddate asc"
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "timesheet")
                    Dim tname As String = custDS.Tables(0).TableName()

                    ' Create second column.
                    Dim myColumn = New System.Data.DataColumn()
                    myColumn.DataType = Type.GetType("System.Boolean")
                    myColumn.ColumnName = "Delete"
                    myColumn.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn)
                    'Dim mycount As Integer = custDS.Tables(0).Rows.Count
                    'Dim i
                    'For i = 0 To mycount - 1
                    '    custDS.Tables(0).Rows(i).Item("Delete") = False
                    '    System.Windows.Forms.Application.DoEvents()
                    'Next i
                    myForms.timesheet.dtgtimesheet.SetDataBinding(custDS, tname)
                    addtablestyle(tname)
                Else
                    myForms.timesheet.dtgtimesheet.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared strtime As String = ""
    Public Shared Sub addtablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.timesheet.dtgtimesheet.Width - 20
            mywidth = mywidth / 5

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            'Dim myno As New DataGridTextBoxColumn()
            'myno.MappingName = "job_no"
            'myno.HeaderText = "Job Number"
            'myno.Width = mywidth
            'ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.
            Dim myname1 As New DataGridBoolColumn()
            myname1.MappingName = "Delete"
            myname1.HeaderText = "Delete"
            myname1.Width = mywidth
            myname1.AllowNull = False
            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "job_tittle"
            myname.HeaderText = "Job Title"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "task"
            myname100.HeaderText = "Task"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)
            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn()
            mydesc.MappingName = "description"
            mydesc.HeaderText = "Job title"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add a second column style.
            Dim mydesc1ccc As New DataGridTextBoxColumn()
            mydesc1ccc.MappingName = "stime"
            mydesc1ccc.HeaderText = "Start time"
            mydesc1ccc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc1ccc)

            ' Add a second column style.
            Dim mydesc2a As New DataGridTextBoxColumn()
            mydesc2a.MappingName = "etime"
            mydesc2a.HeaderText = "End time"
            mydesc2a.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2a)

            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn()
            mydesc2.MappingName = "timespent"
            mydesc2.HeaderText = "Time"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc200 As New DataGridTextBoxColumn()
            mydesc200.MappingName = "ddate"
            mydesc200.HeaderText = "Date"
            mydesc200.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200)


            ' Add the DataGridTableStyle objects to the collection.
            myForms.timesheet.dtgtimesheet.TableStyles.Clear()
            ts1.AllowSorting = False
            myForms.timesheet.dtgtimesheet.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub

#End Region

#Region "it issues"
    Public Shared itindex As String
    Public Shared ffield As String
    Public Shared itsearch As Boolean = False
    Public Shared itno As String
    Public Shared showall As Boolean = False
    Public Delegate Sub mydelegateit()
    Public Shared Sub itinvoke()
        Try
            myForms.itissues.dtgissues.Invoke(New mydelegateit(AddressOf loadit))
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub loadit()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect1 As New ADODB.Connection()
            connect1.Mode = ConnectModeEnum.adModeReadWrite
            connect1.CursorLocation = CursorLocationEnum.adUseClient
            connect1.ConnectionString = connectstr
            connect1.Open()
            Dim str As String
            If itindex = "0" Then
                If myForms.itissues.btndelete.Visible = True Then
                    str = "select it.*,personnel_info.namme from it inner join personnel_info" _
                    & " on it.id_no=personnel_info.id_no " _
                    & " where it.id_no='" & itno & "'"
                Else
                    str = "select * from it where id_no='" & itno & "'"
                End If

                'If itno.Trim().Length < 1 Then
                '    str = "select * from it "
                'End If
            ElseIf itindex = "1" Then
                If myForms.itissues.btndelete.Visible = True Then
                    str = " select it.*,personnel_info.namme from it inner join personnel_info" _
                    & " on it.id_no=personnel_info.id_no " _
                    & " where it.id_no='" & itno & "'"
                    str += " and it.solved='" & True & "'"
                Else
                    str = "select * from it where id_no='" & itno & "'"
                    str += " and solved='" & True & "'"
                End If
                'str = "select * from it where id_no='" & itno & "'"

            Else
                If myForms.itissues.btndelete.Visible = True Then
                    str = " select it.*,personnel_info.namme from it inner join personnel_info" _
                    & " on it.id_no=personnel_info.id_no " _
                    & " where it.id_no='" & itno & "'"
                    str += " and it.solved='" & False & "'"
                Else
                    str = "select * from it where id_no='" & itno & "'"
                    str += " and solved='" & False & "'"
                End If
                'str = "select * from it where id_no='" & itno & "'"

            End If
            If itsearch = True Then
               
                If myForms.itissues.btndelete.Visible = True Then
                    str = " select it.*,personnel_info.namme from it inner join personnel_info" _
                    & " on it.id_no=personnel_info.id_no "
                    If itno.Length < 1 Then
                        str += " where it.id_no<>'" & itno & "'"
                    Else
                        str += " where it.id_no='" & itno & "'"
                    End If


                    Select Case itindex
                        Case "1"
                            str += " and  it.solved =" & True & ""
                        Case "2"
                            str += " and  it.solved =" & False & ""
                        Case Else

                    End Select
                    Dim dfv As String = "it." & ffield
                    If itno.Trim().Length < 1 Then
                        str += " and lower(" & dfv & ") like '%" & myForms.itissues.txtparam.Text.Trim.ToLower() & "%'"
                    Else
                        str += " and lower(" & dfv & ") like '%" & myForms.itissues.txtparam.Text.Trim.ToLower() & "%'"
                    End If
                Else
                        str = " select * from it where id_no='" & itno & "'"
                        Select Case itindex
                            Case "1"
                                str += " and  solved =" & True & ""
                            Case "2"
                                str += " and solved =" & False & ""
                            Case Else

                        End Select
                        If itno.Trim().Length < 1 Then
                            str += " and lower('" & ffield & "') like '%" & myForms.itissues.txtparam.Text.Trim.ToLower() & "%'"
                        Else
                            str += " and lower('" & ffield & "') like '%" & myForms.itissues.txtparam.Text.Trim.ToLower() & "%'"
                        End If
                    End If

                    itsearch = False
                End If
            If showall = True Then
                str = " select it.*,personnel_info.namme from it inner join personnel_info" _
                    & " on it.id_no=personnel_info.id_no "

                showall = False
            End If
            str += ";"

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect1)
                If .BOF = False And .EOF = False Then
                    Dim adap As OleDbDataAdapter = New OleDbDataAdapter
                    Dim dset As DataSet = New DataSet
                    adap.Fill(dset, rs, "it")
                    '---------------
                    Dim myColumn0a = New System.Data.DataColumn
                    myColumn0a.DataType = Type.GetType("System.Boolean")
                    myColumn0a.ColumnName = "mybool"
                    myColumn0a.DefaultValue = False
                    dset.Tables(0).Columns.Add(myColumn0a)

                    Dim inty As Integer = dset.Tables(0).Rows.Count
                    Dim wq As Integer = 0
                    Dim vb As Boolean = dset.Tables(0).Rows(0).Item("solved")
                    For wq = 0 To inty - 1
                        If dset.Tables(0).Rows(wq).Item("solved") = True Then
                            dset.Tables(0).Rows(wq).Item("mybool") = True
                        Else
                            dset.Tables(0).Rows(wq).Item("mybool") = False
                        End If
                        Application.DoEvents()
                    Next
                    '--------
                    Dim tname As String = dset.Tables(0).TableName()
                    myForms.itissues.dtgissues.SetDataBinding(dset, tname)
                    addittablestyle(tname)
                Else
                    myForms.itissues.dtgissues.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect1.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
                        & ex.InnerException().ToString() & vbCrLf _
                        & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub addittablestyle(ByVal tablename As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = tablename
            Dim mywidth As Integer
            mywidth = myForms.itissues.dtgissues.Width - 20
            mywidth = mywidth / 4
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.


            Try
                Dim myname2 As New DataGridTextBoxColumn
                myname2.MappingName = "namme"
                myname2.HeaderText = "Name"
                myname2.Width = mywidth

                ts1.GridColumnStyles.Add(myname2)
            Catch ex As Exception

            End Try
            Dim myname1 As New DataGridTextBoxColumn
            myname1.MappingName = "report_date"
            myname1.HeaderText = "Report Date"
            myname1.Width = mywidth

            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "issues"
            myname.HeaderText = "Issues"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn
            myname100.MappingName = "comments"
            myname100.HeaderText = "Comments"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)

            ' Add a second column style.
            Dim mydesc As New DataGridBoolColumn
            mydesc.MappingName = "mybool"
            mydesc.HeaderText = "Solved"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add the DataGridTableStyle objects to the collection.
            myForms.itissues.dtgissues.TableStyles.Clear()
            myForms.itissues.dtgissues.TableStyles.Add(ts1)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
             & ex.InnerException().ToString() & vbCrLf _
             & ex.StackTrace.ToString())
        End Try
    End Sub
    '------------load users into  cbofor 
    Public Shared Sub loaditcombo()
        Try
            '--------------write code to bind the combo box control
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim rsjob As New ADODB.Recordset()
            With rsjob
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                Dim strj As String
                strj = "SELECT  personnel_info.namme, personnel_info.id_no, seccheck.id_no AS id2 " _
                            & " FROM  seccheck INNER JOIN" _
                            & " personnel_info ON seccheck.id_no = personnel_info.id_no"

                .Open(strj, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    Try
                        myForms.itissues.cbofor.Items.Clear()
                        myForms.itissues.cboid.Items.Clear()
                        myForms.itissues.cbofor.Items.Add("All")
                        myForms.itissues.cboid.Items.Add("")
                    Catch zx As Exception

                    End Try

                    While .EOF = False
                        Try
                            myForms.itissues.cbofor.Items.Add(.Fields("namme").Value)
                            myForms.itissues.cboid.Items.Add(.Fields("id_no").Value)
                        Catch es300 As Exception
                        End Try

                        Application.DoEvents()
                        .MoveNext()
                    End While
                    myForms.itissues.cbofor.SelectedIndex = 0
                Else
                End If
            End With
            Try
                rsjob.Close()
            Catch re As Exception
            End Try
        Catch wq As Exception

        End Try
    End Sub
    Public Delegate Sub mydelegateitcombo()
    Public Shared Sub itcomboinvoke()
        Try
            myForms.itissues.Invoke(New mydelegateitcombo(AddressOf loaditcombo))
        Catch ex As Exception
        End Try
    End Sub
    '---------------------------------------
#End Region

#Region "personnel admin code"
    Public Delegate Sub mydelegate3()
    Public Shared Sub personnelinvoke()
        Try
            myForms.adminform.dtgpersonnel.Invoke(New mydelegate1(AddressOf loadpersonnel))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub loadpersonnel()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect1 As New ADODB.Connection()
            connect1.Mode = ConnectModeEnum.adModeReadWrite
            connect1.CursorLocation = CursorLocationEnum.adUseClient
            connect1.ConnectionString = connectstr
            connect1.Open()

            Dim str As String = "select * from personnel_info "
            str += " where id_no<>'9999999999' "
            str += " order by namme asc"
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect1)
                If .BOF = False And .EOF = False Then

                    Dim adap As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim dset As DataSet = New DataSet()
                    adap.Fill(dset, rs, "pesonnel")
                    Dim tname As String = dset.Tables(0).TableName()


                    myForms.adminform.dtgpersonnel.SetDataBinding(dset, tname)
                    addpersonneltablestyle(tname)
                End If
                Try
                    .Close()
                    connect1.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
                        & ex.InnerException().ToString() & vbCrLf _
                        & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub addpersonneltablestyle(ByVal tablename As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = tablename
            Dim mywidth As Integer
            mywidth = myForms.adminform.dtgpersonnel.Width - 20
            mywidth = mywidth / 4
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim myname1 As New DataGridTextBoxColumn()
            myname1.MappingName = "namme"
            myname1.HeaderText = "Name"
            myname1.Width = mywidth

            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "id_no"
            myname.HeaderText = "Id No"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "hourly_rate"
            myname100.HeaderText = "Hourly Rate"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)
            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn()
            mydesc.MappingName = "gender"
            mydesc.HeaderText = "Job Description"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.adminform.dtgpersonnel.TableStyles.Clear()
            myForms.adminform.dtgpersonnel.TableStyles.Add(ts1)
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
             & ex.InnerException().ToString() & vbCrLf _
             & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared searchstr As String
    Public Shared Sub searchinvoke()
        Try
            myForms.adminform.dtgpersonnel.Invoke(New mydelegate3(AddressOf searchresult))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub searchresult()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect1 As New ADODB.Connection()
            connect1.Mode = ConnectModeEnum.adModeReadWrite
            connect1.CursorLocation = CursorLocationEnum.adUseClient
            connect1.ConnectionString = connectstr
            connect1.Open()

            Dim str As String = "select * from personnel_info  "
            str += " where lower(namme) like '%" & searchstr.Trim().ToLower() & "%'"
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect1)
                If .BOF = False And .EOF = False Then

                    Dim adap As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim dset As DataSet = New DataSet()
                    adap.Fill(dset, rs, "pesonnel")
                    Dim tname As String = dset.Tables(0).TableName()


                    myForms.adminform.dtgpersonnel.SetDataBinding(dset, tname)
                    addpersonneltablestyle(tname)
                End If
                Try
                    .Close()
                    connect1.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ert As Exception
            MessageBox.Show(ert.Message.ToString() & vbCrLf _
            & ert.InnerException().ToString() & vbCrLf _
            & ert.StackTrace.ToString())
        End Try
    End Sub
#End Region

#Region "jobsheet code"
    Public Delegate Sub mydelegate2()
    Public Shared hisid_no As String
    Public Shared Sub loadjobsheet()
        Try
            myForms.jobsheet.dtpedate.Value = Now
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect2 As New ADODB.Connection()
            connect2.Mode = ConnectModeEnum.adModeReadWrite
            connect2.CursorLocation = CursorLocationEnum.adUseClient
            connect2.ConnectionString = connectstr
            connect2.Open()

            Dim str As String = "select * from personnel_info" _
            & " where id_no='" & hisid_no & "'"
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect2)
                If .BOF = False And .EOF = False Then
                    Try
                        myForms.jobsheet.txtname.Text = .Fields("namme").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtemail.Text = .Fields("email").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txthourlyrate.Text = .Fields("hourly_rate").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtphoneno.Text = .Fields("phone_no").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtpin.Text = .Fields("pin_no").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtpostaladdress.Text = .Fields("postal_address").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.cbogender.Text = .Fields("gender").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtbirthday.Text = .Fields("birthday").Value
                    Catch zx As Exception

                    End Try
                    '----------------
                    Try
                        myForms.jobsheet.txtcontractend.Text = .Fields("contract_end").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtnssfno.Text = .Fields("nssf_no").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtnhifno.Text = .Fields("nhif_no").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtmedicalcover.Text = .Fields("medical_cover").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.dtpdoe.Value = CDate(.Fields("dateofemployment").Value)
                    Catch zx As Exception

                    End Try

                    Try
                        myForms.jobsheet.dtpdot.Value = CDate(.Fields("dateoftermination").Value)
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtnextofkin.Text = .Fields("nextofkin").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtcomments.Text = .Fields("comments").Value
                    Catch zx As Exception

                    End Try
                    Try
                        myForms.jobsheet.txtmobileno.Text = .Fields("mobile_no").Value
                    Catch zx As Exception

                    End Try
                    '------------------
                    Dim f As String = .Fields("imagefile").Value
                    Try
                        f = f.Replace("|", "\")
                    Catch b As Exception
                    End Try
                    myForms.jobsheet.imagefilename = f
                    Try
                        Dim MyImage As System.Drawing.Bitmap
                        ' Stretches the image to fit the pictureBox. 
                        myForms.jobsheet.pbimage.SizeMode = PictureBoxSizeMode.StretchImage
                        MyImage = New Bitmap(f)
                        'pbimage.ClientSize = New Size(xSize, ySize)
                        myForms.jobsheet.pbimage.Image = CType(MyImage, Image)
                        'Me.pbimage.Image.FromFile(imagefilename)
                        ' pbimage.Refresh()
                    Catch sa As Exception

                    End Try

                End If
                Try
                    .Close()
                    connect2.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub jobsheetinvoke()
        Try
            myForms.jobsheet.Invoke(New mydelegate2(AddressOf loadjobsheet))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegate99()
    Public Shared Sub gridinvoke()
        Try
            myForms.jobsheet.Invoke(New mydelegate99(AddressOf gridcontrols))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Private Shared Sub gridcontrols()
        Try
            'Create the combo control to be added and set its properties

            myForms.jobsheet.comboControl = New System.Windows.Forms.ComboBox()
            myForms.jobsheet.comboControl.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.jobsheet.comboControl.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
            myForms.jobsheet.comboControl.Dock = DockStyle.Fill
            myForms.jobsheet.comboControl.Visible = True

            myForms.jobsheet.comboid = New System.Windows.Forms.ComboBox()
            myForms.jobsheet.comboid.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.jobsheet.comboid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
            myForms.jobsheet.comboid.Dock = DockStyle.Fill
            myForms.jobsheet.comboid.Visible = True

            myForms.jobsheet.combojob = New System.Windows.Forms.ComboBox()
            myForms.jobsheet.combojob.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.jobsheet.combojob.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
            myForms.jobsheet.combojob.Dock = DockStyle.Fill
            myForms.jobsheet.combojob.Visible = True

            ' TextBoxes
            myForms.jobsheet.txttask = New System.Windows.Forms.Button()
            myForms.jobsheet.txttask.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.jobsheet.txttask.Dock = DockStyle.Fill
            myForms.jobsheet.txttask.Visible = True

            myForms.jobsheet.txtdesc = New System.Windows.Forms.Button()
            myForms.jobsheet.txtdesc.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.jobsheet.txtdesc.Dock = DockStyle.Fill
            myForms.jobsheet.txtdesc.Visible = True

            myForms.jobsheet.txttimespent = New AMS.TextBox.MaskedTextBox()
            myForms.jobsheet.txttimespent.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.jobsheet.txttimespent.Dock = DockStyle.Fill
            'myForms.jobsheet.txttimespent.MaxLength = 8
            myForms.jobsheet.txttimespent.Mask = "##.##"
            myForms.jobsheet.txttimespent.Visible = True

            myForms.jobsheet.dtpddate = New System.Windows.Forms.DateTimePicker()
            myForms.jobsheet.dtpddate.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.jobsheet.dtpddate.Dock = DockStyle.Fill
            myForms.jobsheet.dtpddate.Format = DateTimePickerFormat.Short
            myForms.jobsheet.dtpddate.Visible = True

            '-----------------------------thiis is crap
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim rsjob As New ADODB.Recordset()
            'With rsjob
            '    .CursorLocation = CursorLocationEnum.adUseClient
            '    .CursorType = CursorTypeEnum.adOpenStatic
            '    Dim strj As String
            '    strj = "select  namme,"
            '    strj += " id_no"
            '    strj += " from personnel_info order by namme asc "
            '    .Open(strj, connect)
            '    If .BOF = False And .EOF = False Then
            '        .MoveFirst()
            '        While .EOF = False
            '            Try
            '                myForms.jobsheet.comboControl.Items.Add(.Fields("namme").Value)


            '            Catch es300 As Exception
            '            End Try

            '            Application.DoEvents()
            '            .MoveNext()
            '        End While
            '    Else
            '    End If
            'End With
            'Try
            '    rsjob.Close()
            'Catch re As Exception
            'End Try
            '------------------job tittles
            Dim rsjob1 As New ADODB.Recordset()
            With rsjob1
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                Dim strj As String
                strj = "select  * "
                strj += " from rcljobs where job_tittle<>'" & "" & "' order by job_tittle "
                .Open(strj, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            myForms.jobsheet.combojob.Items.Add(.Fields("job_tittle").Value)
                            myForms.jobsheet.comboid.Items.Add(.Fields("job_no").Value)

                        Catch es300 As Exception
                        End Try
                        Application.DoEvents()
                        .MoveNext()
                    End While
                Else
                End If
            End With
            Try
                rsjob1.Close()
            Catch re As Exception
            End Try
            Try

                connect.Close()
            Catch ex245 As Exception

            End Try

            '------------------
            'myForms.jobsheet.comboControl.Items.AddRange(New Object() {"", "Information Technology", "Computer Science", "Bio Technology", "Electrical Engg", "aaa"})
            'Add the controls to the respective columns in the data grid
            Dim i As Integer
            Dim sType As String
            'Take the text box from the second column of the grid where u will be adding the controls of your choice	
            myForms.jobsheet.datagridtextBox = CType(myForms.jobsheet.dtgtimesheet.TableStyles(0).GridColumnStyles(1), DataGridTextBoxColumn)
            myForms.jobsheet.datagridtextBox1 = CType(myForms.jobsheet.dtgtimesheet.TableStyles(0).GridColumnStyles(2), DataGridTextBoxColumn)
            myForms.jobsheet.datagridtextBox2 = CType(myForms.jobsheet.dtgtimesheet.TableStyles(0).GridColumnStyles(3), DataGridTextBoxColumn)
            myForms.jobsheet.datagridtextBox3 = CType(myForms.jobsheet.dtgtimesheet.TableStyles(0).GridColumnStyles(4), DataGridTextBoxColumn)
            myForms.jobsheet.datagridtextBox4 = CType(myForms.jobsheet.dtgtimesheet.TableStyles(0).GridColumnStyles(5), DataGridTextBoxColumn)
            myForms.jobsheet.datagridtextBox5 = CType(myForms.jobsheet.dtgtimesheet.TableStyles(0).GridColumnStyles(6), DataGridTextBoxColumn)
            'chk.SendToBack()
            'rb.SendToBack()
            'dtp.SendToBack()
            'myForms.jobsheet.comboControl.SendToBack()
            'myForms.jobsheet.datagridtextBox.TextBox.Controls.Add(myForms.jobsheet.comboControl)
            'myForms.jobsheet.comboControl.BringToFront()
            'myForms.jobsheet.datagridtextBox.TextBox.BackColor = Color.White

            '--------------------------txttask
            myForms.jobsheet.combojob.SendToBack()
            myForms.jobsheet.datagridtextBox1.TextBox.Controls.Add(myForms.jobsheet.combojob)
            myForms.jobsheet.combojob.BringToFront()
            myForms.jobsheet.datagridtextBox1.TextBox.BackColor = Color.White
            '--------------------------txttask
            myForms.jobsheet.txttask.SendToBack()
            myForms.jobsheet.datagridtextBox2.TextBox.Controls.Add(myForms.jobsheet.txttask)
            myForms.jobsheet.txttask.BringToFront()
            myForms.jobsheet.datagridtextBox2.TextBox.BackColor = Color.White

            '---------------------------descrip
            myForms.jobsheet.txtdesc.SendToBack()
            myForms.jobsheet.datagridtextBox3.TextBox.Controls.Add(myForms.jobsheet.txtdesc)
            myForms.jobsheet.txtdesc.BringToFront()
            myForms.jobsheet.datagridtextBox3.TextBox.BackColor = Color.White

            '---------------------------timespent
            myForms.jobsheet.txttimespent.Visible = True
            myForms.jobsheet.txttimespent.SendToBack()
            myForms.jobsheet.datagridtextBox4.TextBox.Controls.Add(myForms.jobsheet.txttimespent)
            myForms.jobsheet.txttimespent.BringToFront()
            myForms.jobsheet.datagridtextBox4.TextBox.BackColor = Color.White
            '---------------------------datetime picker
            myForms.jobsheet.dtpddate.SendToBack()
            myForms.jobsheet.datagridtextBox5.TextBox.Controls.Add(myForms.jobsheet.dtpddate)
            myForms.jobsheet.dtpddate.BringToFront()
            myForms.jobsheet.datagridtextBox5.TextBox.BackColor = Color.White

            'For i = 0 To DataTable.Rows.Count - 1
            '    sType = dgMember(i, 0).ToString()
            '    If hitTestGrid.Row = i Then
            '        Select Case hitTestGrid.Row
            '            Case 1
            '                datagridtextBox.TextBox.Controls.Add(dtp)
            '                dtp.BringToFront()
            '            Case 0
            '                datagridtextBox.TextBox.Controls.Add(comboControl)
            '                comboControl.BringToFront()
            '            Case 2
            '                datagridtextBox.TextBox.Controls.Add(chk)
            '                chk.BringToFront()
            '            Case 3
            '                datagridtextBox.TextBox.Controls.Add(rb)
            '                rb.BringToFront()
            '        End Select
            '    End If
            '    datagridtextBox.TextBox.BackColor = Color.White
            'Next i
        Catch ex As Exception

        End Try
    End Sub

#Region "miscllenous"
    Public Shared mid As String
    Public Delegate Sub mydelegate2m()
    Public Shared Sub populatecbos()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = " SELECT namme,id_no" _
            & " FROM personnel_info" _
            & "  where id_no like '%" & mid & "%'"

            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                Try
                    myForms.jobsheet.cboleaves.Items.Clear()
                    myForms.jobsheet.cbodayoff.Items.Clear()
                    myForms.jobsheet.cbosickoff.Items.Clear()
                    myForms.jobsheet.cbotimeoff.Items.Clear()
                    myForms.jobsheet.cbomiscelanous.Items.Clear()
                Catch zxc As Exception
                End Try
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            myForms.jobsheet.cboleaves.Items.Add(.Fields("namme").Value)
                            myForms.jobsheet.cbodayoff.Items.Add(.Fields("namme").Value)
                            myForms.jobsheet.cbosickoff.Items.Add(.Fields("namme").Value)
                            myForms.jobsheet.cbotimeoff.Items.Add(.Fields("namme").Value)
                            myForms.jobsheet.cbomiscelanous.Items.Add(.Fields("id_no").Value)


                        Catch es300 As Exception
                        End Try
                        Application.DoEvents()
                        .MoveNext()
                    End While
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
        Catch we As Exception

        End Try
    End Sub
    Public Shared Sub cbosinvoke()
        Try
            myForms.jobsheet.Invoke(New mydelegate2m(AddressOf populatecbos))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub leavesdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select leaves.*,personnel_info.namme from leaves"
            str += " inner join personnel_info on leaves.idno=personnel_info.id_no "
            str += " and leaves.idno='" & mid & "' "
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "leaves")
                    Dim tname As String = custDS.Tables(0).TableName()
                    myForms.jobsheet.dtgleaves.SetDataBinding(custDS, tname)
                    addtablestyleleaves(tname)
                Else
                    myForms.jobsheet.dtgleaves.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestyleleaves(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.jobsheet.dtgleaves.Width - 20
            mywidth = mywidth / 3

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "description"
            myno.HeaderText = "Description"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "sdate"
            myname.HeaderText = "Start date"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "edate"
            myname100.HeaderText = "End date"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.jobsheet.dtgleaves.TableStyles.Clear()
            myForms.jobsheet.dtgleaves.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub leavesinvoke()
        Try
            myForms.jobsheet.Invoke(New mydelegate3m(AddressOf leavesdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegate3m()
#Region " sick off"
    Public Shared Sub sickoffdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select sickoff.*,personnel_info.namme from sickoff"
            str += " inner join personnel_info on sickoff.idno=personnel_info.id_no "
            str += " and sickoff.idno='" & mid & "' "
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "sickoff")
                    Dim tname As String = custDS.Tables(0).TableName()
                    myForms.jobsheet.dtgsickoff.SetDataBinding(custDS, tname)
                    addtablestylesickoff(tname)
                Else
                    myForms.jobsheet.dtgsickoff.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestylesickoff(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.jobsheet.dtgsickoff.Width - 20
            mywidth = mywidth / 3

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "description"
            myno.HeaderText = "Description"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "sdate"
            myname.HeaderText = "Start date"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "edate"
            myname100.HeaderText = "End date"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.jobsheet.dtgsickoff.TableStyles.Clear()
            myForms.jobsheet.dtgsickoff.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub sickoffinvoke()
        Try
            myForms.jobsheet.Invoke(New mydelegates1(AddressOf sickoffdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegates1()
#End Region
#Region " time off"
    Public Shared Sub timeoffdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select timeoff.*,personnel_info.namme from timeoff"
            str += " inner join personnel_info on timeoff.idno=personnel_info.id_no "
            str += " and timeoff.idno='" & mid & "' "
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "timeoff")
                    Dim tname As String = custDS.Tables(0).TableName()
                    myForms.jobsheet.dtgtimeoff.SetDataBinding(custDS, tname)
                    addtablestyletimeoff(tname)
                Else
                    myForms.jobsheet.dtgtimeoff.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestyletimeoff(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.jobsheet.dtgtimeoff.Width - 20
            mywidth = mywidth / 3

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "description"
            myno.HeaderText = "Description"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "dateoff"
            myname.HeaderText = "Date off"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "timeoff"
            myname100.HeaderText = "Time off"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.jobsheet.dtgtimeoff.TableStyles.Clear()
            myForms.jobsheet.dtgtimeoff.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub timeoffinvoke()
        Try
            myForms.jobsheet.Invoke(New mydelegatet11(AddressOf timeoffdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegatet11()
#End Region
#Region "day off duty"
    Public Shared Sub dayoffdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select dayoff.*,personnel_info.namme from dayoff"
            str += " inner join personnel_info on dayoff.idno=personnel_info.id_no "
            str += " and dayoff.idno='" & mid & "' "
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "dayoff")
                    Dim tname As String = custDS.Tables(0).TableName()
                    myForms.jobsheet.dtgdayoff.SetDataBinding(custDS, tname)
                    addtablestyledayoff(tname)
                Else
                    myForms.jobsheet.dtgdayoff.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestyledayoff(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.jobsheet.dtgdayoff.Width - 20
            mywidth = mywidth / 3

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "description"
            myno.HeaderText = "Description"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "dateoff"
            myname.HeaderText = "Date off"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            '' Add a second column style.
            'Dim myname100 As New DataGridTextBoxColumn()
            'myname100.MappingName = "timeoff"
            'myname100.HeaderText = "Time off"
            'myname100.Width = mywidth
            'ts1.GridColumnStyles.Add(myname100)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.jobsheet.dtgdayoff.TableStyles.Clear()
            myForms.jobsheet.dtgdayoff.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub dayoffinvoke()
        Try
            myForms.jobsheet.Invoke(New mydelegated1(AddressOf dayoffdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegated1()
#End Region
#End Region

#End Region

#End Region

#Region "jobs"

#Region "personnel"
    Public Shared Sub personneldata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select daily_time.task,personnel_info.namme,daily_time.timespent,daily_time.id_no," _
            & "daily_time.job_no,daily_time.description,daily_time.ddate,daily_time.milliseconds,daily_time.ano," _
            & " personnel_info.hourly_rate from daily_time"
            str += " inner join personnel_info on daily_time.id_no=personnel_info.id_no "
            str += " and daily_time.job_no='" & strjobno & "' "
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "jobsummary")
                    Dim tname As String = custDS.Tables(0).TableName()
                    ' Create second column.
                    Dim myColumn = New System.Data.DataColumn()
                    myColumn.DataType = Type.GetType("System.String")
                    myColumn.ColumnName = "Cost"
                    custDS.Tables(0).Columns.Add(myColumn)

                    myForms.CustomerForm2.dtgpersonnel.SetDataBinding(custDS, tname)
                    addtablestylepersonnel(tname)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet()
                    ds = myForms.CustomerForm2.dtgpersonnel.DataSource
                    computecost(ds)
                    Dim strf As String
                    strf = " update grossmargin set personnel='" & myForms.CustomerForm2.txtpersonelcost.TextBoxText.Trim & "'" _
                    & " where job_no='" & strjobno & "'"
                    Try
                        connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                        connect.BeginTrans()
                        connect.Execute(strf)
                        connect.CommitTrans()

                    Catch az As Exception

                    End Try

                Else
                    myForms.CustomerForm2.dtgpersonnel.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestylepersonnel(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.CustomerForm2.dtgpersonnel.Width - 20
            mywidth = mywidth / 4

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "task"
            myname.HeaderText = "Task"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "namme"
            myno.HeaderText = "Done by"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)


            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "timespent"
            myname100.HeaderText = "Time taken(hrs)"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)

            ' Add a second column style.
            Dim myname100x As New DataGridTextBoxColumn()
            myname100x.MappingName = "Cost"
            myname100x.HeaderText = "Cost"
            myname100x.Width = mywidth
            ts1.GridColumnStyles.Add(myname100x)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.CustomerForm2.dtgpersonnel.TableStyles.Clear()
            myForms.CustomerForm2.dtgpersonnel.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub jobsinvoke()
        Try
            myForms.CustomerForm2.Invoke(New mydelegatej1(AddressOf personneldata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegatej1()
    Public Shared strjobno As String
    Public Shared Sub computecost(ByVal ds As System.Data.DataSet)
        Try
            Dim intcount As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim ttime, hrate, ccost As Double
            Dim nb As Double
            Dim f
            For kappa = 0 To intcount - 1
                Application.DoEvents()
                Try
                    ttime = Convert.ToDouble(ds.Tables(0).Rows(kappa).Item("timespent"))
                    hrate = Convert.ToDouble(ds.Tables(0).Rows(kappa).Item("hourly_rate"))
                    nb = ttime * hrate
                    f = Math.Round(Convert.ToDecimal(nb), 2)
                    ds.Tables(0).Rows(kappa).Item("Cost") = f
                    ccost += nb

                Catch sa As Exception
                End Try

            Next
            ccost = Math.Round(Convert.ToDecimal(ccost), 2)
            myForms.CustomerForm2.txtpersonelcost.TextBoxText = ccost
            myForms.CustomerForm2.kpersonnel = ccost
        Catch bc As Exception

        End Try
    End Sub
#End Region

#Region "casuals"
    Public Shared Sub casualdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select *  from casuals where job_no='" & strjobno & "'"

            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "casuals")
                    Dim tname As String = custDS.Tables(0).TableName()

                    myForms.CustomerForm2.dtgcasuals.SetDataBinding(custDS, tname)
                    addtablestylecasual(tname)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet()
                    ds = myForms.CustomerForm2.dtgcasuals.DataSource
                    computecasual(ds)
                    Dim strf As String
                    strf = " update grossmargin set casuals ='" & myForms.CustomerForm2.txtlabour.TextBoxText.Trim & "'" _
                    & " where job_no='" & strjobno & "'"
                    Try
                        connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                        connect.BeginTrans()
                        connect.Execute(strf)
                        connect.CommitTrans()

                    Catch az As Exception

                    End Try
                Else
                    myForms.CustomerForm2.dtgcasuals.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestylecasual(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.CustomerForm2.dtgcasuals.Width - 20
            mywidth = mywidth / 4

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "namme"
            myno.HeaderText = "Name"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "task"
            myname.HeaderText = "Task performed"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)




            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "datehired"
            myname100.HeaderText = "Date hired"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)

            ' Add a second column style.
            Dim myname100x As New DataGridTextBoxColumn()
            myname100x.MappingName = "wagespaid"
            myname100x.HeaderText = "Wages paid"
            myname100x.Width = mywidth
            ts1.GridColumnStyles.Add(myname100x)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.CustomerForm2.dtgcasuals.TableStyles.Clear()
            myForms.CustomerForm2.dtgcasuals.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub casualsinvoke()
        Try
            myForms.CustomerForm2.Invoke(New mydelegatej1(AddressOf casualdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegatec1()
    Public Shared Sub computecasual(ByVal ds As System.Data.DataSet)
        Try
            Dim intcount As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim ttime, hrate, ccost As Double
            Dim nb As Double
            Dim f
            For kappa = 0 To intcount - 1
                Application.DoEvents()
                Try
                    ttime = Convert.ToDouble(ds.Tables(0).Rows(kappa).Item("wagespaid"))
                    ccost += ttime
                Catch za As Exception

                End Try

            Next
            ccost = Math.Round(Convert.ToDecimal(ccost), 2)
            myForms.CustomerForm2.txtlabour.TextBoxText = ccost
            myForms.CustomerForm2.kcasual = ccost
        Catch bc As Exception

        End Try
    End Sub
#End Region

#Region "travel"
    Public Shared Sub traveldata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select *  from travel where job_no='" & strjobno & "'"

            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "travel")
                    Dim tname As String = custDS.Tables(0).TableName()

                    myForms.CustomerForm2.dtgtravel.SetDataBinding(custDS, tname)
                    addtablestyletravel(tname)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet()
                    ds = myForms.CustomerForm2.dtgtravel.DataSource
                    computetravel(ds)
                    Dim strf As String
                    strf = " update grossmargin set travel ='" & myForms.CustomerForm2.txttravel.TextBoxText.Trim & "'" _
                    & " where job_no='" & strjobno & "'"
                    Try
                        connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                        connect.BeginTrans()
                        connect.Execute(strf)
                        connect.CommitTrans()

                    Catch az As Exception

                    End Try
                Else
                    myForms.CustomerForm2.dtgtravel.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestyletravel(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.CustomerForm2.dtgtravel.Width - 20
            mywidth = mywidth / 2

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "description"
            myname.HeaderText = "Description"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)




            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "costincurred"
            myname100.HeaderText = "Cost Incurred"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)


            ' Add the DataGridTableStyle objects to the collection.
            myForms.CustomerForm2.dtgtravel.TableStyles.Clear()
            myForms.CustomerForm2.dtgtravel.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub travelinvoke()
        Try
            myForms.CustomerForm2.Invoke(New mydelegatetr1(AddressOf traveldata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegatetr1()
    Public Shared Sub computetravel(ByVal ds As System.Data.DataSet)
        Try
            Dim intcount As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim ttime, hrate, ccost As Double
            For kappa = 0 To intcount - 1
                Application.DoEvents()
                Try
                    ttime = Convert.ToDouble(ds.Tables(0).Rows(kappa).Item("costincurred"))
                    ccost += ttime
                Catch za As Exception

                End Try

            Next
            ccost = Math.Round(Convert.ToDecimal(ccost), 2)
            myForms.CustomerForm2.txttravel.TextBoxText = ccost
            myForms.CustomerForm2.ktravel = ccost
        Catch bc As Exception

        End Try
    End Sub
#End Region

#Region "accomodation"
    Public Shared Sub accomodationdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select *  from accomodation where job_no='" & strjobno & "'"

            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "accomodationdata")
                    Dim tname As String = custDS.Tables(0).TableName()

                    myForms.CustomerForm2.dtgaccomodation.SetDataBinding(custDS, tname)
                    addtablestyleaccomodation(tname)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet()
                    ds = myForms.CustomerForm2.dtgaccomodation.DataSource
                    computeaccomodation(ds)
                    Dim strf As String
                    strf = " update grossmargin set accomodation ='" & myForms.CustomerForm2.txtaccomodation.TextBoxText.Trim & "'" _
                    & " where job_no='" & strjobno & "'"
                    Try
                        connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                        connect.BeginTrans()
                        connect.Execute(strf)
                        connect.CommitTrans()

                    Catch az As Exception

                    End Try
                Else
                    myForms.CustomerForm2.dtgaccomodation.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestyleaccomodation(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.CustomerForm2.dtgaccomodation.Width - 20
            mywidth = mywidth / 2

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "description"
            myname.HeaderText = "Description"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "costincurred"
            myname100.HeaderText = "Cost Incurred"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)

            ' Add the DataGridTableStyle objects to the collection.
            myForms.CustomerForm2.dtgaccomodation.TableStyles.Clear()
            myForms.CustomerForm2.dtgaccomodation.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub accomodationinvoke()
        Try
            myForms.CustomerForm2.Invoke(New mydelegateaa1(AddressOf accomodationdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegateaa1()
    Public Shared Sub computeaccomodation(ByVal ds As System.Data.DataSet)
        Try
            Dim intcount As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim ttime, hrate, ccost As Double
            For kappa = 0 To intcount - 1
                Application.DoEvents()
                Try
                    ttime = Convert.ToDouble(ds.Tables(0).Rows(kappa).Item("costincurred"))
                    ccost += ttime
                Catch we As Exception

                End Try

            Next
            ccost = Math.Round(Convert.ToDecimal(ccost), 2)
            myForms.CustomerForm2.txtaccomodation.TextBoxText = ccost
            myForms.CustomerForm2.kaccomodation = ccost
        Catch bc As Exception

        End Try
    End Sub
#End Region

#Region "equipments"
    Public Shared Sub ramaniequipdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " select history_equip.*,equip_info.model_name  from" _
            '& " history_equip inner join equip_info on equip_info.equip_id = equip_info.equip_id"
            'str += " and history_equip.job_no='" & erjobno & "'"
            '-------------true sql statement

            Dim str As String = "SELECT   history_equip.equip_id,  history_equip.job_no,  history_equip.other,  history_equip.task,  history_equip.description, " _
                                           & "  history_equip.assigned_by,  history_equip.date_assigned,  history_equip.estimate_release_date,  history_equip.date_released," _
                                           & "  history_equip.autonumber, history_equip.ano ,  equip_info.model_name ,equip_info.hourly_rate" _
                                           & "  FROM equip_info INNER JOIN" _
                                           & "   history_equip ON  equip_info.equip_id =  history_equip.equip_id"
            str += " and history_equip.job_no=" & "N" & "'" & erjobno & "'"
            '-----------
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "historyequip")
                    Dim tname As String = custDS.Tables(0).TableName()
                    ' Create second column.
                    Dim myColumn = New System.Data.DataColumn()
                    myColumn.DataType = Type.GetType("System.String")
                    myColumn.ColumnName = "Cost"
                    custDS.Tables(0).Columns.Add(myColumn)

                    ' Create second column.
                    Dim myColumnv = New System.Data.DataColumn()
                    myColumnv.DataType = Type.GetType("System.String")
                    myColumnv.ColumnName = "time"
                    custDS.Tables(0).Columns.Add(myColumnv)

                    myForms.CustomerForm2.dtgramaniequip.SetDataBinding(custDS, tname)
                    addtablestyleramaniequip(tname)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet()
                    ds = myForms.CustomerForm2.dtgramaniequip.DataSource
                    computeramanicost(ds)
                    Dim strf As String
                    strf = " update grossmargin set ramani ='" & myForms.CustomerForm2.txtramani.TextBoxText.Trim & "'" _
                    & " where job_no='" & strjobno & "'"
                    Try
                        connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                        connect.BeginTrans()
                        connect.Execute(strf)
                        connect.CommitTrans()

                    Catch az As Exception

                    End Try
                Else
                    myForms.CustomerForm2.dtgramaniequip.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestyleramaniequip(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.CustomerForm2.dtgramaniequip.Width - 20
            mywidth = mywidth / 4

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "model_name"
            myname.HeaderText = "Name"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "date_assigned"
            myno.HeaderText = "Out"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)


            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "date_released"
            myname100.HeaderText = "In"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)

            ' Add a second column style.
            Dim myname100n As New DataGridTextBoxColumn()
            myname100n.MappingName = "time"
            myname100n.HeaderText = "Time(hrs)"
            myname100n.Width = mywidth
            ts1.GridColumnStyles.Add(myname100n)

            ' Add a second column style.
            Dim myname100x As New DataGridTextBoxColumn()
            myname100x.MappingName = "Cost"
            myname100x.HeaderText = "Cost"
            myname100x.Width = mywidth
            ts1.GridColumnStyles.Add(myname100x)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.CustomerForm2.dtgramaniequip.TableStyles.Clear()
            myForms.CustomerForm2.dtgramaniequip.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub ramaniequipinvoke()
        Try
            myForms.CustomerForm2.Invoke(New mydelegater1(AddressOf ramaniequipdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegater1()
    Public Shared erjobno As String
    Public Shared Sub computeramanicost(ByVal ds As System.Data.DataSet)
        Try
            Dim intcount As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim ttime, hrate, ccost As Double
            Dim date1 As New System.Windows.Forms.DateTimePicker()
            Dim date2 As New System.Windows.Forms.DateTimePicker()
            Dim datediff As New System.TimeSpan()
            Dim nb As Double
            For kappa = 0 To intcount - 1
                Application.DoEvents()
                Try
                    date1.Value = CDate(ds.Tables(0).Rows(kappa).Item("date_assigned"))
                    date2.Value = CDate(ds.Tables(0).Rows(kappa).Item("date_released"))
                    datediff = date2.Value.Subtract(date1.Value)
                    If datediff.TotalHours > 0 Then
                        hrate = Convert.ToDouble(ds.Tables(0).Rows(kappa).Item("hourly_rate"))
                        nb = _
                        datediff.TotalHours * hrate
                        ds.Tables(0).Rows(kappa).Item("Cost") = Math.Round(Convert.ToDecimal(nb), 2)
                        ds.Tables(0).Rows(kappa).Item("time") = Math.Round(Convert.ToDecimal(datediff.TotalHours), 2)
                        ccost += nb
                    End If


                Catch sa As Exception
                End Try

            Next
            ccost = Math.Round(Convert.ToDecimal(ccost), 2)
            myForms.CustomerForm2.txtramani.TextBoxText = ccost
            myForms.CustomerForm2.kramani = ccost
        Catch bc As Exception

        End Try
    End Sub

#Region "hired"
    Public Shared Sub hiredequipdata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " select history_equip.*,equip_info.model_name  from" _
            '& " history_equip inner join equip_info on equip_info.equip_id = equip_info.equip_id"
            'str += " and history_equip.job_no='" & erjobno & "'"
            '-------------true sql statement

            Dim str As String = "     SELECT   * from hiredequip "
            str += " where  job_no= '" & hiredjobno & "'"
            '-----------
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "hiredequip")
                    Dim tname As String = custDS.Tables(0).TableName()
                    '' Create second column.
                    'Dim myColumn = New System.Data.DataColumn()
                    'myColumn.DataType = Type.GetType("System.String")
                    'myColumn.ColumnName = "Cost"
                    'custDS.Tables(0).Columns.Add(myColumn)

                    myForms.CustomerForm2.dtghiredequip.SetDataBinding(custDS, tname)
                    addtablestylehiredequip(tname)
                    Dim ds As System.Data.DataSet = New System.Data.DataSet()
                    ds = myForms.CustomerForm2.dtghiredequip.DataSource
                    computehiredcost(ds)
                    Dim strf As String
                    strf = " update grossmargin set hired ='" & myForms.CustomerForm2.txthiredequip.TextBoxText.Trim & "'" _
                    & " where job_no='" & strjobno & "'"
                    Try
                        connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                        connect.BeginTrans()
                        connect.Execute(strf)
                        connect.CommitTrans()

                    Catch az As Exception

                    End Try
                Else
                    myForms.CustomerForm2.dtghiredequip.DataSource = Nothing
                End If
                Try
                    .Close()
                    connect.Close()
                Catch er344 As Exception

                End Try
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
              & ex.InnerException().ToString() & vbCrLf _
              & ex.StackTrace.ToString())
        End Try

    End Sub
    Public Shared Sub addtablestylehiredequip(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.CustomerForm2.dtghiredequip.Width - 20
            mywidth = mywidth / 4

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "equipname"
            myname.HeaderText = "Equipment name"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim myno As New DataGridTextBoxColumn()
            myno.MappingName = "description"
            myno.HeaderText = "Description"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)


            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn()
            myname100.MappingName = "assigndate"
            myname100.HeaderText = "Assing date"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)

            ' Add a second column style.
            Dim myname100n As New DataGridTextBoxColumn()
            myname100n.MappingName = "releasedate"
            myname100n.HeaderText = "Release date"
            myname100n.Width = mywidth
            ts1.GridColumnStyles.Add(myname100n)

            ' Add a second column style.
            Dim myname100x As New DataGridTextBoxColumn()
            myname100x.MappingName = "hourly_rate"
            myname100x.HeaderText = "Hourly rate"
            myname100x.Width = mywidth
            ts1.GridColumnStyles.Add(myname100x)

            ' Add a second column style.
            Dim myname100xq As New DataGridTextBoxColumn()
            myname100xq.MappingName = "Cost"
            myname100xq.HeaderText = "Cost"
            myname100xq.Width = mywidth
            ts1.GridColumnStyles.Add(myname100xq)
            ' Add the DataGridTableStyle objects to the collection.
            myForms.CustomerForm2.dtghiredequip.TableStyles.Clear()
            myForms.CustomerForm2.dtghiredequip.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub hiredequipinvoke()
        Try
            myForms.CustomerForm2.Invoke(New mydelegatehir1(AddressOf hiredequipdata))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegatehir1()
    Public Shared hiredjobno As String
    Public Shared Sub computehiredcost(ByVal ds As System.Data.DataSet)
        Try
            Dim intcount As Integer = ds.Tables(0).Rows.Count
            Dim kappa As Integer
            Dim ttime, hrate, ccost As Double
            Dim date1 As New System.Windows.Forms.DateTimePicker()
            Dim date2 As New System.Windows.Forms.DateTimePicker()
            Dim datediff As New System.TimeSpan()
            Dim nb As Double
            For kappa = 0 To intcount - 1
                Application.DoEvents()
                Try
                    date1.Value = CDate(ds.Tables(0).Rows(kappa).Item("assigndate"))
                    date2.Value = CDate(ds.Tables(0).Rows(kappa).Item("releasedate"))
                    datediff = date2.Value.Subtract(date1.Value)
                    If datediff.TotalHours > 0 Then
                        hrate = Convert.ToDouble(ds.Tables(0).Rows(kappa).Item("hourly_rate"))
                        nb = _
                        datediff.TotalHours * hrate
                        ds.Tables(0).Rows(kappa).Item("Cost") = Math.Round(Convert.ToDecimal(nb), 2)
                        ccost += nb
                    End If


                Catch sa As Exception
                End Try

            Next
            ccost = Math.Round(Convert.ToDecimal(ccost), 2)
            myForms.CustomerForm2.txthiredequip.TextBoxText = ccost
            myForms.CustomerForm2.khired = ccost
        Catch bc As Exception

        End Try
    End Sub
#End Region

#End Region

#End Region

#Region "home code"
    Public Shared adminarray() As String
    Public Shared viewadmin As Boolean = False
    Public Delegate Sub mydelegateh1()
    Public Shared Sub homeinvoke()
        Try
            myForms.Main.Invoke(New mydelegateh1(AddressOf homedata))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub homedata()
        Try
            Dim strf As String
            Dim myarray1() As String
            strf = adminarray(0)
            myarray1 = strf.Split(",")
            If myarray1(0) = "1" Then
                'clients/contacts
                myForms.Main.pnlclients.Visible = True
            End If
            strf = adminarray(1)
            myarray1.Initialize()
            myarray1 = strf.Split(",")
            If myarray1(0) = "1" Then
                'jobs
                myForms.Main.pnljobs.Visible = True
            End If
            strf = adminarray(2)
            myarray1.Initialize()
            myarray1 = strf.Split(",")
            If myarray1(0) = "1" Then
                'leads
                myForms.Main.pnlleads.Visible = True
            End If
            strf = adminarray(3)
            myarray1.Initialize()
            myarray1 = strf.Split(",")
            If myarray1(0) = "1" Then
                'equipment
                myForms.Main.pnlequip.Visible = True
                myForms.Main.pnlequipcontrols.Visible = True
            End If
            strf = adminarray(4)
            myarray1.Initialize()
            myarray1 = strf.Split(",")
            If myarray1(0) = "1" Then
                'personnel
                myForms.Main.pnlpersonnel.Visible = True
                myForms.Main.ToolBar2.Visible = True
            End If
            If myarray1(1) = "1" Then
                'personnel
                myForms.Main.mnufilesettings.Enabled = True
                myForms.Main.tlbadmin.Enabled = True
            Else
                myForms.Main.mnufilesettings.Enabled = False
                myForms.Main.tlbadmin.Enabled = False
            End If
            '--------------------------------automate this fact now
            'Dim f = adminarray.GetUpperBound(0)
            'Dim omega As Integer = 0
            'Dim k As Integer
            'For omega = 0 To f
            '    If adminarray(omega) = "1" Then
            '        k = k + 1
            '    End If
            'Next
            'If k > 4 Then
            '    myForms.Main.mnufilesettings.Enabled = True
            '    myForms.Main.tlbadmin.Enabled = True

            '    viewadmin = True
            'Else
            '    viewadmin = False
            '    myForms.Main.mnufilesettings.Enabled = False
            '    myForms.Main.tlbadmin.Enabled = False
            'End If
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "admin"
    Public Delegate Sub mydelegatea1()
    Public Delegate Sub mydelegatea2()
    Public Shared loadadmincontrols As Boolean = False
    Public Shared Sub admininvoke()
        Try
            myForms.admin.Invoke(New mydelegatea1(AddressOf admindata))
        Catch ex As Exception
            '---------------crashes if form doesnot exist
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub admindata()
        Try
           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select seccheck.*,personnel_info.namme,personnel_info.id_no as id2 from seccheck right outer join personnel_info on  " _
            & " seccheck.id_no=personnel_info.id_no" _
            & " where personnel_info.id_no<>'" & "9999999999" & "' " _
            & " order by seccheck.name asc"
            Dim rs As New ADODB.Recordset()
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim custDA As OleDbDataAdapter = New OleDbDataAdapter()
                    Dim custDS As DataSet = New DataSet()
                    custDA.Fill(custDS, rs, "admin")
                    Dim tname As String = custDS.Tables(0).TableName()


                    ' Create columns.
                    Dim myColumn0a = New System.Data.DataColumn()
                    myColumn0a.DataType = Type.GetType("System.Boolean")
                    myColumn0a.ColumnName = "mybool"
                    myColumn0a.DefaultValue = True
                    custDS.Tables(0).Columns.Add(myColumn0a)
                    ' Create columns.
                    Dim myColumn45 = New System.Data.DataColumn()
                    myColumn45.DataType = Type.GetType("System.String")
                    myColumn45.ColumnName = "isadded"

                    custDS.Tables(0).Columns.Add(myColumn45)
                    ' Create columns.
                    Dim myColumn0 = New System.Data.DataColumn()
                    myColumn0.DataType = Type.GetType("System.Boolean")
                    myColumn0.ColumnName = "Delete"
                    myColumn0.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn0)

                    ' Create columns.
                    Dim myColumn = New System.Data.DataColumn()
                    myColumn.DataType = Type.GetType("System.Boolean")
                    myColumn.ColumnName = "Leads"
                    myColumn.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn)

                    Dim myColumn1 = New System.Data.DataColumn()
                    myColumn1.DataType = Type.GetType("System.Boolean")
                    myColumn1.ColumnName = "Clients"
                    myColumn1.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn1)

                    Dim myColumn2 = New System.Data.DataColumn()
                    myColumn2.DataType = Type.GetType("System.Boolean")
                    myColumn2.ColumnName = "Jobs"
                    myColumn2.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn2)

                    Dim myColumn3 = New System.Data.DataColumn()
                    myColumn3.DataType = Type.GetType("System.Boolean")
                    myColumn3.ColumnName = "Equipment"
                    myColumn3.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn3)

                    Dim myColumn4 = New System.Data.DataColumn()
                    myColumn4.DataType = Type.GetType("System.Boolean")
                    myColumn4.ColumnName = "Personnel"
                    myColumn4.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn4)

                    ' Create columns.
                    Dim myColumn0ax = New System.Data.DataColumn()
                    myColumn0ax.DataType = Type.GetType("System.Boolean")
                    myColumn0ax.ColumnName = "uupdate"
                    myColumn0ax.DefaultValue = False
                    custDS.Tables(0).Columns.Add(myColumn0ax)
                    Dim mycount As Integer = custDS.Tables(0).Rows.Count
                    Dim i
                    Dim myarray() As String
                    Dim myarray1() As String
                    For i = 0 To mycount - 1
                        'custDS.Tables(0).Rows(i).Item("Delete") = False
                        'custDS.Tables(0).Rows(i).Item("Clients") = False
                        'custDS.Tables(0).Rows(i).Item("Leads") = False
                        'custDS.Tables(0).Rows(i).Item("Jobs") = False
                        'custDS.Tables(0).Rows(i).Item("Equipment") = False
                        'custDS.Tables(0).Rows(i).Item("Personnel") = False
                        'custDS.Tables(0).Rows(i).Item("mybool") = True
                        Try
                            custDS.Tables(0).Rows(i).Item("isadded") = custDS.Tables(0).Rows(i).Item("id_no")
                        Catch es As Exception
                        End Try
                        Dim strf As String
                        Try
                            myarray = Convert.ToString(custDS.Tables(0).Rows(i).Item("seclevel")).Split(":")
                            strf = myarray(0)
                            myarray1 = strf.Split(",")
                            If myarray1(0) = "1" Then
                                custDS.Tables(0).Rows(i).Item("Clients") = True
                            End If
                            myarray1.Initialize()
                            strf = myarray(1)
                            myarray1 = strf.Split(",")
                            If myarray1(0) = "1" Then
                                custDS.Tables(0).Rows(i).Item("jobs") = True
                            End If
                            myarray1.Initialize()
                            strf = myarray(2)
                            myarray1 = strf.Split(",")
                            If myarray1(0) = "1" Then
                                custDS.Tables(0).Rows(i).Item("Leads") = True
                            End If
                            myarray1.Initialize()
                            strf = myarray(3)
                            myarray1 = strf.Split(",")
                            If myarray1(0) = "1" Then
                                custDS.Tables(0).Rows(i).Item("Equipment") = True
                            End If
                            myarray1.Initialize()
                            strf = myarray(4)
                            myarray1 = strf.Split(",")
                            If myarray1(0) = "1" Then
                                custDS.Tables(0).Rows(i).Item("Personnel") = True
                            End If
                        Catch d As Exception
                        End Try
                        System.Windows.Forms.Application.DoEvents()
                    Next i
                    myForms.admin.dtgusers.SetDataBinding(custDS, tname)
                    Dim ds As New System.Data.DataSet()
                    'ds = myForms.admin.dtgusers.DataSource
                    'ds.Tables(0).Columns("password").ReadOnly = False
                    'ds.AcceptChanges()
                    addadmintablestyle(tname)

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
            'combo control
            myForms.admin.comboControl = New System.Windows.Forms.ComboBox()
            myForms.admin.comboControl.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.admin.comboControl.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple
            myForms.admin.comboControl.Dock = DockStyle.Fill
            myForms.admin.comboControl.Visible = True

            'combo control
            myForms.admin.comboid = New System.Windows.Forms.ComboBox()
            myForms.admin.comboid.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.admin.comboid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
            myForms.admin.comboid.Dock = DockStyle.Fill
            myForms.admin.comboid.Visible = True
            'txttask
            myForms.admin.txttask = New System.Windows.Forms.TextBox
            myForms.admin.txttask.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.admin.txttask.Dock = DockStyle.Fill
            myForms.admin.txttask.Visible = True

            'Take the text box from the second column of the grid where u will be adding the controls of your choice	
            myForms.admin.datagridtextBox = CType(myForms.admin.dtgusers.TableStyles(0).GridColumnStyles(2), DataGridTextBoxColumn)
            myForms.admin.datagridtextBox1 = CType(myForms.admin.dtgusers.TableStyles(0).GridColumnStyles(3), DataGridTextBoxColumn)

            myForms.admin.comboControl.SendToBack()
            myForms.admin.datagridtextBox.TextBox.Controls.Add(myForms.admin.comboControl)
            myForms.admin.comboControl.BringToFront()
            myForms.admin.datagridtextBox.TextBox.BackColor = Color.White

            '--------------------------txttask
            myForms.admin.txttask.SendToBack()
            myForms.admin.datagridtextBox1.TextBox.Controls.Add(myForms.admin.txttask)
            myForms.admin.txttask.BringToFront()
            myForms.admin.datagridtextBox1.TextBox.BackColor = Color.White
            '-------------this binds combo box
            Call rebindcombo()


        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub addadmintablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle()
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.admin.dtgusers.Width - 20
            mywidth = mywidth / 9

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            'Dim myno As New DataGridTextBoxColumn()
            'myno.MappingName = "job_no"
            'myno.HeaderText = "Job Number"
            'myno.Width = mywidth
            'ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.

            ' Add a second column style.
            Dim myname100v As New DataGridBoolColumn()
            myname100v.MappingName = "Delete"
            myname100v.HeaderText = "Delete"
            myname100v.Width = mywidth
            myname100v.AllowNull = False
            ts1.GridColumnStyles.Add(myname100v)

            Dim myname1f As New DataGridTextBoxColumn
            myname1f.MappingName = "namme"
            myname1f.HeaderText = "Personnel"
            myname1f.Width = mywidth

            ts1.GridColumnStyles.Add(myname1f)

            Dim myname1 As New DataGridTextBoxColumn()
            myname1.MappingName = "name"
            myname1.HeaderText = "User name"
            myname1.Width = mywidth

            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn()
            myname.MappingName = "password"
            myname.HeaderText = "Password"
            myname.ReadOnly = False
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridBoolColumn()
            mydesc.MappingName = "Leads"
            mydesc.HeaderText = "Leads"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add a second column style.
            Dim myname100 As New DataGridBoolColumn()
            myname100.MappingName = "Clients"
            myname100.HeaderText = "Clients"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)


            ' Add a second column style.
            Dim mydesc2 As New DataGridBoolColumn()
            mydesc2.MappingName = "Jobs"
            mydesc2.HeaderText = "Jobs"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc200 As New DataGridBoolColumn()
            mydesc200.MappingName = "Equipment"
            mydesc200.HeaderText = "Equipment"
            mydesc200.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200)

            ' Add a second column style.
            Dim mydesc2001 As New DataGridBoolColumn()
            mydesc2001.MappingName = "Personnel"
            mydesc2001.HeaderText = "Personnel"
            mydesc2001.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001)


            ' Add the DataGridTableStyle objects to the collection.
            myForms.admin.dtgusers.TableStyles.Clear()
            myForms.admin.dtgusers.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub rebindcombo()
        Try
            '--------------write code to bind the combo box control

           dim connectstr as string = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim rsjob As New ADODB.Recordset()
            With rsjob
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                Dim strj As String
                strj = "SELECT  personnel_info.namme, personnel_info.id_no, seccheck.id_no AS id2 " _
                            & " FROM  seccheck RIGHT OUTER JOIN" _
                            & " personnel_info ON seccheck.id_no = personnel_info.id_no" _
                            & " WHERE    seccheck.id_no IS NULL "
                .Open(strj, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    Try
                        myForms.admin.comboControl.Items.Clear()
                        myForms.admin.comboid.Items.Clear()
                    Catch zx As Exception

                    End Try

                    'While .EOF = False
                    '    Try
                    '        'myForms.admin.comboControl.Items.Add(.Fields("namme").Value)
                    '        'myForms.admin.comboid.Items.Add(.Fields("id_no").Value)

                    '    Catch es300 As Exception
                    '    End Try

                    '    Application.DoEvents()
                    '    .MoveNext()
                    'End While
                Else
                End If
            End With
            Try
                rsjob.Close()
            Catch re As Exception
            End Try
        Catch wq As Exception

        End Try
    End Sub
    Public Shared Sub comboinvoke()
        Try
            myForms.admin.Invoke(New mydelegatea2(AddressOf rebindcombo))
        Catch ex As Exception
            '------------crashes if doesnot exist
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
#End Region

#Region "equipments"
    Public Shared issearch As Boolean = False
    Public Shared strsearch As String
    Public Delegate Sub mydelegatee1()
    Public Shared Sub equipinvoke()
        Try
            myForms.equipments.Invoke(New mydelegatee1(AddressOf equipdata))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub equipdata()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String
            If issearch = True Then
                str = strsearch
                issearch = False
            Else
                str = "select equip_info.*, equip_finances.equip_id as id2,equip_finances.hourly_rate as gd" _
                     & " from equip_info inner join equip_finances on" _
                     & " equip_info.equip_id=equip_finances.equip_id order by  equip_info.equip_id asc"
            End If

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                myForms.equipments.dtginventories.DataSource = Nothing
                If .BOF = False And .EOF = False Then
                    Dim equipDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim equipDS As DataSet = New DataSet

                    equipDA.Fill(equipDS, rs, "equip")
                    Dim tname As String = equipDS.Tables(0).TableName()


                    ' Create columns.
                    Dim myColumn0a = New System.Data.DataColumn
                    myColumn0a.DataType = Type.GetType("System.Boolean")
                    myColumn0a.ColumnName = "mybool"
                    equipDS.Tables(0).Columns.Add(myColumn0a)

                    Dim myColumn0 = New System.Data.DataColumn
                    myColumn0.DataType = Type.GetType("System.String")
                    myColumn0.ColumnName = "equipment_actions"
                    myColumn0.DefaultValue = "Click here"
                    equipDS.Tables(0).Columns.Add(myColumn0)
                    'Try
                    '    Dim fd As Integer = equipDS.Tables(0).Columns.Count()
                    '    Dim bn As Integer = 0
                    '    For bn = 0 To fd - 1
                    '        equipDS.Tables(0).Columns(bn).DefaultValue = "   "
                    '        Application.DoEvents()
                    '    Next
                    'Catch jc As Exception

                    'End Try

                    myForms.equipments.dtginventories.SetDataBinding(equipDS, tname)
                    addequiptablestyle(tname)
                    equipDS.Dispose()
                    Call mycontrols()
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
    End Sub
    Public Shared Sub mycontrols()
        Try
            myForms.equipments.btnequipactions = New System.Windows.Forms.Button
            myForms.equipments.btnequipactions.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.equipments.btnequipactions.Text = "Equipment Actions"
            myForms.equipments.btnequipactions.Dock = DockStyle.Fill
            myForms.equipments.btnequipactions.Visible = True

            myForms.equipments.datagridtextBox = CType(myForms.equipments.dtginventories.TableStyles(0).GridColumnStyles(0), DataGridTextBoxColumn)

            myForms.equipments.btndeleteequipment.SendToBack()
            myForms.equipments.datagridtextBox.TextBox.Controls.Add(myForms.equipments.btnequipactions)
            myForms.equipments.btndeleteequipment.BringToFront()
            myForms.equipments.datagridtextBox.TextBox.BackColor = Color.White
        Catch qw As Exception

        End Try
    End Sub
    Public Shared Sub addequiptablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.equipments.dtginventories.Width - 20
            myForms.equipments.dtginventories.PreferredRowHeight = 33
            mywidth = mywidth / 13

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.
            Dim myn As New DataGridTextBoxColumn
            myn.MappingName = "equipment_actions"
            myn.HeaderText = "Equipment Actions"
            myn.Width = mywidth

            ts1.GridColumnStyles.Add(myn)

            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "equip_id"
            myno.HeaderText = "Equipment Id"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.

            ' Add a second column style.
            Dim myname100v As New DataGridTextBoxColumn
            myname100v.MappingName = "manufacturer"
            myname100v.HeaderText = "Manufacturer"
            myname100v.Width = mywidth
            ts1.GridColumnStyles.Add(myname100v)

            Dim myname1 As New DataGridTextBoxColumn
            myname1.MappingName = "model_no"
            myname1.HeaderText = "Model No"

            myname1.Width = mywidth

            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "serial_no"
            myname.HeaderText = "Serial No"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "model_name"
            mydesc.HeaderText = "Model Name"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn
            myname100.MappingName = "purchase_date"
            myname100.HeaderText = "Purchase Date"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn
            mydesc2.MappingName = "supplier"
            mydesc2.HeaderText = "Description"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc200 As New DataGridTextBoxColumn
            mydesc200.MappingName = "licence"
            mydesc200.HeaderText = "Licence"
            mydesc200.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200)

            ' Add a second column style.
            Dim mydesc2001 As New DataGridTextBoxColumn

            mydesc2001.MappingName = "guarantee"
            mydesc2001.HeaderText = "Guarantee"
            mydesc2001.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001)

            ' Add a second column style.
            Dim mydesc2001a As New DataGridTextBoxColumn
            mydesc2001a.MappingName = "condition"
            mydesc2001a.HeaderText = "Condition"
            mydesc2001a.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001a)

            ' Add a second column style.
            Dim mydesc2001bb As New DataGridTextBoxColumn
            mydesc2001bb.MappingName = "phone"
            mydesc2001bb.HeaderText = "Phone"
            mydesc2001bb.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001bb)

            ' Add a second column style.
            Dim mydesc2001b As New DataGridTextBoxColumn
            mydesc2001b.MappingName = "type"
            mydesc2001b.HeaderText = "Type"
            mydesc2001b.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001b)

            ' Add a second column style.
            Dim mydesc2001c As New DataGridTextBoxColumn
            mydesc2001c.MappingName = "model_year"
            mydesc2001c.HeaderText = "Model Year"
            mydesc2001c.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001c)

            ' Add a second column style.
            Dim mydesc2001cn As New DataGridTextBoxColumn
            mydesc2001cn.MappingName = "hourly_rate"
            mydesc2001cn.HeaderText = "Hourly  Rate"
            mydesc2001cn.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cn)


            ' Add a second column style.
            Dim mydesc2001cn1 As New DataGridTextBoxColumn
            mydesc2001cn1.MappingName = "description"
            mydesc2001cn1.HeaderText = "Supplier"
            mydesc2001cn1.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cn1)

            ' Add a second column style.
            Dim mydesc2001cn2 As New DataGridTextBoxColumn
            mydesc2001cn2.MappingName = "mouse"
            mydesc2001cn2.HeaderText = "Mouse"
            mydesc2001cn2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cn2)

            Dim mydesc2001cn3 As New DataGridTextBoxColumn
            mydesc2001cn3.MappingName = "keyboard"
            mydesc2001cn3.HeaderText = "Keyboard"
            mydesc2001cn3.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cn3)

            Dim mydesc2001cn4 As New DataGridTextBoxColumn
            mydesc2001cn4.MappingName = "monitor"
            mydesc2001cn4.HeaderText = "Monitor(1)"
            mydesc2001cn4.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cn4)

            Dim mydesc2001cn4v As New DataGridTextBoxColumn
            mydesc2001cn4v.MappingName = "monitor2"
            mydesc2001cn4v.HeaderText = "Monitor(2)"
            mydesc2001cn4v.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cn4v)

            Dim mydesc2001nm As New DataGridTextBoxColumn
            mydesc2001nm.MappingName = "batteries"
            mydesc2001nm.HeaderText = "Batteries"
            mydesc2001nm.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001nm)

            Dim mydesc2001mn As New DataGridTextBoxColumn
            mydesc2001mn.MappingName = "downloadcables"
            mydesc2001mn.HeaderText = "Download cables"
            mydesc2001mn.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001mn)

            Dim mydesc2001mm As New DataGridTextBoxColumn
            mydesc2001mm.MappingName = "unit"
            mydesc2001mm.HeaderText = "Unit"
            mydesc2001mm.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001mm)

            ' Add the DataGridTableStyle objects to the collection.
            myForms.equipments.dtginventories.TableStyles.Clear()
            myForms.equipments.dtginventories.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub

#Region "edit equipments"

#End Region

#Region " tojobs"
    Public Shared hastojobloaded As Boolean = False
    Public Shared jobno As String
    Public Shared globalnamme As String
    Public Delegate Sub mydelegatet1()
    Public Shared Sub loadname()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

            Dim str As String = "SELECT  namme from personnel_info where id_no='" & myForms.id_no & "'"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Try
                        myForms.tojobs.namme = .Fields("namme").Value
                    Catch vb As Exception

                    End Try
                    globalnamme = .Fields("namme").Value
                End If
            End With
        Catch qw As Exception

        End Try
    End Sub
    Public Shared Sub equipjobinvoke()
        Try
            myForms.tojobs.Invoke(New mydelegatet1(AddressOf assignequip))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared showavailableequip As Boolean = True
    Public Shared iaminjobs As Boolean = False
    Public Shared ijobno As String
    Public Shared Sub assignequip()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " SELECT  equip_info.*,  assigned_info.equip_id AS id2,  assigned_info.status " _
            '                    & " FROM  assigned_info RIGHT OUTER JOIN " _
            '                    & " equip_info ON  assigned_info.status =  equip_info.equip_id"
            'str += " and assigned_info.status='" & "0" & "'"
            Dim str As String = "SELECT     equip_finances.hourly_rate, equip_info.*, current_equip.job_no, current_equip.equip_id,current_equip.other, current_equip.task, " _
                      & "  current_equip.description AS description2, current_equip.assigned_by, current_equip.date_assigned, " _
                      & "  current_equip.estimate_release_date, assigned_info.status " _
                      & "  FROM equip_finances INNER JOIN " _
                      & "  equip_info ON  equip_finances.equip_id =  equip_info.equip_id INNER JOIN" _
                      & "  assigned_info ON  equip_info.equip_id =  assigned_info.equip_id LEFT OUTER JOIN" _
                      & "  current_equip ON  equip_info.equip_id =  current_equip.equip_id"
            If showavailableequip = True Then
                Dim p As String = ";"
                str += " where assigned_info.status='" & "0" & "'" & p
            Else
                str += " where assigned_info.status='" & "1" & "';"
            End If

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim tojobDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim tojobDS As DataSet = New DataSet
                    tojobDA.Fill(tojobDS, rs, "tojob")
                    Dim tname As String = tojobDS.Tables(0).TableName()

                    ' Create columns.
                    Dim myColumn0a = New System.Data.DataColumn
                    myColumn0a.DataType = Type.GetType("System.Boolean")
                    myColumn0a.ColumnName = "Assign"
                    If showavailableequip = True Then
                        myColumn0a.DefaultValue = False
                    Else
                        myColumn0a.DefaultValue = True
                    End If
                    tojobDS.Tables(0).Columns.Add(myColumn0a)

                    Dim myColumn0ac = New System.Data.DataColumn
                    myColumn0ac.DataType = Type.GetType("System.Boolean")
                    myColumn0ac.ColumnName = "has_changed"
                    myColumn0ac.DefaultValue = False
                    tojobDS.Tables(0).Columns.Add(myColumn0ac)

                    ' Create columns.
                    Dim myColumn = New System.Data.DataColumn
                    myColumn.DataType = Type.GetType("System.String")
                    myColumn.ColumnName = "date_released"

                    tojobDS.Tables(0).Columns.Add(myColumn)
                    '----------------default properties


                    '----------
                    myForms.tojobs.dtgequip.SetDataBinding(tojobDS, tname)
                    addtojobtablestyle(tname)
                    'Dim mycount As Integer = tojobDS.Tables(0).Rows.Count
                    'Dim i
                    'For i = 0 To mycount - 1
                    '    tojobDS.Tables(0).Rows(i).Item("Assign") = False
                    '    System.Windows.Forms.Application.DoEvents()
                    'Next i
                Else
                    myForms.tojobs.dtgequip.DataSource = Nothing
                End If

            End With
            Try
                rs.Close()
            Catch er34b As Exception
            End Try
            Call populatecbojob()
            Try

                connect.Close()
            Catch er344 As Exception

            End Try
        Catch ex As Exception
        Finally
            hastojobloaded = True
        End Try
    End Sub
    Public Shared Sub addtojobtablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.tojobs.dtgequip.Width - 5
            mywidth = mywidth / 11

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim mynoas As New DataGridBoolColumn
            mynoas.MappingName = "Assign"
            mynoas.HeaderText = "Assign/Deassign"
            mynoas.Width = mywidth
            mynoas.AllowNull = False
            ts1.GridColumnStyles.Add(mynoas)

            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "equip_id"
            myno.HeaderText = "Equipment Id"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "model_name"
            mydesc.HeaderText = "Model Name"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            Dim myname1 As New DataGridTextBoxColumn
            myname1.MappingName = "model_no"
            myname1.HeaderText = "Model No"
            myname1.Width = mywidth
            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname100v As New DataGridTextBoxColumn
            myname100v.MappingName = "manufacturer"
            myname100v.HeaderText = "Manufacturer"
            myname100v.Width = mywidth
            ts1.GridColumnStyles.Add(myname100v)



            '' Add a second column style.
            'Dim myname As New DataGridTextBoxColumn()
            'myname.MappingName = "serial_no"
            'myname.HeaderText = "Serial No"
            'myname.Width = mywidth
            'ts1.GridColumnStyles.Add(myname)



            '' Add a second column style.
            'Dim myname100 As New DataGridTextBoxColumn()
            'myname100.MappingName = "purchase_date"
            'myname100.HeaderText = "Purchase Date"
            'myname100.Width = mywidth
            'ts1.GridColumnStyles.Add(myname100)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn
            mydesc2.MappingName = "description"
            mydesc2.HeaderText = "Description"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            '' Add a second column style.
            'Dim mydesc200 As New DataGridTextBoxColumn()
            'mydesc200.MappingName = "licence"
            'mydesc200.HeaderText = "Licence"
            'mydesc200.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc200)

            '' Add a second column style.
            'Dim mydesc2001 As New DataGridTextBoxColumn()
            'mydesc2001.MappingName = "guarantee"
            'mydesc2001.HeaderText = "Guarantee"
            'mydesc2001.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc2001)

            ' Add a second column style.
            Dim mydesc2001a As New DataGridTextBoxColumn
            mydesc2001a.MappingName = "condition"
            mydesc2001a.HeaderText = "Condition"
            mydesc2001a.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001a)

            ' Add a second column style.
            Dim mydesc2001b As New DataGridTextBoxColumn
            mydesc2001b.MappingName = "type"
            mydesc2001b.HeaderText = "Type"
            mydesc2001b.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001b)

            '' Add a second column style.
            Dim mydesc2001c As New DataGridTextBoxColumn
            mydesc2001c.MappingName = "assigned_by"
            mydesc2001c.HeaderText = "Assigned by"
            mydesc2001c.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001c)

            '' Add a second column style.
            Dim mydesc2001cp As New DataGridTextBoxColumn
            mydesc2001cp.MappingName = "date_assigned"
            mydesc2001cp.HeaderText = "Date Assigned"
            mydesc2001cp.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cp)

            '' Add a second column style.
            Dim mydesc2001cpu As New DataGridTextBoxColumn
            mydesc2001cpu.MappingName = "estimate_release_date"
            mydesc2001cpu.HeaderText = "Estimate Release Date"
            mydesc2001cpu.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001cpu)


            ' Add the DataGridTableStyle objects to the collection.
            myForms.tojobs.dtgequip.TableStyles.Clear()
            myForms.tojobs.dtgequip.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub populatecbojob()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = " SELECT rcljobs.*,  clients.name" _
            & " FROM clients INNER JOIN" _
            & "  rcljobs ON clients.client_no = rcljobs.client_no "
            str += " and lower(rcljobs.job_status) like  '%" & "current" & "%'"
            Dim p As String = ";"
            'load one job no only
            If iaminjobs = True Then
                p = "   and rcljobs.job_no like " & "N" & "'" & ijobno & "'" & p
                iaminjobs = False
            End If
            str += p
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            myForms.tojobs.cbojob.Items.Add(Convert.ToString(.Fields("job_no").Value) & " : " & _
                            Convert.ToString(.Fields("job_tittle").Value) & " : " & _
                            Convert.ToString(.Fields("name").Value))
                            myForms.tojobs.comboControl.Items.Add(.Fields("job_no").Value)
                        Catch es300 As Exception
                        End Try
                        Application.DoEvents()
                        .MoveNext()
                    End While
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
        Catch we As Exception

        End Try
    End Sub
    Public Delegate Sub mydelegatet2()
    Public Shared Sub jobinvoke()
        Try
            myForms.tojobs.Invoke(New mydelegatet2(AddressOf populatejobgrid))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub populatejobgrid()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = "select " _
                & " rcljobs.client_no,rcljobs.job_no,rcljobs.job_tittle,rcljobs.job_status,rcljobs.techres," _
                & " clients.client_no,rcljobs.amount,clients.name" _
                & " from rcljobs inner join clients on rcljobs.client_no = clients.client_no and " _
                & " rcljobs.job_no =" _
                & " '" & jobno & "'"

            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    Dim tojobDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim tojobDS As DataSet = New DataSet
                    tojobDA.Fill(tojobDS, rs, "jobs")
                    Dim tname As String = tojobDS.Tables(0).TableName()
                    myForms.tojobs.dtgjobs.SetDataBinding(tojobDS, tname)
                    addcurrjobtablestyle(tname)

                End If
            End With

        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub addcurrjobtablestyle(ByVal tname As String)
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = myForms.tojobs.dtgjobs.Width
            mywidth = mywidth / 7
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            '' Add a second column style.
            'Dim mydesc1 As New DataGridTextBoxColumn()
            'mydesc1.MappingName = "client_no"
            'mydesc1.HeaderText = "Client Number"
            'mydesc1.Width = mywidth
            'ts1.GridColumnStyles.Add(mydesc1)

            ' Add a second column style.
            Dim mydesc4 As New DataGridTextBoxColumn
            mydesc4.MappingName = "name"
            mydesc4.HeaderText = "Name"
            mydesc4.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc4)

            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "job_no"
            myno.HeaderText = "Job Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "job_tittle"
            myname.HeaderText = "Job Title"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "job_status"
            mydesc.HeaderText = "Job Status"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add a second column style.
            Dim mydesc6 As New DataGridTextBoxColumn
            mydesc6.MappingName = "techres"
            mydesc6.HeaderText = "Technician Responsible"
            mydesc6.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc6)

            ' Add a second column style.
            Dim mydesc66 As New DataGridTextBoxColumn
            mydesc66.MappingName = "amount"
            mydesc66.HeaderText = "Amount"
            mydesc66.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc66)

            ' Add the DataGridTableStyle objects to the collection.
            myForms.tojobs.dtgjobs.TableStyles.Clear()
            myForms.tojobs.dtgjobs.TableStyles.Add(ts1)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
    End Sub
    '" SELECT rcljobs.*,  clients.name" _
    '" FROM clients INNER JOIN" _
    '"  rcljobs ON clients.client_no = rcljobs.client_no" _
#End Region

#Region "history"
    Public Delegate Sub mydelegatehist1()
    Public Shared Sub histcboinvoke()
        Try
            myForms.historry.Invoke(New mydelegatehist1(AddressOf cboload))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub cboload()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " SELECT  equip_info.*,  assigned_info.equip_id AS id2,  assigned_info.status " _
            '                    & " FROM  assigned_info RIGHT OUTER JOIN " _
            '                    & " equip_info ON  assigned_info.status =  equip_info.equip_id"
            'str += " and assigned_info.status='" & "0" & "'"
            Dim str As String = "SELECT *  from equip_info order by equip_id asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            myForms.historry.cboequip.Items.Add(Convert.ToString(.Fields("equip_id").Value) & " : " & _
                            Convert.ToString(.Fields("model_name").Value) & " : " & _
                            Convert.ToString(.Fields("model_no").Value))
                        Catch es300 As Exception
                        End Try
                        Application.DoEvents()
                        .MoveNext()
                    End While

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
    End Sub
    Public Shared historyjobno As String
    Public Delegate Sub mydelegatehist2()
    Public Shared Sub histequipdetailsinvoke()
        Try
            myForms.historry.Invoke(New mydelegatehist2(AddressOf equipdetailsload))
            Call equipdetailsinvoke()
            Call addhistcontrols()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub equipdetailsload()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " SELECT  equip_info.*,  assigned_info.equip_id AS id2,  assigned_info.status " _
            '                    & " FROM  assigned_info RIGHT OUTER JOIN " _
            '                    & " equip_info ON  assigned_info.status =  equip_info.equip_id"
            'str += " and assigned_info.status='" & "0" & "'"
            Dim str As String = "SELECT *  from history_equip" _
            & " where equip_id='" & historyjobno.Trim & "' order by equip_id asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then

                    Dim tojobDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim tojobDS As DataSet = New DataSet
                    tojobDA.Fill(tojobDS, rs, "equip history")
                    Dim tname As String = tojobDS.Tables(0).TableName()

                    Dim myColumn0a = New System.Data.DataColumn
                    myColumn0a.DataType = Type.GetType("System.Boolean")
                    myColumn0a.ColumnName = "Edit"
                    myColumn0a.DefaultValue = False
                    tojobDS.Tables(0).Columns.Add(myColumn0a)

                    myForms.historry.dtgequiphistory.SetDataBinding(tojobDS, tname)
                    addequiphistorytablestyle(tname)
                Else
                    myForms.historry.dtgequiphistory.DataSource = Nothing
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
    End Sub
    Public Shared Sub addequiphistorytablestyle(ByVal tname As String)
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = myForms.historry.dtgequiphistory.Width
            mywidth = mywidth / 8
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            '' Add a second column style.
            Dim mydesc1 As New DataGridBoolColumn
            mydesc1.MappingName = "Edit"
            mydesc1.HeaderText = "Edit"
            mydesc1.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc1)

            ' Add a second column style.
            Dim mydesc4 As New DataGridTextBoxColumn
            mydesc4.MappingName = "equip_id"
            mydesc4.HeaderText = "Equipment Id"
            mydesc4.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc4)

            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "job_no"
            myno.HeaderText = "Job Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)

            ' Add a second column style.
            'Dim myname As New DataGridTextBoxColumn()
            'myname.MappingName = "other"
            'myname.HeaderText = "Oher"
            'myname.Width = mywidth
            'ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "task"
            mydesc.HeaderText = "Task"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add a second column style.
            Dim mydesc6 As New DataGridTextBoxColumn
            mydesc6.MappingName = "description"
            mydesc6.HeaderText = "Description"
            mydesc6.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc6)

            ' Add a second column style.
            Dim mydesc66 As New DataGridTextBoxColumn
            mydesc66.MappingName = "assigned_by"
            mydesc66.HeaderText = "Assigned_by"
            mydesc66.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc66)

            ' Add a second column style.
            Dim mydesc666 As New DataGridTextBoxColumn
            mydesc666.MappingName = "estimate_release_date"
            mydesc666.HeaderText = "Estimate Release Date"
            mydesc666.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc666)

            ' Add a second column style.
            Dim mydesc666c As New DataGridTextBoxColumn
            mydesc666c.MappingName = "date_assigned"
            mydesc666c.HeaderText = "Date Assigned"
            mydesc666c.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc666c)

            ' Add a second column style.
            Dim mydesc6666 As New DataGridTextBoxColumn
            mydesc6666.MappingName = "date_released"
            mydesc6666.HeaderText = "Date Released"
            mydesc6666.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc6666)

            ' Add the DataGridTableStyle objects to the collection.
            myForms.historry.dtgequiphistory.TableStyles.Clear()
            myForms.historry.dtgequiphistory.TableStyles.Add(ts1)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
    End Sub
    Public Shared Sub addhistcontrols()
        Try
            'combo job
            myForms.historry.combojobno = New System.Windows.Forms.ComboBox
            myForms.historry.combojobno.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.historry.combojobno.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            myForms.historry.combojobno.Dock = DockStyle.Fill
            myForms.historry.combojobno.Visible = True

            'combo assigned
            myForms.historry.comboassignedby = New System.Windows.Forms.ComboBox
            myForms.historry.comboassignedby.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.historry.comboassignedby.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            myForms.historry.comboassignedby.Dock = DockStyle.Fill
            myForms.historry.comboassignedby.Visible = True

            'date assigned
            myForms.historry.dtpdas = New System.Windows.Forms.DateTimePicker
            myForms.historry.dtpdas.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.historry.dtpdas.Dock = DockStyle.Fill
            myForms.historry.dtpdas.Format = DateTimePickerFormat.Short
            myForms.historry.dtpdas.Visible = True

            'date released
            myForms.historry.dtpdre = New System.Windows.Forms.DateTimePicker
            myForms.historry.dtpdre.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.historry.dtpdre.Dock = DockStyle.Fill
            myForms.historry.dtpdre.Format = DateTimePickerFormat.Short
            myForms.historry.dtpdre.Visible = True

            ' estimated released date
            myForms.historry.dtperd = New System.Windows.Forms.DateTimePicker
            myForms.historry.dtperd.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.historry.dtperd.Dock = DockStyle.Fill
            myForms.historry.dtperd.Format = DateTimePickerFormat.Short
            myForms.historry.dtperd.Visible = True

            ' task
            myForms.historry.txttask = New System.Windows.Forms.TextBox
            myForms.historry.txttask.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.historry.txttask.Multiline = True
            myForms.historry.txttask.Dock = DockStyle.Fill
            myForms.historry.txttask.Visible = True
            'description
            myForms.historry.rtbdesc = New System.Windows.Forms.RichTextBox
            myForms.historry.rtbdesc.Cursor = System.Windows.Forms.Cursors.Arrow
            myForms.historry.rtbdesc.Multiline = True
            myForms.historry.rtbdesc.Dock = DockStyle.Fill
            myForms.historry.rtbdesc.Visible = True













            myForms.historry.datagridtextBox = CType(myForms.historry.dtgequiphistory.TableStyles(0).GridColumnStyles(2), DataGridTextBoxColumn)
            myForms.historry.datagridtextBox1 = CType(myForms.historry.dtgequiphistory.TableStyles(0).GridColumnStyles(3), DataGridTextBoxColumn)
            myForms.historry.datagridtextBox2 = CType(myForms.historry.dtgequiphistory.TableStyles(0).GridColumnStyles(4), DataGridTextBoxColumn)
            myForms.historry.datagridtextBox3 = CType(myForms.historry.dtgequiphistory.TableStyles(0).GridColumnStyles(5), DataGridTextBoxColumn)
            myForms.historry.datagridtextBox4 = CType(myForms.historry.dtgequiphistory.TableStyles(0).GridColumnStyles(6), DataGridTextBoxColumn)
            myForms.historry.datagridtextBox5 = CType(myForms.historry.dtgequiphistory.TableStyles(0).GridColumnStyles(7), DataGridTextBoxColumn)
            myForms.historry.datagridtextBox6 = CType(myForms.historry.dtgequiphistory.TableStyles(0).GridColumnStyles(8), DataGridTextBoxColumn)


            '--------------------------txttask
            myForms.historry.txttask.SendToBack()
            myForms.historry.datagridtextBox1.TextBox.Controls.Add(myForms.historry.txttask)
            myForms.historry.txttask.BringToFront()
            myForms.historry.datagridtextBox1.TextBox.BackColor = Color.White

            '--------------------------combojobno
            'myForms.historry.combojobno.SendToBack()
            'myForms.historry.datagridtextBox.TextBox.Controls.Add(myForms.historry.combojobno)
            'myForms.historry.combojobno.BringToFront()
            'myForms.historry.datagridtextBox.TextBox.BackColor = Color.White

            '--------------------------description
            myForms.historry.rtbdesc.SendToBack()
            myForms.historry.datagridtextBox2.TextBox.Controls.Add(myForms.historry.rtbdesc)
            myForms.historry.rtbdesc.BringToFront()
            myForms.historry.datagridtextBox2.TextBox.BackColor = Color.White

            '--------------------------assigned by
            'myForms.historry.comboassignedby.SendToBack()
            'myForms.historry.datagridtextBox3.TextBox.Controls.Add(myForms.historry.comboassignedby)
            'myForms.historry.comboassignedby.BringToFront()
            'myForms.historry.datagridtextBox3.TextBox.BackColor = Color.White

            '--------------------------estimate release date
            myForms.historry.dtperd.SendToBack()
            myForms.historry.datagridtextBox4.TextBox.Controls.Add(myForms.historry.dtperd)
            myForms.historry.dtperd.BringToFront()
            myForms.historry.datagridtextBox4.TextBox.BackColor = Color.White

            '--------------------------date    assigned       
            myForms.historry.dtpdas.SendToBack()
            myForms.historry.datagridtextBox5.TextBox.Controls.Add(myForms.historry.dtpdas)
            myForms.historry.dtpdas.BringToFront()
            myForms.historry.datagridtextBox5.TextBox.BackColor = Color.White

            '--------------------------date    assigned       
            myForms.historry.dtpdre.SendToBack()
            myForms.historry.datagridtextBox6.TextBox.Controls.Add(myForms.historry.dtpdre)
            myForms.historry.dtpdre.BringToFront()
            myForms.historry.datagridtextBox6.TextBox.BackColor = Color.White

        Catch et As Exception

        End Try
    End Sub
    Public Delegate Sub mydelegatehist3()
    Public Shared Sub equipdetailsinvoke()
        Try
            myForms.equipments.Invoke(New mydelegatehist3(AddressOf equipdetails))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub equipdetails()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = " select * from equip_info" _
            & " where equip_id='" & historyjobno.Trim & "' order by equip_id asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim equipDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim equipDS As DataSet = New DataSet
                    equipDA.Fill(equipDS, rs, "equip details")
                    Dim tname As String = equipDS.Tables(0).TableName()


                    myForms.historry.dtgequipdetails.SetDataBinding(equipDS, tname)
                    addequipdetailstablestyle(tname)
                Else
                    myForms.historry.dtgequipdetails.DataSource = Nothing
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
    End Sub
    Public Shared Sub addequipdetailstablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.historry.dtgequipdetails.Width - 20
            mywidth = mywidth / 10

            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "equip_id"
            myno.HeaderText = "Equipment Id"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.

            ' Add a second column style.
            Dim myname100v As New DataGridTextBoxColumn
            myname100v.MappingName = "manufacturer"
            myname100v.HeaderText = "Manufacturer"
            myname100v.Width = mywidth
            ts1.GridColumnStyles.Add(myname100v)

            Dim myname1 As New DataGridTextBoxColumn
            myname1.MappingName = "model_no"
            myname1.HeaderText = "Model No"
            myname1.Width = mywidth

            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "serial_no"
            myname.HeaderText = "Serial No"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "model_name"
            mydesc.HeaderText = "Model Name"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn
            myname100.MappingName = "purchase_date"
            myname100.HeaderText = "Purchase Date"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn
            mydesc2.MappingName = "description"
            mydesc2.HeaderText = "Description"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc200 As New DataGridTextBoxColumn
            mydesc200.MappingName = "licence"
            mydesc200.HeaderText = "Licence"
            mydesc200.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200)

            ' Add a second column style.
            Dim mydesc2001 As New DataGridTextBoxColumn
            mydesc2001.MappingName = "guarantee"
            mydesc2001.HeaderText = "Guarantee"
            mydesc2001.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001)

            ' Add a second column style.
            Dim mydesc2001a As New DataGridTextBoxColumn
            mydesc2001a.MappingName = "condition"
            mydesc2001a.HeaderText = "Condition"
            mydesc2001a.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001a)

            ' Add a second column style.
            Dim mydesc2001b As New DataGridTextBoxColumn
            mydesc2001b.MappingName = "type"
            mydesc2001b.HeaderText = "Type"
            mydesc2001b.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001b)

            ' Add a second column style.
            Dim mydesc2001c As New DataGridTextBoxColumn
            mydesc2001c.MappingName = "model_year"
            mydesc2001c.HeaderText = "Model Year"
            mydesc2001c.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001c)

            ' Add the DataGridTableStyle objects to the collection.
            myForms.historry.dtgequipdetails.TableStyles.Clear()
            myForms.historry.dtgequipdetails.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "Maintenance"
    Public Delegate Sub mydelegateserv1()
    Public Shared Sub servcboinvoke()
        Try
            myForms.maintenace.Invoke(New mydelegateserv1(AddressOf cboservload))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub cboservload()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " SELECT  equip_info.*,  assigned_info.equip_id AS id2,  assigned_info.status " _
            '                    & " FROM  assigned_info RIGHT OUTER JOIN " _
            '                    & " equip_info ON  assigned_info.status =  equip_info.equip_id"
            'str += " and assigned_info.status='" & "0" & "'"
            Dim str As String = "SELECT *  from equip_info order by equip_id asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()
                    While .EOF = False
                        Try
                            myForms.maintenace.cboequip.Items.Add(Convert.ToString(.Fields("equip_id").Value) & " : " & _
                            Convert.ToString(.Fields("model_name").Value) & " : " & _
                            Convert.ToString(.Fields("model_no").Value))
                        Catch es300 As Exception
                        End Try
                        Application.DoEvents()
                        .MoveNext()
                    End While

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
    End Sub
    Public Shared servno As String
    Public Delegate Sub mydelegateserv2()
    Public Shared Sub servequipdetailsinvoke()
        Try
            myForms.maintenace.Invoke(New mydelegateserv2(AddressOf servequipdetailsload))
            Call servdetailsinvoke()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub servequipdetailsload()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            'Dim str As String = " SELECT  equip_info.*,  assigned_info.equip_id AS id2,  assigned_info.status " _
            '                    & " FROM  assigned_info RIGHT OUTER JOIN " _
            '                    & " equip_info ON  assigned_info.status =  equip_info.equip_id"
            'str += " and assigned_info.status='" & "0" & "'"
            Dim str As String = "SELECT *  from equip_info" _
            & " where equip_id='" & servno.Trim & "' order by equip_id asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then

                    Dim servDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim servDS As DataSet = New DataSet
                    servDA.Fill(servDS, rs, "maintenace")
                    Dim tname As String = servDS.Tables(0).TableName()

                    myForms.maintenace.dtgequipdetails.SetDataBinding(servDS, tname)
                    addequipservtablestyle(tname)
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
    End Sub
    Public Shared Sub addequipservtablestyle(ByVal tname As String)
        Dim currentcursor As Cursor = Cursor.Current
        Try
            Cursor.Current = Cursors.WaitCursor
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = tname
            Dim mywidth As Integer
            mywidth = myForms.maintenace.dtgequipdetails.Width
            mywidth = mywidth / 8
            ' Add a GridColumnStyle and set its MappingName
            ' to the name of a DataColumn in the DataTable.
            ' Set the HeaderText and Width properties. 
            ' Add a second column style.

            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "equip_id"
            myno.HeaderText = "Equipment Id"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.

            ' Add a second column style.
            Dim myname100v As New DataGridTextBoxColumn
            myname100v.MappingName = "manufacturer"
            myname100v.HeaderText = "Manufacturer"
            myname100v.Width = mywidth
            ts1.GridColumnStyles.Add(myname100v)

            Dim myname1 As New DataGridTextBoxColumn
            myname1.MappingName = "model_no"
            myname1.HeaderText = "Model No"
            myname1.Width = mywidth

            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "serial_no"
            myname.HeaderText = "Serial No"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "model_name"
            mydesc.HeaderText = "Model Name"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)

            ' Add a second column style.
            Dim myname100 As New DataGridTextBoxColumn
            myname100.MappingName = "purchase_date"
            myname100.HeaderText = "Purchase Date"
            myname100.Width = mywidth
            ts1.GridColumnStyles.Add(myname100)


            ' Add a second column style.
            Dim mydesc2 As New DataGridTextBoxColumn
            mydesc2.MappingName = "description"
            mydesc2.HeaderText = "Description"
            mydesc2.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2)

            ' Add a second column style.
            Dim mydesc200 As New DataGridTextBoxColumn
            mydesc200.MappingName = "licence"
            mydesc200.HeaderText = "Licence"
            mydesc200.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc200)

            ' Add a second column style.
            Dim mydesc2001 As New DataGridTextBoxColumn
            mydesc2001.MappingName = "guarantee"
            mydesc2001.HeaderText = "Guarantee"
            mydesc2001.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001)

            ' Add a second column style.
            Dim mydesc2001a As New DataGridTextBoxColumn
            mydesc2001a.MappingName = "condition"
            mydesc2001a.HeaderText = "Condition"
            mydesc2001a.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001a)

            ' Add a second column style.
            Dim mydesc2001b As New DataGridTextBoxColumn
            mydesc2001b.MappingName = "type"
            mydesc2001b.HeaderText = "Type"
            mydesc2001b.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001b)

            ' Add a second column style.
            Dim mydesc2001c As New DataGridTextBoxColumn
            mydesc2001c.MappingName = "model_year"
            mydesc2001c.HeaderText = "Model Year"
            mydesc2001c.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc2001c)

            ' Add the DataGridTableStyle objects to the collection.
            myForms.maintenace.dtgequipdetails.TableStyles.Clear()
            myForms.maintenace.dtgequipdetails.TableStyles.Add(ts1)
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        Finally

            Cursor.Current = currentcursor
        End Try
    End Sub
    Public Delegate Sub mydelegateserv3()
    Public Shared Sub servdetailsinvoke()
        Try
            myForms.maintenace.Invoke(New mydelegateserv3(AddressOf servdetails))
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString() & vbCrLf _
            & ex.InnerException().ToString() & vbCrLf _
            & ex.StackTrace.ToString())
        End Try
    End Sub
    Public Shared Sub servdetails()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim str As String = " select * from maintenance_info" _
            & " where equip_id='" & servno.Trim & "' order by equip_id asc"
            Dim rs As New ADODB.Recordset
            With rs
                .CursorLocation = CursorLocationEnum.adUseClient
                .CursorType = CursorTypeEnum.adOpenStatic
                .Open(str, connect)
                If .BOF = False And .EOF = False Then
                    Dim equipDA As OleDbDataAdapter = New OleDbDataAdapter
                    Dim equipDS As DataSet = New DataSet
                    equipDA.Fill(equipDS, rs, "equip details")
                    Dim tname As String = equipDS.Tables(0).TableName()


                    myForms.maintenace.dtgequipmaintenace.SetDataBinding(equipDS, tname)
                    addservdetailstablestyle(tname)
                Else
                    myForms.maintenace.dtgequipmaintenace.DataSource = Nothing
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
    End Sub
    Public Shared Sub addservdetailstablestyle(ByVal namme As String)
        Try
            ' Create a new DataGridTableStyle and set
            ' its MappingName to the TableName of a DataTable. 
            Dim ts1 As New DataGridTableStyle
            ts1.MappingName = namme
            Dim mywidth As Integer
            mywidth = myForms.maintenace.dtgequipdetails.Width - 20
            mywidth = mywidth / 5


            ' Add a second column style.
            Dim myname100v As New DataGridTextBoxColumn
            myname100v.MappingName = "equip_id"
            myname100v.HeaderText = "equipment Id"
            myname100v.Width = mywidth
            ts1.GridColumnStyles.Add(myname100v)

            Dim myname1 As New DataGridTextBoxColumn
            myname1.MappingName = "service_date"
            myname1.HeaderText = "Service Date"
            myname1.Width = mywidth

            ts1.GridColumnStyles.Add(myname1)

            ' Add a second column style.
            Dim myname As New DataGridTextBoxColumn
            myname.MappingName = "description"
            myname.HeaderText = "Description"
            myname.Width = mywidth
            ts1.GridColumnStyles.Add(myname)

            ' Add a second column style.
            Dim mydesc As New DataGridTextBoxColumn
            mydesc.MappingName = "cost_incurred"
            mydesc.HeaderText = "Cost Incurred"
            mydesc.Width = mywidth
            ts1.GridColumnStyles.Add(mydesc)


            Dim myno As New DataGridTextBoxColumn
            myno.MappingName = "invoice_no"
            myno.HeaderText = "Invoice Number"
            myno.Width = mywidth
            ts1.GridColumnStyles.Add(myno)
            ' Add a second column style.
            ' Add the DataGridTableStyle objects to the collection.
            myForms.maintenace.dtgequipmaintenace.TableStyles.Clear()
            myForms.maintenace.dtgequipmaintenace.TableStyles.Add(ts1)
        Catch ex As Exception
        End Try
    End Sub
#End Region

#End Region

#Region "gross margin"
    Public Shared gccost As String
    Public Shared Sub gross()
        Try

            Try
                '-------------finances
                casualsinvoke()
                accomodationinvoke()
                travelinvoke()
                '----------------- 
                '-----------equipments
                ramaniequipinvoke()
                hiredequipinvoke()
                '--------------------
                '-----------personnel
                jobsinvoke()
                '-----------
                Try
                    Try
                        myForms.CustomerForm2.totalkost = Convert.ToDouble(myForms.CustomerForm2.kaccomodation) + _
                        Convert.ToDouble(myForms.CustomerForm2.kcasual)
                    Catch

                    End Try
                    Try
                        myForms.CustomerForm2.totalkost += Convert.ToDouble(myForms.CustomerForm2.khired)
                    Catch

                    End Try
                    Try
                        myForms.CustomerForm2.totalkost += Convert.ToDouble(myForms.CustomerForm2.kpersonnel)
                    Catch

                    End Try
                    Try
                        myForms.CustomerForm2.totalkost += Convert.ToDouble(myForms.CustomerForm2.kramani)

                    Catch

                    End Try
                    Try
                        myForms.CustomerForm2.totalkost += Convert.ToDouble(myForms.CustomerForm2.ktravel)

                    Catch

                    End Try

                    If Convert.ToDouble(myForms.CustomerForm2.totalincome) <> 0 Then
                        Dim ccost As Double = ((Convert.ToDouble(myForms.CustomerForm2.totalincome) - _
                         Convert.ToDouble(myForms.CustomerForm2.totalkost)) / Convert.ToDouble(myForms.CustomerForm2.totalincome)) * 100
                        myForms.CustomerForm2.lblgrossmargin.Text = Math.Round(Convert.ToDecimal(ccost), 2)
                        If ccost > 75 Then
                            myForms.CustomerForm2.lblgrossmargin.BackColor = Color.Green
                        ElseIf ccost < 75 And ccost > 60 Then
                            myForms.CustomerForm2.lblgrossmargin.BackColor = Color.Orange
                        Else
                            myForms.CustomerForm2.lblgrossmargin.BackColor = Color.Red
                        End If
                        gccost = myForms.CustomerForm2.lblgrossmargin.Text
                        Dim strf As String
                        strf = " update rcljobs set  grossmargin='" & myForms.CustomerForm2.lblgrossmargin.Text & "'"
                        strf += " where job_no='" & strjobno & "';"
                        strf += "update grossmargin set personnel='" & myForms.CustomerForm2.kpersonnel & "'"
                        strf += " ,casual='" & myForms.CustomerForm2.kcasual & "'"
                        strf += " ,accomodation='" & myForms.CustomerForm2.kaccomodation & "'"
                        strf += " ,travel='" & myForms.CustomerForm2.ktravel & "'"
                        strf += " ,ramani='" & myForms.CustomerForm2.kramani & "'"
                        strf += " ,hired='" & myForms.CustomerForm2.khired & "'"
                        strf += " where job_no='" & strjobno & "'"
                        Try
                            Dim connectstr As String = "DSN=" & myForms.qconnstr
                            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
                            Dim connect As New ADODB.Connection
                            connect.Mode = ConnectModeEnum.adModeReadWrite
                            connect.CursorLocation = CursorLocationEnum.adUseClient
                            connect.ConnectionString = connectstr
                            connect.Open()
                            connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                            connect.BeginTrans()
                            connect.Execute(strf)
                            connect.CommitTrans()
                        Catch az As Exception

                        End Try
                    End If

                Catch zx As Exception

                End Try
            Catch ex As Exception

            End Try
        Catch vbnm As Exception
        End Try

    End Sub
    Public Shared wwhich As String
    Public Shared Sub loopgross()
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim strf As String
            Dim rs As New ADODB.Recordset
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.CursorType = CursorTypeEnum.adOpenKeyset
            rs.Open("select grossmargin.* ,rcljobs.amount from grossmargin inner join rcljobs on " _
            & " grossmargin.job_no=rcljobs.job_no order by grossmargin.job_no desc", _
            connect)
            If rs.EOF = False And rs.BOF = False Then
                Dim jno, amount As String
                Dim gm As Double
                Dim gmm As String
                rs.MoveFirst()
                While rs.EOF = False
                    Try
                        Application.DoEvents()
                        jno = rs.Fields("job_no").Value
                        amount = rs.Fields("amount").Value
                        gm = 0
                        Try
                            gm = Convert.ToDouble(rs.Fields("personnel").Value)
                        Catch fg As Exception
                        End Try
                        Try
                            gm += Convert.ToDouble(rs.Fields("casual").Value)
                        Catch fg As Exception
                        End Try
                        Try
                            gm += Convert.ToDouble(rs.Fields("accomodation").Value)
                        Catch fg As Exception
                        End Try
                        Try
                            gm += Convert.ToDouble(rs.Fields("travel").Value)
                        Catch fg As Exception
                        End Try
                        Try
                            gm += Convert.ToDouble(rs.Fields("ramani").Value)
                        Catch fg As Exception
                        End Try
                        Try
                            gm += Convert.ToDouble(rs.Fields("hired").Value)
                        Catch fg As Exception
                        End Try
                        Try
                            gm = ((Convert.ToDouble(amount) - gm) / Convert.ToDouble(amount)) * 100
                            gm = Math.Round(gm, 2)
                        Catch we As Exception
                            Try
                                gm = ""
                            Catch ex As Exception

                            End Try

                        End Try

                        Try
                            If Double.IsInfinity(gm) = True _
                            Or Double.IsNaN(gm) = True _
                            Or Double.IsNegativeInfinity(gm) = True _
                            Or Double.IsPositiveInfinity(gm) = True Then
                                gmm = ""
                            Else
                                gmm = gm.ToString()
                            End If
                        Catch ex As Exception

                        End Try


                        strf = " update rcljobs set grossmargin='" & gmm.ToString() & "'"
                        strf += " where job_no='" & jno & "'"
                        Try
                            connect.IsolationLevel = connect.IsolationLevel.adXactSerializable
                            connect.BeginTrans()
                            connect.Execute(strf)
                            connect.CommitTrans()
                        Catch sd As Exception

                        End Try
                    Catch cv As Exception
                    End Try
                    rs.MoveNext()
                End While
            End If
            Select Case wwhich
                Case "0"
                    'btnjobsearch_Click(Me, e)
                    Try
                        Dim tjd As System.Threading.Thread = New System.Threading.Thread( _
                        AddressOf myForms.Main.jseinvoke)
                        tjd.IsBackground = True
                        tjd.Start()
                    Catch xa As Exception

                    End Try
                Case "1"
                    'btnCurrentJobs_Click(Me, e)
                    Dim tjd As System.Threading.Thread = New System.Threading.Thread( _
                       AddressOf myForms.Main.cjinvoke)
                    tjd.IsBackground = True
                    tjd.Start()
                Case "2"
                    'btnCompletedJobs_Click(Me, e)
                    Try
                        Dim tjd As System.Threading.Thread = New System.Threading.Thread( _
                        AddressOf myForms.Main.jcinvoke)
                        tjd.IsBackground = True
                        tjd.Start()
                    Catch xa As Exception

                    End Try
                Case "3"
                    'btnjobdelivered_Click(Me, e)
                    Dim tjd As System.Threading.Thread = New System.Threading.Thread( _
                       AddressOf myForms.Main.jdinvoke)
                    tjd.IsBackground = True
                    tjd.Start()
                Case "4"
                    'btnjobshowall_Click(Me, e)
                    Dim tjd As System.Threading.Thread = New System.Threading.Thread( _
                       AddressOf myForms.Main.jsinvoke)
                    tjd.IsBackground = True
                    tjd.Start()
                Case Else
            End Select
        Catch ex As Exception
        End Try
    End Sub
#End Region

#Region "general functions"
    Public Shared clientno As String
    Public Shared Sub newlno2()
        Dim number, number1, number2 As String
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim rs As New ADODB.Recordset

            With rs
                'add new client number
                .Open(Source:="select leads_no  from clients" _
                & " where client_no ='" & clientno & "' ", _
                activeconnection:=connect, cursortype:=CursorTypeEnum.adOpenForwardOnly)
                number = clientno
                'If Convert.IsDBNull(.Fields("leads_no").Value) = False Then
                If .BOF = False And .EOF = False Then
                    If Convert.IsDBNull(.Fields("leads_no").Value) = False Then
                        number1 = .Fields("leads_no").Value
                        Dim l
                        l = number1.Length - 5
                        number2 = number1.Substring(5, l)
                        number2 = CStr(CLng(number2) + 1)
                        Select Case number2.Length
                            Case 1
                                number2 = "000" & number2
                            Case 2
                                number2 = "00" & number2
                            Case 3
                                number2 = "0" & number2
                            Case Else

                        End Select

                        number += "2" & number2
                    Else
                        number += "2" & "0001"

                    End If
                Else
                    number += "2" & "0001"
                End If

                .Close()
            End With

        Catch ex As Exception


        End Try
        Try
            myForms.CustomerForm2.txtJobNo.Text = number
        Catch zx As Exception

        End Try
    End Sub
    Public Shared Sub newlnoinvoke()
        Try
            myForms.CustomerForm2.Invoke(New mydelegatenew1(AddressOf newlno2))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegatenew1()
    Public Shared Sub neno() 'eqipment number
        Dim number, number1, number2 As String
        Try
            Dim connectstr As String = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()
            Dim rs As New ADODB.Recordset

            With rs
                'add new client number
                .Open(Source:="select leads_no  from clients" _
                & " where client_no ='" & clientno & "' ", _
                activeconnection:=connect, cursortype:=CursorTypeEnum.adOpenForwardOnly)
                number = clientno
                'If Convert.IsDBNull(.Fields("leads_no").Value) = False Then
                If .BOF = False And .EOF = False Then
                    If Convert.IsDBNull(.Fields("leads_no").Value) = False Then
                        number1 = .Fields("leads_no").Value
                        Dim l
                        l = number1.Length - 5
                        number2 = number1.Substring(5, l)
                        number2 = CStr(CLng(number2) + 1)
                        Select Case number2.Length
                            Case 1
                                number2 = "000" & number2
                            Case 2
                                number2 = "00" & number2
                            Case 3
                                number2 = "0" & number2
                            Case Else

                        End Select

                        number += "2" & number2
                    Else
                        number += "2" & "0001"

                    End If
                Else
                    number += "2" & "0001"
                End If

                .Close()
            End With

        Catch ex As Exception


        End Try
        Try
            myForms.CustomerForm2.txtJobNo.Text = number
        Catch zx As Exception

        End Try
    End Sub
    Public Shared Sub nenoinvoke()
        Try
            'myForms.Invoke(New mydelegateneno1(AddressOf neno))
        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
    Public Delegate Sub mydelegateneno1()
#End Region

End Class
