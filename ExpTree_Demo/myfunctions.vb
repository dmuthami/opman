Imports ADODB
Imports System
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Xml

'Imports System.Data.OleDb
'Imports System.Data.OleDb.OleDbConnection

Imports System.ComponentModel
'Imports System.Data.SqlClient

Imports System.Text.StringBuilder
Module myfunctions
    'Public login As frmLogin = Nothing

    Public myclientno As String
    Public myfrmAddClientsform As Integer = 0
    Public jobsform As Boolean
    Public addjobs As Boolean
    Public editjob As Boolean
    Public addcontacts As Boolean
    Public editcontacts As Boolean
    Public hasloadedjobsheet As Boolean = False

    'leads
    Public addleads As Boolean = False
    Public editleads As Boolean = False

    Public editclients As Boolean
    Public isloading As Boolean
    Public myclientname As String
    Public lcg As Boolean
    Public lej As Boolean
    Public isrownumberzero As String
    'refresh grid decllarations
    Public refreshclients As Boolean = False
    Public refreshjobs As Boolean = False
    Public refreshcontacts As Boolean = False
    Public refreshleads As Boolean = False
    Public refreshleadshome As Boolean = False
    'Public conn As New o()

    'contacts

    Public myjobno As String
    Public myjobtitle As String

#Region "connection"

    'Public Sub dbconnect()
    '    Try
    '        If connect.State <> Nothing Then

    '            Exit Try
    '        End If
    '        lcg = False
    '        lej = False
    '        Dim connectstr As String
    '        Dim nv As NameValueCollection
    '        nv = ConfigurationSettings.AppSettings()
    '        connectstr = nv("connectionstring")
    '        'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
    '        connect.Mode = ConnectModeEnum.adModeReadWrite
    '        connect.CursorLocation = CursorLocationEnum.adUseClient
    '        connect.ConnectionString = connectstr
    '        connect.Open()

    '    Catch t As Exception
    '        MsgBox(t.Message.ToString)
    '    End Try
    'End Sub

    ''Public Sub GetCellValue(ByVal myGrid As DataGrid)
    ''    Try
    ''        Dim myCell As New DataGridCell()
    ''        ' Use an arbitrary cell.
    ''        myCell.RowNumber = myGrid.CurrentRowIndex
    ''        'isrownumberzero = CStr(myGrid.CurrentRowIndex)
    ''        myCell.ColumnNumber = 0
    ''        myclientno = myGrid(myCell)
    ''        myCell.ColumnNumber = 1
    ''        myclientname = myGrid(myCell)
    ''        Dim myform As New frmEditClients()
    ''        If myform.ShowInTaskbar = True Then

    ''        End If
    ''    Catch ex As Exception
    ''        myclientno = "0"
    ''    End Try
    ''End Sub
    'Public Sub disconnect()
    '    Try
    '        With connect
    '            .Close()

    '        End With
    '        connect = Nothing
    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Function loadgridcontact1()
    '    'Dim currentcursor As Cursor = Cursor.Current
    '    Try
    '        'Cursor.Current = Cursors.WaitCursor
    '        Dim dadap As New OdbcDataAdapter("select * from contacts where client_no='" & myclientno & "'", "DSN=RCL_DB")
    '        Dim dset As New DataSet()
    '        dadap.Fill(dset, "contacts")
    '        Dim datview As DataView = dset.Tables("contacts").DefaultView
    '        ' dtgContacts.DataSource = datview
    '        loadgridcontact1 = datview





    '    Catch t As Exception
    '        MessageBox.Show("error" & t.InnerException.ToString, "Error" _
    '        , MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    Finally
    '        'Cursor.Current = currentcursor
    '    End Try
    'End Function
    Public Sub GetCellValuedtgjobs(ByVal myGrid As DataGrid)
        Try
            Dim myCell As New DataGridCell()
            ' Use an arbitrary cell.
            myCell.RowNumber = myGrid.CurrentRowIndex
            'isrownumberzero = CStr(myGrid.CurrentRowIndex)
            myCell.ColumnNumber = 0
            myjobno = myGrid(myCell)
            myCell.ColumnNumber = 1
            myjobtitle = myGrid(myCell)

        Catch ex As Exception
            MessageBox.Show(text:=ex.ToString)
        End Try
    End Sub

#End Region

#Region "other functions"
    Public Function rml(ByVal b() As String) As String
        'procedure removes multiple lines in a textbox
        Dim i As Integer
        i = b.GetUpperBound(0)
        Dim str As String
        Dim j As Integer
        'rml = " " & b(0)
        For j = 0 To i
            str = b(j)
            str = str.Trim
            If str <> "" Then
                If j = 0 Then
                    rml = str
                Else
                    rml = rml & " " & " " & str
                End If
            End If
        Next
    End Function
    Public Function editclientstatus(ByVal cno As String)
        Try
            Dim connectstr As String
             connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

            Dim rs As New ADODB.Recordset()
            Dim status As String
            Dim str = "PHO"
            With rs
                'add new client number
                .Open(Source:="select distinct status  from leads" _
                & " where client_no ='" & cno & "'", _
                activeconnection:=connect, cursortype:=CursorTypeEnum.adOpenForwardOnly)
                If .BOF = False And .EOF = False Then
                    .MoveFirst()

                    While .EOF = False
                        status = .Fields("status").Value
                        Select Case status
                            Case "Suspect"
                                str = "Suspect"
                            Case "Prospect"
                                If str = "Proposal" Or str = "PHO" Then
                                    str = "Prospect"
                                End If
                            Case "Proposal"
                                If str = "PHO" Then
                                    str = "Proposal"
                                End If
                            Case "PHO"
                                str = "PHO"
                            Case Else

                        End Select
                        .MoveNext()
                        Application.DoEvents()
                    End While
                Else
                    str = ""
                End If
            End With
            Return str
            Try
                connect.Close()
            Catch bn As Exception

            End Try
        Catch ex As Exception

        End Try
    End Function
    Public Function newlno(ByVal clientno As String) As String
        Try
            Dim connectstr As String
             connectstr = "DSN=" & myForms.qconnstr
            'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
            Dim connect As New ADODB.Connection()
            connect.Mode = ConnectModeEnum.adModeReadWrite
            connect.CursorLocation = CursorLocationEnum.adUseClient
            connect.ConnectionString = connectstr
            connect.Open()

            Dim rs As New ADODB.Recordset()
            Dim number, number1, number2 As String
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
            Return number.ToString()
        Catch ex As Exception

        End Try
    End Function



#End Region

End Module

