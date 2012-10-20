Imports System
Imports ADODB

Public Class datemanipulation
    Public Shared Sub tdate() '-------transmit date
        ' Use the StrArg field as an argument.
        Dim connectstr As String
         connectstr = "DSN=" & myForms.qconnstr
        'connectstr = "DSN=RCL_DB;server=192.168.1.60;initial catalog=RCL_DB;"
        Dim connect As New ADODB.Connection()
        connect.Mode = ConnectModeEnum.adModeReadWrite
        connect.CursorLocation = CursorLocationEnum.adUseClient
        connect.ConnectionString = connectstr
        connect.Open()
        Try
            Dim rs As New ADODB.Recordset()
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.CursorType = CursorTypeEnum.adOpenStatic
            Dim str As String
            str = "select *  from storedate "
            rs.Open(str, connect)
            If rs.BOF = False And rs.EOF = False Then
                myForms.Main.currentdate = rs.Fields("curdate").Value
            Else
                myForms.Main.currentdate = ""
            End If

        Catch we As Exception

        End Try
        Try
            connect.Close()
        Catch qw As Exception

        End Try
    End Sub
    Public Delegate Sub mydelegate()
    Public Shared Sub curdateinvoke()
        Try
            myForms.Main.Invoke(New mydelegate(AddressOf tdate))
        Catch ex As Exception
            '-----------crashes
            'MessageBox.Show(ex.Message.ToString() & vbCrLf _
            '& ex.InnerException().ToString() & vbCrLf _
            '& ex.StackTrace.ToString())
        End Try
    End Sub
#Region "validate time"
    Public Shared dbtime As String 'database time okay
    Public Shared ctime As String 'current time
    Public Shared Function datediff() As String
        Try
            Dim ddiff As New System.TimeSpan()
            Dim d1 As New DateTimePicker()
            Dim d2 As New DateTimePicker()
            d1.Value = dbtime 'database time
            d2.Value = ctime 'displayed time

            ddiff = d1.Value.Subtract(d2.Value)
            datediff = ddiff.TotalDays
        Catch ex As Exception
            datediff = ""
        End Try
    End Function
#End Region



End Class
