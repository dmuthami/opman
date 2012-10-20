Imports System
Imports System.Diagnostics
Imports System.Threading

'Namespace processdemo

Module ProcessDemo
    'environment class provides information to manipulate the current environment and platform
    Public Sub Mainpro()
        Try
            Dim args As String()
            Dim appName As String
            'returns an array for the command line arguments
            args = Environment.GetCommandLineArgs()
            appName = args(0)

            'If (args.Length <> 2) Then
            '    Console.WriteLine("Usage: " + appName + " <executable>")
            '    Exit Sub
            'End If

            Dim executableFilename As String
            Dim a() As String
            a = appName.Split("\")
            Dim i As Integer
            i = a.GetUpperBound(0)
            executableFilename = a(i)

            Dim process As Process
            process = New Process()
            process.StartInfo.FileName = executableFilename
            process.Start()

            process.WaitForInputIdle()

            Thread.Sleep(1000)
            If (Not process.CloseMainWindow()) Then
                process.Kill()
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Sub
    Public Sub closeprogram()
        Try
            Dim args As String()
            Dim appName As String
            'Provides information about, and means to manipulate, the current environment and platform
            'returns an array for the command line arguments
            args = Environment.GetCommandLineArgs()
            appName = args(0)

            'If (args.Length <> 2) Then
            '    Console.WriteLine("Usage: " + appName + " <executable>")
            '    Exit Sub
            'End If

            Dim executableFilename As String
            Dim a() As String
            a = appName.Split("\")
            Dim i As Integer
            i = a.GetUpperBound(0)
            executableFilename = a(i)
            a.Initialize()
            a = executableFilename.Split(".")

            Dim myprocess As New Process() 'Create an instance of the Process component class in code. 
            Dim myprocesses() As Process
            'binding to existing processes
            myprocesses = Process.GetProcessesByName(a(0)) '' Returns array containing all instances of "RCL_DB".
            Dim iprocess As Integer
            iprocess = myprocesses.GetUpperBound(0)
            Dim j As Integer
            If iprocess >= 1 Then
                For j = 0 To iprocess - 1 'remove all instances leaving only one of them
                    Application.DoEvents()
                    myprocesses(j).Kill() 'literally destroy the processes
                Next
            Else
                If isloading = False Then
                    myprocesses(0).Kill() 'remove all instances
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

End Module

'End Namespace

