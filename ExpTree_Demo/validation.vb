Imports System
Public Class validation
    Public Shared Function _validatetextbox(ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        _validatetextbox = True
        Try
            Select Case e.KeyChar
                Case "'"
                    e.Handled = True 'it indicates the event is handled.

                    'Case "%"
                    '    e.Handled = True 'it indicates the event is handled.

                    'Case "\"
                    '    e.Handled = True 'it indicates the event is handled.

                    'Case """"
                    '    e.Handled = True 'it indicates the event is handled.

                Case Else
                    _validatetextbox = False

            End Select
        Catch we As Exception

        End Try
    End Function
End Class
