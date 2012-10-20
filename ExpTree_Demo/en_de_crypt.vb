Imports System
Imports System.Text
Public Class en_de_crypt
    Public Function EnDeCrypt(ByVal strTextIn As String, ByVal strPassword As String, ByVal blnEnDeCrypt As Boolean) As String
        Try
            ' Declarations 
            Dim intForX As Integer 'used for the for .. next loops 
            Dim strOutput As String 'holds the output 
            Dim intArrayCount As Integer 'holds the arraycount 
            Dim intTempValue As Integer 'holds a temporary 

            ' Create an array to hold the individual values from the password 
            Dim intPasswordValue() As Integer
            ReDim intPasswordValue(Len(strPassword))

            ' Fill the array 
            For intForX = 1 To Len(strPassword)
                Application.DoEvents()
                intPasswordValue(intForX) = Asc(Mid(strPassword, 1, 1))
            Next intForX

            'Be sure the outputstring is empty 
            strOutput = ""

            ' start on the first array position 
            intArrayCount = 1

            'En or Decrypt the strTextIn 
            For intForX = 1 To Len(strTextIn)
                Application.DoEvents()
                If blnEnDeCrypt = True Then 'we are encrypting 
                    ' shifted the letters based on the password 
                    intTempValue = Asc(Mid(strTextIn, intForX, 1)) - _
                                   intPasswordValue(intArrayCount)
                    'Be sure the value is valid asc 
                    If intTempValue < 1 Then
                        intTempValue = intTempValue + 255
                    End If

                Else 'we are decrypting 
                    intTempValue = Asc(Mid(strTextIn, intForX, 1)) + _
                                intPasswordValue(intArrayCount)
                    'Be sure the value is valid asc 
                    If intTempValue > 256 Then
                        intTempValue = intTempValue - 255
                    End If

                End If

                ' add to the output string 
                strOutput = strOutput & Chr(intTempValue)

                ' goto the next array position 
                intArrayCount = intArrayCount + 1

                ' if we are at the end of the array goto the begining 
                If intArrayCount = UBound(intPasswordValue) Then
                    intArrayCount = 1
                End If

            Next intForX

            ' Return the output 
            Dim strbuilder As New StringBuilder(strOutput)
            EnDeCrypt = strOutput
            Dim bcvc As String = "sdfff"
        Catch ex As Exception

        End Try


    End Function
End Class
