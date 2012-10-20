Imports System.Security.Cryptography
Imports System.IO
Imports System
Imports System.Text




Public Class crypt

#Region "private members"
    'We need these Global Objects, So that we can Access that Memory Stream again
    'Although the mCryptoProv object is not necessary to be Global. Since we are using
    'custom Key and IV
    Dim mCryptProv As SymmetricAlgorithm
    Dim mMemStr As MemoryStream
    Private txtkey As String = "save"
    Private txtIV As String = "saved"
    Public txtData As String = ""
#End Region

    Public Sub encrypt()
        'Check if something Algo is Selected 
        'If Me.cmbAlgo.SelectedIndex < 0 Then
        '    MsgBox("Select an Algo to use to Encrypt or Decrypt.", MsgBoxStyle.Critical)
        '    Exit Sub
        'End If
        'Since, The selected algo may not be Symmetric, So Try it
        Try
            Me.mCryptProv = SymmetricAlgorithm.Create("Rijndael")
        Catch ee As Exception
            MsgBox("Exception : " & ee.Message, MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Try
            'Each Algo has different Size of Key and Block, So let's try
            mCryptProv.BlockSize = 256
            mCryptProv.KeySize = 256
        Catch
            'Hummmmmn, Invalid Block or Key Size. Show the Default Ones
            Dim mStr As String
            mStr = "Invalid Block or Key Size." & vbCrLf & "Valid Sizes are : " & vbCrLf
            mStr += "Block Size " & mCryptProv.LegalBlockSizes(0).MinSize & " - " & mCryptProv.LegalBlockSizes(0).MaxSize & " With Increment of " & mCryptProv.LegalBlockSizes(0).SkipSize & vbCrLf
            mStr += "Key  Size " & mCryptProv.LegalKeySizes(0).MinSize & " - " & mCryptProv.LegalKeySizes(0).MaxSize & " With Increment of " & mCryptProv.LegalKeySizes(0).SkipSize & vbCrLf
            MsgBox(mStr, MsgBoxStyle.Critical)
            Exit Sub
            'Show the Valid Block Size 
        End Try
        'Create Memory Stream
        mMemStr = New MemoryStream
        'Create the Encryptor, Passing that the Key and IV
        Dim mEncryptor As ICryptoTransform = mCryptProv.CreateEncryptor(Me.GetKey, Me.GetIV)
        'Create the Crypto Stream, Passing that the Memory Stream and Encryptor
        Dim mCryptStr As CryptoStream = New CryptoStream(mMemStr, mEncryptor, CryptoStreamMode.Write)

        'Create Stream Writer to Write Encrypted Data
        Dim mStrWri As New StreamWriter(mCryptStr)
        mStrWri.Write(Me.txtData)      'Encrypt the TextBox Data
        mStrWri.Flush()     'Make Sure everything is written
        mCryptStr.FlushFinalBlock()

        'The Data has been Encrypted in the Memory, Now get that and Display in the TexTBox
        'Create Byte Array of the Length of MemoryStream
        Dim mBytes(mMemStr.Length - 1) As Byte
        mMemStr.Position = 0        'Move Cursor to Start of MemoryStream
        mMemStr.Read(mBytes, 0, mMemStr.Length)
        'Encode the Byte Array to UTF8 and display
        Dim mEnc As New Text.UTF8Encoding
        Me.txtData = mEnc.GetString(mBytes)

        'Me.btnDecrypt.Enabled = True
        'Me.btnEncrypt.Enabled = False
    End Sub
    Public Sub dencrypt()
        'Set the Key and IV and also, Algorithm to Decrypt
        'If Me.cmbAlgo.SelectedIndex < 0 Then
        '    MsgBox("Select an Algo to use to Encrypt or Decrypt.", MsgBoxStyle.Critical)
        '    Exit Sub
        'End If
        'Since, The selected algo may not be Symmetric, So Try it
        Try
            Me.mCryptProv = SymmetricAlgorithm.Create("Rijndael")
        Catch ee As Exception
            MsgBox("Exception : " & ee.Message, MsgBoxStyle.Critical)
            Exit Sub
        End Try

        Try
            'Each Algo has different Size of Key and Block, So let's try
            'mCryptProv.BlockSize = Me.udBlockSize.Value
            'mCryptProv.KeySize = Me.udKeySize.Value
            mCryptProv.BlockSize = 256
            mCryptProv.KeySize = 256
        Catch
            'Hummmmmn, Invalid Block or Key Size. Show the Default Ones
            Dim mStr As String
            mStr = "Invalid Block or Key Size." & vbCrLf & "Valid Sizes are : " & vbCrLf
            mStr += "Block Size " & mCryptProv.LegalBlockSizes(0).MinSize & " - " & mCryptProv.LegalBlockSizes(0).MaxSize & " With Increment of " & mCryptProv.LegalBlockSizes(0).SkipSize & vbCrLf
            mStr += "Key  Size " & mCryptProv.LegalKeySizes(0).MinSize & " - " & mCryptProv.LegalKeySizes(0).MaxSize & " With Increment of " & mCryptProv.LegalKeySizes(0).SkipSize & vbCrLf
            MsgBox(mStr, MsgBoxStyle.Critical)
            Exit Sub
            'Show the Valid Block Size 
        End Try

        '---------------original message
        ''Try to Decrypt values in the memory
        ''First Create the Decryptor, By Passing Key and IV
        'Dim mDecrypt As ICryptoTransform = Me.mCryptProv.CreateDecryptor(Me.GetKey, Me.GetIV)
        ''Create Crypto Stream

        'Dim sdata As Byte()
        'sdata = System.Text.Encoding.UTF8.GetBytes(txtData)
        'Dim txtSource As New System.Windows.Forms.TextBox
        'txtSource.Text = txtData
        'mMemStr = New MemoryStream
        'mMemStr.Write(sdata, 0, sdata.Length)



        ''Move the Memory Pointer to the Beginning of Memory Stream
        'mMemStr.Position = 0

        'Dim mCSReader As New CryptoStream(mMemStr, mDecrypt, CryptoStreamMode.Read)

        ''Create Stream Reader
        'Dim mStrRead As New StreamReader(mCSReader, System.Text.UTF8Encoding.UTF8, False, mMemStr.Length)

        'Try
        '    Me.txtData = mStrRead.ReadToEnd()
        'Catch ee As CryptographicException
        '    MsgBox("Exception : " & ee.Message)
        '    Exit Sub
        'End Try


        'mCSReader.Close()
        'mStrRead.Close()
        'mMemStr.Close()
        'Me.mCryptProv.Clear()
        ''Me.btnDecrypt.Enabled = False
        ''Me.btnEncrypt.Enabled = True
        '--------------------------
        '-----------new code
        Dim roundtrip As String
        Dim textConverter As New ASCIIEncoding
        Dim myRijndael As New RijndaelManaged
        Dim fromEncrypt() As Byte
        Dim encrypted() As Byte
        Dim toEncrypt() As Byte
      
        Dim sdata As Byte()
        sdata = System.Text.Encoding.UTF8.GetBytes(txtData)

        'Get encrypted array of bytes.
        encrypted = sdata

        'This is where the message would be transmitted to a recipient
        ' who already knows your secret key. Optionally, you can
        ' also encrypt your secret key using a public key algorithm
        ' and pass it to the mesage recipient along with the RijnDael
        ' encrypted message.            
        'Get a decryptor that uses the same key and IV as the encryptor.
        Dim decryptor As ICryptoTransform = myRijndael.CreateDecryptor(Me.GetKey, Me.GetIV)

        'Now decrypt the previously encrypted message using the decryptor
        ' obtained in the above step.
        Dim msDecrypt As New MemoryStream(encrypted)
        Dim csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)

        fromEncrypt = New Byte(encrypted.Length) {}

        'Read the data out of the crypto stream.
        csDecrypt.Read(fromEncrypt, 0, fromEncrypt.Length)

        'Convert the byte array back into a string.
        roundtrip = textConverter.GetString(fromEncrypt)

    End Sub

#Region " IV and Key properties"
    Private ReadOnly Property GetKey() As Byte()
        Get
            'Dim thisSize As Integer = (Me.udKeySize.Value / 8) - 1
            Dim thisSize As Integer = (256 / 8) - 1
            Dim temp As Integer
            Dim thisKey(thisSize) As Byte
            If Me.txtkey.Length < 1 Then
                Return thisKey
            End If
            Dim lastBound As Integer = Me.txtkey.Length
            If lastBound > thisSize Then lastBound = thisSize
            For temp = 0 To lastBound - 1
                thisKey(temp) = Convert.ToByte(txtkey.Chars(temp))
            Next
            Return thisKey
        End Get
    End Property
    Private ReadOnly Property GetIV() As Byte()
        Get
            'Convert Bits to Bytes
            'Dim thisSize As Integer = (Me.udBlockSize.Value / 8) - 1
            Dim thisSize As Integer = (256 / 8) - 1
            Dim thisIV(thisSize) As Byte
            If Me.txtIV.Length < 1 Then
                Return thisIV
            End If

            Dim temp As Integer
            Dim lastBound As Integer = Me.txtIV.Length
            If lastBound > thisSize Then lastBound = thisSize
            For temp = 0 To lastBound - 1
                thisIV(temp) = Convert.ToByte(txtIV.Chars(temp))
            Next
            Return thisIV
        End Get
    End Property
#End Region

End Class
