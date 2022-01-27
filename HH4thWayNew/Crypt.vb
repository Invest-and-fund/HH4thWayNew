Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text

Public Class Crypt
    Private Shared ReadOnly key As Byte() = New Byte(23) {1, 2, 3, 4, 5, 6,
  7, 8, 9, 0, 1, 2,
  3, 4, 5, 6, 7, 8,
  9, 0, 1, 2, 3, 4}

    Public Shared Function EncryptQuery(input As String) As String
        Dim s As String
        Dim inputArray As Byte() = UTF8Encoding.UTF8.GetBytes(input)
        Dim tripleDES As New TripleDESCryptoServiceProvider()
        tripleDES.GenerateKey()
        tripleDES.Key = key
        tripleDES.Mode = CipherMode.ECB
        tripleDES.Padding = PaddingMode.PKCS7
        Dim cTransform As ICryptoTransform = tripleDES.CreateEncryptor()
        Dim resultArray As Byte() = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length)
        tripleDES.Clear()
        s = Convert.ToBase64String(resultArray, 0, resultArray.Length)
        s = s.Replace(" ", "+")
        Return s
    End Function

    Public Shared Function DecryptQuery(input As String) As String
        If Not IsNothing(input) Then
            If input.Length > 0 Then
                Dim inputArray As Byte() = Convert.FromBase64String(input.Replace(" ", "+"))
                Dim tripleDES As New TripleDESCryptoServiceProvider()
                tripleDES.Key = key
                tripleDES.Mode = CipherMode.ECB
                tripleDES.Padding = PaddingMode.PKCS7
                Dim cTransform As ICryptoTransform = tripleDES.CreateDecryptor()
                Dim resultArray As Byte() = cTransform.TransformFinalBlock(inputArray, 0, inputArray.Length)
                tripleDES.Clear()
                Dim retStr As String = UTF8Encoding.UTF8.GetString(resultArray)
                Return retStr
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "HK2GB92K6IU08AS"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
             &H65, &H64, &H76, &H65, &H64, &H65,
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function

    Public Shared Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "HK2GB92K6IU08AS"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
             &H65, &H64, &H76, &H65, &H64, &H65,
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function


End Class

