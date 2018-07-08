Module MActivation
    'https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/String_Hashing_in_VBA

    Sub TestHash()
        'run this to test md5, sha1, sha2/256, sha2/384, or sha2/512
        Dim sIn As String, bB64 As Boolean, sH As String
        Dim App As String, Ver As String, ProdID As String, AppSaltNo

        'insert the text to hash within the sIn quotes
        'note that a private string could be joined to sIn at this point
        'sIn = "abcdefghujklmnopqrstuvwxyqrstuvwxy"
        App = "BillMacros"
        Ver = "20180511"
        Ver = Left(Ver, 4)
        ProdID = OSInfo
        AppSaltNo = "A53B3" 'Secret string
        sIn = ProdID & App & Ver & AppSaltNo

        'select as required
        bB64 = False   'output hex
        'bB64 = True   'output base-64

        'enable any one
        sH = MD5(sIn, bB64)
        'sH = SHA1(sIn, bB64)
        'sH = SHA256(sIn, bB64)
        'sH = SHA384(sIn, bB64)
        'sH = SHA512(sIn, bB64)

        MsgBox(sIn & vbNewLine & sH & vbNewLine & Len(sH) & " characters in length")

    End Sub

    Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = False) As String
        'Set a reference to mscorlib 4.0 64-bit

        'Test with empty string input:
        'Hex:   d41d8cd98f00...etc
        'Base-64: 1B2M2Y8Asg...etc

        Dim oT As Object, oMD5 As Object
        Dim TextToHash() As Byte
        Dim bytes() As Byte

        oT = CreateObject("System.Text.UTF8Encoding")
        oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

        TextToHash = oT.GetBytes_4(sIn)
        bytes = oMD5.ComputeHash_2((TextToHash))

        If bB64 = True Then
            MD5 = ConvToBase64String(bytes)
        Else
            MD5 = ConvToHexString(bytes)
        End If

        oT = Nothing
        oMD5 = Nothing

    End Function

    Public Function SHA1(sIn As String, Optional bB64 As Boolean = False) As String
        'Set a reference to mscorlib 4.0 64-bit

        'Test with empty string input:
        '40 Hex:   da39a3ee5e6...etc
        '28 Base-64:   2jmj7l5rSw0yVb...etc

        Dim oT As Object, oSHA1 As Object
        Dim TextToHash() As Byte
        Dim bytes() As Byte

        oT = CreateObject("System.Text.UTF8Encoding")
        oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")

        TextToHash = oT.GetBytes_4(sIn)
        bytes = oSHA1.ComputeHash_2((TextToHash))

        If bB64 = True Then
            SHA1 = ConvToBase64String(bytes)
        Else
            SHA1 = ConvToHexString(bytes)
        End If

        oT = Nothing
        oSHA1 = Nothing

    End Function

    Public Function SHA256(sIn As String, Optional bB64 As Boolean = False) As String
        'Set a reference to mscorlib 4.0 64-bit

        'Test with empty string input:
        '64 Hex:   e3b0c44298f...etc
        '44 Base-64:   47DEQpj8HBSa+/...etc

        Dim oT As Object, oSHA256 As Object
        Dim TextToHash() As Byte, bytes() As Byte

        oT = CreateObject("System.Text.UTF8Encoding")
        oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")

        TextToHash = oT.GetBytes_4(sIn)
        bytes = oSHA256.ComputeHash_2((TextToHash))

        If bB64 = True Then
            SHA256 = ConvToBase64String(bytes)
        Else
            SHA256 = ConvToHexString(bytes)
        End If

        oT = Nothing
        oSHA256 = Nothing

    End Function

    Public Function SHA384(sIn As String, Optional bB64 As Boolean = False) As String
        'Set a reference to mscorlib 4.0 64-bit

        'Test with empty string input:
        '96 Hex:   38b060a751ac...etc
        '64 Base-64:   OLBgp1GsljhM2T...etc

        Dim oT As Object, oSHA384 As Object
        Dim TextToHash() As Byte, bytes() As Byte

        oT = CreateObject("System.Text.UTF8Encoding")
        oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")

        TextToHash = oT.GetBytes_4(sIn)
        bytes = oSHA384.ComputeHash_2((TextToHash))

        If bB64 = True Then
            SHA384 = ConvToBase64String(bytes)
        Else
            SHA384 = ConvToHexString(bytes)
        End If

        oT = Nothing
        oSHA384 = Nothing

    End Function

    Public Function SHA512(sIn As String, Optional bB64 As Boolean = False) As String
        'Set a reference to mscorlib 4.0 64-bit

        'Test with empty string input:
        '128 Hex:   cf83e1357eefb8bd...etc
        '88 Base-64:   z4PhNX7vuL3xVChQ...etc

        Dim oT As Object, oSHA512 As Object
        Dim TextToHash() As Byte, bytes() As Byte

        oT = CreateObject("System.Text.UTF8Encoding")
        oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")

        TextToHash = oT.GetBytes_4(sIn)
        bytes = oSHA512.ComputeHash_2((TextToHash))

        If bB64 = True Then
            SHA512 = ConvToBase64String(bytes)
        Else
            SHA512 = ConvToHexString(bytes)
        End If

        oT = Nothing
        oSHA512 = Nothing

    End Function

    Private Function ConvToBase64String(vIn As Object) As Object

        Dim oD As Object

        oD = CreateObject("MSXML2.DOMDocument")
        With oD
            .LoadXML("<root />")
            .DocumentElement.DataType = "bin.base64"
            .DocumentElement.nodeTypedValue = vIn
        End With
        ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")

        oD = Nothing

    End Function

    Private Function ConvToHexString(vIn As Object) As Object

        Dim oD As Object

        oD = CreateObject("MSXML2.DOMDocument")

        With oD
            .LoadXML("<root />")
            .DocumentElement.DataType = "bin.Hex"
            .DocumentElement.nodeTypedValue = vIn
        End With
        ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")

        oD = Nothing

    End Function

End Module
