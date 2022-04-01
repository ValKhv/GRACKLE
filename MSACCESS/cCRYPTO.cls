VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCRYPTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
' This class implements general cryptography methods
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit


Private Const DEFAULT_BUFFER_SIZE As Long = 32768
Private CRC_32(0 To 255) As Long

' DES ALGORITHM
Private Const INITIALIZATION_VECTOR = "zlrs$5kd"  'Always 8 characters
Private Const TRIPLE_DES_KEY = ">tlF8adk=35K{dsa" 'Always 16 characters
'****************************************************************************************************************************************************
' AES ALGORITHM
Private Const clOneMask = 16515072          '000000 111111 111111 111111
Private Const clTwoMask = 258048            '111111 000000 111111 111111
Private Const clThreeMask = 4032            '111111 111111 000000 111111
Private Const clFourMask = 63               '111111 111111 111111 000000

Private Const clHighMask = 16711680         '11111111 00000000 00000000
Private Const clMidMask = 65280             '00000000 11111111 00000000
Private Const clLowMask = 255               '00000000 00000000 11111111

Private Const cl2Exp18 = 262144             '2 to the 18th power
Private Const cl2Exp12 = 4096               '2 to the 12th
Private Const cl2Exp6 = 64                  '2 to the 6th
Private Const cl2Exp8 = 256                 '2 to the 8th
Private Const cl2Exp16 = 65536              '2 to the 16th
'****************************************************************************************************************************************************



Public Enum HashOutputFormat
    OUTPUT_HEX = 0
    OUTPUT_BASE64 = 1
End Enum

'=====================================================================================================================================================
' Calc32ulate CRC32 sum for the string
'=====================================================================================================================================================
Public Function CRC32_String(Stri As String) As String
  CRC32_String = CStr(Hex$(Calc32(Stri)))
End Function
'=====================================================================================================================================================
' Calc32ulate CRC32 sum for the file
'=====================================================================================================================================================
Public Function CRC32_File(file As String, Optional BufferSize As Long = DEFAULT_BUFFER_SIZE) As String
Dim FileSize As Long, FileNumber As Integer, buffer As String, code As String, modCrc As Long

On Error GoTo ErrHandle
'---------------------------------------------------
  If Len(Dir$(file)) > 0 And BufferSize > 0 Then
    FileSize = FileLen(file)
    FileNumber = FreeFile()
    Open file For Binary Access Read As #FileNumber
    buffer = String(BufferSize, 0)
    While (FileSize - Seek(FileNumber) + 1) >= BufferSize
        Get #FileNumber, , buffer
        code = code & buffer
    Wend
    buffer = String(FileSize - Seek(FileNumber) + 1, 0)
    Get #FileNumber, , buffer
    code = code & buffer
    Close #FileNumber
    modCrc = Calc32(code)
    CRC32_File = CStr(Hex$(modCrc))
  Else
    CRC32_File = ""
  End If
'---------------------------------------------------
ExitHere:
  Exit Function
'-----------------
ErrHandle:
  ErrPrint "CRC32_File", Err.Number, Err.Description
  Err.Clear
End Function

'=====================================================================================================================================================
' Calculate MD5 sum for the string
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function MD5_String(ByVal sIn As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim oT As Object, oMD5 As Object
Dim TextToHash() As Byte
Dim bytes() As Byte
        
On Error GoTo ErrHandle
'------------------------
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oMD5.ComputeHash_2((TextToHash))
 
    If bB64 = OUTPUT_BASE64 Then
       MD5_String = Conv_To_Base64String(bytes)
    Else
       MD5_String = Conv_To_HexString(bytes)
    End If
'------------------------
ExitHere:
    Set oT = Nothing
    Set oMD5 = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "MD5_String", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================================================
' Calculate SHA1 sum for the string
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA1_String(sIn As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim oT As Object, oSHA1 As Object
Dim TextToHash() As Byte
Dim bytes() As Byte
    
    On Error GoTo ErrHandle
'--------------------------------------
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA1.ComputeHash_2((TextToHash))
        
    If bB64 = OUTPUT_BASE64 Then
       SHA1_String = Conv_To_Base64String(bytes)
    Else
       SHA1_String = Conv_To_HexString(bytes)
    End If
'--------------------------------------
ExitHere:
    Set oT = Nothing
    Set oSHA1 = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA1_String", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Calculate SHA256 sum for the string
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA256_String(sIn As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim oT As Object, oSHA256 As Object
Dim TextToHash() As Byte, bytes() As Byte
    
    On Error GoTo ErrHandle
'--------------------------------------
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA256.ComputeHash_2((TextToHash))
    
    If bB64 = OUTPUT_BASE64 Then
       SHA256_String = Conv_To_Base64String(bytes)
    Else
       SHA256_String = Conv_To_HexString(bytes)
    End If
'--------------------------------------
ExitHere:
    Set oT = Nothing
    Set oSHA256 = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA256_String", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Calculate SHA384 sum for the string
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA384_String(sIn As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim oT As Object, oSHA384 As Object
Dim TextToHash() As Byte, bytes() As Byte
    
    On Error GoTo ErrHandle
'--------------------------------------
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA384.ComputeHash_2((TextToHash))
    
    If bB64 = OUTPUT_BASE64 Then
       SHA384_String = Conv_To_Base64String(bytes)
    Else
       SHA384_String = Conv_To_HexString(bytes)
    End If
'--------------------------------------
ExitHere:
    Set oT = Nothing
    Set oSHA384 = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA384_String", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Calculate SHA384 sum for the string
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA512_String(sIn As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim oT As Object, oSHA512 As Object
Dim TextToHash() As Byte, bytes() As Byte

   On Error GoTo ErrHandle
'--------------------------------------
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA512.ComputeHash_2((TextToHash))
    
    If bB64 = OUTPUT_BASE64 Then
       SHA512_String = Conv_To_Base64String(bytes)
    Else
       SHA512_String = Conv_To_HexString(bytes)
    End If
'--------------------------------------
ExitHere:
    Set oT = Nothing
    Set oSHA512 = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA512_String", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Calculate MD5 sum for the File
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function MD5_File(sFullPath As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim enc, bytes, outstr As String, POS As Integer
   
   On Error GoTo ErrHandle
'--------------------------------------
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath)
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = OUTPUT_BASE64 Then
       MD5_File = Conv_To_Base64String(bytes)
    Else
       MD5_File = Conv_To_HexString(bytes)
    End If
'--------------------------------------
ExitHere:
    Set enc = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "MD5_File", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Calculate SHA1 sum for the File
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA1_File(sFullPath As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim enc, bytes, outstr As String, POS As Integer
    
   On Error GoTo ErrHandle
'--------------------------------------
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    bytes = GetFileBytes(sFullPath) ' Convert the string to a byte array and hash it, returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = OUTPUT_BASE64 Then
       SHA1_File = Conv_To_Base64String(bytes)
    Else
       SHA1_File = Conv_To_HexString(bytes)
    End If
'--------------------------------------
ExitHere:
    Set enc = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA1_File", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Calculate SHA256_File sum for the File
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA256_File(sFullPath As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim enc, bytes, outstr As String, POS As Integer
 
   On Error GoTo ErrHandle
'--------------------------------------
    Set enc = CreateObject("System.Security.Cryptography.SHA256Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = OUTPUT_BASE64 Then
       SHA256_File = Conv_To_Base64String(bytes)
    Else
       SHA256_File = Conv_To_HexString(bytes)
    End If
        
'--------------------------------------
ExitHere:
    Set enc = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA256_File", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Calculate SHA384_File sum for the File
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA384_File(sFullPath As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim enc, bytes, outstr As String, POS As Integer
    
    On Error GoTo ErrHandle
'--------------------------------------
    Set enc = CreateObject("System.Security.Cryptography.SHA384Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = OUTPUT_BASE64 Then
       SHA384_File = Conv_To_Base64String(bytes)
    Else
       SHA384_File = Conv_To_HexString(bytes)
    End If
'--------------------------------------
ExitHere:
    Set enc = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA384_File", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================================================
' Calculate SHA512_File sum for the File
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function SHA512_File(sFullPath As String, Optional bB64 As HashOutputFormat = OUTPUT_HEX) As String
Dim enc, bytes, outstr As String, POS As Integer
    
    On Error GoTo ErrHandle
'--------------------------------------
    Set enc = CreateObject("System.Security.Cryptography.SHA512Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = OUTPUT_BASE64 Then
       SHA512_File = Conv_To_Base64String(bytes)
    Else
       SHA512_File = Conv_To_HexString(bytes)
    End If
    
'--------------------------------------
ExitHere:
    Set enc = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "SHA512_File", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Encrypt string with 3DES Alghorithm
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function EncryptStringTripleDES(plain_string As String, Optional sPass As String) As Variant
Dim encryption_object As Object, TRDESKey As String, TRDESINITVector As String
Dim plain_byte_data() As Byte, encrypted_byte_data() As Byte, encrypted_base64_string As String
    
    
On Error GoTo ErrHandle
'-----------------------------------------------------
Call Create_DESKEY(sPass, TRDESKey, TRDESINITVector)
    EncryptStringTripleDES = Null
    plain_byte_data = CreateObject("System.Text.UTF8Encoding").GetBytes_4(plain_string)
    
    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.key = CreateObject("System.Text.UTF8Encoding").GetBytes_4(TRDESKey)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(TRDESINITVector)
    encrypted_byte_data = _
            encryption_object.CreateEncryptor().TransformFinalBlock(plain_byte_data, 0, UBound(plain_byte_data) + 1)
    'encrypted_base64_string = BytesToBase64(encrypted_byte_data)
    encrypted_base64_string = Conv_To_Base64String(encrypted_byte_data)
'------------------------------------------------------
ExitHere:
    EncryptStringTripleDES = encrypted_base64_string
    Set encryption_object = Nothing
    Exit Function
'-------------------------------
ErrHandle:
    ErrPrint "EncryptStringTripleDES", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Decrypt string with 3DES Alghorithm
' Working with mscorlib 4.0 64-bit
'=====================================================================================================================================================
Public Function DecryptStringTripleDES(encrypted_string As String, Optional sPass As String) As Variant
Dim encryption_object As Object, TRDESKey As String, TRDESINITVector As String
Dim encrypted_byte_data() As Byte, plain_byte_data() As Byte, plain_string As String


On Error GoTo ErrHandle
'-----------------------------------
Call Create_DESKEY(sPass, TRDESKey, TRDESINITVector)
    DecryptStringTripleDES = Null
    encrypted_byte_data = ConvBase64_To_Bytes(encrypted_string)
    
    Set encryption_object = CreateObject("System.Security.Cryptography.TripleDESCryptoServiceProvider")
    encryption_object.key = CreateObject("System.Text.UTF8Encoding").GetBytes_4(TRDESKey)
    encryption_object.IV = CreateObject("System.Text.UTF8Encoding").GetBytes_4(TRDESINITVector)
    plain_byte_data = encryption_object.CreateDecryptor().TransformFinalBlock(encrypted_byte_data, 0, UBound(encrypted_byte_data) + 1)
            
    plain_string = CreateObject("System.Text.UTF8Encoding").GetString(plain_byte_data)
    
'-----------------------------------
ExitHere:
    DecryptStringTripleDES = plain_string '!!!!!!!!!!!!!!!!!!!
    Set encryption_object = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint "DecryptStringTripleDES", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function


'=====================================================================================================================================================
' Get Encrypt String with AES
' Dependency on "functions Base64" and "functions SHA1"
'=====================================================================================================================================================
Public Function GetEncryptAES(sMessage As String, sPassword As String, Optional iRounds As Integer = 1000) As String
    GetEncryptAES = Encode64(EncryptAES(SHA1_String(sMessage) & ":" & sMessage, sPassword, iRounds))
End Function

'=====================================================================================================================================================
' Get Decrypt String with AES
' Dependency on "functions Base64" and "functions SHA1"
'=====================================================================================================================================================
Public Function GetDecryptAES(sMessage As String, sPassword As String, Optional iRounds As Integer = 1000, Optional asString As Boolean = False) As String
Dim dMessage As String, HASH As String, SEPARATOR As Integer
    
    dMessage = DecryptAES(Decode64(sMessage), sPassword, iRounds)
    If InStr(dMessage, ":") < 1 Then
        If asString = True Then
            GetDecryptAES = "#invalidpassword"
        Else
            Err.Raise 640, "functions AES", "Invalid password/passphrase"
        End If
    End If
    HASH = Left(dMessage, InStr(dMessage, ":") - 1)
    
    GetDecryptAES = Right(dMessage, Len(dMessage) - (Len(HASH) + 1))
    If HASH <> SHA1_String(GetDecryptAES) Then
        If asString = True Then
            GetDecryptAES = "#invalidpassword"
        Else
            Err.Raise 640, "functions AES", "Invalid password/passphrase"
        End If
    End If
    
End Function

Public Function EncryptAES(sMessage As String, sPassword As String, Optional iRounds As Integer = 1) As String
    EncryptAES = sMessage
    Do Until iRounds < 1
        EncryptAES = AES(EncryptAES, True, sPassword)
        iRounds = iRounds - 1
    Loop
End Function

Public Function DecryptAES(sMessage As String, sPassword As String, Optional iRounds As Integer = 1) As String
    DecryptAES = sMessage
    Do Until iRounds < 1
        DecryptAES = AES(DecryptAES, False, sPassword)
        iRounds = iRounds - 1
    Loop
    DecryptAES = nTrim2(DecryptAES)
End Function




'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Error Handler
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True, _
                                                                                                  Optional sModName As String = "cCRYPTO") As String
Dim sRes As String
Const ERR_CHAR As String = "#"
Const ERR_REPEAT As Integer = 60

sRes = String(ERR_REPEAT, ERR_CHAR) & vbCrLf & "ERROR OF [" & sModName & ": " & FuncName & "]" & vbTab & "ERR#" & ErrNumber & vbTab & Now() & _
       vbCrLf & ErrDescription & vbCrLf & String(ERR_REPEAT, ERR_CHAR)
If bDebug Then Debug.Print sRes
'----------------------------------------------------------
ExitHere:
       Beep
       ErrPrint = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' CALCULATORS
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' CRC32 Main Calc32ulator
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Calc32(Stri As String) As Long
Dim tCRC32 As Long, I As Long, m As Integer, n As Integer
  Call CRC_Initialization
  tCRC32 = &HFFFFFFFF
  For I = 1 To Len(Stri)
    m = Asc(Mid$(Stri, I, 1))
    n = (tCRC32 Xor m) And &HFF
    tCRC32 = CRC_32(n) Xor (sHr(tCRC32, 8) And &HFFFFFF)
  Next I&
  Calc32 = -(tCRC32 + 1)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Calc32ulate CRC_32
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CRC_Initialization()
  CRC_32(0) = &H0: CRC_32(1) = &H77073096: CRC_32(2) = &HEE0E612C: CRC_32(3) = &H990951BA
  CRC_32(4) = &H76DC419: CRC_32(5) = &H706AF48F: CRC_32(6) = &HE963A535: CRC_32(7) = &H9E6495A3
  CRC_32(8) = &HEDB8832: CRC_32(9) = &H79DCB8A4: CRC_32(10) = &HE0D5E91E: CRC_32(11) = &H97D2D988
  CRC_32(12) = &H9B64C2B: CRC_32(13) = &H7EB17CBD: CRC_32(14) = &HE7B82D07: CRC_32(15) = &H90BF1D91
  CRC_32(16) = &H1DB71064: CRC_32(17) = &H6AB020F2: CRC_32(18) = &HF3B97148: CRC_32(19) = &H84BE41DE
  CRC_32(20) = &H1ADAD47D: CRC_32(21) = &H6DDDE4EB: CRC_32(22) = &HF4D4B551: CRC_32(23) = &H83D385C7
  CRC_32(24) = &H136C9856: CRC_32(25) = &H646BA8C0: CRC_32(26) = &HFD62F97A: CRC_32(27) = &H8A65C9EC
  CRC_32(28) = &H14015C4F: CRC_32(29) = &H63066CD9: CRC_32(30) = &HFA0F3D63: CRC_32(31) = &H8D080DF5
  CRC_32(32) = &H3B6E20C8: CRC_32(33) = &H4C69105E: CRC_32(34) = &HD56041E4: CRC_32(35) = &HA2677172
  CRC_32(36) = &H3C03E4D1: CRC_32(37) = &H4B04D447: CRC_32(38) = &HD20D85FD: CRC_32(39) = &HA50AB56B
  CRC_32(40) = &H35B5A8FA: CRC_32(41) = &H42B2986C: CRC_32(42) = &HDBBBC9D6: CRC_32(43) = &HACBCF940
  CRC_32(44) = &H32D86CE3: CRC_32(45) = &H45DF5C75: CRC_32(46) = &HDCD60DCF: CRC_32(47) = &HABD13D59
  CRC_32(48) = &H26D930AC: CRC_32(49) = &H51DE003A: CRC_32(50) = &HC8D75180: CRC_32(51) = &HBFD06116
  CRC_32(52) = &H21B4F4B5: CRC_32(53) = &H56B3C423: CRC_32(54) = &HCFBA9599: CRC_32(55) = &HB8BDA50F
  CRC_32(56) = &H2802B89E: CRC_32(57) = &H5F058808: CRC_32(58) = &HC60CD9B2: CRC_32(59) = &HB10BE924
  CRC_32(60) = &H2F6F7C87: CRC_32(61) = &H58684C11: CRC_32(62) = &HC1611DAB: CRC_32(63) = &HB6662D3D
  CRC_32(64) = &H76DC4190: CRC_32(65) = &H1DB7106: CRC_32(66) = &H98D220BC: CRC_32(67) = &HEFD5102A
  CRC_32(68) = &H71B18589: CRC_32(69) = &H6B6B51F: CRC_32(70) = &H9FBFE4A5: CRC_32(71) = &HE8B8D433
  CRC_32(72) = &H7807C9A2: CRC_32(73) = &HF00F934: CRC_32(74) = &H9609A88E: CRC_32(75) = &HE10E9818
  CRC_32(76) = &H7F6A0DBB: CRC_32(77) = &H86D3D2D: CRC_32(78) = &H91646C97: CRC_32(79) = &HE6635C01
  CRC_32(80) = &H6B6B51F4: CRC_32(81) = &H1C6C6162: CRC_32(82) = &H856530D8: CRC_32(83) = &HF262004E
  CRC_32(84) = &H6C0695ED: CRC_32(85) = &H1B01A57B: CRC_32(86) = &H8208F4C1: CRC_32(87) = &HF50FC457
  CRC_32(88) = &H65B0D9C6: CRC_32(89) = &H12B7E950: CRC_32(90) = &H8BBEB8EA: CRC_32(91) = &HFCB9887C
  CRC_32(92) = &H62DD1DDF: CRC_32(93) = &H15DA2D49: CRC_32(94) = &H8CD37CF3: CRC_32(95) = &HFBD44C65
  CRC_32(96) = &H4DB26158: CRC_32(97) = &H3AB551CE: CRC_32(98) = &HA3BC0074: CRC_32(99) = &HD4BB30E2
  CRC_32(100) = &H4ADFA541: CRC_32(101) = &H3DD895D7: CRC_32(102) = &HA4D1C46D: CRC_32(103) = &HD3D6F4FB
  CRC_32(104) = &H4369E96A: CRC_32(105) = &H346ED9FC: CRC_32(106) = &HAD678846: CRC_32(107) = &HDA60B8D0
  CRC_32(108) = &H44042D73: CRC_32(109) = &H33031DE5: CRC_32(110) = &HAA0A4C5F: CRC_32(111) = &HDD0D7CC9
  CRC_32(112) = &H5005713C: CRC_32(113) = &H270241AA: CRC_32(114) = &HBE0B1010: CRC_32(115) = &HC90C2086
  CRC_32(116) = &H5768B525: CRC_32(117) = &H206F85B3: CRC_32(118) = &HB966D409: CRC_32(119) = &HCE61E49F
  CRC_32(120) = &H5EDEF90E: CRC_32(121) = &H29D9C998: CRC_32(122) = &HB0D09822: CRC_32(123) = &HC7D7A8B4
  CRC_32(124) = &H59B33D17: CRC_32(125) = &H2EB40D81: CRC_32(126) = &HB7BD5C3B: CRC_32(127) = &HC0BA6CAD
  CRC_32(128) = &HEDB88320: CRC_32(129) = &H9ABFB3B6: CRC_32(130) = &H3B6E20C: CRC_32(131) = &H74B1D29A
  CRC_32(132) = &HEAD54739: CRC_32(133) = &H9DD277AF: CRC_32(134) = &H4DB2615: CRC_32(135) = &H73DC1683
  CRC_32(136) = &HE3630B12: CRC_32(137) = &H94643B84: CRC_32(138) = &HD6D6A3E: CRC_32(139) = &H7A6A5AA8
  CRC_32(140) = &HE40ECF0B: CRC_32(141) = &H9309FF9D: CRC_32(142) = &HA00AE27: CRC_32(143) = &H7D079EB1
  CRC_32(144) = &HF00F9344: CRC_32(145) = &H8708A3D2: CRC_32(146) = &H1E01F268: CRC_32(147) = &H6906C2FE
  CRC_32(148) = &HF762575D: CRC_32(149) = &H806567CB: CRC_32(150) = &H196C3671: CRC_32(151) = &H6E6B06E7
  CRC_32(152) = &HFED41B76: CRC_32(153) = &H89D32BE0: CRC_32(154) = &H10DA7A5A: CRC_32(155) = &H67DD4ACC
  CRC_32(156) = &HF9B9DF6F: CRC_32(157) = &H8EBEEFF9: CRC_32(158) = &H17B7BE43: CRC_32(159) = &H60B08ED5
  CRC_32(160) = &HD6D6A3E8: CRC_32(161) = &HA1D1937E: CRC_32(162) = &H38D8C2C4: CRC_32(163) = &H4FDFF252
  CRC_32(164) = &HD1BB67F1: CRC_32(165) = &HA6BC5767: CRC_32(166) = &H3FB506DD: CRC_32(167) = &H48B2364B
  CRC_32(168) = &HD80D2BDA: CRC_32(169) = &HAF0A1B4C: CRC_32(170) = &H36034AF6: CRC_32(171) = &H41047A60
  CRC_32(172) = &HDF60EFC3: CRC_32(173) = &HA867DF55: CRC_32(174) = &H316E8EEF: CRC_32(175) = &H4669BE79
  CRC_32(176) = &HCB61B38C: CRC_32(177) = &HBC66831A: CRC_32(178) = &H256FD2A0: CRC_32(179) = &H5268E236
  CRC_32(180) = &HCC0C7795: CRC_32(181) = &HBB0B4703: CRC_32(182) = &H220216B9: CRC_32(183) = &H5505262F
  CRC_32(184) = &HC5BA3BBE: CRC_32(185) = &HB2BD0B28: CRC_32(186) = &H2BB45A92: CRC_32(187) = &H5CB36A04
  CRC_32(188) = &HC2D7FFA7: CRC_32(189) = &HB5D0CF31: CRC_32(190) = &H2CD99E8B: CRC_32(191) = &H5BDEAE1D
  CRC_32(192) = &H9B64C2B0: CRC_32(193) = &HEC63F226: CRC_32(194) = &H756AA39C: CRC_32(195) = &H26D930A
  CRC_32(196) = &H9C0906A9: CRC_32(197) = &HEB0E363F: CRC_32(198) = &H72076785: CRC_32(199) = &H5005713
  CRC_32(200) = &H95BF4A82: CRC_32(201) = &HE2B87A14: CRC_32(202) = &H7BB12BAE: CRC_32(203) = &HCB61B38
  CRC_32(204) = &H92D28E9B: CRC_32(205) = &HE5D5BE0D: CRC_32(206) = &H7CDCEFB7: CRC_32(207) = &HBDBDF21
  CRC_32(208) = &H86D3D2D4: CRC_32(209) = &HF1D4E242: CRC_32(210) = &H68DDB3F8: CRC_32(211) = &H1FDA836E
  CRC_32(212) = &H81BE16CD: CRC_32(213) = &HF6B9265B: CRC_32(214) = &H6FB077E1: CRC_32(215) = &H18B74777
  CRC_32(216) = &H88085AE6: CRC_32(217) = &HFF0F6A70: CRC_32(218) = &H66063BCA: CRC_32(219) = &H11010B5C
  CRC_32(220) = &H8F659EFF: CRC_32(221) = &HF862AE69: CRC_32(222) = &H616BFFD3: CRC_32(223) = &H166CCF45
  CRC_32(224) = &HA00AE278: CRC_32(225) = &HD70DD2EE: CRC_32(226) = &H4E048354: CRC_32(227) = &H3903B3C2
  CRC_32(228) = &HA7672661: CRC_32(229) = &HD06016F7: CRC_32(230) = &H4969474D: CRC_32(231) = &H3E6E77DB
  CRC_32(232) = &HAED16A4A: CRC_32(233) = &HD9D65ADC: CRC_32(234) = &H40DF0B66: CRC_32(235) = &H37D83BF0
  CRC_32(236) = &HA9BCAE53: CRC_32(237) = &HDEBB9EC5: CRC_32(238) = &H47B2CF7F: CRC_32(239) = &H30B5FFE9
  CRC_32(240) = &HBDBDF21C: CRC_32(241) = &HCABAC28A: CRC_32(242) = &H53B39330: CRC_32(243) = &H24B4A3A6
  CRC_32(244) = &HBAD03605: CRC_32(245) = &HCDD70693: CRC_32(246) = &H54DE5729: CRC_32(247) = &H23D967BF
  CRC_32(248) = &HB3667A2E: CRC_32(249) = &HC4614AB8: CRC_32(250) = &H5D681B02: CRC_32(251) = &H2A6F2B94
  CRC_32(252) = &HB40BBE37: CRC_32(253) = &HC30C8EA1: CRC_32(254) = &H5A05DF1B: CRC_32(255) = &H2D02EF8D
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Using as shrink m,n for CRC32 Calc32ulation
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function sHr(n As Long, m As Integer) As Long
Dim q As Variant
  If (m > 31) Then
    sHr = 0
    Exit Function
  End If
  If (n >= 0) Then
    sHr = n \ (2 ^ m)
  Else
    q = n And &H7FFFFFFF
    q = q \ (2 ^ m)
    sHr = q Or (2 ^ (31 - m))
  End If
End Function

Private Function GetFileSize(sFilePath As String, nSize As Long) As Boolean
Dim fs As Object, f As Object
    
    On Error Resume Next
'-----------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FileExists(sFilePath) Then
        Set f = fs.GetFile(sFilePath)
        nSize = f.SIZE
        GetFileSize = True
        Exit Function
    End If
    
    Set f = Nothing
    Set fs = Nothing
    
End Function

Private Function GetFileBytes(ByVal sPath As String) As Byte()
    
    Dim lngFileNum As Long, bytRtnVal() As Byte, bTest
    
    lngFileNum = FreeFile
    
    If LenB(Dir(sPath)) Then ''// Does file exist?
        
        Open sPath For Binary Access Read As lngFileNum
        
        'a zero length file content will give error 9 here
        
        ReDim bytRtnVal(0 To LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53 'File not found
    End If
    
    GetFileBytes = bytRtnVal
    
    Erase bytRtnVal

End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------
' Generate the triple key
'---------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Create_DESKEY(ByVal sPass As String, ByRef The3DES_KEY As String, ByRef The3DES_InitVector As String)
Dim sCode As String, iL As Integer

Const minLen As Integer = 24

On Error GoTo ErrHandle
'------------------------------------
If sPass = "" Then
    The3DES_KEY = TRIPLE_DES_KEY
    The3DES_InitVector = INITIALIZATION_VECTOR
    Exit Sub
End If
'-----------------------------------
sCode = MD5_String(sPass, OUTPUT_HEX)
iL = Len(sCode): If iL < minLen Then sCode = sCode & String(minLen - iL, "a")
      
    The3DES_InitVector = Left(sCode, 8)
    The3DES_KEY = Right(sCode, 16)
'------------------------------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint "Create_DESKEY", Err.Number, Err.Description
    Err.Clear
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------------
' The Main Calculation Function for AES Algorithm
' Based on http://bytes.com/topic/access/answers/940196-help-aes-encryption-algorithm-vba-vbscript
'---------------------------------------------------------------------------------------------------------------------------------------------------
Private Function AES(sMessage As String, isEncode As Boolean, sPassword As String) As String
Dim sbox(), sboxinv(), rcon(), g2(), g3(), g9(), g11(), g13(), g14()
Dim expandedKey, block(16), aesKey(32), I, isDone, J
Dim sPlain, sPass, sCipher, sTemp, nonce(16), priorCipher(16)
Dim x, R, y, temp(4), intTemp


On Error GoTo ErrHandle
'---------------------------------------------------------
    g2 = Array( _
        &H0, &H2, &H4, &H6, &H8, &HA, &HC, &HE, &H10, &H12, &H14, &H16, &H18, &H1A, &H1C, &H1E, _
        &H20, &H22, &H24, &H26, &H28, &H2A, &H2C, &H2E, &H30, &H32, &H34, &H36, &H38, &H3A, &H3C, &H3E, _
        &H40, &H42, &H44, &H46, &H48, &H4A, &H4C, &H4E, &H50, &H52, &H54, &H56, &H58, &H5A, &H5C, &H5E, _
        &H60, &H62, &H64, &H66, &H68, &H6A, &H6C, &H6E, &H70, &H72, &H74, &H76, &H78, &H7A, &H7C, &H7E, _
        &H80, &H82, &H84, &H86, &H88, &H8A, &H8C, &H8E, &H90, &H92, &H94, &H96, &H98, &H9A, &H9C, &H9E, _
        &HA0, &HA2, &HA4, &HA6, &HA8, &HAA, &HAC, &HAE, &HB0, &HB2, &HB4, &HB6, &HB8, &HBA, &HBC, &HBE, _
        &HC0, &HC2, &HC4, &HC6, &HC8, &HCA, &HCC, &HCE, &HD0, &HD2, &HD4, &HD6, &HD8, &HDA, &HDC, &HDE, _
        &HE0, &HE2, &HE4, &HE6, &HE8, &HEA, &HEC, &HEE, &HF0, &HF2, &HF4, &HF6, &HF8, &HFA, &HFC, &HFE, _
        &H1B, &H19, &H1F, &H1D, &H13, &H11, &H17, &H15, &HB, &H9, &HF, &HD, &H3, &H1, &H7, &H5, _
        &H3B, &H39, &H3F, &H3D, &H33, &H31, &H37, &H35, &H2B, &H29, &H2F, &H2D, &H23, &H21, &H27, &H25, _
        &H5B, &H59, &H5F, &H5D, &H53, &H51, &H57, &H55, &H4B, &H49, &H4F, &H4D, &H43, &H41, &H47, &H45, _
        &H7B, &H79, &H7F, &H7D, &H73, &H71, &H77, &H75, &H6B, &H69, &H6F, &H6D, &H63, &H61, &H67, &H65, _
        &H9B, &H99, &H9F, &H9D, &H93, &H91, &H97, &H95, &H8B, &H89, &H8F, &H8D, &H83, &H81, &H87, &H85, _
        &HBB, &HB9, &HBF, &HBD, &HB3, &HB1, &HB7, &HB5, &HAB, &HA9, &HAF, &HAD, &HA3, &HA1, &HA7, &HA5, _
        &HDB, &HD9, &HDF, &HDD, &HD3, &HD1, &HD7, &HD5, &HCB, &HC9, &HCF, &HCD, &HC3, &HC1, &HC7, &HC5, _
        &HFB, &HF9, &HFF, &HFD, &HF3, &HF1, &HF7, &HF5, &HEB, &HE9, &HEF, &HED, &HE3, &HE1, &HE7, &HE5)
 
    g3 = Array( _
        &H0, &H3, &H6, &H5, &HC, &HF, &HA, &H9, &H18, &H1B, &H1E, &H1D, &H14, &H17, &H12, &H11, _
        &H30, &H33, &H36, &H35, &H3C, &H3F, &H3A, &H39, &H28, &H2B, &H2E, &H2D, &H24, &H27, &H22, &H21, _
        &H60, &H63, &H66, &H65, &H6C, &H6F, &H6A, &H69, &H78, &H7B, &H7E, &H7D, &H74, &H77, &H72, &H71, _
        &H50, &H53, &H56, &H55, &H5C, &H5F, &H5A, &H59, &H48, &H4B, &H4E, &H4D, &H44, &H47, &H42, &H41, _
        &HC0, &HC3, &HC6, &HC5, &HCC, &HCF, &HCA, &HC9, &HD8, &HDB, &HDE, &HDD, &HD4, &HD7, &HD2, &HD1, _
        &HF0, &HF3, &HF6, &HF5, &HFC, &HFF, &HFA, &HF9, &HE8, &HEB, &HEE, &HED, &HE4, &HE7, &HE2, &HE1, _
        &HA0, &HA3, &HA6, &HA5, &HAC, &HAF, &HAA, &HA9, &HB8, &HBB, &HBE, &HBD, &HB4, &HB7, &HB2, &HB1, _
        &H90, &H93, &H96, &H95, &H9C, &H9F, &H9A, &H99, &H88, &H8B, &H8E, &H8D, &H84, &H87, &H82, &H81, _
        &H9B, &H98, &H9D, &H9E, &H97, &H94, &H91, &H92, &H83, &H80, &H85, &H86, &H8F, &H8C, &H89, &H8A, _
        &HAB, &HA8, &HAD, &HAE, &HA7, &HA4, &HA1, &HA2, &HB3, &HB0, &HB5, &HB6, &HBF, &HBC, &HB9, &HBA, _
        &HFB, &HF8, &HFD, &HFE, &HF7, &HF4, &HF1, &HF2, &HE3, &HE0, &HE5, &HE6, &HEF, &HEC, &HE9, &HEA, _
        &HCB, &HC8, &HCD, &HCE, &HC7, &HC4, &HC1, &HC2, &HD3, &HD0, &HD5, &HD6, &HDF, &HDC, &HD9, &HDA, _
        &H5B, &H58, &H5D, &H5E, &H57, &H54, &H51, &H52, &H43, &H40, &H45, &H46, &H4F, &H4C, &H49, &H4A, _
        &H6B, &H68, &H6D, &H6E, &H67, &H64, &H61, &H62, &H73, &H70, &H75, &H76, &H7F, &H7C, &H79, &H7A, _
        &H3B, &H38, &H3D, &H3E, &H37, &H34, &H31, &H32, &H23, &H20, &H25, &H26, &H2F, &H2C, &H29, &H2A, _
        &HB, &H8, &HD, &HE, &H7, &H4, &H1, &H2, &H13, &H10, &H15, &H16, &H1F, &H1C, &H19, &H1A)
 
    g9 = Array( _
        &H0, &H9, &H12, &H1B, &H24, &H2D, &H36, &H3F, &H48, &H41, &H5A, &H53, &H6C, &H65, &H7E, &H77, _
        &H90, &H99, &H82, &H8B, &HB4, &HBD, &HA6, &HAF, &HD8, &HD1, &HCA, &HC3, &HFC, &HF5, &HEE, &HE7, _
        &H3B, &H32, &H29, &H20, &H1F, &H16, &HD, &H4, &H73, &H7A, &H61, &H68, &H57, &H5E, &H45, &H4C, _
        &HAB, &HA2, &HB9, &HB0, &H8F, &H86, &H9D, &H94, &HE3, &HEA, &HF1, &HF8, &HC7, &HCE, &HD5, &HDC, _
        &H76, &H7F, &H64, &H6D, &H52, &H5B, &H40, &H49, &H3E, &H37, &H2C, &H25, &H1A, &H13, &H8, &H1, _
        &HE6, &HEF, &HF4, &HFD, &HC2, &HCB, &HD0, &HD9, &HAE, &HA7, &HBC, &HB5, &H8A, &H83, &H98, &H91, _
        &H4D, &H44, &H5F, &H56, &H69, &H60, &H7B, &H72, &H5, &HC, &H17, &H1E, &H21, &H28, &H33, &H3A, _
        &HDD, &HD4, &HCF, &HC6, &HF9, &HF0, &HEB, &HE2, &H95, &H9C, &H87, &H8E, &HB1, &HB8, &HA3, &HAA, _
        &HEC, &HE5, &HFE, &HF7, &HC8, &HC1, &HDA, &HD3, &HA4, &HAD, &HB6, &HBF, &H80, &H89, &H92, &H9B, _
        &H7C, &H75, &H6E, &H67, &H58, &H51, &H4A, &H43, &H34, &H3D, &H26, &H2F, &H10, &H19, &H2, &HB, _
        &HD7, &HDE, &HC5, &HCC, &HF3, &HFA, &HE1, &HE8, &H9F, &H96, &H8D, &H84, &HBB, &HB2, &HA9, &HA0, _
        &H47, &H4E, &H55, &H5C, &H63, &H6A, &H71, &H78, &HF, &H6, &H1D, &H14, &H2B, &H22, &H39, &H30, _
        &H9A, &H93, &H88, &H81, &HBE, &HB7, &HAC, &HA5, &HD2, &HDB, &HC0, &HC9, &HF6, &HFF, &HE4, &HED, _
        &HA, &H3, &H18, &H11, &H2E, &H27, &H3C, &H35, &H42, &H4B, &H50, &H59, &H66, &H6F, &H74, &H7D, _
        &HA1, &HA8, &HB3, &HBA, &H85, &H8C, &H97, &H9E, &HE9, &HE0, &HFB, &HF2, &HCD, &HC4, &HDF, &HD6, _
        &H31, &H38, &H23, &H2A, &H15, &H1C, &H7, &HE, &H79, &H70, &H6B, &H62, &H5D, &H54, &H4F, &H46)
 
    g11 = Array( _
        &H0, &HB, &H16, &H1D, &H2C, &H27, &H3A, &H31, &H58, &H53, &H4E, &H45, &H74, &H7F, &H62, &H69, _
        &HB0, &HBB, &HA6, &HAD, &H9C, &H97, &H8A, &H81, &HE8, &HE3, &HFE, &HF5, &HC4, &HCF, &HD2, &HD9, _
        &H7B, &H70, &H6D, &H66, &H57, &H5C, &H41, &H4A, &H23, &H28, &H35, &H3E, &HF, &H4, &H19, &H12, _
        &HCB, &HC0, &HDD, &HD6, &HE7, &HEC, &HF1, &HFA, &H93, &H98, &H85, &H8E, &HBF, &HB4, &HA9, &HA2, _
        &HF6, &HFD, &HE0, &HEB, &HDA, &HD1, &HCC, &HC7, &HAE, &HA5, &HB8, &HB3, &H82, &H89, &H94, &H9F, _
        &H46, &H4D, &H50, &H5B, &H6A, &H61, &H7C, &H77, &H1E, &H15, &H8, &H3, &H32, &H39, &H24, &H2F, _
        &H8D, &H86, &H9B, &H90, &HA1, &HAA, &HB7, &HBC, &HD5, &HDE, &HC3, &HC8, &HF9, &HF2, &HEF, &HE4, _
        &H3D, &H36, &H2B, &H20, &H11, &H1A, &H7, &HC, &H65, &H6E, &H73, &H78, &H49, &H42, &H5F, &H54, _
        &HF7, &HFC, &HE1, &HEA, &HDB, &HD0, &HCD, &HC6, &HAF, &HA4, &HB9, &HB2, &H83, &H88, &H95, &H9E, _
        &H47, &H4C, &H51, &H5A, &H6B, &H60, &H7D, &H76, &H1F, &H14, &H9, &H2, &H33, &H38, &H25, &H2E, _
        &H8C, &H87, &H9A, &H91, &HA0, &HAB, &HB6, &HBD, &HD4, &HDF, &HC2, &HC9, &HF8, &HF3, &HEE, &HE5, _
        &H3C, &H37, &H2A, &H21, &H10, &H1B, &H6, &HD, &H64, &H6F, &H72, &H79, &H48, &H43, &H5E, &H55, _
        &H1, &HA, &H17, &H1C, &H2D, &H26, &H3B, &H30, &H59, &H52, &H4F, &H44, &H75, &H7E, &H63, &H68, _
        &HB1, &HBA, &HA7, &HAC, &H9D, &H96, &H8B, &H80, &HE9, &HE2, &HFF, &HF4, &HC5, &HCE, &HD3, &HD8, _
        &H7A, &H71, &H6C, &H67, &H56, &H5D, &H40, &H4B, &H22, &H29, &H34, &H3F, &HE, &H5, &H18, &H13, _
        &HCA, &HC1, &HDC, &HD7, &HE6, &HED, &HF0, &HFB, &H92, &H99, &H84, &H8F, &HBE, &HB5, &HA8, &HA3)
 
    g13 = Array( _
        &H0, &HD, &H1A, &H17, &H34, &H39, &H2E, &H23, &H68, &H65, &H72, &H7F, &H5C, &H51, &H46, &H4B, _
        &HD0, &HDD, &HCA, &HC7, &HE4, &HE9, &HFE, &HF3, &HB8, &HB5, &HA2, &HAF, &H8C, &H81, &H96, &H9B, _
        &HBB, &HB6, &HA1, &HAC, &H8F, &H82, &H95, &H98, &HD3, &HDE, &HC9, &HC4, &HE7, &HEA, &HFD, &HF0, _
        &H6B, &H66, &H71, &H7C, &H5F, &H52, &H45, &H48, &H3, &HE, &H19, &H14, &H37, &H3A, &H2D, &H20, _
        &H6D, &H60, &H77, &H7A, &H59, &H54, &H43, &H4E, &H5, &H8, &H1F, &H12, &H31, &H3C, &H2B, &H26, _
        &HBD, &HB0, &HA7, &HAA, &H89, &H84, &H93, &H9E, &HD5, &HD8, &HCF, &HC2, &HE1, &HEC, &HFB, &HF6, _
        &HD6, &HDB, &HCC, &HC1, &HE2, &HEF, &HF8, &HF5, &HBE, &HB3, &HA4, &HA9, &H8A, &H87, &H90, &H9D, _
        &H6, &HB, &H1C, &H11, &H32, &H3F, &H28, &H25, &H6E, &H63, &H74, &H79, &H5A, &H57, &H40, &H4D, _
        &HDA, &HD7, &HC0, &HCD, &HEE, &HE3, &HF4, &HF9, &HB2, &HBF, &HA8, &HA5, &H86, &H8B, &H9C, &H91, _
        &HA, &H7, &H10, &H1D, &H3E, &H33, &H24, &H29, &H62, &H6F, &H78, &H75, &H56, &H5B, &H4C, &H41, _
        &H61, &H6C, &H7B, &H76, &H55, &H58, &H4F, &H42, &H9, &H4, &H13, &H1E, &H3D, &H30, &H27, &H2A, _
        &HB1, &HBC, &HAB, &HA6, &H85, &H88, &H9F, &H92, &HD9, &HD4, &HC3, &HCE, &HED, &HE0, &HF7, &HFA, _
        &HB7, &HBA, &HAD, &HA0, &H83, &H8E, &H99, &H94, &HDF, &HD2, &HC5, &HC8, &HEB, &HE6, &HF1, &HFC, _
        &H67, &H6A, &H7D, &H70, &H53, &H5E, &H49, &H44, &HF, &H2, &H15, &H18, &H3B, &H36, &H21, &H2C, _
        &HC, &H1, &H16, &H1B, &H38, &H35, &H22, &H2F, &H64, &H69, &H7E, &H73, &H50, &H5D, &H4A, &H47, _
        &HDC, &HD1, &HC6, &HCB, &HE8, &HE5, &HF2, &HFF, &HB4, &HB9, &HAE, &HA3, &H80, &H8D, &H9A, &H97)
 
    g14 = Array( _
        &H0, &HE, &H1C, &H12, &H38, &H36, &H24, &H2A, &H70, &H7E, &H6C, &H62, &H48, &H46, &H54, &H5A, _
        &HE0, &HEE, &HFC, &HF2, &HD8, &HD6, &HC4, &HCA, &H90, &H9E, &H8C, &H82, &HA8, &HA6, &HB4, &HBA, _
        &HDB, &HD5, &HC7, &HC9, &HE3, &HED, &HFF, &HF1, &HAB, &HA5, &HB7, &HB9, &H93, &H9D, &H8F, &H81, _
        &H3B, &H35, &H27, &H29, &H3, &HD, &H1F, &H11, &H4B, &H45, &H57, &H59, &H73, &H7D, &H6F, &H61, _
        &HAD, &HA3, &HB1, &HBF, &H95, &H9B, &H89, &H87, &HDD, &HD3, &HC1, &HCF, &HE5, &HEB, &HF9, &HF7, _
        &H4D, &H43, &H51, &H5F, &H75, &H7B, &H69, &H67, &H3D, &H33, &H21, &H2F, &H5, &HB, &H19, &H17, _
        &H76, &H78, &H6A, &H64, &H4E, &H40, &H52, &H5C, &H6, &H8, &H1A, &H14, &H3E, &H30, &H22, &H2C, _
        &H96, &H98, &H8A, &H84, &HAE, &HA0, &HB2, &HBC, &HE6, &HE8, &HFA, &HF4, &HDE, &HD0, &HC2, &HCC, _
        &H41, &H4F, &H5D, &H53, &H79, &H77, &H65, &H6B, &H31, &H3F, &H2D, &H23, &H9, &H7, &H15, &H1B, _
        &HA1, &HAF, &HBD, &HB3, &H99, &H97, &H85, &H8B, &HD1, &HDF, &HCD, &HC3, &HE9, &HE7, &HF5, &HFB, _
        &H9A, &H94, &H86, &H88, &HA2, &HAC, &HBE, &HB0, &HEA, &HE4, &HF6, &HF8, &HD2, &HDC, &HCE, &HC0, _
        &H7A, &H74, &H66, &H68, &H42, &H4C, &H5E, &H50, &HA, &H4, &H16, &H18, &H32, &H3C, &H2E, &H20, _
        &HEC, &HE2, &HF0, &HFE, &HD4, &HDA, &HC8, &HC6, &H9C, &H92, &H80, &H8E, &HA4, &HAA, &HB8, &HB6, _
        &HC, &H2, &H10, &H1E, &H34, &H3A, &H28, &H26, &H7C, &H72, &H60, &H6E, &H44, &H4A, &H58, &H56, _
        &H37, &H39, &H2B, &H25, &HF, &H1, &H13, &H1D, &H47, &H49, &H5B, &H55, &H7F, &H71, &H63, &H6D, _
        &HD7, &HD9, &HCB, &HC5, &HEF, &HE1, &HF3, &HFD, &HA7, &HA9, &HBB, &HB5, &H9F, &H91, &H83, &H8D)
 
    sbox = Array( _
        &H63, &H7C, &H77, &H7B, &HF2, &H6B, &H6F, &HC5, &H30, &H1, &H67, &H2B, &HFE, &HD7, &HAB, &H76, _
        &HCA, &H82, &HC9, &H7D, &HFA, &H59, &H47, &HF0, &HAD, &HD4, &HA2, &HAF, &H9C, &HA4, &H72, &HC0, _
        &HB7, &HFD, &H93, &H26, &H36, &H3F, &HF7, &HCC, &H34, &HA5, &HE5, &HF1, &H71, &HD8, &H31, &H15, _
        &H4, &HC7, &H23, &HC3, &H18, &H96, &H5, &H9A, &H7, &H12, &H80, &HE2, &HEB, &H27, &HB2, &H75, _
        &H9, &H83, &H2C, &H1A, &H1B, &H6E, &H5A, &HA0, &H52, &H3B, &HD6, &HB3, &H29, &HE3, &H2F, &H84, _
        &H53, &HD1, &H0, &HED, &H20, &HFC, &HB1, &H5B, &H6A, &HCB, &HBE, &H39, &H4A, &H4C, &H58, &HCF, _
        &HD0, &HEF, &HAA, &HFB, &H43, &H4D, &H33, &H85, &H45, &HF9, &H2, &H7F, &H50, &H3C, &H9F, &HA8, _
        &H51, &HA3, &H40, &H8F, &H92, &H9D, &H38, &HF5, &HBC, &HB6, &HDA, &H21, &H10, &HFF, &HF3, &HD2, _
        &HCD, &HC, &H13, &HEC, &H5F, &H97, &H44, &H17, &HC4, &HA7, &H7E, &H3D, &H64, &H5D, &H19, &H73, _
        &H60, &H81, &H4F, &HDC, &H22, &H2A, &H90, &H88, &H46, &HEE, &HB8, &H14, &HDE, &H5E, &HB, &HDB, _
        &HE0, &H32, &H3A, &HA, &H49, &H6, &H24, &H5C, &HC2, &HD3, &HAC, &H62, &H91, &H95, &HE4, &H79, _
        &HE7, &HC8, &H37, &H6D, &H8D, &HD5, &H4E, &HA9, &H6C, &H56, &HF4, &HEA, &H65, &H7A, &HAE, &H8, _
        &HBA, &H78, &H25, &H2E, &H1C, &HA6, &HB4, &HC6, &HE8, &HDD, &H74, &H1F, &H4B, &HBD, &H8B, &H8A, _
        &H70, &H3E, &HB5, &H66, &H48, &H3, &HF6, &HE, &H61, &H35, &H57, &HB9, &H86, &HC1, &H1D, &H9E, _
        &HE1, &HF8, &H98, &H11, &H69, &HD9, &H8E, &H94, &H9B, &H1E, &H87, &HE9, &HCE, &H55, &H28, &HDF, _
        &H8C, &HA1, &H89, &HD, &HBF, &HE6, &H42, &H68, &H41, &H99, &H2D, &HF, &HB0, &H54, &HBB, &H16)
 
    sboxinv = Array( _
        &H52, &H9, &H6A, &HD5, &H30, &H36, &HA5, &H38, &HBF, &H40, &HA3, &H9E, &H81, &HF3, &HD7, &HFB, _
        &H7C, &HE3, &H39, &H82, &H9B, &H2F, &HFF, &H87, &H34, &H8E, &H43, &H44, &HC4, &HDE, &HE9, &HCB, _
        &H54, &H7B, &H94, &H32, &HA6, &HC2, &H23, &H3D, &HEE, &H4C, &H95, &HB, &H42, &HFA, &HC3, &H4E, _
        &H8, &H2E, &HA1, &H66, &H28, &HD9, &H24, &HB2, &H76, &H5B, &HA2, &H49, &H6D, &H8B, &HD1, &H25, _
        &H72, &HF8, &HF6, &H64, &H86, &H68, &H98, &H16, &HD4, &HA4, &H5C, &HCC, &H5D, &H65, &HB6, &H92, _
        &H6C, &H70, &H48, &H50, &HFD, &HED, &HB9, &HDA, &H5E, &H15, &H46, &H57, &HA7, &H8D, &H9D, &H84, _
        &H90, &HD8, &HAB, &H0, &H8C, &HBC, &HD3, &HA, &HF7, &HE4, &H58, &H5, &HB8, &HB3, &H45, &H6, _
        &HD0, &H2C, &H1E, &H8F, &HCA, &H3F, &HF, &H2, &HC1, &HAF, &HBD, &H3, &H1, &H13, &H8A, &H6B, _
        &H3A, &H91, &H11, &H41, &H4F, &H67, &HDC, &HEA, &H97, &HF2, &HCF, &HCE, &HF0, &HB4, &HE6, &H73, _
        &H96, &HAC, &H74, &H22, &HE7, &HAD, &H35, &H85, &HE2, &HF9, &H37, &HE8, &H1C, &H75, &HDF, &H6E, _
        &H47, &HF1, &H1A, &H71, &H1D, &H29, &HC5, &H89, &H6F, &HB7, &H62, &HE, &HAA, &H18, &HBE, &H1B, _
        &HFC, &H56, &H3E, &H4B, &HC6, &HD2, &H79, &H20, &H9A, &HDB, &HC0, &HFE, &H78, &HCD, &H5A, &HF4, _
        &H1F, &HDD, &HA8, &H33, &H88, &H7, &HC7, &H31, &HB1, &H12, &H10, &H59, &H27, &H80, &HEC, &H5F, _
        &H60, &H51, &H7F, &HA9, &H19, &HB5, &H4A, &HD, &H2D, &HE5, &H7A, &H9F, &H93, &HC9, &H9C, &HEF, _
        &HA0, &HE0, &H3B, &H4D, &HAE, &H2A, &HF5, &HB0, &HC8, &HEB, &HBB, &H3C, &H83, &H53, &H99, &H61, _
        &H17, &H2B, &H4, &H7E, &HBA, &H77, &HD6, &H26, &HE1, &H69, &H14, &H63, &H55, &H21, &HC, &H7D)
 
    rcon = Array( _
        &H8D, &H1, &H2, &H4, &H8, &H10, &H20, &H40, &H80, &H1B, &H36, &H6C, &HD8, &HAB, &H4D, &H9A, _
        &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, &HC5, &H91, &H39, _
        &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, &H25, &H4A, &H94, &H33, &H66, &HCC, &H83, &H1D, &H3A, _
        &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, &H8, &H10, &H20, &H40, &H80, &H1B, &H36, &H6C, &HD8, _
        &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, _
        &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, &H25, &H4A, &H94, &H33, &H66, &HCC, _
        &H83, &H1D, &H3A, &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, &H8, &H10, &H20, &H40, &H80, &H1B, _
        &H36, &H6C, &HD8, &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, &H6A, &HD4, &HB3, _
        &H7D, &HFA, &HEF, &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, &H25, &H4A, &H94, _
        &H33, &H66, &HCC, &H83, &H1D, &H3A, &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, &H8, &H10, &H20, _
        &H40, &H80, &H1B, &H36, &H6C, &HD8, &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, &HC6, &H97, &H35, _
        &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, &H61, &HC2, &H9F, _
        &H25, &H4A, &H94, &H33, &H66, &HCC, &H83, &H1D, &H3A, &H74, &HE8, &HCB, &H8D, &H1, &H2, &H4, _
        &H8, &H10, &H20, &H40, &H80, &H1B, &H36, &H6C, &HD8, &HAB, &H4D, &H9A, &H2F, &H5E, &HBC, &H63, _
        &HC6, &H97, &H35, &H6A, &HD4, &HB3, &H7D, &HFA, &HEF, &HC5, &H91, &H39, &H72, &HE4, &HD3, &HBD, _
        &H61, &HC2, &H9F, &H25, &H4A, &H94, &H33, &H66, &HCC, &H83, &H1D, &H3A, &H74, &HE8, &HCB)
 
 '--------------------------------------------------------------------
    For I = 0 To 15
        nonce(I) = 0
    Next
 '--------------------------------------------------------------------
    For I = 0 To (Len(sPassword) - 1)
        aesKey(I) = Asc(Mid(sPassword, I + 1, 1))
    Next
 '--------------------------------------------------------------------
    For I = Len(sPassword) To 31
        aesKey(I) = 0
    Next
 '--------------------------------------------------------------------
    expandedKey = expandKey(aesKey, sbox, rcon)
 
    sPlain = sMessage
    sCipher = ""
    J = 0
    isDone = False

    Do Until isDone
        sTemp = Mid(sPlain, J * 16 + 1, 16)

        If Len(sTemp) < 16 Then
            For I = Len(sTemp) To 15
                sTemp = sTemp & Chr(0)
            Next
        End If

        For I = 0 To 15
            block(I) = Asc(Mid(sTemp, (I Mod 4) * 4 + (I \ 4) + 1, 1))
        Next

        If (J + 1) * 16 >= Len(sPlain) Then
            isDone = True
        End If

        J = J + 1

        If isEncode Then
            R = 0
            For I = 0 To 15
                block(I) = block(I) Xor nonce(I) Xor expandedKey((I Mod 4) * 4 + (I \ 4))
            Next

            For x = 1 To 13
                block(0) = sbox(block(0))
                block(1) = sbox(block(1))
                block(2) = sbox(block(2))
                block(3) = sbox(block(3))

                intTemp = sbox(block(4))
                block(4) = sbox(block(5))
                block(5) = sbox(block(6))
                block(6) = sbox(block(7))
                block(7) = intTemp

                intTemp = sbox(block(8))
                block(8) = sbox(block(10))
                block(10) = intTemp
                intTemp = sbox(block(9))
                block(9) = sbox(block(11))
                block(11) = intTemp

                intTemp = sbox(block(12))
                block(12) = sbox(block(15))
                block(15) = sbox(block(14))
                block(14) = sbox(block(13))
                block(13) = intTemp

                R = x * 16
                For I = 0 To 3
                    temp(0) = block(I)
                    temp(1) = block(I + 4)
                    temp(2) = block(I + 8)
                    temp(3) = block(I + 12)

                    block(I) = g2(temp(0)) Xor temp(3) Xor temp(2) Xor g3(temp(1)) Xor expandedKey(R + I * 4)
                    block(I + 4) = g2(temp(1)) Xor temp(0) Xor temp(3) Xor g3(temp(2)) Xor expandedKey(R + I * 4 + 1)
                    block(I + 8) = g2(temp(2)) Xor temp(1) Xor temp(0) Xor g3(temp(3)) Xor expandedKey(R + I * 4 + 2)
                    block(I + 12) = g2(temp(3)) Xor temp(2) Xor temp(1) Xor g3(temp(0)) Xor expandedKey(R + I * 4 + 3)
                Next
            Next

            block(0) = sbox(block(0)) Xor expandedKey(224)
            block(1) = sbox(block(1)) Xor expandedKey(228)
            block(2) = sbox(block(2)) Xor expandedKey(232)
            block(3) = sbox(block(3)) Xor expandedKey(236)

            intTemp = sbox(block(4)) Xor expandedKey(237)
            block(4) = sbox(block(5)) Xor expandedKey(225)
            block(5) = sbox(block(6)) Xor expandedKey(229)
            block(6) = sbox(block(7)) Xor expandedKey(233)
            block(7) = intTemp

            intTemp = sbox(block(8)) Xor expandedKey(234)
            block(8) = sbox(block(10)) Xor expandedKey(226)
            block(10) = intTemp
            intTemp = sbox(block(9)) Xor expandedKey(238)
            block(9) = sbox(block(11)) Xor expandedKey(230)
            block(11) = intTemp

            intTemp = sbox(block(12)) Xor expandedKey(231)
            block(12) = sbox(block(15)) Xor expandedKey(227)
            block(15) = sbox(block(14)) Xor expandedKey(239)
            block(14) = sbox(block(13)) Xor expandedKey(235)
            block(13) = intTemp

            For I = 0 To 15
                nonce(I) = block(I)
            Next
        Else
            For I = 0 To 15
                priorCipher(I) = block(I)
            Next

            block(0) = sboxinv(block(0) Xor expandedKey(224))
            block(1) = sboxinv(block(1) Xor expandedKey(228))
            block(2) = sboxinv(block(2) Xor expandedKey(232))
            block(3) = sboxinv(block(3) Xor expandedKey(236))

            intTemp = sboxinv(block(4) Xor expandedKey(225))
            block(4) = sboxinv(block(7) Xor expandedKey(237))
            block(7) = sboxinv(block(6) Xor expandedKey(233))
            block(6) = sboxinv(block(5) Xor expandedKey(229))
            block(5) = intTemp

            intTemp = sboxinv(block(8) Xor expandedKey(226))
            block(8) = sboxinv(block(10) Xor expandedKey(234))
            block(10) = intTemp
            intTemp = sboxinv(block(9) Xor expandedKey(230))
            block(9) = sboxinv(block(11) Xor expandedKey(238))
            block(11) = intTemp

            intTemp = sboxinv(block(12) Xor expandedKey(227))
            block(12) = sboxinv(block(13) Xor expandedKey(231))
            block(13) = sboxinv(block(14) Xor expandedKey(235))
            block(14) = sboxinv(block(15) Xor expandedKey(239))
            block(15) = intTemp

            For x = 13 To 1 Step -1
                R = x * 16

                For I = 0 To 3
                    temp(0) = block(I) Xor expandedKey(R + I * 4)
                    temp(1) = block(I + 4) Xor expandedKey(R + I * 4 + 1)
                    temp(2) = block(I + 8) Xor expandedKey(R + I * 4 + 2)
                    temp(3) = block(I + 12) Xor expandedKey(R + I * 4 + 3)

                    block(I) = g14(temp(0)) Xor g9(temp(3)) Xor g13(temp(2)) Xor g11(temp(1))
                    block(I + 4) = g14(temp(1)) Xor g9(temp(0)) Xor g13(temp(3)) Xor g11(temp(2))
                    block(I + 8) = g14(temp(2)) Xor g9(temp(1)) Xor g13(temp(0)) Xor g11(temp(3))
                    block(I + 12) = g14(temp(3)) Xor g9(temp(2)) Xor g13(temp(1)) Xor g11(temp(0))
                Next

                block(0) = sboxinv(block(0))
                block(1) = sboxinv(block(1))
                block(2) = sboxinv(block(2))
                block(3) = sboxinv(block(3))

                intTemp = sboxinv(block(4))
                block(4) = sboxinv(block(7))
                block(7) = sboxinv(block(6))
                block(6) = sboxinv(block(5))
                block(5) = intTemp

                intTemp = sboxinv(block(8))
                block(8) = sboxinv(block(10))
                block(10) = intTemp
                intTemp = sboxinv(block(9))
                block(9) = sboxinv(block(11))
                block(11) = intTemp

                intTemp = sboxinv(block(12))
                block(12) = sboxinv(block(13))
                block(13) = sboxinv(block(14))
                block(14) = sboxinv(block(15))
                block(15) = intTemp
            Next

            R = 0
            For I = 0 To 15
                block(I) = block(I) Xor expandedKey((I Mod 4) * 4 + (I \ 4)) Xor nonce(I)
                nonce(I) = priorCipher(I)
            Next
        End If

        For I = 0 To 15
            sCipher = sCipher & Chr(block((I Mod 4) * 4 + (I \ 4)))
        Next
    Loop
'---------------------------------------------------
ExitHere:
    AES = sCipher
    Exit Function
'-------------------
ErrHandle:
    ErrPrint "AES", Err.Number, Err.Description
    Err.Clear
End Function
Private Function keyScheduleCore(ByRef row(), ByVal a, ByRef sbox(), ByRef rcon())
    Dim result(4), I
 
    For I = 0 To 3
        result(I) = sbox(row((I + 5) Mod 4))
    Next
 
    result(0) = result(0) Xor rcon(a)
    keyScheduleCore = result
End Function
Private Function expandKey(ByRef key(), ByRef sbox(), ByRef rcon())
    Dim rConIter, temp(), I, result(240)
 
    ReDim temp(4)
    rConIter = 1
 
    For I = 0 To 31
        result(I) = key(I)
    Next
 
    For I = 32 To 239 Step 4
        temp(0) = result(I - 4)
        temp(1) = result(I - 3)
        temp(2) = result(I - 2)
        temp(3) = result(I - 1)
 
        If I Mod 32 = 0 Then
            temp = keyScheduleCore(temp, rConIter, sbox, rcon)
            rConIter = rConIter + 1
        End If
 
        If I Mod 32 = 16 Then
            temp(0) = sbox(temp(0))
            temp(1) = sbox(temp(1))
            temp(2) = sbox(temp(2))
            temp(3) = sbox(temp(3))
        End If
 
        result(I) = result(I - 32) Xor temp(0)
        result(I + 1) = result(I - 31) Xor temp(1)
        result(I + 2) = result(I - 30) Xor temp(2)
        result(I + 3) = result(I - 29) Xor temp(3)
    Next
 
    expandKey = result
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' CONVERTERS
Private Function Conv_To_Base64String(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    Conv_To_Base64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function


Private Function Conv_To_HexString(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    Conv_To_HexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function
Private Function ConvBase64_To_Bytes(varStr As String) As Byte()
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
         .DataType = "bin.base64"
         .Text = varStr
         ConvBase64_To_Bytes = .nodeTypedValue
    End With
 End Function
Private Function nTrim2(TheString As String) As String
Dim I As Long
    Dim iPos As Long
    iPos = Len(TheString)
    For I = iPos To 0 Step -1
        iPos = I
        If Mid$(TheString, I, 1) <> Chr$(0) Then Exit For
    Next
    nTrim2 = Left$(TheString, iPos)
End Function
Private Function Encode64(sString As String) As String

    Dim bTrans(63) As Byte, lPowers8(255) As Long, lPowers16(255) As Long, bOut() As Byte, bIn() As Byte
    Dim lChar As Long, lTrip As Long, iPad As Integer, lLen As Long, lTemp As Long, lPos As Long, lOutSize As Long

    For lTemp = 0 To 63                                 'Fill the translation table.
        Select Case lTemp
            Case 0 To 25
                bTrans(lTemp) = 65 + lTemp              'A - Z
            Case 26 To 51
                bTrans(lTemp) = 71 + lTemp              'a - z
            Case 52 To 61
                bTrans(lTemp) = lTemp - 4               '1 - 0
            Case 62
                bTrans(lTemp) = 43                      'Chr(43) = "+"
            Case 63
                bTrans(lTemp) = 47                      'Chr(47) = "/"
        End Select
    Next lTemp

    For lTemp = 0 To 255                                'Fill the 2^8 and 2^16 lookup tables.
        lPowers8(lTemp) = lTemp * cl2Exp8
        lPowers16(lTemp) = lTemp * cl2Exp16
    Next lTemp

    iPad = Len(sString) Mod 3                           'See if the length is divisible by 3
    If iPad Then                                        'If not, figure out the end pad and resize the input.
        iPad = 3 - iPad
        sString = sString & String(iPad, Chr(0))
    End If

    bIn = StrConv(sString, vbFromUnicode)               'Load the input string.
    lLen = ((UBound(bIn) + 1) \ 3) * 4                  'Length of resulting string.
    lTemp = lLen \ 72                                   'Added space for vbCrLfs.
    lOutSize = ((lTemp * 2) + lLen) - 1                 'Calculate the size of the output buffer.
    ReDim bOut(lOutSize)                                'Make the output buffer.

    lLen = 0                                            'Reusing this one, so reset it.

    For lChar = LBound(bIn) To UBound(bIn) Step 3
        lTrip = lPowers16(bIn(lChar)) + lPowers8(bIn(lChar + 1)) + bIn(lChar + 2)    'Combine the 3 bytes
        lTemp = lTrip And clOneMask                     'Mask for the first 6 bits
        bOut(lPos) = bTrans(lTemp \ cl2Exp18)           'Shift it down to the low 6 bits and get the value
        lTemp = lTrip And clTwoMask                     'Mask for the second set.
        bOut(lPos + 1) = bTrans(lTemp \ cl2Exp12)       'Shift it down and translate.
        lTemp = lTrip And clThreeMask                   'Mask for the third set.
        bOut(lPos + 2) = bTrans(lTemp \ cl2Exp6)        'Shift it down and translate.
        bOut(lPos + 3) = bTrans(lTrip And clFourMask)   'Mask for the low set.
        If lLen = 68 Then                               'Ready for a newline
            bOut(lPos + 4) = 13                         'Chr(13) = vbCr
            bOut(lPos + 5) = 10                         'Chr(10) = vbLf
            lLen = 0                                    'Reset the counter
            lPos = lPos + 6
        Else
            lLen = lLen + 4
            lPos = lPos + 4
        End If
    Next lChar

    If bOut(lOutSize) = 10 Then lOutSize = lOutSize - 2 'Shift the padding chars down if it ends with CrLf.

    If iPad = 1 Then                                    'Add the padding chars if any.
        bOut(lOutSize) = 61                             'Chr(61) = "="
    ElseIf iPad = 2 Then
        bOut(lOutSize) = 61
        bOut(lOutSize - 1) = 61
    End If

    Encode64 = StrConv(bOut, vbUnicode)                 'Convert back to a string and return it.

End Function
Private Function Decode64(sString As String) As String

    Dim bOut() As Byte, bIn() As Byte, bTrans(255) As Byte, lPowers6(63) As Long, lPowers12(63) As Long
    Dim lPowers18(63) As Long, lQuad As Long, iPad As Integer, lChar As Long, lPos As Long, sOUT As String
    Dim lTemp As Long

    sString = Replace(sString, vbCr, vbNullString)      'Get rid of the vbCrLfs.  These could be in...
    sString = Replace(sString, vbLf, vbNullString)      'either order.

    lTemp = Len(sString) Mod 4                          'Test for valid input.
    If lTemp Then
        Call Err.Raise(vbObjectError, "MyDecode", "Input string is not valid Base64.")
    End If

    If InStrRev(sString, "==") Then                     'InStrRev is faster when you know it's at the end.
        iPad = 2                                        'Note:  These translate to 0, so you can leave them...
    ElseIf InStrRev(sString, "=") Then                  'in the string and just resize the output.
        iPad = 1
    End If

    For lTemp = 0 To 255                                'Fill the translation table.
        Select Case lTemp
            Case 65 To 90
                bTrans(lTemp) = lTemp - 65              'A - Z
            Case 97 To 122
                bTrans(lTemp) = lTemp - 71              'a - z
            Case 48 To 57
                bTrans(lTemp) = lTemp + 4               '1 - 0
            Case 43
                bTrans(lTemp) = 62                      'Chr(43) = "+"
            Case 47
                bTrans(lTemp) = 63                      'Chr(47) = "/"
        End Select
    Next lTemp

    For lTemp = 0 To 63                                 'Fill the 2^6, 2^12, and 2^18 lookup tables.
        lPowers6(lTemp) = lTemp * cl2Exp6
        lPowers12(lTemp) = lTemp * cl2Exp12
        lPowers18(lTemp) = lTemp * cl2Exp18
    Next lTemp

    bIn = StrConv(sString, vbFromUnicode)               'Load the input byte array.
    ReDim bOut((((UBound(bIn) + 1) \ 4) * 3) - 1)       'Prepare the output buffer.

    For lChar = 0 To UBound(bIn) Step 4
        
        lQuad = lPowers18(bTrans(bIn(lChar))) + lPowers12(bTrans(bIn(lChar + 1))) + _
        lPowers6(bTrans(bIn(lChar + 2))) + bTrans(bIn(lChar + 3))           'Rebuild the bits.
        
        lTemp = lQuad And clHighMask                    'Mask for the first byte
        bOut(lPos) = lTemp \ cl2Exp16                   'Shift it down
        lTemp = lQuad And clMidMask                     'Mask for the second byte
        bOut(lPos + 1) = lTemp \ cl2Exp8                'Shift it down
        bOut(lPos + 2) = lQuad And clLowMask            'Mask for the third byte
        lPos = lPos + 3
    Next lChar

    sOUT = StrConv(bOut, vbUnicode)                     'Convert back to a string.
    If iPad Then sOUT = Left$(sOUT, Len(sOUT) - iPad)   'Chop off any extra bytes.
    Decode64 = sOUT

End Function
