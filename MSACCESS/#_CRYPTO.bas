Attribute VB_Name = "#_CRYPTO"
'*********************************.ze$$e. **********************************************************************************************************
'              .ed$$$eee..      .$$$$$$$P""              ########  #######       #### ####### ##   ##  ##     #######
'           z$$$$$$$$$$$$$$$$$ee$$$$$$"                  ##        ##    ##     ## ## ##      ##  ##   ##     ##
'        .d$$$$$$$$$$$$$$$$$$$$$$$$$"                    ##        ##   ##     ##  ## ##      ####     ##     ####
'      .$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$e..                ##   #### ######     ####### ##      ## ##    ##     ##
'    .$$****""""***$$$$$$$$$$$$$$$$$$$$$$$$$$$be.        ##     ## ##   ##   ##    ## ##      ##   ##  ##     ##
'                     ""**$$$$$$$$$$$$$$$$$$$$$$$L       ######### ##    ## ##     ## ####### ##    ## #####  #######
'                       z$$$$$$$$$$$$$$$$$$$$$$$$$
'                     .$$$$$$$$P**$$$$$$$$$$$$$$$$              ##     ##  #####       ####
'                    d$$$$$$$"              4$$$$$               ##    ##  ##  ##     ## ##
'                  z$$$$$$$$$                $$$P"                ##   ##  ####      ##  ##
'                 d$$$$$$$$$F                $P"                   ##  ##  ##  ##   #######
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##      ####  ###### ##  ##  ##### ###### ####
'                  *$$$$$$$$"                                        ####  #####  ##     ##     ##     ##  ##  ## ##  ##  ##  ##  ##  ##
'                    "***""               _____________                                      ## ##     ####      ###  #####   ##  ##  ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2017/03/19 |                                    #   ##     ## ##      ##  ##      ##  ##  ##
' The module contains frequently used functions and is part of the G-VBA library             ##  ##### ##   ##    ##  ##      ##   ####
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Public Const INITIALIZATION_VECTOR = "12345678" '??8???
Public Const TRIPLE_DES_KEY = "motorakudekirunj" '??16???


'=====================================================================================================================================================
' Pseudo-GUID (HASH) generation
'=====================================================================================================================================================
Public Function GetHASH(Optional sPrefix As String = "HHHH") As String
  GetHASH = sPrefix & "#" & UserName & "#" & GenRandomStr(11, True, True, True) & "#" & _
                 Format(Now(), "yyyymmddhhnnss") & "#" & HDDSerial
End Function
'=====================================================================================================================================================
' CRC32 For String
'=====================================================================================================================================================
Public Function CRC32String(str As String) As String
Dim cP As New cCRYPTO
    If str = "" Then Exit Function
    
    CRC32String = cP.CRC32_String(str) '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function
'=====================================================================================================================================================
' CRC32 For File
'=====================================================================================================================================================
Public Function CRC32File(FilePath As String) As String
Dim cP As New cCRYPTO

    If FilePath = "" Then Exit Function
    CRC32File = cP.CRC32_File(FilePath)  '!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function
'=====================================================================================================================================================
' MD5 For String
'=====================================================================================================================================================
Public Function MD5String(str As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO
    If str = "" Then Exit Function
    
    MD5String = cP.MD5_String(str, outputformat)   '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA1 For String
'=====================================================================================================================================================
Public Function SHA1String(str As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If str = "" Then Exit Function
    
    SHA1String = cP.SHA1_String(str, outputformat)    '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA256 For String
'=====================================================================================================================================================
Public Function SHA256String(str As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If str = "" Then Exit Function
    
    SHA256String = cP.SHA256_String(str, outputformat)     '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA384 For String
'=====================================================================================================================================================
Public Function SHA384String(str As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If str = "" Then Exit Function
    
    SHA384String = cP.SHA384_String(str, outputformat)     '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA512 For String
'=====================================================================================================================================================
Public Function SHA512String(str As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If str = "" Then Exit Function
    
    SHA512String = cP.SHA512_String(str, outputformat)      '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' MD5 For File
'=====================================================================================================================================================
Public Function MD5File(FilePath As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If FilePath = "" Then Exit Function
    
    MD5File = cP.MD5_File(FilePath, outputformat)       '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA1 For File
'=====================================================================================================================================================
Public Function SHA1File(FilePath As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If FilePath = "" Then Exit Function
    
    SHA1File = cP.SHA1_File(FilePath, outputformat)        '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA256 For File
'=====================================================================================================================================================
Public Function SHA256File(FilePath As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If FilePath = "" Then Exit Function
    
    SHA256File = cP.SHA256_File(FilePath, outputformat)         '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA384 For File
'=====================================================================================================================================================
Public Function SHA384File(FilePath As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If FilePath = "" Then Exit Function
    
    SHA384File = cP.SHA384_File(FilePath, outputformat)          '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' SHA512 For File
'=====================================================================================================================================================
Public Function SHA512File(FilePath As String, Optional outputformat As HashOutputFormat = OUTPUT_HEX) As String
Dim cP As New cCRYPTO

   If FilePath = "" Then Exit Function
    
    SHA512File = cP.SHA512_File(FilePath, outputformat)           '!!!!!!!!!!!!!!!!!!!!
    
    Set cP = Nothing
End Function

'=====================================================================================================================================================
' Encrypt string with 3DES Methods
'=====================================================================================================================================================
Public Function TripleDES_Encrypt(str As String, Optional sPass As String) As String
Dim cP As New cCRYPTO
   If str = "" Then Exit Function
      TripleDES_Encrypt = cP.EncryptStringTripleDES(str, sPass)
   Set cP = Nothing
End Function
'=====================================================================================================================================================
' Decrypt string with 3DES Methods
'=====================================================================================================================================================
Public Function TripleDES_Decrypt(str As String, Optional sPass As String) As String
Dim cP As New cCRYPTO
   If str = "" Then Exit Function
      TripleDES_Decrypt = cP.DecryptStringTripleDES(str, sPass)
   Set cP = Nothing
End Function

'=====================================================================================================================================================
' Encrypt string with AES Methods
'=====================================================================================================================================================
Public Function AES_Encrypt(str As String, sPass As String, Optional iRound As Integer = 5) As String
Dim cP As New cCRYPTO
   If str = "" Then Exit Function
      AES_Encrypt = cP.GetEncryptAES(str, sPass, iRound)
   Set cP = Nothing
End Function
'=====================================================================================================================================================
' Decrypt string with 3DES Methods
'=====================================================================================================================================================
Public Function AES_Decrypt(str As String, sPass As String, Optional iRound As Integer = 5) As String
Dim cP As New cCRYPTO
   If str = "" Then Exit Function
      AES_Decrypt = cP.GetDecryptAES(str, sPass, iRound)
   Set cP = Nothing
End Function



