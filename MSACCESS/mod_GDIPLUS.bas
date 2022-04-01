Attribute VB_Name = "mod_GDIPLUS"

'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' The code below is part of the aerc library provided in the basGDIPlus module vailable on
' https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas
' The module is shipped as is with minor modifications. See source and copyright below. Included to VBA-G module for support
' some of the GDI+ library features with minor changes
' see copyright - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' The text is provided AS IS without warranty or obligation:
'
' ||    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ||    ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' ||    WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' ||    DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' ||    DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' ||    (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' ||    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ||    ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' ||    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' ||    SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'******************************************************************************************************************************************************
Option Explicit

' STRUCTURES AND DECLARATIONS FOR GUID+

#If Win64 Then
Private Type PICTDESC
    cbSizeOfStruct                      As Long
    PicType                             As Long
    hImage                              As LongPtr
    xExt                                As Long
    yExt                                As Long
End Type

Public Type GDIPStartupInput
    GdiplusVersion                      As Long
    DebugEventCallback                  As LongPtr
    SuppressBackgroundThread            As LongPtr
    SuppressExternalCodecs              As LongPtr
End Type


#Else

Private Type PICTDESC
    cbSizeOfStruct                      As Long
    PicType                             As Long
    hImage                              As Long
    xExt                                As Long
    yExt                                As Long
End Type

Public Type GDIPStartupInput
    GdiplusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type


#End If



#If Win64 Then

    'API-Declarations: ----------------------------------------------------------------------------
    'Some calls are commented out, so don't use the current version.

    ' G.A.: olepro32 in oleaut32 geändert. Olepro32 ist in x64 nicht verfügbar.
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fPictureOwnsHandle As LongPtr, ByRef IPic As Object) As Long

    'Retrieve GUID-Type from string :
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, ByRef pclsid As GUID) As Long

    'Memory functions:
    '#Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As LongPtr, ByVal dwBytes As LongPtr) As Long
    '#Private Declare PtrSafe Function GlobalSize Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    '# Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    '#Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long
    '#Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    '#Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByRef Source As Byte, ByVal Length As LongPtr)

    'Modules API:
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

    'Timer API:
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As Long
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

    'OLE-Stream functions :
    Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As LongPtr, ByRef ppstm As Any) As Long
    'Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As LongPtr) As Long


    'GDIPlus Flat-API declarations:

    'Initialization GDIP:
    Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (ByRef token As LongPtr, ByRef inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    'Tear down GDIP:
    Private Declare PtrSafe Function GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr) As Long
    'Load GDIP-Image from file :
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal fileName As LongPtr, BITMAP As LongPtr) As Long
    'Create GDIP- graphical area from Windows-DeviceContext:
    Private Declare PtrSafe Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hdc As LongPtr, ByRef GpGraphics As LongPtr) As Long
    'Delete GDIP graphical area :
    'Private Declare PtrSafe Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As LongPtr) As Long
    'Copy GDIP-Image to graphical area:
    'Private Declare PtrSafe Function GdipDrawImageRect Lib "GDIPlus" (ByVal graphics As LongPtr, ByVal Image As LongPtr, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal height As Single) As Long
    'Clear allocated bitmap memory from GDIP :
    Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As LongPtr) As Long
    'Retrieve windows bitmap handle from GDIP-Image:
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As LongPtr, ByRef hbmReturn As LongPtr, ByVal background As LongPtr) As Long
    'Retrieve Windows-Icon-Handle from GDIP-Image:
    Public Declare PtrSafe Function GdipCreateHICONFromBitmap Lib "GDIPlus" (ByVal BITMAP As LongPtr, ByRef hbmReturn As LongPtr) As Long
    'Scaling GDIP-Image size:
    Private Declare PtrSafe Function GdipGetImageThumbnail Lib "GDIPlus" (ByVal Image As LongPtr, ByVal thumbWidth As LongPtr, ByVal thumbHeight As LongPtr, ByRef thumbImage As LongPtr, Optional ByVal callback As LongPtr = 0, Optional ByVal callbackData As LongPtr = 0) As Long
    'Retrieve GDIP-Image from Windows-Bitmap-Handle:
    Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As LongPtr, ByVal hPal As LongPtr, ByRef BITMAP As LongPtr) As Long
    'Retrieve GDIP-Image from Windows-Icon-Handle:
    Private Declare PtrSafe Function GdipCreateBitmapFromHICON Lib "GDIPlus" (ByVal hIcon As LongPtr, ByRef BITMAP As LongPtr) As Long
    'Retrieve width of a GDIP-Image (Pixel):
    Private Declare PtrSafe Function GdipGetImageWidth Lib "GDIPlus" (ByVal Image As LongPtr, ByRef Width As LongPtr) As Long
    'Retrieve height of a GDIP-Image (Pixel):
    Private Declare PtrSafe Function GdipGetImageHeight Lib "GDIPlus" (ByVal Image As LongPtr, ByRef height As LongPtr) As Long
    'Save GDIP-Image to file in seletable format:
    'Private Declare PtrSafe Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As LongPtr, ByVal fileName As LongPtr, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
    'Save GDIP-Image in OLE-Stream with seletable format:
    Private Declare PtrSafe Function GdipSaveImageToStream Lib "GDIPlus" (ByVal Image As LongPtr, ByVal stream As IUnknown, ByRef clsidEncoder As GUID, ByRef encoderParams As Any) As Long
    'Retrieve GDIP-Image from OLE-Stream-Object:
    Private Declare PtrSafe Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal stream As IUnknown, ByRef Image As LongPtr) As Long
    'Create a gdip image from scratch
    Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "GDIPlus" (ByVal Width As Long, ByVal height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByRef scan0 As Any, ByRef BITMAP As Long) As Long
    'Get the DC of an gdip image
    Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As LongPtr, ByRef graphics As LongPtr) As Long
    'Blit the contents of an gdip image to another image DC using positioning
    'Private Declare PtrSafe Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal graphics As LongPtr, ByVal Image As LongPtr, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long


    '-----------------------------------------------------------------------------------------
    'Global module variable:
    Private lGDIP                       As LongPtr
    '-----------------------------------------------------------------------------------------

    Private TempVarGDIPlus              As LongPtr
#Else

    'API-Declarations: ----------------------------------------------------------------------------

    'Convert a windows bitmap to OLE-Picture :
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As Long
    'Retrieve GUID-Type from string :
    Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long

    'Memory functions:
    'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    'Private Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
    'Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    'Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Byte, ByVal Length As Long)

    'Modules API:
    Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
    Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

    'Timer API:
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long


    'OLE-Stream functions :
    Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
    'Private Declare Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As Long) As Long

    'GDIPlus Flat-API declarations:

    'Initialization GDIP:
    Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    'Tear down GDIP:
    Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
    'Load GDIP-Image from file :
    Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal fileName As Long, BITMAP As Long) As Long
    'Create GDIP- graphical area from Windows-DeviceContext:
    Private Declare Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hdc As Long, GpGraphics As Long) As Long
    'Delete GDIP graphical area :
    Private Declare Function GdipDeleteGraphics Lib "GDIPlus" (ByVal graphics As Long) As Long
    'Copy GDIP-Image to graphical area:
    Private Declare Function GdipDrawImageRect Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal height As Single) As Long
    'Clear allocated bitmap memory from GDIP :
    Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
    'Retrieve windows bitmap handle from GDIP-Image:
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As Long, hbmReturn As Long, ByVal background As Long) As Long
    'Retrieve Windows-Icon-Handle from GDIP-Image:
    Public Declare Function GdipCreateHICONFromBitmap Lib "GDIPlus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
    'Scaling GDIP-Image size:
    Private Declare Function GdipGetImageThumbnail Lib "GDIPlus" (ByVal Image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
    'Retrieve GDIP-Image from Windows-Bitmap-Handle:
    Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
    'Retrieve GDIP-Image from Windows-Icon-Handle:
    Private Declare Function GdipCreateBitmapFromHICON Lib "GDIPlus" (ByVal hIcon As Long, BITMAP As Long) As Long
    'Retrieve width of a GDIP-Image (Pixel):
    Private Declare Function GdipGetImageWidth Lib "GDIPlus" (ByVal Image As Long, Width As Long) As Long
    'Retrieve height of a GDIP-Image (Pixel):
    Private Declare Function GdipGetImageHeight Lib "GDIPlus" (ByVal Image As Long, height As Long) As Long
    'Save GDIP-Image to file in seletable format:
    Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal fileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
    'Save GDIP-Image in OLE-Stream with seletable format:
    Private Declare Function GdipSaveImageToStream Lib "GDIPlus" (ByVal Image As Long, ByVal stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
    'Retrieve GDIP-Image from OLE-Stream-Object:
    Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal stream As IUnknown, Image As Long) As Long
    'Create a gdip image from scratch
    Private Declare Function GdipCreateBitmapFromScan0 Lib "GDIPlus" (ByVal Width As Long, ByVal height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
    'Get the DC of an gdip image
    Private Declare Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As Long, graphics As Long) As Long
    'Blit the contents of an gdip image to another image DC using positioning
    Private Declare Function GdipDrawImageRectRectI Lib "GDIPlus" (ByVal graphics As Long, ByVal Image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long

    '-----------------------------------------------------------------------------------------
    'Global module variable:
    Private lGDIP                       As Long

#End If

'//Declares from mTDSample
'Public Declare PtrSafe Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
'Public Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const GUID_IPicture  As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"    'IPicture

Private tVarTimer()                     As Long
Private lCounter                        As Long
Private bSharedLoad                     As Boolean

'***********************

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////######////#///#///####////#####///////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////#/////////#///#///#///#/////#////////#////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////#//###////#///#///#///#/////#//////#####//////////////////////////////////////////////////////////////////////////////
'////////////////////////////////#////#////#///#///#///#/////#////////#////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////###### ////###/////###////#####///////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' This section is shipped as is with minor modifications. See source and copyright below
' The function is implemented in accordance with the aegit/aerc library
' //https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' Included to VBA-G module for support some of the GDI+ library features with minor changes
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'======================================================================================================================================================
'Create an Bitmap Struvture from Byte-Array PicBin()
' The function is implemented in accordance with the aegit/aerc library
' //https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' Included to VBA-G module for support some of the GDI+ library features with minor changes
'======================================================================================================================================================
Public Function ArrayToBitmap(ByRef PicBin() As Byte, Optional ByVal iconHeight As Long, Optional ByVal iconWidth As Long) As LongPtr

    On Error GoTo 0

    Dim IStm                            As IUnknown
    '    #If Win64 Then
    Dim lBitmap                         As LongPtr
    Dim hBmp                            As LongPtr
    '    #Else
    '        Dim lBitmap As Long
    '        Dim hBmp As Long
    '    #End If

    Dim ret                             As Long


    If Not InitGDIP Then
Debug.Print "Exit"
        Exit Function
    End If

    ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)  'Create stream from memory stack

    If ret = 0 Then    'OK, start GDIP :
        'Convert stream to GDIP-Image :
        ret = GdipLoadImageFromStream(IStm, lBitmap)
        ScaleBitmap lBitmap, iconHeight, iconWidth
        GdipCreateHICONFromBitmap lBitmap, hBmp

        ArrayToBitmap = hBmp

        'Clear memory ...
        GdipDisposeImage lBitmap
    End If

End Function
'======================================================================================================================================================
'Create an OLE-Picture from Byte-Array PicBin()
' The function is implemented in accordance with the aegit/aerc library
' //https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' Included to VBA-G module for support some of the GDI+ library features with minor changes
'======================================================================================================================================================
Public Function ArrayToPicture(ByRef PicBin() As Byte) As Picture

    On Error GoTo 0

    Dim IStm                            As IUnknown
    #If Win64 Then
        Dim lBitmap                     As LongPtr
        Dim hBmp                        As LongPtr
    #Else
        Dim lBitmap                     As Long
        Dim hBmp                        As Long
    #End If

    Dim ret                             As Long

    If Not InitGDIP Then
        Exit Function
    End If

    ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)  'Create stream from memory stack

    If ret = 0 Then                                        ' OK, start GDIP
        
        ret = GdipLoadImageFromStream(IStm, lBitmap)       ' Convert stream to GDIP-Image
        If ret = 0 Then
            GdipCreateHBITMAPFromBitmap lBitmap, hBmp, 0&  ' Get Windows-Bitmap from GDIP-Image
            If hBmp <> 0 Then
                Set ArrayToPicture = BitmapToPicture(hBmp) ' Convert bitmap to picture object
            End If
        End If
        
        GdipDisposeImage lBitmap                           ' Clear memory ...
    End If

End Function
'======================================================================================================================================================
' Picture upload function - provided in 64-be option
' The function is implemented in accordance with the aegit/aerc library
' //https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' Included to VBA-G module for support some of the GDI+ library features with minor changes
'======================================================================================================================================================
Public Function LoadPicturehBitmap(ByRef sFileName As String, Optional ByVal iconHeight As Long, Optional ByVal iconWidth As Long) As LongPtr
    #If Win64 Then
        Dim hBmp                        As LongPtr
        Dim hPic                        As LongPtr
    #Else
        Dim hBmp                        As Long
        Dim hPic                        As Long
    #End If

    On Error GoTo 0

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromFile(StrPtr(sFileName), hPic) = 0 Then
        ScaleBitmap hPic, iconHeight, iconWidth
        GdipCreateHBITMAPFromBitmap hPic, hBmp, 0&
        LoadPicturehBitmap = hBmp
    End If

End Function
'======================================================================================================================================================
' Function used by external processes to load the icon as a resource
' The function is implemented in accordance with the aegit/aerc library
' //https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' Included to VBA-G module for support some of the GDI+ library features with minor changes
'======================================================================================================================================================
Public Function LoadPicturehIcon(ByRef sFileName As String, Optional ByVal iconHeight As Long, Optional ByVal iconWidth As Long) As LongPtr
    Dim hBmp                            As LongPtr
    Dim hPic                            As LongPtr

    On Error GoTo 0

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromFile(StrPtr(sFileName), hBmp) = 0 Then
        ScaleBitmap hBmp, iconHeight, iconWidth

        GdipCreateHICONFromBitmap hBmp, hPic
        LoadPicturehIcon = hPic
    End If

End Function
'======================================================================================================================================================
'Create a picture object from an Access  attachment
'strTable:              Table containing picture file attachments
'strAttachmentField:    Name of the attachment column in the table
'strPkField:            Name of primary key field in the image table.
'dtPkField:             Data type of primary key field (or some other field with unique values).
'varPkValue:            PK value of the record with the image to return
'? AttachmentToPicture("ribbonimages","imageblob","cloudy.png").Width
' The function is implemented in accordance with the aegit/aerc library
' //https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' Included to VBA-G module for support some of the GDI+ library features with minor changes
'======================================================================================================================================================
Public Function AttachmentToPicture(ByVal strTable As String, ByVal strAttachmentField As String, ByVal strPkField As String, _
                                           ByVal dtPkField As DataTypeEnum, ByVal varPkValue As Variant, Optional ByVal Addin As Boolean) As StdPicture
Dim SQL                          As String
Dim bIn()                           As Byte

    '    #If Win64 Then
    '              Dim nOffset           As LongPtr
    '              Dim nSize             As LongPtr
    '    #Else
    Dim nOffset                         As Long
    Dim nSize                           As Long
    '    #End If
    SQL = "SELECT " & strAttachmentField & ".FileData AS data " & _
             "FROM " & strTable & " WHERE " & strPkField & "="

    Select Case dtPkField
        Case dbText
            SQL = SQL & "'" & varPkValue & "'"
        Case Else
            SQL = SQL & varPkValue
    End Select

On Error Resume Next
    If Addin Then
        bIn = CodeDb.OpenRecordset(SQL, dbOpenSnapshot)(0)
    Else
        bIn = DBEngine(0)(0).OpenRecordset(SQL, dbOpenSnapshot)(0)
    End If
    If Err.Number = 0 Then
        On Error GoTo 0
        Dim bin2()                      As Byte
        nOffset = bIn(0)    'First byte of Field2.FileData identifies offset to the file data block
        nSize = UBound(bIn)
        ReDim bin2(nSize - nOffset)
        CopyMemory bin2(0), bIn(nOffset), nSize - nOffset   'Copy file into new byte array starting at nOffset
        Set AttachmentToPicture = ArrayToPicture(bin2)
        Erase bin2
        Erase bIn
    Else
Debug.Assert False     'Query failed. Check function arguments.
    End If
End Function

'======================================================================================================================================================
' Load Attached Resource to Icon Handler (optimized only for 64 bit Office)
' The function is implemented in accordance with the aegit/aerc library
' //https://github.com/peterennis/aegit/blob/master/aerc/src/basGDIPlus.bas - (c) mossSOFT / Sascha Trowitzsch rev. 04/2009
' Included to VBA-G module for support some of the GDI+ library features with minor changes
'======================================================================================================================================================
Public Function AttachmentTohIcon(ByVal strTable As String, ByVal strAttachmentField As String, ByVal strPkField As String, _
                    ByVal dtPkField As DataTypeEnum, ByVal varPkValue As Variant, Optional ByVal iconHeight As Long, Optional ByVal iconWidth As Long, _
                                                                                                            Optional ByVal Addin As Boolean) As LongPtr
Dim SQL As String, bIn() As Byte
Dim nOffset As Long, nSize As Long

    SQL = "SELECT " & strAttachmentField & ".FileData AS data " & _
             "FROM " & strTable & " WHERE " & strPkField & "="

    Select Case dtPkField
        Case dbText
            SQL = SQL & "'" & varPkValue & "'"
        Case Else
            SQL = SQL & varPkValue
    End Select

    On Error Resume Next
    
    If Addin Then
        bIn = CodeDb.OpenRecordset(SQL, dbOpenSnapshot)(0)
    Else
        bIn = DBEngine(0)(0).OpenRecordset(SQL, dbOpenSnapshot)(0)
    End If
    
    If Err.Number = 0 Then
        On Error GoTo 0
        Dim bin2()                      As Byte
        nOffset = bIn(0)    'First byte of Field2.FileData identifies offset to the file data block
        nSize = UBound(bIn)
        ReDim bin2(nSize - nOffset)
        
        '// CopyMemory is slow on O365 x64
        CopyMemory bin2(0), bIn(nOffset), nSize - nOffset   'Copy file into new byte array starting at nOffset

        AttachmentTohIcon = ArrayToBitmap(bin2, iconHeight, iconWidth)

        Erase bin2
        Erase bIn
    Else
Debug.Assert False     'Query failed. Check function arguments.
    End If
End Function





















'------------------------------------------------------------------------------------------------------------------------------------------------------
'Initialize GDI+
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function InitGDIP() As Boolean

    On Error GoTo 0

    Dim TGDP                            As GDIPStartupInput
    Dim hMod                            As Long


    If lGDIP = 0 Then
        #If Win64 Then
            If TempVarGDIPlus = 0 Then                   'If lGDIP is broken due to unhandled errors restore it from the Tempvars collection
                TGDP.GdiplusVersion = 1
                hMod = GetModuleHandle("GDIPlus.dll")    'ogl.dll not yet loaded?
                If hMod = 0 Then
                    hMod = LoadLibrary("GDIPlus.dll")
                    bSharedLoad = False
                Else
                    bSharedLoad = True
                End If
                GdiplusStartup lGDIP, TGDP              'Get a personal instance of GDIPlus
                TempVarGDIPlus = lGDIP

            Else
                lGDIP = TempVarGDIPlus
            End If

        #Else
            If IsNull(TempVars("GDIPlusHandle")) Then
                TGDP.GdiplusVersion = 1
                hMod = GetModuleHandle("GDIPlus.dll")
                If hMod = 0 Then
                    hMod = LoadLibrary("GDIPlus.dll")
                    bSharedLoad = False
                Else
                    bSharedLoad = True
                End If
                GdiplusStartup lGDIP, TGDP
                TempVars("GDIPlusHandle") = lGDIP
            Else
                lGDIP = TempVars("GDIPlusHandle")
            End If

        #End If

    End If

    InitGDIP = (lGDIP <> 0)

    AutoShutDown

End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
'Clear GDI+
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ShutDownGDIP()

    On Error GoTo 0

    If lGDIP <> 0 Then

        Dim lngDummy                    As Long
        Dim lngDummyTimer               As Long
        For lngDummy = 0 To lCounter - 1
            lngDummyTimer = tVarTimer(lngDummy)

            If lngDummyTimer <> 0 Then
                If KillTimer(0&, CLng(lngDummyTimer)) Then
                    tVarTimer(lngDummy) = 0
                End If

            End If
        Next

        GdiplusShutdown lGDIP
        lGDIP = 0

        #If Win64 Then
            TempVarGDIPlus = 0
        #Else
            TempVars("GDIPlusHandle") = Null
        #End If

        If Not bSharedLoad Then FreeLibrary GetModuleHandle("GDIPlus.dll")

    End If

End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
'Scheduled ShutDown of GDI+ handle to avoid memory leaks
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AutoShutDown()
'Set to 5 seconds for next shutdown
'That's IMO appropriate for looped routines  - but configure for your own purposes

    On Error GoTo 0

    If lGDIP <> 0 Then
        ReDim Preserve tVarTimer(lCounter)
        tVarTimer(lCounter) = SetTimer(0&, 0&, 5000, AddressOf TimerProc)
        
    End If
    lCounter = lCounter + 1

End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
'Callback for AutoShutDown
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error GoTo 0
    ShutDownGDIP
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
'Load image file with GDIP
'It's equivalent to the method LoadPicture() in OLE-Automation library (stdole2.tlb)
'Allowed format: bmp, gif, jp(e)g, tif, png, wmf, emf, ico
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function LoadPictureGDIP(ByRef sFileName As String) As StdPicture
    #If Win64 Then
        Dim hBmp                        As LongPtr
        Dim hPic                        As LongPtr
    #Else
        Dim hBmp                        As Long
        Dim hPic                        As Long
    #End If

    On Error GoTo 0

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromFile(StrPtr(sFileName), hPic) = 0 Then

        GdipCreateHBITMAPFromBitmap hPic, hBmp, 0&

        If hBmp <> 0 Then
            Set LoadPictureGDIP = BitmapToPicture(hBmp)
            GdipDisposeImage hPic
        End If
    End If

End Function


'------------------------------------------------------------------------------------------------------------------------------------------------------
' Scaling helper function
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ScaleBitmap(ByRef hBmp As LongPtr, Optional ByVal iconHeight As Long, Optional ByVal iconWidth As Long)
    Dim Thumbnail                       As LongPtr
    Dim bmpHeight                       As LongPtr
    Dim bmpWidth                        As LongPtr

    If iconHeight <> 0 Or iconWidth <> 0 Then
        GdipGetImageHeight hBmp, bmpHeight
        GdipGetImageWidth hBmp, bmpWidth
        If iconHeight = 0 Then
            iconHeight = Int((iconWidth / bmpWidth) * bmpHeight)
        End If
        If iconWidth = 0 Then
            iconWidth = Int((iconHeight / bmpHeight) * bmpWidth)
        End If

        GdipGetImageThumbnail hBmp, iconWidth, iconHeight, Thumbnail
        hBmp = Thumbnail
    End If


End Function


'------------------------------------------------------------------------------------------------------------------------------------------------------
' Help function to get a OLE-Picture from Windows-Bitmap-Handle. If bIsIcon = TRUE, an Icon-Handle is committed
'------------------------------------------------------------------------------------------------------------------------------------------------------
#If Win64 Then
    Private Function BitmapToPicture(ByVal hBmp As LongPtr, Optional ByRef bIsIcon As Boolean = False) As StdPicture

        On Error GoTo 0

        Dim TPicConv                        As PICTDESC
        Dim uId                             As GUID

        With TPicConv
            If bIsIcon Then
                .cbSizeOfStruct = 16
                .PicType = 3    'PicType Icon
            Else
                .cbSizeOfStruct = Len(TPicConv)
                .PicType = 1    'PicType Bitmap
            End If
            .hImage = hBmp
        End With

    CLSIDFromString StrPtr(GUID_IPicture), uId
    OleCreatePictureIndirect TPicConv, uId, True, BitmapToPicture

    End Function
#Else
    Private Function BitmapToPicture(ByVal hBmp As Long, Optional bIsIcon As Boolean = False) As StdPicture

        On Error GoTo 0

        Dim TPicConv As PICTDESC, uId       As GUID

        With TPicConv
            If bIsIcon Then
                .cbSizeOfStruct = 16
                .PicType = 3    'PicType Icon
            Else
                .cbSizeOfStruct = Len(TPicConv)
                .PicType = 1    'PicType Bitmap
            End If
            .hImage = hBmp
        End With

        CLSIDFromString StrPtr(GUID_IPicture), uId
        OleCreatePictureIndirect TPicConv, uId, True, BitmapToPicture

    End Function
#End If







