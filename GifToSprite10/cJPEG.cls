VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'===================================================================================
'
' cJPG Class Module (Formerly named cDIBSection.cls)
' -----------------
'
' Created By  : vbAccelerator
'
' Last Update : ?
'
' VB Versions : 5.0 / 6.0
'
' Requires    : modJPEG.bas by Steve McMahon
'               IJL10.DLL (Intel)
'
' Description : This module is intended as an easy interface to Intel's IJL
'               (Intel JPG Library) for use in Visual Basic 5.0 / 6.0
'
' NOTE        : The JPEG_Show function is used to display the loaded JPEG.
'               However, if you set the AutoRedraw property of the object
'               showing the picture to TRUE, the picture will not correctly
'               be displayed.  And if you set the AutoRedraw property to
'               FALSE and then minimize the form it's in or put another
'               window infront of it... then bring the focus back to the
'               form containing the picture, the picture disapears.
'               THE WORK AROUND is to put the JPEG_Show in the Paint event
'               of the object containing the picture.
'
'-----------------------------------------------------------------------------------
' vbAccelerator Copyright© 1999 by Steve McMahon (http://vbaccelerator.com)
' IJL.DLL Copyright© 1999 by Intel
'
' IMPORTANT   : Intel is not responsible for any errors in this code, and
'               should not be mentioned in any Help, About, or support in
'               any product using the Intel library
'
'===================================================================================


' ---- Constants ----
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  Color table in RGBs
Private Const CF_BITMAP = 2
Private Const CF_DIB = 8

' ---- Enumerations / Types ----
Private Type SAFEARRAYBOUND
  cElements       As Long
  lLbound         As Long
End Type

Private Type SAFEARRAY2D
  cDims           As Integer
  fFeatures       As Integer
  cbElements      As Long
  Clocks          As Long
  pvData          As Long
  Bounds(0 To 1)  As SAFEARRAYBOUND
End Type

Private Type RGBQUAD
  rgbBlue         As Byte
  rgbGreen        As Byte
  rgbRed          As Byte
  rgbReserved     As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
  biSize          As Long
  biWidth         As Long
  biHeight        As Long
  biPlanes        As Integer
  biBitCount      As Integer
  biCompression   As Long
  biSizeImage     As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed       As Long
  biClrImportant  As Long
End Type

Private Type BITMAPINFO
  bmiHeader       As BITMAPINFOHEADER
  bmiColors       As RGBQUAD
End Type

Private Type BITMAP
  bmType          As Long
  bmWidth         As Long
  bmHeight        As Long
  bmWidthBytes    As Long
  bmPlanes        As Integer
  bmBitsPixel     As Integer
  bmBits          As Long
End Type

' ---- API Functions / Subs ----
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "USER32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function OpenClipboard Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "USER32" () As Long
Private Declare Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "USER32" () As Long

' Note : The following declaration is not the VB API Viewer - Modify lplpVoid to be Byref so we get the pointer back
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long

' ---- Variables ----

Private m_Loaded  As Boolean    ' Tells weather the class contains a loaded picture
Private m_hDIb    As Long       ' Handle to the current DIBSection
Private m_hBmpOld As Long       ' Handle to the old bitmap in the DC, for clear up
Private m_hDC     As Long       ' Handle to the Device context holding the DIBSection
Private m_lPtr    As Long       ' Address of memory pointing to the DIBSection's bits
Private m_tBI     As BITMAPINFO ' Type containing the Bitmap information


'================================================================================
'                             CLASS INITIALIZATIONS
'================================================================================


Private Sub Class_Terminate()
  
  ClearUp
  
End Sub


'================================================================================
'                          CLASS PROPERTIES
'================================================================================


Public Property Get PictureLoaded() As Boolean
  
  PictureLoaded = m_Loaded
  
End Property

Public Property Let PictureLoaded(ByVal NewValue As Boolean)
  
  m_Loaded = NewValue
  
End Property

Public Property Get DIBSectionBitsPtr() As Long
On Error Resume Next
  
  DIBSectionBitsPtr = m_lPtr
  
End Property

Public Property Get hdc() As Long
On Error Resume Next
  
  hdc = m_hDC
  
End Property

Public Property Get hDib() As Long
On Error Resume Next
  
  hDib = m_hDIb
  
End Property

Public Property Get Height() As Long
On Error Resume Next
  
  Height = m_tBI.bmiHeader.biHeight
  
End Property

Public Property Get Width() As Long
On Error Resume Next
  
  Width = m_tBI.bmiHeader.biWidth
  
End Property


'================================================================================
'                               CLASS METHODS
'================================================================================


' Copy picture to clipboard as a bitmap
Public Function CopyToClipboard() As Boolean
On Error Resume Next
  
  Dim lhDCDesktop As Long
  Dim lhDC As Long
  Dim lhBmpOld As Long
  Dim hObj As Long
  Dim lFmt As Long
  Dim B() As Byte
  Dim tBI As BITMAPINFO
  Dim lPtr As Long
  Dim hDibCopy As Long

  lhDCDesktop = GetDC(GetDesktopWindow())
  If (lhDCDesktop <> 0) Then
    lhDC = CreateCompatibleDC(lhDCDesktop)
    If (lhDC <> 0) Then
      
      ' Create a compatible bitmap and copy to the clipboard
      hObj = CreateCompatibleBitmap(lhDCDesktop, Width, Height)
      If (hObj <> 0) Then
        lhBmpOld = SelectObject(lhDC, hObj)
        PaintPicture lhDC
        SelectObject lhDC, lhBmpOld
        lFmt = CF_BITMAP
        
        ' Now set the clipboard to the bitmap
        If (OpenClipboard(0) <> 0) Then
          EmptyClipboard
          If (SetClipboardData(lFmt, hObj) <> 0) Then
            CopyToClipboard = True
          End If
          CloseClipboard
        End If
      End If
      DeleteDC lhDC
    End If
    DeleteDC lhDCDesktop
  End If

End Function

Public Function CreateDIB(ByVal lhDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByRef hDib As Long) As Boolean
On Error Resume Next
  
  With m_tBI.bmiHeader
    .biSize = Len(m_tBI.bmiHeader)
    .biWidth = lWidth
    .biHeight = lHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = BytesPerScanLine * .biHeight
  End With
  
  hDib = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
  CreateDIB = (hDib <> 0)
  
End Function

Public Function CreateFromPicture(ByRef PictureSource As StdPicture)
On Error GoTo ErrorTrap
  
  Dim lhDC As Long
  Dim lhDCDesktop As Long
  Dim lhBmpOld As Long
  Dim tBMP As BITMAP
  
  m_Loaded = False
  
  GetObjectAPI PictureSource.Handle, Len(tBMP), tBMP
  If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
    lhDCDesktop = GetDC(GetDesktopWindow())
    If (lhDCDesktop <> 0) Then
      lhDC = CreateCompatibleDC(lhDCDesktop)
      DeleteDC lhDCDesktop
      If (lhDC <> 0) Then
        lhBmpOld = SelectObject(lhDC, PictureSource.Handle)
        LoadPictureBlt lhDC
        SelectObject lhDC, lhBmpOld
        DeleteObject lhDC
      End If
    End If
  End If
  
  m_Loaded = True
  
  Exit Function
  
ErrorTrap:
  
  Err.Clear
  
End Function

Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
On Error Resume Next
  
  ClearUp
  m_hDC = CreateCompatibleDC(0)
  If (m_hDC <> 0) Then
    If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
      m_hBmpOld = SelectObject(m_hDC, m_hDIb)
      Create = True
    Else
      DeleteObject m_hDC
      m_hDC = 0
    End If
  End If
  
End Function

Public Property Get BytesPerScanLine() As Long
On Error Resume Next
  
  ' Scans must align on dword boundaries:
  BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
  
End Property

Public Sub LoadPictureBlt(ByVal lhDC As Long, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1, Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
On Error Resume Next
  
  If lSrcWidth < 0 Then
    lSrcWidth = m_tBI.bmiHeader.biWidth
  End If
  
  If lSrcHeight < 0 Then
    lSrcHeight = m_tBI.bmiHeader.biHeight
  End If
  
  BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
  
End Sub

Public Sub PaintPicture(ByVal lhDC As Long, Optional ByVal lDestLeft As Long = 0, Optional ByVal lDestTop As Long = 0, Optional ByVal lDestWidth As Long = -1, Optional ByVal lDestHeight As Long = -1, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
On Error Resume Next
  
  If (lDestWidth < 0) Then
    lDestWidth = m_tBI.bmiHeader.biWidth
  End If
  
  If (lDestHeight < 0) Then
    lDestHeight = m_tBI.bmiHeader.biHeight
  End If
  
  BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
  
End Sub

Public Sub RandomiseBits(Optional ByVal bGray As Boolean = False)
On Error Resume Next

  Dim bDib() As Byte
  Dim X As Long, Y As Long
  Dim lC As Long
  Dim tSA As SAFEARRAY2D
  Dim xEnd As Long
    
  ' Get the bits in the from DIB section
  With tSA
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
  End With
  CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
  
  ' random:
  Randomize Timer
  
  xEnd = (Width - 1) * 3
  If (bGray) Then
    For Y = 0 To m_tBI.bmiHeader.biHeight - 1
      For X = 0 To xEnd Step 3
        lC = Rnd * 255
        bDib(X, Y) = lC
        bDib(X + 1, Y) = lC
        bDib(X + 2, Y) = lC
      Next X
    Next Y
  Else
    For X = 0 To xEnd Step 3
      For Y = 0 To m_tBI.bmiHeader.biHeight - 1
        bDib(X, Y) = 0
        bDib(X + 1, Y) = Rnd * 255
        bDib(X + 2, Y) = Rnd * 255
      Next Y
    Next X
  End If
  
  ' Clear the temporary array descriptor
  ' NOTE : This does not appear to be necessary, but for safety do it anyway
  CopyMemory ByVal VarPtrArray(bDib), 0&, 4
  
End Sub

Public Sub ClearUp()
On Error Resume Next
  
  m_Loaded = False
  
  If (m_hDC <> 0) Then
    If (m_hDIb <> 0) Then
      SelectObject m_hDC, m_hBmpOld
      DeleteObject m_hDIb
    End If
    DeleteObject m_hDC
  End If
  
  m_hDC = 0
  m_hDIb = 0
  m_hBmpOld = 0
  m_lPtr = 0
  
End Sub

Public Function Resample(ByVal lNewHeight As Long, ByVal lNewWidth As Long) As cJPEG
On Error Resume Next
  
  Dim cDib As cJPEG
    
  Set cDib = New cJPEG
  If cDib.Create(lNewWidth, lNewHeight) Then
    If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
      ' Change in size, do resample:
      ResampleDib cDib
    Else
      ' No size change so just return a copy:
      cDib.LoadPictureBlt m_hDC
    End If
    Set Resample = cDib
  End If
  
End Function

Private Function ResampleDib(ByRef cDibTo As cJPEG) As Boolean
On Error Resume Next

  Dim bDibFrom() As Byte
  Dim bDibTo() As Byte
  Dim tSAFrom As SAFEARRAY2D
  Dim tSATo As SAFEARRAY2D
  Dim xScale As Single
  Dim yScale As Single
  Dim X As Long
  Dim Y As Long
  Dim xEnd As Long
  Dim xOut As Long
  Dim fX As Single
  Dim fY As Single
  Dim ifY As Long
  Dim ifX As Long
  Dim dX As Single
  Dim dY As Single
  Dim r As Long
  Dim R1 As Single
  Dim R2 As Single
  Dim R3 As Single
  Dim R4 As Single
  Dim G As Long
  Dim G1 As Single
  Dim G2 As Single
  Dim G3 As Single
  Dim G4 As Single
  Dim B As Long
  Dim B1 As Single
  Dim B2 As Single
  Dim B3 As Single
  Dim B4 As Single
  Dim iR1 As Long
  Dim iG1 As Long
  Dim iB1 As Long
  Dim iR2 As Long
  Dim iG2 As Long
  Dim iB2 As Long
  
  ' Get the bits in the from DIB section:
  With tSAFrom
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
  End With
  CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4
  
  ' Get the bits in the to DIB section:
  With tSATo
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = cDibTo.Height
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = cDibTo.BytesPerScanLine()
    .pvData = cDibTo.DIBSectionBitsPtr
  End With
  CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4
  
  xScale = (Width - 1) / cDibTo.Width
  yScale = (Height - 1) / cDibTo.Height
  xEnd = cDibTo.Width - 1
    
  For Y = 0 To cDibTo.Height - 1
    fY = Y * yScale
    ifY = Int(fY)
    dY = fY - ifY
    For X = 0 To xEnd
      fX = X * xScale
      ifX = Int(fX)
      dX = fX - ifX
      ifX = ifX * 3
      
      ' Interpolate using the four nearest pixels in the source
      B1 = bDibFrom(ifX, ifY): G1 = bDibFrom(ifX + 1, ifY): R1 = bDibFrom(ifX + 2, ifY)
      B2 = bDibFrom(ifX + 3, ifY): G2 = bDibFrom(ifX + 4, ifY): R2 = bDibFrom(ifX + 5, ifY)
      B3 = bDibFrom(ifX, ifY + 1): G3 = bDibFrom(ifX + 1, ifY + 1): R3 = bDibFrom(ifX + 2, ifY + 1)
      B4 = bDibFrom(ifX + 3, ifY + 1): G4 = bDibFrom(ifX + 4, ifY + 1): R4 = bDibFrom(ifX + 5, ifY + 1)
      
      ' Interplate in x direction:
      iR1 = R1 * (1 - dY) + R3 * dY: iG1 = G1 * (1 - dY) + G3 * dY: iB1 = B1 * (1 - dY) + B3 * dY
      iR2 = R2 * (1 - dY) + R4 * dY: iG2 = G2 * (1 - dY) + G4 * dY: iB2 = B2 * (1 - dY) + B4 * dY
      ' Interpolate in y:
      r = iR1 * (1 - dX) + iR2 * dX: G = iG1 * (1 - dX) + iG2 * dX: B = iB1 * (1 - dX) + iB2 * dX
      
      ' Set output:
      If (r < 0) Then r = 0
      If (r > 255) Then r = 255
      If (G < 0) Then G = 0
      If (G > 255) Then G = 255
      If (B < 0) Then B = 0
      If (B > 255) Then B = 255
      xOut = X * 3
      bDibTo(xOut, Y) = B
      bDibTo(xOut + 1, Y) = G
      bDibTo(xOut + 2, Y) = r
    Next X
  Next Y
  
  ' Clear the temporary array descriptor
  ' NOTE : This does not appear to be necessary, but for safety do it anyway
  CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
  CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4
  
End Function
