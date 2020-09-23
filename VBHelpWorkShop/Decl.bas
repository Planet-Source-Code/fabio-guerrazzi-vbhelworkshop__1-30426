Attribute VB_Name = "Decl"
' CaptureClient - Captures the client area of a form.
' CaptureScreen - Captures the entire screen.
' PrintPictureToFitPage - prints any picture as big as possible on
' the page.
'
' NOTES
'    - No error trapping is included in these routines.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Option Explicit
Option Base 0

Public CTLMode As Integer
Public CTLSize As Single
Public CTLZoom As Single


Public Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Public Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Dim Bx As Single, By As Single

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Public Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Sub SaveAsBitmap(Pic As Object, FileName As String)
  Dim P As Picture
  Refresh
  DoEvents
  Set P = CaptureWindow(Pic.hWnd, False, 0, 0, Pic.Width, Pic.Height)
  Set Pic.Picture = P
  SavePicture Pic.Picture, FileName
  Set P = Nothing
End Sub

Public Function CaptureScreen() As Picture
  Dim hWndScreen As Long

   ' Get a handle to the desktop window.
   hWndScreen = GetDesktopWindow()

   ' Call CaptureWindow to capture the entire desktop give the handle
   ' and return the resulting Picture object.

   Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function


  Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim r As Long
  Dim hDCSrc As Long
  Dim hPal As Long
  Dim hPalPrev As Long
  Dim RasterCapsScrn As Long
  Dim HasPaletteScrn As Long
  Dim PaletteSizeScrn As Long
  Dim LogPal As LOGPALETTE

   ' Depending on the value of Client get the proper device context.
   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                    ' window.
   End If

   ' Create a memory device context for the copy process.
   hDCMemory = CreateCompatibleDC(hDCSrc)
   ' Create a bitmap and place it in the memory DC.
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   ' Get screen properties.
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                      ' capabilities.
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                        ' support.
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                        ' palette.

   ' If the screen has a palette make a copy and realize it.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette.
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it.
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

   ' Copy the on-screen image into the memory DC.
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

' Remove the new copy of the  on-screen image.
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system.
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles. Then return the resulting picture
   ' object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CaptureActiveWindow() As Picture

    Dim hWndActive As Long
    Dim r As Long
    
    Dim RectActive As RECT
    
    ' Get a handle to the active/foreground window.
    hWndActive = GetForegroundWindow()
    
    ' Get the dimensions of the window.
    r = GetWindowRect(hWndActive, RectActive)
    
    ' Call CaptureWindow to capture the active window given its
    ' handle and return the Resulting Picture object.
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function

Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long

   Dim Pic As PicBmp
   ' IPicture requires a reference to "Standard OLE Types."
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID

   ' Fill in with IDispatch Interface ID.
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   ' Fill Pic with necessary parts.
   With Pic
      .Size = Len(Pic)          ' Length of structure.
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With

   ' Create Picture object.
   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

   ' Return the new Picture object.
   Set CreateBitmapPicture = IPic
End Function


Public Function CaptureArea(xmin!, ymin!, xmax!, ymax!) As Picture
  
  Dim hWndScreen As Long

   ' Get a handle to the desktop window.
   hWndScreen = GetDesktopWindow()

   ' Call CaptureWindow to capture the entire desktop give the handle
   ' and return the resulting Picture object.

   Set CaptureArea = CaptureWindow(hWndScreen, False, xmin, ymin, xmax, ymax)
   
End Function


