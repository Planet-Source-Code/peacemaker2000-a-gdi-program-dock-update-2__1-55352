VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'Kein
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lstNames 
      Height          =   1035
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":002B
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   120
      Pattern         =   "*.png"
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const bytMaxSize As Byte = 128
Private Const bytMinSize As Byte = 64
Private Const gdiBicubic = 7

Dim funcBlend32bpp As BLENDFUNCTION
Dim bmpInfo As BITMAPINFO

Dim dcMemory As Long, bmpMemory As Long
Dim lngHeight As Long, lngWidth As Long, lngBitmap As Long, lngImage As Long, lngGDI As Long, lngReturn As Long, lngCursor As Long
Dim sngIndex As Single, sngUBound As Single, sngStep As Single, sngStartTop As Single, sngStartLeft As Single
Dim lngFont As Long, lngBrush As Long, lngFontFamily As Long, lngCurrentFont As Long, lngFormat As Long
Dim sngHeight As Single, sngWidth As Single, sngLeft As Single, sngTop As Single, sngFrom() As Single
Dim apiWindow As POINTAPI, apiPoint As POINTAPI, apiMouse As POINTAPI
Dim bDrawn As Boolean, bHandle As Boolean
Dim gdiPosition As TASKBAR_POSITION
Dim gdipInit As GDIPLUS_STARTINPUT
Dim gdipColors As GDIPLUS_COLORS
Dim rctText As RECTF

Private Sub Form_Click()
MsgBox sngIndex
Unload Me
End Sub

Private Sub Form_Initialize()
gdipInit.GDIPlusVersion = 1
If GdiplusStartup(lngGDI, gdipInit, ByVal 0&) <> 0 Then
    MsgBox "Error loading GDI+!", vbCritical
    Unload Me
End If

' Set the position of the Program Dock
gdiPosition = vbTop
Me.Height = Screen.Height
Me.Width = Screen.Width
If Right(App.Path, 1) = "\" Then
    File1.Path = App.Path & "Icons\"
Else
    File1.Path = App.Path & "\Icons\"
End If
ReDim Preserve sngFrom(File1.ListCount)
sngUBound = File1.ListCount - 1
If gdiPosition = vbBottom Then sngStartTop = (Me.Height / Screen.TwipsPerPixelX) - bytMaxSize
If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngStartLeft = ((Me.Width / Screen.TwipsPerPixelX) - (File1.ListCount * bytMinSize)) / 2
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngStartTop = ((Me.Height / Screen.TwipsPerPixelY) - (File1.ListCount * bytMinSize)) / 2
If gdiPosition = vbRight Then sngStartLeft = (Me.Width / Screen.TwipsPerPixelX) - bytMaxSize

bmpInfo.bmpHeader.Size = Len(bmpInfo.bmpHeader)
bmpInfo.bmpHeader.BitCount = 32
bmpInfo.bmpHeader.Height = Me.ScaleHeight
bmpInfo.bmpHeader.Width = Me.ScaleWidth
bmpInfo.bmpHeader.Planes = 1
bmpInfo.bmpHeader.SizeImage = bmpInfo.bmpHeader.Width * bmpInfo.bmpHeader.Height * (bmpInfo.bmpHeader.BitCount / 8)

dcMemory = CreateCompatibleDC(Me.hdc)
bmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
SelectObject dcMemory, bmpMemory
GdipCreateFromHDC dcMemory, lngImage

If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngLeft = sngStartLeft
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngTop = sngStartTop
For a = 0 To File1.ListCount - 1
    sngHeight = bytMinSize
    sngWidth = bytMinSize
    If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMinSize
    If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMinSize
    LoadPictureGDIPlus File1.Path & "\" & File1.List(a), CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
    If gdiPosition = vbBottom Or gdiPosition = vbTop Then
        sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
        sngLeft = sngLeft + sngWidth
    ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
        sngFrom(a) = sngTop * Screen.TwipsPerPixelY
        sngTop = sngTop + sngHeight
    End If
Next a
If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngFrom(UBound(sngFrom)) = sngTop * Screen.TwipsPerPixelY
UpdateGDIPlus
End Sub

Private Sub DrawText(strText As String, X As Single, Y As Single, Optional strFont As String = "Tahoma", Optional bytFontSize As Byte = 22, Optional bytBorderSize As Byte = 3)
GdipCreateFromHDC dcMemory, lngFont
GdipCreateSolidFill Black, lngBrush
GdipCreateFontFamilyFromName StrConv(strFont, vbUnicode), 0, lngFontFamily
GdipCreateFont lngFontFamily, bytFontSize, FontStyleBold, UnitPoint, lngCurrentFont
GdipCreateStringFormat 0, 0, lngFormat
GdipSetStringFormatAlign lngFormat, StringAlignmentCenter
GdipSetStringFormatLineAlign lngFormat, StringAlignmentNear

' Draw the border
rctText.Left = Y - bytBorderSize
rctText.Top = X
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
rctText.Left = Y + bytBorderSize
rctText.Top = X
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
rctText.Left = Y
rctText.Top = X - bytBorderSize
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush
rctText.Left = Y
rctText.Top = X + bytBorderSize
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush

' Draw the Text
GdipCreateSolidFill White, lngBrush
rctText.Left = Y
rctText.Top = X
rctText.Right = Me.ScaleWidth
rctText.Bottom = 36
GdipDrawString lngFont, StrConv(strText, vbUnicode), -1, lngCurrentFont, rctText, lngFormat, lngBrush

GdipDeleteStringFormat lngFormat
GdipDeleteFont lngCurrentFont
GdipDeleteFontFamily lngFontFamily
GdipDeleteBrush lngBrush
GdipDeleteGraphics lngFont
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lngImage Then
    GdipReleaseDC lngImage, dcMemory
    GdipDeleteGraphics lngImage
End If
If lngBitmap Then GdipDisposeImage lngBitmap
If lngGDI Then GdiplusShutdown lngGDI
End Sub

Private Function RestoreGDIPlus()
DeleteObject bmpMemory
bmpMemory = CreateDIBSection(dcMemory, bmpInfo, DIB_RGB_COLORS, ByVal 0, 0, 0)
SelectObject dcMemory, bmpMemory
GdipCreateFromHDC dcMemory, lngImage
End Function

Private Function UpdateGDIPlus()
lngReturn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
SetWindowLong Me.hwnd, GWL_EXSTYLE, lngReturn Or WS_EX_LAYERED
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE

apiPoint.X = 0
apiPoint.Y = 0
apiWindow.X = Screen.Width / Screen.TwipsPerPixelX
apiWindow.Y = Screen.Height / Screen.TwipsPerPixelY

funcBlend32bpp.AlphaFormat = AC_SRC_ALPHA
funcBlend32bpp.BlendFlags = 0
funcBlend32bpp.BlendOp = AC_SRC_OVER
funcBlend32bpp.SourceConstantAlpha = 255

GdipDisposeImage lngBitmap
GdipDeleteGraphics lngImage
UpdateLayeredWindow Me.hwnd, Me.hdc, ByVal 0&, apiWindow, dcMemory, apiPoint, 0, funcBlend32bpp, ULW_ALPHA
End Function

Private Function LoadPictureGDIPlus(strFilename As String, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
GdipLoadImageFromFile StrPtr(strFilename), lngBitmap
If Width = -1 Or Height = -1 Then
    GdipGetImageHeight lngBitmap, Height
    GdipGetImageWidth lngBitmap, Width
End If
' For better quality use:
' GdipSetInterpolationMode lngImage, gdiBicubic
GdipDrawImageRectI lngImage, lngBitmap, Left, Top, Width, Height
' This is the damn line which i forgot
' and which caused the memory leak
GdipDisposeImage lngBitmap
End Function

Private Sub Timer1_Timer()
lngReturn = GetCursorPos(apiMouse)

If gdiPosition = vbBottom Then bHandle = apiMouse.Y < (Me.Height / Screen.TwipsPerPixelY) - bytMaxSize Or apiMouse.X < sngStartLeft Or apiMouse.X > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX
If gdiPosition = vbTop Then bHandle = apiMouse.Y > bytMaxSize Or apiMouse.X < sngStartLeft Or apiMouse.X > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX
If gdiPosition = vbRight Then bHandle = apiMouse.X < (Me.Width / Screen.TwipsPerPixelY) - bytMaxSize Or apiMouse.Y < sngStartTop Or apiMouse.Y > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX
If gdiPosition = vbLeft Then bHandle = apiMouse.X > bytMaxSize Or apiMouse.Y < sngStartTop Or apiMouse.Y > sngFrom(UBound(sngFrom)) / Screen.TwipsPerPixelX

If bHandle Then
    ' Check bDrawn so the program doen't
    ' draw the same picture more than once
    If bDrawn = False Then
        If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngLeft = sngStartLeft
        If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngTop = sngStartTop
        RestoreGDIPlus
        For a = 0 To File1.ListCount - 1
            sngHeight = bytMinSize
            sngWidth = bytMinSize
            If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMinSize
            If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMinSize
            LoadPictureGDIPlus File1.Path & "\" & File1.List(a), CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
            If gdiPosition = vbBottom Or gdiPosition = vbTop Then
                sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
                sngLeft = sngLeft + sngWidth
            ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
                sngFrom(a) = sngTop * Screen.TwipsPerPixelY
                sngTop = sngTop + sngHeight
            End If
        Next a
        If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
        If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngFrom(UBound(sngFrom)) = sngTop * Screen.TwipsPerPixelY
        UpdateGDIPlus
        bDrawn = True
    End If
    Exit Sub
End If

bDrawn = False
lngCursor = LoadCursor(0, 32512&)
If (lngCursor > 0) Then SetCursor lngCursor
For a = 0 To sngUBound
    If gdiPosition = vbBottom Or gdiPosition = vbTop Then bHandle = apiMouse.X >= sngFrom(a) / Screen.TwipsPerPixelX And apiMouse.X <= sngFrom(a + 1) / Screen.TwipsPerPixelX
    If gdiPosition = vbLeft Or gdiPosition = vbRight Then bHandle = apiMouse.Y >= sngFrom(a) / Screen.TwipsPerPixelY And apiMouse.Y <= sngFrom(a + 1) / Screen.TwipsPerPixelY
    If bHandle Then
        sngIndex = a
        Exit For
    End If
Next a
If gdiPosition = vbBottom Or gdiPosition = vbTop Then
    sngLeft = sngStartLeft
    sngStep = ((apiMouse.X * Screen.TwipsPerPixelX) - sngFrom(sngIndex)) / (2 * Screen.TwipsPerPixelX)
ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
    sngTop = sngStartTop
    sngStep = ((apiMouse.Y * Screen.TwipsPerPixelX) - sngFrom(sngIndex)) / (2 * Screen.TwipsPerPixelY)
End If

RestoreGDIPlus
For a = 0 To sngUBound
    If a <> sngIndex And (a <> sngIndex - 1 And a <> sngIndex + 1) Then
        sngHeight = bytMinSize
        sngWidth = bytMinSize
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMinSize
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMinSize
    ElseIf a = sngIndex - 1 Then
        sngHeight = bytMaxSize - sngStep
        sngWidth = bytMaxSize - sngStep
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - (bytMaxSize - sngStep)
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - (bytMaxSize - sngStep)
    ElseIf a = sngIndex Then
        sngHeight = bytMaxSize
        sngWidth = bytMaxSize
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - bytMaxSize
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - bytMaxSize
    ElseIf a = sngIndex + 1 Then
        sngHeight = bytMinSize + sngStep
        sngWidth = bytMinSize + sngStep
        If gdiPosition = vbBottom Then sngTop = sngStartTop + bytMaxSize - (bytMinSize + sngStep)
        If gdiPosition = vbRight Then sngLeft = sngStartLeft + bytMaxSize - (bytMinSize + sngStep)
    End If
    DrawText lstNames.List(sngIndex), sngTop + bytMaxSize, 0
    LoadPictureGDIPlus File1.Path & "\" & File1.List(a), CLng(sngLeft), CLng(sngTop), CLng(sngWidth), CLng(sngHeight)
    If gdiPosition = vbBottom Or gdiPosition = vbTop Then
        sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
        sngLeft = sngLeft + sngWidth
    ElseIf gdiPosition = vbLeft Or gdiPosition = vbRight Then
        sngFrom(a) = sngTop * Screen.TwipsPerPixelY
        sngTop = sngTop + sngHeight
    End If
Next a
If gdiPosition = vbBottom Or gdiPosition = vbTop Then sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
If gdiPosition = vbLeft Or gdiPosition = vbRight Then sngFrom(UBound(sngFrom)) = sngTop * Screen.TwipsPerPixelY
UpdateGDIPlus
End Sub
