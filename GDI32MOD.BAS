Attribute VB_Name = "GDI32Mod"
'// Made hBitmap Type, Functions, etc.
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SwapBuffers Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Type BITMAPFILEHEADER
  bfType As Integer
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type

Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type
Public Const IMAGE_BITMAP = 0, LR_LOADFROMFILE = &H10
Public Const SRCCOPY = &HCC0020, SRCERASE = &H440328, SRCINVERT = &H660046, SRCAND = &H8800C6
Public Type hBitmap
  Width As Long
  Height As Long
  hdc As Long
  hbmp As Long
  Bits As Long
End Type

Public Sub LoadBitmap(lbFilename As String, lbhBitmap As hBitmap, Optional hHDC As Long)
Dim BMFileNum As Integer
Dim FileHeader As BITMAPFILEHEADER
Dim InfoHeader As BITMAPINFOHEADER
  If Dir(lbFilename) = "" Then
    MsgBox "File Does not Exist!"
    Exit Sub
  End If
  
  FileNum = FreeFile
  Open lbFilename For Binary Access Read As #FileNum
    Get #FileNum, , FileHeader
    Get #FileNum, , InfoHeader
  Close #FileNum
  If IsMissing(hHDC) Then
    lbhBitmap.hdc = CreateCompatibleDC(0)
  Else
    lbhBitmap.hdc = CreateCompatibleDC(hHDC)
  End If
  
  '// LoadImage can also AntiAlisa a Picture as it is loaded into
  '// Memory
  lbhBitmap.hbmp = LoadImage(ByVal 0&, lbFilename, 0, 0, 0, LR_LOADFROMFILE)
  SelectObject lbhBitmap.hdc, lbhBitmap.hbmp
  lbhBitmap.Bits = InfoHeader.biBitCount
  lbhBitmap.Height = InfoHeader.biHeight
  lbhBitmap.Width = InfoHeader.biWidth
End Sub
Public Sub DeleteBitmap(dbhBitmap As hBitmap)
  dbhBitmap.Bits = 0
  dbhBitmap.Height = 0
  dbhBitmap.Width = 0
  DeleteDC dbhBitmap.hdc
  DeleteObject dbhBitmap.hbmp
End Sub
Public Sub CreateBitmap(lbhBitmap As hBitmap, Width As Long, Height As Long, CompatibleHDC As Long)
  lbhBitmap.Width = Width
  lbhBitmap.Height = Height
  lbhBitmap.hdc = CreateCompatibleDC(CompatibleHDC)
  lbhBitmap.hbmp = CreateCompatibleBitmap(CompatibleHDC, Width, Height)
  SelectObject lbhBitmap.hdc, lbhBitmap.hbmp
End Sub
