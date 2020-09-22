Attribute VB_Name = "modAFS"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&
                                        
                                        
                                        
Public Const WM_SYSCOMMAND = &H112

Dim CurRgn, TempRgn As Long

Public Function ShapeForm(bg As Form)
Dim x, y As Integer

CurRgn = CreateRectRgn(0, 0, bg.ScaleWidth, bg.ScaleHeight)

x = 0
y = bg.ScaleHeight - 1

            TempRgn = CreateRectRgn(x, y - 11, x + 1, y + 11)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 1, y - 9, x + 2, y + 9)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 2, y - 8, x + 3, y + 8)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 3, y - 7, x + 5, y + 7)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 5, y - 6, x + 8, y + 6)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 8, y - 4, x + 9, y + 4)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 9, y - 2, x + 10, y + 2)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 10, y - 1, x + 11, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 11, y, x + 13, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            
x = 0
y = 0

            TempRgn = CreateRectRgn(x, y, x + 1, y + 12)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 1, y, x + 2, y + 10)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 2, y, x + 3, y + 9)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 3, y, x + 5, y + 8)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 5, y, x + 8, y + 7)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 8, y, x + 9, y + 5)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 9, y, x + 10, y + 3)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 10, y, x + 11, y + 2)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x + 11, y, x + 12, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)

x = bg.ScaleWidth
y = bg.ScaleHeight

            TempRgn = CreateRectRgn(x - 5, y - 1, x + 1, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x - 3, y - 2, x + 1, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x - 2, y - 3, x + 1, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x - 1, y - 5, x + 1, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)

x = bg.ScaleWidth
y = 0

            TempRgn = CreateRectRgn(x - 5, y - 1, x - 3, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x - 3, y - 1, x - 2, y + 2)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x - 2, y - 1, x - 1, y + 3)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)
            TempRgn = CreateRectRgn(x - 1, y - 1, x, y + 5)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
            DeleteObject (success)

success = SetWindowRgn(bg.hwnd, CurRgn, True)
DeleteObject (CurRgn)
DeleteObject (success)
End Function
