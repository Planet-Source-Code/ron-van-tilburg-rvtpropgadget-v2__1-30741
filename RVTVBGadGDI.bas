Attribute VB_Name = "mPropGDI"
Option Explicit

'The references to GDI32 functions Used

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type BITMAP '14 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

'Pen Constants
Public Const PS_SOLID = 0

Public Type LOGBRUSH
  lbStyle As Long
  lbColor As Long
  lbHatch As Long
End Type

'Brush Types
Public Const BS_SOLID = 0

Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'The references to USER32 functions Used
'Edge Constants
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENOUTER = &H2

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

'DrawText Constants
Public Const DT_CENTER = &H1
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4

Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'Rectancle Functions
Public Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Public Declare Function InflateRect Lib "user32" (ByRef lpRect As RECT, ByVal dx As Long, ByVal dy As Long) As Long

'The references to olepro32 functions Used
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

'===================================  Bitmap used for Rendering Dithered Pattern Fill ========================

Public LastXY              As POINTAPI

Private BM                 As BITMAP
Private BitMapBits(0 To 7) As Long

'==================================   BASIC STACK FOR GDI OBJECTS ============================================

Private GDIstk(0 To 19) As Long
Private GDIsp           As Long

Private Sub PushGDI(Obj As Long)
  If GDIsp < UBound(GDIstk) Then
    GDIstk(GDIsp) = Obj: GDIsp = GDIsp + 1
  Else
    MsgBox "Error PushGDI(): GDIStk is Full"
  End If
End Sub

Private Function PopGDI() As Long
  If GDIsp > LBound(GDIstk) Then
    GDIsp = GDIsp - 1: PopGDI = GDIstk(GDIsp)
  Else
    MsgBox "Error PopGDI(): GDIStk is Empty"
    PopGDI = -1
  End If
End Function
'============================================================================================================

Public Function RectCX(ByRef RT As RECT) As Long      'xcoord of rectangle centre
  RectCX = (RT.Left + RT.Right) \ 2
End Function

Public Function RectCY(ByRef RT As RECT) As Long      'ycoord of rectangle centre
  RectCY = (RT.Top + RT.Bottom) \ 2
End Function

Public Sub MoveXY(ByVal hDC As Long, x As Long, y As Long)
  Call MoveToEx(hDC, x, y, LastXY)
End Sub

Public Function OLEtoRGB(ByVal Color As OLE_COLOR) As Long
  If OleTranslateColor(Color, 0, OLEtoRGB) Then
    OLEtoRGB = &H0    'black
  End If
End Function

Public Function NewSolidPen(ByVal ColorRef As OLE_COLOR) As Long
  NewSolidPen = CreatePen(PS_SOLID, 1, OLEtoRGB(ColorRef))
End Function

Public Sub SelectNewSolidPen(ByVal hDC As Long, ByVal ColorRef As OLE_COLOR)
  Call PushGDI(SelectObject(hDC, NewSolidPen(ColorRef)))
End Sub

Public Sub SelectPrevPen(hDC As Long)
  Call DeleteObject(SelectObject(hDC, PopGDI()))
End Sub

Public Function NewSolidBrush(ByVal ColorRef As OLE_COLOR) As Long
  Dim NewBR As LOGBRUSH
  
  NewBR.lbColor = OLEtoRGB(ColorRef)
  NewBR.lbStyle = BS_SOLID
  NewSolidBrush = CreateBrushIndirect(NewBR)
End Function

'BM should be a 2 color 8x8 bitmap, 0=WHITE=Pen,1=BLACK=Paper
Public Sub PatternFillRect(ByVal hDC As Long, ByRef RT As RECT)
  Dim lPBR As Long, hPBM As Long
    
  If BitMapBits(0) <> &HAAAAAAAA Then                     'only do this the first time
    BitMapBits(0) = &HAAAAAAAA: BitMapBits(1) = &H55555555
    BitMapBits(2) = &HAAAAAAAA: BitMapBits(3) = &H55555555
    BitMapBits(4) = &HAAAAAAAA: BitMapBits(5) = &H55555555
    BitMapBits(6) = &HAAAAAAAA: BitMapBits(7) = &H55555555
    With BM
      .bmWidth = 8
      .bmHeight = 8
      .bmPlanes = 1
      .bmBitsPixel = 1
      .bmWidthBytes = 4
      .bmBits = VarPtr(BitMapBits(0))
    End With
  End If
  
  hPBM = CreateBitmapIndirect(BM)
  If hPBM Then
    lPBR = CreatePatternBrush(hPBM)
    If lPBR Then
      Call FillRect(hDC, RT, lPBR)
      Call DeleteObject(lPBR)
    End If
    Call DeleteObject(hPBM)
  End If
End Sub

Public Sub DrawEllipse(ByVal hDC As Long, _
                       ByVal x0 As Long, ByVal y0 As Long, _
                       ByVal x1 As Long, ByVal y1 As Long, _
                       Optional ByVal ColorRef As Long = -1)
  
  If ColorRef <> -1 Then Call SelectNewSolidPen(hDC, ColorRef)
  Call Ellipse(hDC, x0, y0, x1, y1)
  If ColorRef <> -1 Then Call SelectPrevPen(hDC)
End Sub


