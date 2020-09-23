VERSION 5.00
Begin VB.UserControl RVTPropGadget 
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   20
   ToolboxBitmap   =   "RVTPropGadget.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   -30
      Top             =   1890
   End
End
Attribute VB_Name = "RVTPropGadget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'rvtPropGadget - Gauges, Sliders and Scrollers, all in only one small control
'©2001 Ron van Tilburg - all rights reserved

'Source given for Educational purposes - any commercial use requires permission from the Author
'No representations are given as to the suitability of this code for any particular purpose.

'To compile: requires RVTVBGadGDI.bas, mMouseWheel.bas for external library references
'Subclassing can lead to GPFs in the IDE so take care - only exit the program by ending it

'Tested on Win98 SE 2, VB6

'RVT 17 Dec 2001
' v 1.03  Added MouseWheel support of MS_MOUSEWHEEL MSGS  (only works on Win98,NT4 and above)
' v 1.04  Also added Min,Max,Value,LargeChange and SmallChange properties to be compatible with MS Sliders
'         Only the ValueChanged EVENT is now different (but at least the compiler will see it)

'Elements making up the Type of Proportional Gadget

Private Const GAUGE     As Long = &H10&       'A display Only   Lo-Hi Level PropGadget
Private Const SLIDER    As Long = &H20&       'An interactive   Lo-Hi Level PropGadget
Private Const SCROLLER  As Long = &H40&       'An interactive   Hi-Lo Level PropGadget

Private Const TYPEMASK  As Long = &H70&       'Is it a Slider or a Scroller or a Gauge?

Private Const VERT      As Long = &H1&
Private Const HORZ      As Long = &H2&
Private Const MOVEMASK  As Long = &H3&        'Does it move VERTically or HORiZontally or Both?

Private Const SCALETICKSIZE As Long = 4       'For showing ScaleTicks below or to right of gadget

Public Enum rvtPropGadgetType
  VGauge = GAUGE + VERT                 '&H11
  HGauge = GAUGE + HORZ                 '&H12
  VSlider = SLIDER + VERT               '&H21
  HSlider = SLIDER + HORZ               '&H22
  VScroller = SCROLLER + VERT           '&H41
  HScroller = SCROLLER + HORZ           '&H42
  HVScroller = SCROLLER + VERT + HORZ   '&H43
End Enum

Public Enum rvtScrollingArrows      'Sliders and Gauges do not have scrollingarrows
  None = 0                          'For any Prop Type
  AtTopAndBottom = 1                'Only for VScroller
  AtTop = 2                         'Only for VScroller
  AtBottom = 3                      'Only for VScroller
  AtLeftAndRight = 4                'Only for HScroller
  AtLeft = 5                        'Only for HScroller
  AtRight = 6                       'Only for HScroller
  AtEdges = 7                       'Only for HVScroller
End Enum

'Property Variables:
Private m_PropGadgetType  As rvtPropGadgetType
Private m_ScrollingArrows As rvtScrollingArrows
Private m_Borderless      As Boolean

Private m_VLargeChange   As Long
Private m_VSmallChange   As Long
Private m_VMax           As Long
Private m_VMin           As Long
Private m_VValue         As Long
Private m_PrevVValue     As Long

Private m_HLargeChange   As Long
Private m_HSmallChange   As Long
Private m_HMax           As Long
Private m_HMin           As Long
Private m_HValue         As Long
Private m_PrevHValue     As Long

Private m_ShowValue      As Boolean     'ONLY for Gauges
Private m_ScaleSteps     As Long        '0=NONE

'Internal variables
Private RTGadget         As RECT        'The Extremeties of the Gadget Itself
Private RTUp             As RECT        'The Up Arrow
Private RTDown           As RECT        'The Down Arrow
Private RTLeft           As RECT        'The Left Arrow
Private RTRight          As RECT        'The Right Arrow
Private RTTrack          As RECT        'The Track of the Slider
Private RTKnob           As RECT        'The Slider Knob

Private KnobH            As Long        'used for drawing knob
Private KnobW            As Long
Private KnobX            As Long        'used for Tracking knob
Private KnobY            As Long

Private HasFocus         As Boolean

Private Const MOUSEUP     As Long = 0   'Constants used when Tracking Keys or Mouse Actions
Private Const UPBUTTON    As Long = 1
Private Const DOWNBUTTON  As Long = 2
Private Const LEFTBUTTON  As Long = 3
Private Const RIGHTBUTTON As Long = 4
Private Const TRACK       As Long = 5
Private Const KNOB        As Long = 6
Private ButtonX           As Long
Private LastShift         As Integer
Private LastX             As Long
Private LastY             As Long

'Subclassing Variables
'mWndProcOrg holds the original address of the Window Procedure for this window. This is used to
'route messages to the original procedure after you process them.

Private mWndProcOrg      As Long
Private mHWndSubClassed  As Long                         'Handle (hWnd) of the subclassed window.

'Event Declarations:
'We are only interested in one thing: Did Mouse or Key Action on the PropGadget result in a change of values?
Event ValueChanged(ByVal NewVValue As Long, ByVal NewHValue As Long) 'The Only Event Notified

Private Sub UserControl_GotFocus()
  HasFocus = True
  Call Refresh
End Sub

Private Sub UserControl_LostFocus()
  HasFocus = False
  Call Refresh
End Sub

Private Sub UserControl_Paint()
 Dim RT As RECT
   
  With UserControl
    .Cls                                                      'Get Ready by clearing area
     
    Call SetRect(RT, 0, 0, .ScaleWidth, .ScaleHeight)
    
    If m_ScaleSteps <> 0 Then           'We will eventually Draw a Scale alongside
      If (m_PropGadgetType And VERT) = VERT Then RT.Right = RT.Right - SCALETICKSIZE
      If (m_PropGadgetType And HORZ) = HORZ Then RT.Bottom = RT.Bottom - SCALETICKSIZE
    End If
    
    If Not m_Borderless Then
      Call DrawGadgetBorder(RT)         'Get The OuterFrame Done
      Call InflateRect(RT, -2, -2)      'Shrink by 2 all round
    End If
       
    RTGadget = RT
    If m_ScrollingArrows = rvtScrollingArrows.None Then      'Gauges and Sliders NEVER have Arrows
      RTTrack = RT
    Else                                                     'We have the Scroller Family with Arrows
      Select Case m_PropGadgetType
        Case rvtPropGadgetType.VScroller:
          
          Select Case m_ScrollingArrows
            Case rvtScrollingArrows.AtTopAndBottom:
              Call SetRect(RTUp, RT.Left, RT.Top, RT.Right, RT.Top + 12)
              Call SetRect(RTDown, RT.Left, RT.Bottom - 12, RT.Right, RT.Bottom)
              Call SetRect(RTTrack, RT.Left, RTUp.Bottom + 1, RT.Right, RTDown.Top - 1)
             
            Case rvtScrollingArrows.AtTop:
              Call SetRect(RTUp, RT.Left, RT.Top, RT.Right, RTUp.Top + 12)
              Call SetRect(RTDown, RT.Left, RTUp.Bottom, RT.Right, RTUp.Bottom + 12)
              Call SetRect(RTTrack, RT.Left, RTDown.Bottom + 1, RT.Right, RT.Bottom)
            
            Case rvtScrollingArrows.AtBottom:
              Call SetRect(RTDown, RT.Left, RT.Bottom - 12, RT.Right, RT.Bottom)
              Call SetRect(RTUp, RT.Left, RTDown.Top - 13, RT.Right, RTDown.Top - 1)
              Call SetRect(RTTrack, RT.Left, RT.Top, RT.Right, RTUp.Top - 1)
          End Select
          Call DrawArrowButton(.hDC, RTUp, UPBUTTON)     'Up
          Call DrawArrowButton(.hDC, RTDown, DOWNBUTTON)   'Down
        
        Case rvtPropGadgetType.HScroller:
      
          Select Case m_ScrollingArrows
            Case rvtScrollingArrows.AtLeftAndRight:
              Call SetRect(RTLeft, RT.Left, RT.Top, RT.Left + 12, RT.Bottom)
              Call SetRect(RTRight, RT.Right - 12, RT.Top, RT.Right, RT.Bottom)
              Call SetRect(RTTrack, RTLeft.Right + 1, RT.Top, RTRight.Left - 1, RT.Bottom)
            
            Case rvtScrollingArrows.AtLeft:
              Call SetRect(RTLeft, RT.Left, RT.Top, RT.Left + 12, RT.Bottom)
              Call SetRect(RTRight, RTLeft.Right + 1, RT.Top, RTLeft.Right + 12, RT.Bottom)
              Call SetRect(RTTrack, RTRight.Right + 1, RT.Top, RT.Right, RT.Bottom)
            
            Case rvtScrollingArrows.AtRight:
              Call SetRect(RTRight, RT.Right - 12, RT.Top, RT.Right, RT.Bottom)
              Call SetRect(RTLeft, RTRight.Left - 13, RT.Top, RTRight.Left - 1, RT.Bottom)
              Call SetRect(RTTrack, RT.Left, RT.Top, RTLeft.Left - 1, RT.Bottom)
          End Select
          Call DrawArrowButton(.hDC, RTLeft, LEFTBUTTON)   'Left
          Call DrawArrowButton(.hDC, RTRight, RIGHTBUTTON)  'Right
        
        Case rvtPropGadgetType.HVScroller:                                 'Only case is AtEdges
          Call SetRect(RTUp, RT.Left + 12, RT.Top, RT.Right - 12, RT.Top + 12)
          Call SetRect(RTDown, RT.Left + 12, RT.Bottom - 12, RT.Right - 12, RT.Bottom)
          Call SetRect(RTLeft, RT.Left, RT.Top + 12, RT.Left + 12, RT.Bottom - 12)
          Call SetRect(RTRight, RT.Right - 12, RT.Top + 12, RT.Right, RT.Bottom - 12)
          Call SetRect(RTTrack, RTLeft.Right + 1, RTUp.Bottom + 1, RTRight.Left - 1, RTDown.Top - 1)
          
          Call DrawArrowButton(.hDC, RTUp, UPBUTTON)     'Up
          Call DrawArrowButton(.hDC, RTDown, DOWNBUTTON)   'Down
          Call DrawArrowButton(.hDC, RTLeft, LEFTBUTTON)   'Left
          Call DrawArrowButton(.hDC, RTRight, RIGHTBUTTON)  'Right
      End Select
    End If
    
    'We Draw The Track in a Patterned Brush to Make It Appear a little lighter than the Buttons
    Call DrawTrackAndKnob(.hDC)
    
    If m_ScaleSteps <> 0 Then Call DrawScaleTicks
  End With
End Sub

Private Sub DrawGadgetBorder(ByRef RT As RECT)
  Dim RTx As RECT
    
  Call DrawEdge(UserControl.hDC, RT, BDR_RAISEDINNER, BF_RECT)            '3D Look Only
  Call SetRect(RTx, RT.Left + 1, RT.Top + 1, RT.Right - 1, RT.Bottom - 1)
  Call DrawEdge(UserControl.hDC, RTx, BDR_SUNKENOUTER, BF_RECT)
End Sub

Private Sub DrawButtonNormal(ByRef RT As RECT)
  If UserControl.Enabled Then
    If m_ScrollingArrows <> rvtScrollingArrows.None Then
      Call DrawEdge(UserControl.hDC, RT, BDR_RAISEDINNER, BF_RECT)        'The Button unselected
    End If
  End If
End Sub

Private Sub DrawButtonSelected(ByRef RT As RECT)
  If UserControl.Enabled Then
    If m_ScrollingArrows <> rvtScrollingArrows.None Then
      Call DrawEdge(UserControl.hDC, RT, BDR_SUNKENOUTER, BF_RECT)       'The Button selected
    End If
  End If
End Sub

Private Sub DrawArrowButton(ByVal hDC As Long, RT As RECT, ByVal Which As Integer)  'Which  1=Up 2=Down 3=Left 4=Right
  Dim x0 As Long, y0 As Long, i As Long
  
  If UserControl.Enabled Then
    Call DrawButtonNormal(RT)                       'The Button
    x0 = RectCX(RT)
    y0 = RectCY(RT)
    
    Call SelectNewSolidPen(hDC, vb3DDKShadow)
    Select Case Which
      Case UPBUTTON:                                           'UP
        y0 = y0 - 3
        For i = 0 To 4
          Call MoveXY(hDC, x0 - i, y0 + i)
          Call LineTo(hDC, x0 + i, y0 + i)
        Next
        
      Case DOWNBUTTON:                                           'Down
        y0 = y0 + 2
        For i = 0 To 4
          Call MoveXY(hDC, x0 - i, y0 - i)
          Call LineTo(hDC, x0 + i, y0 - i)
        Next
          
      Case LEFTBUTTON:                                           'Left
        x0 = x0 - 3
        For i = 0 To 4
          Call MoveXY(hDC, x0 + i, y0 - i)
          Call LineTo(hDC, x0 + i, y0 + i)
        Next
      
      Case RIGHTBUTTON:                                           'Right
        x0 = x0 + 2
        For i = 0 To 4
          Call MoveXY(hDC, x0 - i, y0 - i)
          Call LineTo(hDC, x0 - i, y0 + i)
        Next
    End Select
    Call SelectPrevPen(hDC)
  End If
End Sub

Private Sub DrawTrackAndKnob(ByVal hDC As Long)
  Dim lsBR As Long, RTx As RECT, cx As Long, cy As Long, x As Long, y As Long, s As String, tmpC As Long
  
  'The Track is the Travelling space for the knob.
  'For Scrollers the entire track is rendered as a pattern over the entire space
  'For Sliders we render this with a thinner guide
  'For Gauges (display only) the KNOB represents the Value contained
   
  lsBR = NewSolidBrush(UserControl.BackColor)
  RTx = RTTrack
  Select Case m_PropGadgetType
    Case rvtPropGadgetType.VSlider:
      Call InflateRect(RTx, -5, 0)
      Call FillRect(hDC, RTTrack, lsBR)
      
    Case rvtPropGadgetType.HSlider:
      Call InflateRect(RTx, 0, -5)
      Call FillRect(hDC, RTTrack, lsBR)
  End Select
  
  tmpC = UserControl.ForeColor            'cache temporarily
  UserControl.ForeColor = vb3DLight       'Patterned Brush BackColor/vb3DLight
  Call PatternFillRect(hDC, RTx)          '8x8, 1 pixel on-off checkerboard
  UserControl.ForeColor = tmpC            'uncache
  
  If UserControl.Enabled Then             'Now Draw the Knob -  It needs to fit Proportionally inside RTTrack
    If (m_PropGadgetType And TYPEMASK) = GAUGE Then
      
      Select Case m_PropGadgetType
        Case rvtPropGadgetType.VGauge:
          cy = ScaledCYG(RTTrack)
          Call SetRect(RTKnob, RTTrack.Left, RTTrack.Bottom - (cy + 2), _
                               RTTrack.Right, RTTrack.Bottom)
          If cy > 1 Then
            Call FillRect(hDC, RTKnob, lsBR)
            Call DrawEdge(hDC, RTKnob, BDR_RAISEDINNER, BF_RECT)       'The Knob
          End If
          s = Str$(m_VValue): s = Right$(s, Len(s) - 1)
          
        Case rvtPropGadgetType.HGauge:
          cx = ScaledCXG(RTTrack)
          Call SetRect(RTKnob, RTTrack.Left, RTTrack.Top, _
                               RTTrack.Left + (cx + 2), RTTrack.Bottom)
          If cx > 1 Then
            Call FillRect(hDC, RTKnob, lsBR)
            Call DrawEdge(hDC, RTKnob, BDR_RAISEDINNER, BF_RECT)       'The Knob
          End If
          s = Str$(m_HValue): s = Right$(s, Len(s) - 1)
      End Select
      
      If m_ShowValue Then Call DrawText(hDC, s, -1, RTGadget, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
   Else
      Select Case m_PropGadgetType
          Case rvtPropGadgetType.VSlider:
          cy = ScaledCY(RTTrack)
          Call SetRect(RTKnob, RTTrack.Left, RTTrack.Bottom - (cy + KnobH), _
                               RTTrack.Right, RTTrack.Bottom - cy)
        
        Case rvtPropGadgetType.VScroller:
          cy = ScaledCY(RTTrack)
          Call SetRect(RTKnob, RTTrack.Left, RTTrack.Top + cy, _
                               RTTrack.Right, RTTrack.Top + cy + KnobH)
        
        Case rvtPropGadgetType.HScroller, rvtPropGadgetType.HSlider:
          cx = ScaledCX(RTTrack)
          Call SetRect(RTKnob, RTTrack.Left + cx, RTTrack.Top, _
                               RTTrack.Left + cx + KnobW, RTTrack.Bottom)
          
        Case rvtPropGadgetType.HVScroller
          cx = ScaledCX(RTTrack)
          cy = ScaledCY(RTTrack)
          Call SetRect(RTKnob, RTTrack.Left + cx, RTTrack.Top + cy, _
                               RTTrack.Left + cx + KnobW, RTTrack.Top + cy + KnobH)
      End Select
      
      If HasFocus Then Call FillRect(hDC, RTKnob, lsBR)          'SOLID if IT HAS Focus otherwise patterned
      Call DrawEdge(hDC, RTKnob, BDR_RAISEDINNER, BF_RECT)       'The Knob
      
      x = RectCX(RTKnob)                                         'and the little Knob Handle ;-)
      y = RectCY(RTKnob)
      
      Call DrawEllipse(hDC, x - 1, y - 1, x + 3, y + 3, vb3DHighlight)
      Call DrawEllipse(hDC, x - 2, y - 2, x + 2, y + 2, vb3DDKShadow)
    End If
  End If
  Call DeleteObject(lsBR)     'Kill SolidBrush
End Sub

Private Sub DrawScaleTicks()
  Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, OldFC As OLE_COLOR
  
  With UserControl
    If (m_PropGadgetType And VERT) = VERT Then
      x1 = .ScaleWidth - SCALETICKSIZE
      x2 = x1 + SCALETICKSIZE - 1
      For i = 0 To m_ScaleSteps
        y1 = RTTrack.Top + ((RTTrack.Bottom - RTTrack.Top - KnobH) * i) \ m_ScaleSteps
        Call MoveXY(.hDC, x1, y1): Call LineTo(.hDC, x2, y1)
      Next
    End If
  
    If (m_PropGadgetType And HORZ) = HORZ Then
      y1 = .ScaleHeight - SCALETICKSIZE
      y2 = y1 + SCALETICKSIZE - 1
      For i = 0 To m_ScaleSteps
        x1 = RTTrack.Left + ((RTTrack.Right - RTTrack.Left - KnobW) * i) \ m_ScaleSteps
        Call MoveXY(.hDC, x1, y1): Call LineTo(.hDC, x1, y2)
      Next
    End If
  End With
End Sub

Private Function ScaledCXG(RT As RECT) As Long                'for Gauges only From an HValue to an offset
  Dim w As Long, r As Long
  
  w = RT.Right - RT.Left
  r = m_HMax - m_HMin
  ScaledCXG = ((m_HValue - m_HMin) * (w - 2)) \ r
  KnobW = 0
End Function

Private Function ScaledCYG(RT As RECT) As Long                'for Gauges only From a VVAlue to an offset
  Dim h As Long, r As Long
  
  h = RT.Bottom - RT.Top
  r = m_VMax - m_VMin
  ScaledCYG = ((m_VValue - m_VMin) * (h - 2)) \ r
  KnobH = 0
End Function

Private Function ScaledCX(RT As RECT) As Long                     'From an HValue to an offset
  Dim w As Long, r As Long
  
  w = RT.Right - RT.Left
  r = m_HMax - m_HMin
  KnobW = 1 + 2 * (w \ r)
  If KnobW < 13 Then KnobW = 13
  ScaledCX = ((m_HValue - m_HMin) * (w - KnobW)) \ r
End Function

Private Function UnScaledCX(RT As RECT, ByVal x As Long) As Long  'From an offset to an HValue
  Dim w As Long, r As Long
  
  w = RT.Right - RT.Left
  r = m_HMax - m_HMin
  If x < RT.Left Then x = RT.Left
  If x > RT.Right - KnobW Then x = RT.Right - KnobW
  UnScaledCX = m_HMin + ((x - RT.Left) * r) / (w - KnobW)
End Function

Private Function ScaledCY(RT As RECT) As Long                     'From a VVAlue to an offset
  Dim h As Long, r As Long
  
  h = RT.Bottom - RT.Top
  r = m_VMax - m_VMin
  KnobH = 1 + 2 * (h \ r)
  If KnobH < 13 Then KnobH = 13
  ScaledCY = ((m_VValue - m_VMin) * (h - KnobH)) \ r
End Function

Private Function UnScaledCY(RT As RECT, ByVal y As Long) As Long  'From an offset to a VValue
  Dim h As Long, r As Long
  
  h = RT.Bottom - RT.Top
  r = m_VMax - m_VMin
  If y < RT.Top Then y = RT.Top
  If y > RT.Bottom - KnobH Then y = RT.Bottom - KnobH
  UnScaledCY = m_VMin + ((y - RT.Top) * r) / (h - KnobH)
End Function

Public Sub Refresh()        'Redraw from scratch
  Call UserControl.Refresh
End Sub

'============================   EVENT HANDLING ROUTINES ====================================================
'Gauges do not respond to events or generate any - they are display only
'
'For the others we can handle Keys and Mouse
'KEYS:
'   Scroll arrows move SmallChange amounts
'   if Shifted or Alted then largeChange amounts
'   if CTRLed then move the max range  (ie to and from extremes)
'
'MOUSE:
'   Hitting any arrow button moves SmallChange amounts (exactly as if a Key was Pressed (can be shifted))
'   Hitting the Track moves LargeChange amounts for Scrollers, but SmallChange for Sliders
'   Hitting and Dragging the knob moves the Knob in its track for variable amounts

'   IF Held down Keys and Mouse will repeat actions until released
'Here we go.....

'Hitting a key will depress the relevant ArrowButton (if it exists)
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If (m_PropGadgetType And TYPEMASK) <> GAUGE Then
    Select Case KeyCode
      Case vbKeyUp:
        Select Case m_PropGadgetType
          Case rvtPropGadgetType.VScroller, _
               rvtPropGadgetType.HVScroller, _
               rvtPropGadgetType.VSlider
            Call AdjUp(UPBUTTON, Shift): Call ValueChanged
        End Select
      
      Case vbKeyDown:
        Select Case m_PropGadgetType
          Case rvtPropGadgetType.VScroller, _
               rvtPropGadgetType.HVScroller, _
               rvtPropGadgetType.VSlider:
            Call AdjDown(DOWNBUTTON, Shift): Call ValueChanged
         End Select
      
      Case vbKeyLeft:
         Select Case m_PropGadgetType
          Case rvtPropGadgetType.HScroller, _
               rvtPropGadgetType.HVScroller, _
               rvtPropGadgetType.HSlider
            Call AdjLeft(LEFTBUTTON, Shift): Call ValueChanged
         End Select
      
      Case vbKeyRight:
         Select Case m_PropGadgetType
          Case rvtPropGadgetType.HScroller, _
               rvtPropGadgetType.HVScroller, _
               rvtPropGadgetType.HSlider:
            Call AdjRight(RIGHTBUTTON, Shift): Call ValueChanged
         End Select
    End Select
  End If
End Sub

Public Sub PretendMouseKey(ByVal n As Long)       'Triggered from MOUSEWHEEL EVENTS
  If (m_PropGadgetType And TYPEMASK) <> GAUGE Then
    If n < 0 Then
      Do While n < 0
        Select Case m_PropGadgetType
          Case rvtPropGadgetType.VScroller, _
               rvtPropGadgetType.HVScroller, _
               rvtPropGadgetType.VSlider
            Call AdjUp(UPBUTTON, 0): Call ValueChanged
          
         Case rvtPropGadgetType.HScroller, _
              rvtPropGadgetType.HVScroller, _
              rvtPropGadgetType.HSlider
           Call AdjLeft(LEFTBUTTON, 0): Call ValueChanged
        End Select
        n = n + 1
      Loop
    ElseIf n > 0 Then
      Do While n > 0
        Select Case m_PropGadgetType
          Case rvtPropGadgetType.VScroller, _
               rvtPropGadgetType.HVScroller, _
               rvtPropGadgetType.VSlider:
            Call AdjDown(DOWNBUTTON, 0): Call ValueChanged
         
         Case rvtPropGadgetType.HScroller, _
              rvtPropGadgetType.HVScroller, _
              rvtPropGadgetType.HSlider:
           Call AdjRight(RIGHTBUTTON, 0): Call ValueChanged
        End Select
        n = n - 1
      Loop
    End If
    Call Refresh
  End If
End Sub

'Releasing a Key fixes the rendering of the ArrowButtons (if it exists)
'We dont really do anything else...
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  If (m_PropGadgetType And TYPEMASK) <> GAUGE Then
    ButtonX = 0
    Call Refresh
  End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If (m_PropGadgetType And TYPEMASK) <> GAUGE Then
    LastX = x
    LastY = y
    If Button = vbLeftButton Then
      'First check if we are on the track or knob
      If OnButton(RTKnob, x, y) Then
        ButtonX = KNOB
        KnobX = x - RTKnob.Left
        KnobY = y - RTKnob.Top
        m_PrevVValue = m_VValue
        m_PrevHValue = m_HValue
        RaiseEvent ValueChanged(m_VValue, m_HValue)        'Always issue the first values even if current
        
      ElseIf OnButton(RTTrack, x, y) Then
        If (m_PropGadgetType And TYPEMASK) = SLIDER Then
          If (Shift And 5) > 0 Then Shift = 0    'Small Change or MaxChange
        Else
          Shift = (Shift And 2) + 1              'LargeChange or MaxChange
        End If
        If x < RTKnob.Left Then
          Call AdjLeft(TRACK, Shift)
        ElseIf x > RTKnob.Right Then
          Call AdjRight(TRACK, Shift)
        End If
        If y < RTKnob.Top Then
          Call AdjUp(TRACK, Shift)
        ElseIf y > RTKnob.Bottom Then
          Call AdjDown(TRACK, Shift)
        End If
        Call LogAction
      Else            'we must have hit a button (if there is one; there are none on a Slider)
                      'cant actually end up here if there are no buttons  ....
        Select Case m_PropGadgetType
          Case rvtPropGadgetType.VScroller:
            If OnButton(RTUp, x, y) Then
              Call AdjUp(UPBUTTON, Shift): Call LogAction
            ElseIf OnButton(RTDown, x, y) Then
              Call AdjDown(DOWNBUTTON, Shift): Call LogAction
            End If
            
          Case rvtPropGadgetType.HScroller:
            If OnButton(RTLeft, x, y) Then
              Call AdjLeft(LEFTBUTTON, Shift): Call LogAction
            ElseIf OnButton(RTRight, x, y) Then
              Call AdjRight(RIGHTBUTTON, Shift): Call LogAction
            End If
          
          Case rvtPropGadgetType.HVScroller:
            If OnButton(RTUp, x, y) Then
              Call AdjUp(UPBUTTON, Shift): Call LogAction
            ElseIf OnButton(RTDown, x, y) Then
              Call AdjDown(DOWNBUTTON, Shift): Call LogAction
            ElseIf OnButton(RTLeft, x, y) Then
              Call AdjLeft(LEFTBUTTON, Shift): Call LogAction
            ElseIf OnButton(RTRight, x, y) Then
              Call AdjRight(RIGHTBUTTON, Shift): Call LogAction
            End If
        End Select
      End If
    End If
  End If
End Sub

'Mouse move is only done for the knob, all others are ignored
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If (m_PropGadgetType And TYPEMASK) <> GAUGE Then
    If (Button And vbLeftButton) = vbLeftButton Then
      If ButtonX = KNOB Then
        Select Case m_PropGadgetType
          Case rvtPropGadgetType.VSlider:
            m_VValue = UnScaledCY(RTTrack, RTTrack.Bottom - KnobH - (y - KnobY))
            
          Case rvtPropGadgetType.VScroller:
            m_VValue = UnScaledCY(RTTrack, y - KnobY)
          
          Case rvtPropGadgetType.HScroller, _
               rvtPropGadgetType.HSlider:
            m_HValue = UnScaledCX(RTTrack, x - KnobX)
          
          Case rvtPropGadgetType.HVScroller:
            m_VValue = UnScaledCY(RTTrack, y - KnobY)
            m_HValue = UnScaledCX(RTTrack, x - KnobX)
        End Select
        Call ValueChanged
      End If
    End If
  End If
End Sub

Private Sub LogAction()     'setup the timer to do something while the mouse is down
  Call ValueChanged
  Timer1.Interval = 500
  Timer1.Enabled = True
End Sub

'The timer is only on for a mousedriven button hit and at this time its down
Private Sub Timer1_Timer()
  
  Select Case ButtonX
    Case UPBUTTON:      Call AdjUp(ButtonX, LastShift)
    Case DOWNBUTTON:    Call AdjDown(ButtonX, LastShift)
    Case LEFTBUTTON:    Call AdjLeft(ButtonX, LastShift)
    Case RIGHTBUTTON:   Call AdjRight(ButtonX, LastShift)
    Case TRACK:
      If LastX < RTKnob.Left Then
        Call AdjLeft(ButtonX, LastShift)
      ElseIf LastX > RTKnob.Right Then
        Call AdjRight(ButtonX, LastShift)
      End If
      If LastY < RTKnob.Top Then
        Call AdjUp(ButtonX, LastShift)
      ElseIf LastY > RTKnob.Bottom Then
        Call AdjDown(ButtonX, LastShift)
      End If
  End Select
  Call ValueChanged
  If Timer1.Interval > 125 Then Timer1.Interval = 125
End Sub

'Once again Releasing the mouse just stops the action . GO to Start, do not collect $200
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If (m_PropGadgetType And TYPEMASK) <> GAUGE Then
    Timer1.Enabled = False
    ButtonX = MOUSEUP
    Call Refresh      'Reset the imagery completely
  End If
End Sub

Private Function OnButton(ByRef RT As RECT, ByVal x As Single, ByVal y As Single) As Boolean
  If x >= RT.Left Then
    If x <= RT.Right Then
      If y >= RT.Top Then
        If y <= RT.Bottom Then
          OnButton = True
        End If
      End If
    End If
  End If
End Function

'We only get to these if we have hit something correct like a button or the right key
Private Sub AdjUp(ByVal Button As Long, ByVal Shift As Integer)
  ButtonX = Button
  LastShift = Shift
  Call DrawButtonSelected(RTUp)
  If (m_PropGadgetType And TYPEMASK) = SCROLLER Then
    If Shift And 5 Then
      Call Adj_VValue(-m_VLargeChange)
    ElseIf Shift And 2 Then
      Call Adj_VValue(-m_VMax)
    Else
      Call Adj_VValue(-m_VSmallChange)
    End If
  Else                          'opposite sense for VSLIDER
    If Shift And 5 Then
      Call Adj_VValue(m_VLargeChange)
    ElseIf Shift And 2 Then
      Call Adj_VValue(m_VMax)
    Else
      Call Adj_VValue(m_VSmallChange)
    End If
  End If
End Sub

Private Sub AdjDown(ByVal Button As Long, ByVal Shift As Integer)
  ButtonX = Button
  LastShift = Shift
  Call DrawButtonSelected(RTDown)
  If (m_PropGadgetType And TYPEMASK) = SCROLLER Then
    If Shift And 5 Then
      Call Adj_VValue(m_VLargeChange)
    ElseIf Shift And 2 Then
      Call Adj_VValue(m_VMax)
    Else
      Call Adj_VValue(m_VSmallChange)
    End If
  Else                          'opposite sense for VSLIDER
    If Shift And 5 Then
      Call Adj_VValue(-m_VLargeChange)
    ElseIf Shift And 2 Then
      Call Adj_VValue(-m_VMax)
    Else
      Call Adj_VValue(-m_VSmallChange)
    End If
  End If
End Sub

Private Sub AdjLeft(ByVal Button As Long, ByVal Shift As Integer)
  ButtonX = Button
  LastShift = Shift
  Call DrawButtonSelected(RTLeft)
  If Shift And 5 Then
    Call Adj_HValue(-m_HLargeChange)
  ElseIf Shift And 2 Then
    Call Adj_HValue(-m_HMax)
  Else
    Call Adj_HValue(-m_HSmallChange)
  End If
End Sub

Private Sub AdjRight(ByVal Button As Long, ByVal Shift As Integer)
  ButtonX = Button
  LastShift = Shift
  Call DrawButtonSelected(RTRight)
  If Shift And 5 Then
    Call Adj_HValue(m_HLargeChange)
  ElseIf Shift And 2 Then
    Call Adj_HValue(m_HMax)
  Else
    Call Adj_HValue(m_HSmallChange)
  End If
End Sub

Private Sub Adj_VValue(ByVal Delta As Long)
  m_PrevVValue = m_VValue
  m_VValue = m_VValue + Delta
  If m_VValue < m_VMin Then
    m_VValue = m_VMin
  ElseIf m_VValue > m_VMax Then
    m_VValue = m_VMax
  End If
End Sub

Private Sub Adj_HValue(ByVal Delta As Long)
  m_PrevHValue = m_HValue
  m_HValue = m_HValue + Delta
  If m_HValue < m_HMin Then
    m_HValue = m_HMin
  ElseIf m_HValue > m_HMax Then
    m_HValue = m_HMax
  End If
End Sub

Private Sub ValueChanged()
  If (m_PrevVValue <> m_VValue) Or (m_PrevHValue <> m_HValue) Then  'we did in fact change values
    Call DrawTrackAndKnob(UserControl.hDC)
    RaiseEvent ValueChanged(m_VValue, m_HValue)
  End If
End Sub

'================== INITIALISATION AND PROPERTY MANAGEMENT =================================================

Private Sub UserControl_Initialise()
  Call LocalInit
End Sub

Private Sub UserControl_Resize()
  Call Refresh
End Sub

Private Sub UserControl_Terminate()
  Timer1.Enabled = False
  Call UnSubClass
End Sub

Private Sub LocalInit()

  UserControl.ScaleMode = vbPixels
  UserControl.FillColor = UserControl.BackColor
  Timer1.Enabled = False
  Call SubClass
End Sub

Private Sub UserControl_InitProperties()
  
  UserControl.BackColor = Parent.BackColor
  UserControl.ForeColor = Parent.ForeColor
  Set UserControl.Font = Parent.Font
  
  m_PropGadgetType = rvtPropGadgetType.VScroller
  m_ScrollingArrows = rvtScrollingArrows.AtTopAndBottom
  m_Borderless = False
  
  m_VLargeChange = 10
  m_VSmallChange = 1
  m_VMax = 100
  m_VMin = 0
  m_VValue = 50
  
  m_HLargeChange = 10
  m_HSmallChange = 1
  m_HMax = 100
  m_HMin = 0
  m_HValue = 50
  
  m_ShowValue = False
  m_ScaleSteps = 0
  Call LocalInit
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BackColor = PropBag.ReadProperty("BackColor", Parent.BackColor)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", Parent.ForeColor)
  Set UserControl.Font = PropBag.ReadProperty("Font", Parent.Font)
  
  m_Borderless = PropBag.ReadProperty("Borderless", False)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  
  m_PropGadgetType = PropBag.ReadProperty("PropGadgetType", rvtPropGadgetType.VScroller)
  m_ScrollingArrows = PropBag.ReadProperty("ScrollingArrows", rvtScrollingArrows.AtTopAndBottom)
  
  m_VLargeChange = PropBag.ReadProperty("VLargeChange", 10)
  m_VSmallChange = PropBag.ReadProperty("VSmallChange", 1)
  m_VMax = PropBag.ReadProperty("VMax", 100)
  m_VMin = PropBag.ReadProperty("VMin", 0)
  m_VValue = PropBag.ReadProperty("VValue", 50)
  
  m_HLargeChange = PropBag.ReadProperty("HLargeChange", 10)
  m_HSmallChange = PropBag.ReadProperty("HSmallChange", 1)
  m_HMax = PropBag.ReadProperty("HMax", 100)
  m_HMin = PropBag.ReadProperty("HMin", 0)
  m_HValue = PropBag.ReadProperty("HValue", 50)
  
  m_ShowValue = PropBag.ReadProperty("ShowValue", False)
  m_ScaleSteps = PropBag.ReadProperty("ScaleSteps", 0)
  
  Call LocalInit
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, Parent.BackColor)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, Parent.ForeColor)
  Call PropBag.WriteProperty("Font", UserControl.Font, Parent.Font)
  
  Call PropBag.WriteProperty("Borderless", m_Borderless, False)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  
  Call PropBag.WriteProperty("PropGadgetType", m_PropGadgetType, rvtPropGadgetType.VScroller)
  Call PropBag.WriteProperty("ScrollingArrows", m_ScrollingArrows, rvtScrollingArrows.AtTopAndBottom)
  
  Call PropBag.WriteProperty("VLargeChange", m_VLargeChange, 10)
  Call PropBag.WriteProperty("VSmallChange", m_VSmallChange, 1)
  Call PropBag.WriteProperty("VMax", m_VMax, 100)
  Call PropBag.WriteProperty("VMin", m_VMin, 0)
  Call PropBag.WriteProperty("VValue", m_VValue, 50)
  
  Call PropBag.WriteProperty("HLargeChange", m_HLargeChange, 10)
  Call PropBag.WriteProperty("HSmallChange", m_HSmallChange, 1)
  Call PropBag.WriteProperty("HMax", m_HMax, 100)
  Call PropBag.WriteProperty("HMin", m_HMin, 0)
  Call PropBag.WriteProperty("HValue", m_HValue, 50)
  
  Call PropBag.WriteProperty("ShowValue", m_ShowValue, False)
  Call PropBag.WriteProperty("ScaleSteps", m_ScaleSteps, 0)
End Sub

    ' THE PROPERTIES
    
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor = New_BackColor
  PropertyChanged "BackColor"
  Call Refresh
End Property
    
Public Property Get ForeColor() As OLE_COLOR
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
  Call Refresh
End Property
    
Public Property Get Font() As Font
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
  Call Refresh
End Property
    
Public Property Get Borderless() As Boolean
  Borderless = m_Borderless
End Property

Public Property Let Borderless(ByVal New_Borderless As Boolean)
  m_Borderless = New_Borderless
  PropertyChanged "Borderless"
  Call Refresh
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled = New_Enabled
  PropertyChanged "Enabled"
  Call Refresh
End Property

Public Property Get VLargeChange() As Long
  VLargeChange = m_VLargeChange
End Property

Public Property Let VLargeChange(ByVal New_VLargeChange As Long)
  If New_VLargeChange > 0 _
  And New_VLargeChange > m_VSmallChange _
  And New_VLargeChange < (m_VMax - m_VMin) Then
    m_VLargeChange = New_VLargeChange
    PropertyChanged "VLargeChange"
  End If
End Property


Public Property Get VSmallChange() As Long
  VSmallChange = m_VSmallChange
End Property

Public Property Let VSmallChange(ByVal New_VSmallChange As Long)
  If New_VSmallChange > 0 And New_VSmallChange < m_VLargeChange Then
    m_VSmallChange = New_VSmallChange
    PropertyChanged "VSmallChange"
  End If
End Property


Public Property Get HLargeChange() As Long
  HLargeChange = m_HLargeChange
End Property
'
Public Property Let HLargeChange(ByVal New_HLargeChange As Long)
  If New_HLargeChange > 0 _
  And New_HLargeChange > m_HSmallChange _
  And New_HLargeChange < (m_HMax - m_HMin) Then
    m_HLargeChange = New_HLargeChange
    PropertyChanged "HLargeChange"
  End If
End Property


Public Property Get HSmallChange() As Long
  HSmallChange = m_HSmallChange
End Property

Public Property Let HSmallChange(ByVal New_HSmallChange As Long)
  If New_HSmallChange > 0 And New_HSmallChange < m_HLargeChange Then
    m_HSmallChange = New_HSmallChange
    PropertyChanged "HSmallChange"
  End If
End Property


Public Property Get VMax() As Long
  VMax = m_VMax
End Property
'
Public Property Let VMax(ByVal New_VMax As Long)
  If New_VMax > m_VMin And New_VMax <= 32767 Then
    m_VMax = New_VMax
    PropertyChanged "VMax"
    Call Refresh
  End If
End Property


Public Property Get VMin() As Long
  VMin = m_VMin
End Property

Public Property Let VMin(ByVal New_VMin As Long)
  If New_VMin >= 0 And New_VMin < m_VMax Then
    m_VMin = New_VMin
    PropertyChanged "VMin"
    Call Refresh
  End If
End Property


Public Property Get VValue() As Long
  VValue = m_VValue
End Property

Public Property Let VValue(ByVal New_VValue As Long)
  If New_VValue >= m_VMin And New_VValue <= m_VMax Then
    m_VValue = New_VValue
    PropertyChanged "VValue"
    Call Refresh
  End If
End Property


Public Property Get HMax() As Long
  HMax = m_HMax
End Property

Public Property Let HMax(ByVal New_HMax As Long)
  If New_HMax > m_HMin And New_HMax <= 32767 Then
    m_HMax = New_HMax
    PropertyChanged "HMax"
    Call Refresh
  End If
End Property


Public Property Get HMin() As Long
  HMin = m_HMin
End Property

Public Property Let HMin(ByVal New_HMin As Long)
  If New_HMin >= 0 And New_HMin < m_HMax Then
    m_HMin = New_HMin
    PropertyChanged "HMin"
    Call Refresh
  End If
End Property

Public Property Get HValue() As Long
  HValue = m_HValue
End Property

Public Property Let HValue(ByVal New_HValue As Long)
  If New_HValue >= m_HMin And New_HValue <= m_HMax Then
    m_HValue = New_HValue
    PropertyChanged "HValue"
    Call Refresh
  End If
End Property


Public Property Get ScaleSteps() As Long
  ScaleSteps = m_ScaleSteps
End Property

Public Property Let ScaleSteps(ByVal New_ScaleSteps As Long)
  If New_ScaleSteps >= 0 And New_ScaleSteps <= 32767 Then
    m_ScaleSteps = New_ScaleSteps
    PropertyChanged "ScaleSteps"
    Call Refresh
  End If
End Property


Public Property Get ShowValue() As Boolean                      'For Gauges and Sliders only
  ShowValue = m_ShowValue
End Property

Public Property Let ShowValue(ByVal New_ShowValue As Boolean)   'For Gauges and Sliders only
  m_ShowValue = New_ShowValue
  PropertyChanged "ShowValue"
  Call Refresh
End Property

Public Property Get PropGadgetType() As rvtPropGadgetType
  PropGadgetType = m_PropGadgetType
End Property

Public Property Let PropGadgetType(ByVal New_PropGadgetType As rvtPropGadgetType)
  
  m_PropGadgetType = New_PropGadgetType
  Call VerifyCombinations
  PropertyChanged "PropGadgetType"
  Call Refresh
End Property


Public Property Get ScrollingArrows() As rvtScrollingArrows
  ScrollingArrows = m_ScrollingArrows
End Property

Public Property Let ScrollingArrows(ByVal New_ScrollingArrows As rvtScrollingArrows)
  
  m_ScrollingArrows = New_ScrollingArrows
  Call VerifyCombinations
  PropertyChanged "ScrollingArrows"
  Call Refresh
End Property

Private Sub VerifyCombinations()
  Select Case m_PropGadgetType
    Case rvtPropGadgetType.VSlider, _
         rvtPropGadgetType.HSlider, _
         rvtPropGadgetType.VGauge, _
         rvtPropGadgetType.HGauge:
      If m_ScrollingArrows <> None Then m_ScrollingArrows = None
      
    Case rvtPropGadgetType.VScroller:
      Select Case m_ScrollingArrows
        Case AtLeftAndRight, AtLeft, AtRight, AtEdges: m_ScrollingArrows = AtTopAndBottom
      End Select
    
    Case rvtPropGadgetType.HScroller:
      Select Case m_ScrollingArrows
        Case AtTopAndBottom, AtTop, AtBottom, AtEdges: m_ScrollingArrows = AtLeftAndRight
      End Select
    
    Case rvtPropGadgetType.HVScroller:
      Select Case m_ScrollingArrows
        Case None, AtEdges:
        Case Else:
          m_ScrollingArrows = AtEdges
      End Select
  End Select
End Sub

'These Properties are automatically available for compatibility with normal MS Scrollbars
'2D gadgets wont be set properly

Public Property Let LargeChange(vdata As Long)
  If (m_PropGadgetType And VERT) = VERT Then
    VLargeChange = vdata
  Else
    HLargeChange = vdata
  End If
End Property

Public Property Get LargeChange() As Long
  If (m_PropGadgetType And VERT) = VERT Then
    LargeChange = VLargeChange
  Else
    LargeChange = HLargeChange
  End If
End Property

Public Property Let SmallChange(vdata As Long)
  If (m_PropGadgetType And VERT) = VERT Then
    VSmallChange = vdata
  Else
    HSmallChange = vdata
  End If
End Property

Public Property Get SmallChange() As Long
  If (m_PropGadgetType And VERT) = VERT Then
    SmallChange = VSmallChange
  Else
    SmallChange = HSmallChange
  End If
End Property

Public Property Let Max(vdata As Long)
  If (m_PropGadgetType And VERT) = VERT Then
    VMax = vdata
  Else
    HMax = vdata
  End If
End Property

Public Property Get Max() As Long
  If (m_PropGadgetType And VERT) = VERT Then
    Max = VMax
  Else
    Max = HMax
  End If
End Property

Public Property Let Min(vdata As Long)
  If (m_PropGadgetType And VERT) = VERT Then
    VMin = vdata
  Else
    HMin = vdata
  End If
End Property

Public Property Get Min() As Long
  If (m_PropGadgetType And VERT) = VERT Then
    Min = VMin
  Else
    Min = HMin
  End If
End Property

Public Property Let Value(vdata As Long)
  If (m_PropGadgetType And VERT) = VERT Then
    VValue = vdata
  Else
    HValue = vdata
  End If
End Property

Public Property Get Value() As Long
  If (m_PropGadgetType And VERT) = VERT Then
    Value = VValue
  Else
    Value = HValue
  End If
End Property

'=================================== SUBCLASSING CODE FOR USERCONTROL (Support MouseWheel on 98 and NT4)=========
'see also mMouseWheel.bas

Private Sub SubClass()
  '-------------------------------------------------------------
  'Initiates the subclassing of this UserControl's window (hwnd).
  'Records the original WinProc of the window in mWndProcOrg.
  'Places a pointer to the object in the window's UserData area.
  '-------------------------------------------------------------

  'Exit if the window is already subclassed.
  If mWndProcOrg <> 0 Then Exit Sub

  'Redirect the window's messages from this control's default Window Procedure to the SubWndProc function
  'in your .BAS module and record the address of the previous Window Procedure for this window in mWndProcOrg.
  mWndProcOrg = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubWndProc)

  'Record your window handle in case SetWindowLong gave you a new one.
  'You will need this handle so that you can unsubclass.
  mHWndSubClassed = hWnd

  'Store a pointer to this object in the UserData section of this window that will be used later to get
  'the pointer to the control based on the handle (hwnd) of the window getting the message.
  Call SetWindowLong(hWnd, GWL_USERDATA, ObjPtr(Me))
  
  'Get the Size of a Wheel Scroll in lines
  gucWheelScrollLines = SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, pulScrollLines, 0)
End Sub

Private Sub UnSubClass()
  '-----------------------------------------------------------------------------------------------
  'Unsubclasses this UserControl's window (hwnd), setting the address of the Windows Procedure
  'back to the address it was at before it was subclassed.
  '------------------------------------------------------------------------------------------------
  
  If mWndProcOrg = 0 Then Exit Sub  'Ensures that you don't try to unsubclass the window when it is not subclassed.
  SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcOrg     'Reset the window's function back to the original address.
  mWndProcOrg = 0                   '0 Indicates that you are no longer subclassed.
End Sub

Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '--------------------------------------------------------------
  'Process the window's messages that are sent to your UserControl. The WindowProc function is declared as
  'a "Friend" function so that the .BAS module can call the function but the function cannot be seen from
  'outside the UserControl project.
  '--------------------------------------------------------------

  'We are only intersetsed in picking up Mousewheel messages. We handle them as if the correct scroll key
  'had been repeatedly pressed by the user
  
  Dim ScrollAmt As Long
 
  Select Case uMsg
    Case WM_MOUSEWHEEL:
      If (wParam And (MK_SHIFT Or MK_CONTROL)) = 0 Then   ' Don't handle zoom and datazoom.
        
        gcWheelDelta = gcWheelDelta - (wParam And &HFFFF0000) / 65536
        If Abs(gcWheelDelta) >= WHEEL_DELTA Then
            
          ScrollAmt = gcWheelDelta / WHEEL_DELTA
    
          Do While gcWheelDelta < -WHEEL_DELTA
            gcWheelDelta = gcWheelDelta + WHEEL_DELTA
          Loop
          Do While gcWheelDelta > WHEEL_DELTA
            gcWheelDelta = gcWheelDelta - WHEEL_DELTA
          Loop
            
          Call PretendMouseKey(ScrollAmt)
        End If
      End If
  End Select
  
  'Forwards the window's messages that came in to the original Window Procedure that handles the messages
  'and returns the result back to the SubWndProc function.
  WindowProc = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, ByVal lParam)
End Function
