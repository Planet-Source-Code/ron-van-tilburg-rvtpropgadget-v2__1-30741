Attribute VB_Name = "mMouseWheel"
Option Explicit

'Fairly standard Usercontol Subclassing as described in MS HOWTO: Subclass a UserControl Q179398
'Adapted only enough to do Moosewheel handling

'API Declarations used for subclassing.
Public Declare Sub CopyMemory _
   Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Public Declare Function SetWindowLong _
   Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong _
   Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function CallWindowProc _
   Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
                                         ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'For mousewheel values
Public Declare Function SystemParametersInfo _
   Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
                                               ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

'Constants for GetWindowLong() and SetWindowLong() APIs.
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

'MOUSEWHEEL CONSTANTS
Public Const WM_MOUSEWHEEL As Long = &H20A

' Key State Masks for Mouse Messages
Public Const MK_LBUTTON As Integer = &H1
Public Const MK_RBUTTON As Integer = &H2
Public Const MK_SHIFT   As Integer = &H4
Public Const MK_CONTROL As Integer = &H8
Public Const MK_MBUTTON As Integer = &H10

Public Const WHEEL_DELTA As Long = 120
Public Const SPI_GETWHEELSCROLLLINES As Long = 104

'MouseWheel Variables

Public gcWheelDelta        As Integer      'wheel delta from roll
Public gucWheelScrollLines As Integer      'number of lines to scroll on a wheel rotation
Public pulScrollLines      As Long

'Used to hold a reference to the control to call its procedure.
'NOTE: "UserControl1" is the UserControl.Name Property at design-time of the .CTL file.
'      ('As Object' or 'As Control' does not work)
Dim ctlShadowControl As RVTPropGadget
Dim ptrObject As Long                     'Used as a pointer to the UserData section of a window.

'The address of this function is used for subclassing. Messages will be sent here and then forwarded to the
'UserControl's WindowProc function. The HWND determines to which control the message is sent.

Public Function SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long

  On Error Resume Next

  'Get pointer to the control's VTable from the window's UserData section. The VTable is an internal
  'structure that contains pointers to the methods and properties of the control.
  ptrObject = GetWindowLong(hWnd, GWL_USERDATA)

  'Copy the memory that points to the VTable of our original control to the shadow copy of the control you use to
  'call the original control's WindowProc Function. This way, when you call the method of the shadow control,
  'you are actually calling the original controls' method.
  CopyMemory ctlShadowControl, ptrObject, 4

  'Call the WindowProc function in the instance of the UserControl.
  SubWndProc = ctlShadowControl.WindowProc(hWnd, Msg, wParam, lParam)

  'Destroy the Shadow Control Copy
  CopyMemory ctlShadowControl, 0&, 4
  Set ctlShadowControl = Nothing
  
  On Error GoTo 0
End Function



