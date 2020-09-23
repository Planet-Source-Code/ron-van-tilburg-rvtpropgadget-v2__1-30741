VERSION 5.00
Begin VB.Form fTestProp 
   Caption         =   "rvtPropGadget - Gauges, Sliders, and Scrollers - All in one small control"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00DCDC9A&
      Caption         =   "HVScroller,  MS Scrollbars"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   8130
      TabIndex        =   21
      Top             =   120
      Width           =   2685
      Begin VB.VScrollBar VScroll1 
         Height          =   4005
         LargeChange     =   10
         Left            =   2100
         Max             =   100
         TabIndex        =   25
         Top             =   390
         Width           =   435
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   405
         LargeChange     =   10
         Left            =   120
         Max             =   100
         TabIndex        =   24
         Top             =   4530
         Width           =   2445
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   1425
         Index           =   10
         Left            =   360
         TabIndex        =   22
         ToolTipText     =   "HVScroller,Arrows=None,ScaleSteps=5"
         Top             =   570
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   2514
         BackColor       =   14474394
         PropGadgetType  =   67
         ScrollingArrows =   0
         ScaleSteps      =   5
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   1425
         Index           =   9
         Left            =   330
         TabIndex        =   23
         ToolTipText     =   "Arrows=AtEdges"
         Top             =   2340
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   2514
         BackColor       =   14474394
         PropGadgetType  =   67
         ScrollingArrows =   7
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "V and H Scrollers"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   5430
      TabIndex        =   14
      Top             =   120
      Width           =   2685
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   11
         Left            =   150
         TabIndex        =   15
         ToolTipText     =   "HScroller, ScrollingARrows=None"
         Top             =   2910
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   8438015
         PropGadgetType  =   66
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   6
         Left            =   150
         TabIndex        =   16
         ToolTipText     =   "VScroller, ScrollingArrows=None"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   8438015
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   7
         Left            =   1020
         TabIndex        =   17
         ToolTipText     =   "Borderless"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   8438015
         Borderless      =   -1  'True
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   8
         Left            =   1950
         TabIndex        =   18
         ToolTipText     =   "ScaleSteps=5,ScrollingArrows=Bottom"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   8438015
         ScrollingArrows =   3
         ScaleSteps      =   5
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   12
         Left            =   150
         TabIndex        =   19
         ToolTipText     =   "ScrollingArrows=LeftandRight"
         Top             =   3690
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   8438015
         PropGadgetType  =   66
         ScrollingArrows =   4
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   13
         Left            =   150
         TabIndex        =   20
         ToolTipText     =   "ScaleSteps=4,Arrows=AtLeft,Borderless"
         Top             =   4500
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   8438015
         PropGadgetType  =   66
         ScrollingArrows =   5
         ScaleSteps      =   4
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H001BAD6F&
      Caption         =   "V and H Sliders"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5265
      Left            =   2730
      TabIndex        =   7
      Top             =   90
      Width           =   2685
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   14
         Left            =   150
         TabIndex        =   8
         ToolTipText     =   "HSlider"
         Top             =   2910
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   1813871
         PropGadgetType  =   34
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   3
         Left            =   150
         TabIndex        =   9
         ToolTipText     =   "VSlider"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   1813871
         PropGadgetType  =   33
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   4
         Left            =   1020
         TabIndex        =   10
         ToolTipText     =   "Borderless=True"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   1813871
         Borderless      =   -1  'True
         PropGadgetType  =   33
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   5
         Left            =   1950
         TabIndex        =   11
         ToolTipText     =   "ScaleSteps=5"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   1813871
         Borderless      =   -1  'True
         PropGadgetType  =   33
         ScrollingArrows =   0
         ScaleSteps      =   5
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   15
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Borderless"
         Top             =   3690
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   1813871
         Borderless      =   -1  'True
         PropGadgetType  =   34
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   16
         Left            =   150
         TabIndex        =   13
         ToolTipText     =   "ScaleSteps=4"
         Top             =   4500
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   1813871
         PropGadgetType  =   34
         ScrollingArrows =   0
         ScaleSteps      =   4
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009C9CDE&
      Caption         =   "V and H Gauges"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   5265
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   2685
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   17
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "HGauge"
         Top             =   2910
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   10263774
         PropGadgetType  =   18
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   0
         Left            =   150
         TabIndex        =   1
         ToolTipText     =   "VGauge"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   10263774
         PropGadgetType  =   17
         ScrollingArrows =   0
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         ToolTipText     =   "ShowValue=True + Borderless=True"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   10263774
         Borderless      =   -1  'True
         PropGadgetType  =   17
         ScrollingArrows =   0
         ShowValue       =   -1  'True
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   2205
         Index           =   2
         Left            =   1950
         TabIndex        =   3
         ToolTipText     =   "ScaleSteps=5"
         Top             =   240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   3889
         BackColor       =   10263774
         ForeColor       =   -2147483634
         PropGadgetType  =   17
         ScrollingArrows =   0
         ShowValue       =   -1  'True
         ScaleSteps      =   5
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   18
         Left            =   150
         TabIndex        =   5
         ToolTipText     =   "ShowValue=True"
         Top             =   3690
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   10263774
         PropGadgetType  =   18
         ScrollingArrows =   0
         ShowValue       =   -1  'True
      End
      Begin TestProp.RVTPropGadget RVTProp 
         Height          =   495
         Index           =   19
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "ScaleSteps=4"
         Top             =   4500
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   10263774
         ForeColor       =   -2147483634
         PropGadgetType  =   18
         ScrollingArrows =   0
         ShowValue       =   -1  'True
         ScaleSteps      =   4
      End
   End
End
Attribute VB_Name = "fTestProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HScroll1_Change()
  Call RVTProp_ValueChanged(19, 0, HScroll1.Value)
End Sub

Private Sub RVTProp_ValueChanged(Index As Integer, ByVal NewVValue As Long, ByVal NewHValue As Long)
  Dim i As Long
  
  i = Index
  If Index >= 0 And Index <= 10 Then
    VScroll1.Value = NewVValue
    For i = 0 To 10
      RVTProp(i).VValue = NewVValue
    Next
  End If
  
  If Index >= 9 And Index <= 19 Then
    HScroll1.Value = NewHValue
    For i = 9 To 19
      RVTProp(i).HValue = NewHValue
    Next
  End If
End Sub

Private Sub VScroll1_Change()
  Call RVTProp_ValueChanged(0, VScroll1.Value, 0)
End Sub
