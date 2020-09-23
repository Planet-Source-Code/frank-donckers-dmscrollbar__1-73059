VERSION 5.00
Begin VB.UserControl DMGraphicalScrollBar 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   PropertyPages   =   "DMGraphicalScrollBar.ctx":0000
   ScaleHeight     =   4020
   ScaleWidth      =   7770
   Begin VB.PictureBox picTmpHor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2400
      Picture         =   "DMGraphicalScrollBar.ctx":0021
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2400
      Picture         =   "DMGraphicalScrollBar.ctx":0A2D
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1620
      Top             =   1320
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2100
      Top             =   1320
   End
   Begin VB.Timer tmrRight 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2100
      Top             =   1800
   End
   Begin VB.Timer tmrLeft 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1620
      Top             =   1800
   End
   Begin VB.Timer Tracking2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3540
      Top             =   1980
   End
   Begin VB.Timer Tracking1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3060
      Top             =   1980
   End
   Begin VB.Timer Tracking3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4020
      Top             =   1980
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   180
      ScaleHeight     =   3015
      ScaleWidth      =   780
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   780
      Begin VB.PictureBox picScroller 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "DMGraphicalScrollBar.ctx":1439
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   1140
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picRight 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "DMGraphicalScrollBar.ctx":19BB
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picLeft 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "DMGraphicalScrollBar.ctx":1F3D
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   4
         Top             =   2340
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picDN 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "DMGraphicalScrollBar.ctx":24BF
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   2
         Top             =   2670
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picUP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "DMGraphicalScrollBar.ctx":2A41
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   315
      End
   End
End
Attribute VB_Name = "DMGraphicalScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
' Programmer:        Donckers Frank
'                    DarkManSoft@Gmail.com
'
' Description:       Control Custom Scrollbar

'
'=====================================================
' POINTAPI
'=====================================================
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

'=====================================================
' Events
'=====================================================
Public Event Change()
Public Event Scroll()
Public Event DblClick()

'=====================================================
' Enum BorderStyles
'=====================================================
Public Enum BorderStyles
    [None] = 0
    [Fixed Single] = 1
End Enum

'=====================================================
' Enum ArrowStyles
'=====================================================
Public Enum ArrowStyles
    [ArrowRaised] = 0
    [ArrowEngraved] = 1
    [TriangleRaised] = 2
    [TriangleEngraved] = 3
    [WafelsRaised] = 4
    [WafelsEngraved] = 5
    [CirclesRaised] = 6
    [CirclesEngraved] = 7
End Enum

'=====================================================
' Enum GripperStyles
'=====================================================
Public Enum GripperStyles
    [NoGripper] = 0
    [DotEngraved] = 1
    [DotRaised] = 2
    [LineEngraved] = 3
    [LineRaised] = 4
    [BoxEngraved] = 5
    [BoxRaised] = 6
    [WafelEngraved] = 7
    [WafelRaised] = 8
    [DiamondEngraved] = 9
    [DiamondRaised] = 10
    [CircleEngraved] = 11
    [CircleRaised] = 12
End Enum

'=====================================================
' Enum ArrowTypes
'=====================================================
Public Enum ArrowTypes
    [ArrowUp] = 0
    [ArrowDown] = 1
    [ArrowLeft] = 2
    [ArrowRight] = 3
End Enum

'=====================================================
' Enum Orientations
'=====================================================
Public Enum Orientations
    [Horizontal] = 0
    [Vertical] = 1
End Enum

'=====================================================
' Enum Orientations
'=====================================================
Public Enum Styles
    [Graphical] = 0
    [Flat] = 1
    [UserGraphics] = 2
End Enum

'=====================================================
' Default Property Values
'=====================================================
' Bar
Const m_def_BarBorderColor = &H808080
Const m_def_BarBackColor = &HC0C0C0
Const m_def_BarBorderStyle = 1
' Buttons
Const m_def_ButtonsHeightMax = 315
Const m_def_ButtonsWidthMax = 315
Const m_def_ButtonsArrowStyle = 1
Const m_def_ButtonsBorderStyle = 1
Const m_def_ButtonsBackColor = &H808080
Const m_def_ButtonsBorderColor = &H808080
Const m_def_ButtonsArrowColor = &H0&
' Scroller
Const m_def_ScrollerBackColor = &H808080
Const m_def_ScrollerBorderColor = &H808080
Const m_def_ScrollerGripColor = &H404040
Const m_def_ScrollerGripperStyle = 0
Const m_def_ScrollerBorderStyle = 1
Const m_def_ScrollerHeightMax = 315
Const m_def_ScrollerWidthMax = 315
Const m_def_ScrollInterval = 100
' Disabled
Const m_def_DisabledBackColor = &HC0FFFF
Const m_def_DisabledBorderColor = &H81BECB
Const m_def_Locked = False
' Values
Const m_def_Value = 0
Const m_def_MinValue = 0
Const m_def_MaxValue = 10
Const m_def_SmallChange = 1
Const m_def_LargeChange = 2
' Looks
Const m_def_Style = 0
Const m_def_ButtonsVisible = True
Const m_def_ToolTipText = ""
Const m_def_Orientation = 0
Const m_def_Enabled = True

'=====================================================
' Property Variables
'=====================================================
' Bar
Dim m_BarBorderColor                    As OLE_COLOR
Dim m_BarBackColor                      As OLE_COLOR
Dim m_PicBackVertical                   As Picture
Dim m_PicBackVerticalDiasabled          As Picture
Dim m_PicBackHorizontal                 As Picture
Dim m_PicBackHorizontalDisabled         As Picture
Dim m_PicBack                           As Picture 'xxxx
Dim m_BarBorderStyle                    As BorderStyles
' Buttons
Dim m_ButtonsBackColor                  As OLE_COLOR
Dim m_ButtonsBorderColor                As OLE_COLOR
Dim m_ButtonsArrowColor                 As OLE_COLOR
Dim m_ButtonsHeightMax                  As Long
Dim m_ButtonsWidthMax                   As Long
Dim m_ButtonsArrowStyle                 As ArrowStyles
Dim m_ButtonsBorderStyle                As BorderStyles
Dim m_PicButtonTop_UP                   As Picture
Dim m_PicButtonTop_DOWN                 As Picture
Dim m_PicButtonTop_DISABLED             As Picture
Dim m_PicButtonTop_HOOVER               As Picture
Dim m_PicButtonBottom_UP                As Picture
Dim m_PicButtonBottom_DOWN              As Picture
Dim m_PicButtonBottom_DISABLED          As Picture
Dim m_PicButtonBottom_HOOVER            As Picture
Dim m_PicButtonLeft_UP                  As Picture
Dim m_PicButtonLeft_DOWN                As Picture
Dim m_PicButtonLeft_DISABLED            As Picture
Dim m_PicButtonLeft_HOOVER              As Picture
Dim m_PicButtonRight_UP                 As Picture
Dim m_PicButtonRight_DOWN               As Picture
Dim m_PicButtonRight_DISABLED           As Picture
Dim m_PicButtonRight_HOOVER             As Picture
' Scroller
Dim m_ScrollerBackColor                 As OLE_COLOR
Dim m_ScrollerBorderColor               As OLE_COLOR
Dim m_ScrollerGripColor                 As OLE_COLOR
Dim m_ScrollerGripperStyle              As GripperStyles
Dim m_ScrollerBorderStyle               As BorderStyles
Dim m_ScrollerHeightMax                 As Long
Dim m_ScrollerWidthMax                  As Long
Dim m_PicScrollerVertical_UP            As Picture
Dim m_PicScrollerVertical_DOWN          As Picture
Dim m_PicScrollerVertical_HOOVER        As Picture
Dim m_PicScrollerVertical_DISABLED      As Picture
Dim m_PicScrollerHorizontal_UP          As Picture
Dim m_PicScrollerHorizontal_DOWN        As Picture
Dim m_PicScrollerHorizontal_DISABLED    As Picture
Dim m_PicScrollerHorizontal_HOOVER      As Picture
Dim m_PicScroller_UP                    As Picture 'xxxx
Dim m_PicScroller_DOWN                  As Picture 'xxxx
Dim m_PicScroller_DISABLED              As Picture 'xxxx
Dim m_PicScroller_HOOVER                As Picture 'xxxx
Dim m_ScrollInterval                    As Integer
' Disabled
Dim m_DisabledBackColor                 As OLE_COLOR
Dim m_DisabledBorderColor               As OLE_COLOR
Dim m_Locked                            As Boolean
' Values
Dim m_Value                             As Long
Dim m_MaxValue                          As Long
Dim m_MinValue                          As Long
Dim m_SmallChange                       As Long
Dim m_LargeChange                       As Long
' Looks
Dim m_Style                             As Styles
Dim m_ButtonsVisible                    As Boolean
Dim m_ToolTipText                       As String
Dim m_Orientation                       As Orientations
Dim m_Enabled                           As Boolean


'=====================================================
' Misc Variables
'=====================================================
Private bAddedToIDE        As Boolean
Private Const nMaxValue    As Double = 2147483647
'// Variable for scrolling
Private m_MouseX           As Long
Private m_MouseY           As Long
Private m_Sliding          As Boolean
Private m_OldPosn          As Long
Private m_ValueChanged     As Boolean
Private OldScaleMode As Byte
Private cControl As Control
Private StartCol As Double, EndCol As Double
Private RedI As Single, BlueI As Single, GreenI As Single
Private RedStart As Integer, GreenStart As Integer, BlueStart As Integer
Private RedEnd As Double, GreenEnd As Double, BlueEnd As Double
Private i, ii, iii As Integer
Private NewColor As Single
Private MidX, MidY As Long
Private MouseDnUp, MouseDnDown, MouseDnLeft, MouseDnRight As Boolean
'Property Variables:


'=====================================================
' Draw the scrollbar
'=====================================================
Private Sub DrawTheBar()
    On Error Resume Next
    ' You could make a scrollbar that has only 2 buttons and 1 scroller
    ' and just change the possitions of them,
    ' but, becouse horizontal buttons mostly have there shades, etc...
    ' different then vertical, different pictures are used in this
    ' scrollbar to hold it a bit simple and you don't always have to
    ' switch the pictures or images.
    ' The program also switches the back/scroller pictures to horizontal/vertical,
    ' so when you make 1 scrollbar with all pictures set you don't need to
    ' switch the pictures when you switch the orientation
    SetPictureVisability
    If m_Style = 2 Then
        If m_Orientation = Vertical Then
            If Not m_PicBackVertical Is Nothing Then
                Set m_PicBack = m_PicBackVertical
            Else
                m_PicBack.Picture = Nothing
            End If
            If Not m_PicScrollerVertical_UP Is Nothing Then
                Set m_PicScroller_UP = m_PicScrollerVertical_UP
            Else
                Set m_PicScroller_UP = Nothing
            End If
            If Not m_PicScrollerVertical_DOWN Is Nothing Then
                Set m_PicScroller_DOWN = m_PicScrollerVertical_DOWN
            Else
                Set m_PicScroller_DOWN = Nothing
            End If
            If Not m_PicScrollerVertical_HOOVER Is Nothing Then
                Set m_PicScroller_HOOVER = m_PicScrollerVertical_HOOVER
            Else
                Set m_PicScroller_HOOVER = Nothing
            End If
            If Not m_PicButtonTop_UP Is Nothing Then
                picUP.PaintPicture m_PicButtonTop_UP, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
            Else
                picUP.Picture = Nothing
            End If
            If Not m_PicButtonBottom_UP Is Nothing Then
                picDN.PaintPicture m_PicButtonBottom_UP, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
            Else
                picDN.Picture = Nothing
            End If
       Else
            If Not m_PicBackHorizontal Is Nothing Then
                Set m_PicBack = m_PicBackHorizontal
            Else
                Set m_PicBack = Nothing
            End If
            If Not m_PicScrollerHorizontal_UP Is Nothing Then
                Set m_PicScroller_UP = m_PicScrollerHorizontal_UP
            Else
                Set m_PicScroller_UP = Nothing
            End If
            If Not m_PicScrollerHorizontal_DOWN Is Nothing Then
                Set m_PicScroller_DOWN = m_PicScrollerHorizontal_DOWN
            Else
                Set m_PicScroller_DOWN = Nothing
            End If
            If Not m_PicScrollerHorizontal_HOOVER Is Nothing Then
                Set m_PicScroller_HOOVER = m_PicScrollerHorizontal_HOOVER
            Else
                Set m_PicScroller_HOOVER = Nothing
            End If
            If Not m_PicButtonLeft_UP Is Nothing Then
                picLeft.PaintPicture m_PicButtonLeft_UP, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
            Else
                picLeft.Picture = Nothing
            End If
            If Not m_PicButtonRight_UP Is Nothing Then
                picRight.PaintPicture m_PicButtonRight_UP, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
            Else
                picRight.Picture = Nothing
            End If
        End If
        If Not m_PicBack Is Nothing Then
            picBG.PaintPicture m_PicBack, 0, 0, picBG.ScaleWidth, picBG.ScaleHeight
        Else
            picBG.Picture = Nothing
        End If
        If Not m_PicScroller_UP Is Nothing Then
            picScroller.PaintPicture m_PicScroller_UP, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
        Else
            picScroller.Picture = Nothing
        End If
    Else
        If m_Orientation = Vertical Then
            ' Show / Hide buttons
            picUP.Visible = True
            picDN.Visible = True
            picRight.Visible = False
            picLeft.Visible = False
            picUP.Picture = LoadPicture("")
            picDN.Picture = LoadPicture("")
            picScroller.Picture = LoadPicture("")
            picBG.Picture = LoadPicture("")
            DrawPicBack m_BarBackColor, ShiftColors(m_BarBackColor, 170), m_BarBorderColor
            DrawButton picUP, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowUp, m_ButtonsArrowColor, m_ButtonsBorderColor
            DrawButton picDN, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowDown, m_ButtonsArrowColor, m_ButtonsBorderColor
            DrawScroller m_ScrollerBackColor, ShiftColors(m_ScrollerBackColor, 170), m_ScrollerBorderColor, m_ScrollerGripColor
        Else
            ' Show / Hide buttons
            picUP.Visible = False
            picDN.Visible = False
            picRight.Visible = True
            picLeft.Visible = True
            picLeft.Picture = LoadPicture("")
            picRight.Picture = LoadPicture("")
            picScroller.Picture = LoadPicture("")
            picBG.Picture = LoadPicture("")
            DrawPicBack m_BarBackColor, ShiftColors(m_BarBackColor, 170), m_BarBorderColor
            DrawButton picLeft, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowLeft, m_ButtonsArrowColor, m_ButtonsBorderColor
            DrawButton picRight, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowRight, m_ButtonsArrowColor, m_ButtonsBorderColor
            DrawScroller m_ScrollerBackColor, ShiftColors(m_ScrollerBackColor, 170), m_ScrollerBorderColor, m_ScrollerGripColor
        End If
    End If
End Sub

'=====================================================
' Draw the scrollbar disabled
'=====================================================
Private Sub DrawTheBarDisabled()
    ' You could make a scrollbar that has only 2 buttons and 1 scroller
    ' and just change the possitions of them,
    ' but, becouse horizontal buttons mostly have there shades, etc...
    ' different then vertical, different pictures are used in this
    ' scrollbar to hold it a bit simple and you don't always have to
    ' switch the pictures or images.
    On Error Resume Next
    SetPictureVisability
    If m_Style = 2 Then
        If m_Orientation = Vertical Then
            If Not m_PicBackVerticalDiasabled Is Nothing Then
                Set m_PicBack = m_PicBackVerticalDiasabled
            Else
                Set m_PicBack = Nothing
            End If
            If Not m_PicScrollerVertical_DISABLED Is Nothing Then
                Set m_PicScroller_UP = m_PicScrollerVertical_DISABLED
            Else
                Set m_PicScroller_UP = Nothing
            End If
            If Not m_PicButtonTop_DISABLED Is Nothing Then
                picUP.PaintPicture m_PicButtonTop_DISABLED, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
            Else
                picUP.Picture = Nothing
            End If
            If Not m_PicButtonBottom_DISABLED Is Nothing Then
                picDN.PaintPicture m_PicButtonBottom_DISABLED, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
            Else
                picDN.Picture = Nothing
            End If
        Else
            If Not m_PicBackHorizontalDisabled Is Nothing Then
                Set m_PicBack = m_PicBackHorizontalDisabled
            Else
                Set m_PicBack = Nothing
            End If
            If Not m_PicScrollerHorizontal_DISABLED Is Nothing Then
                Set m_PicScroller_UP = m_PicScrollerHorizontal_DISABLED
            Else
                Set m_PicScroller_UP = Nothing
            End If
            If Not m_PicButtonLeft_DISABLED Is Nothing Then
                picLeft.PaintPicture m_PicButtonLeft_DISABLED, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
            Else
                picLeft.Picture = Nothing
            End If
            If Not m_PicButtonRight_DISABLED Is Nothing Then
                picRight.PaintPicture m_PicButtonRight_DISABLED, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
            Else
                picRight.Picture = Nothing
            End If
        End If
        If Not m_PicBack Is Nothing Then
            picBG.PaintPicture m_PicBack, 0, 0, picBG.ScaleWidth, picBG.ScaleHeight
        Else
            picBG.Picture = Nothing
        End If
        If Not m_PicScroller_UP Is Nothing Then
            picScroller.PaintPicture m_PicScroller_UP, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
        Else
            picScroller.Picture = Nothing
        End If
    Else
        If m_Orientation = Vertical Then
            '// Show / Hide PictureBoxes
            picUP.Visible = True
            picDN.Visible = True
            picRight.Visible = False
            picLeft.Visible = False
            picUP.Picture = LoadPicture("")
            picDN.Picture = LoadPicture("")
            picScroller.Picture = LoadPicture("")
            picBG.Picture = LoadPicture("")
            DrawPicBack m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), m_DisabledBorderColor
            DrawButton picUP, m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), ArrowUp, m_DisabledBorderColor, m_DisabledBorderColor
            DrawButton picDN, m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), ArrowDown, m_DisabledBorderColor, m_DisabledBorderColor
            DrawScroller m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), m_DisabledBorderColor, m_DisabledBorderColor
        Else
            '// Show / Hide PictureBoxes
            picUP.Visible = False
            picDN.Visible = False
            picRight.Visible = True
            picLeft.Visible = True
            picLeft.Picture = LoadPicture("")
            picRight.Picture = LoadPicture("")
            picScroller.Picture = LoadPicture("")
            picBG.Picture = LoadPicture("")
            DrawPicBack m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), m_DisabledBorderColor
            DrawButton picLeft, m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), ArrowLeft, m_DisabledBorderColor, m_DisabledBorderColor
            DrawButton picRight, m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), ArrowRight, m_DisabledBorderColor, m_DisabledBorderColor
            DrawScroller m_DisabledBackColor, ShiftColors(m_DisabledBackColor, 170), m_DisabledBorderColor, m_DisabledBorderColor
        End If
    End If
End Sub

Private Sub Command1_Click()
CheckBackHor
End Sub

Private Sub picBG_DblClick()
    RaiseEvent DblClick
End Sub

'=====================================================
' Background
'=====================================================

Private Sub picBG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_Orientation = Vertical Then
      If y > picScroller.Top Then
         ' clicked below the scroller
         Value = m_Value + m_LargeChange
      Else
         ' clicked above the scroller
         Value = m_Value - m_LargeChange
      End If
   Else
      If x > picScroller.Left Then
         ' clicked right of the scroller
         Value = m_Value + m_LargeChange
      Else
         ' clicked left of the scroller
         Value = m_Value - m_LargeChange
      End If
   End If
End Sub


'=====================================================
' Button Up
'=====================================================
Private Sub picUP_Click()
   '// Decr the Value by SmallChange Property
   Value = m_Value - m_SmallChange
   MouseDnUp = False
End Sub

Private Sub picUP_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picUP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    MouseDnUp = True
    If m_Style = 0 Then
       DrawButton picUP, ShiftColors(m_ButtonsBackColor, 170), m_ButtonsBackColor, ArrowUp, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 1 Then
       DrawButton picUP, ShiftColors(m_ButtonsBackColor, -20), ShiftColors(m_ButtonsBackColor, -20), ArrowUp, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 2 Then
        If Not m_PicButtonTop_DOWN Is Nothing Then
            picUP.PaintPicture m_PicButtonTop_DOWN, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
        Else
            picUP.PaintPicture m_PicButtonTop_UP, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
        End If
    End If
    tmrUp.Enabled = True
End Sub

Private Sub picUP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    MouseDnUp = False
    If m_Style = 0 Or m_Style = 1 Then
       If CheckMouseOver(picUP.hWnd) = True Then
           DrawButton picUP, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowUp, m_ButtonsArrowColor, m_ButtonsBorderColor
       Else
           DrawButton picUP, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowUp, m_ButtonsArrowColor, m_ButtonsBorderColor
       End If
    Else
       If CheckMouseOver(picUP.hWnd) = True Then
           If Not m_PicButtonTop_HOOVER Is Nothing Then
               picUP.PaintPicture m_PicButtonTop_HOOVER, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
           Else
               If Not m_PicButtonTop_UP Is Nothing Then
                   picUP.PaintPicture m_PicButtonTop_UP, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
               Else
                   picUP.Picture = Nothing
               End If
           End If
       Else
           If Not m_PicButtonTop_UP Is Nothing Then
               picUP.PaintPicture m_PicButtonTop_UP, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
           Else
               picUP.Picture = Nothing
           End If
       End If
    End If
    tmrUp.Enabled = False
End Sub

Private Sub picUP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If MouseDnUp = True Then Exit Sub
    If CheckMouseOver(picUP.hWnd) = True Then
        If x > 60 And x < picUP.Width - 60 And y > 60 And y < picUP.Height - 60 Then
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picUP, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowUp, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonTop_HOOVER Is Nothing Then
                    picUP.PaintPicture m_PicButtonTop_HOOVER, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
                Else
                    If Not m_PicButtonTop_UP Is Nothing Then
                        picUP.PaintPicture m_PicButtonTop_UP, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
                    Else
                        picUP.Picture = Nothing
                    End If
                End If
            End If
        Else
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picUP, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowUp, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonTop_UP Is Nothing Then
                    picUP.PaintPicture m_PicButtonTop_UP, 0, 0, picUP.ScaleWidth, picUP.ScaleHeight
                Else
                    picUP.Picture = Nothing
                End If
            End If
        End If
    End If
End Sub

Private Sub tmrUp_Timer()
   Value = m_Value - m_SmallChange
End Sub


'=====================================================
' Button Down
'=====================================================
Private Sub picDN_Click()
    Value = m_Value + m_SmallChange
    MouseDnDown = False
End Sub

Private Sub picDN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    MouseDnDown = True
    If m_Style = 0 Then
        DrawButton picDN, ShiftColors(m_ButtonsBackColor, 170), m_ButtonsBackColor, ArrowDown, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 1 Then
        DrawButton picDN, ShiftColors(m_ButtonsBackColor, -20), ShiftColors(m_ButtonsBackColor, -20), ArrowDown, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 2 Then
       If Not m_PicButtonBottom_DOWN Is Nothing Then
           picDN.PaintPicture m_PicButtonBottom_DOWN, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
       Else
           picDN.PaintPicture m_PicButtonBottom_UP, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
       End If
    End If
    tmrDown.Enabled = True
End Sub

Private Sub picDN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    MouseDnDown = False
    If m_Style = 0 Or m_Style = 1 Then
       If CheckMouseOver(picDN.hWnd) = True Then
           DrawButton picDN, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowDown, m_ButtonsArrowColor, m_ButtonsBorderColor
       Else
           DrawButton picDN, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowDown, m_ButtonsArrowColor, m_ButtonsBorderColor
       End If
    Else
       If CheckMouseOver(picDN.hWnd) = True Then
           If Not m_PicButtonBottom_HOOVER Is Nothing Then
               picDN.PaintPicture m_PicButtonBottom_HOOVER, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
           Else
               If Not m_PicButtonBottom_UP Is Nothing Then
                   picDN.PaintPicture m_PicButtonBottom_UP, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
               Else
                   picDN.Picture = Nothing
               End If
           End If
       Else
           If Not m_PicButtonBottom_UP Is Nothing Then
               picDN.PaintPicture m_PicButtonBottom_UP, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
           Else
               picDN.Picture = Nothing
           End If
       End If
    End If
    tmrDown.Enabled = False
End Sub

Private Sub picDN_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If MouseDnDown = True Then Exit Sub
    If CheckMouseOver(picDN.hWnd) = True Then
        If x > 60 And x < picDN.Width - 60 And y > 60 And y < picDN.Height - 60 Then
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picDN, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowDown, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonBottom_HOOVER Is Nothing Then
                    picDN.PaintPicture m_PicButtonBottom_HOOVER, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
                Else
                    If Not m_PicButtonBottom_UP Is Nothing Then
                        picDN.PaintPicture m_PicButtonBottom_UP, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
                    Else
                        picDN.Picture = Nothing
                    End If
                End If
            End If
        Else
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picDN, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowDown, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonBottom_UP Is Nothing Then
                    picDN.PaintPicture m_PicButtonBottom_UP, 0, 0, picDN.ScaleWidth, picDN.ScaleHeight
                Else
                    picDN.Picture = Nothing
                End If
            End If
        End If
    End If
End Sub

Private Sub tmrDown_Timer()
    Value = m_Value + m_SmallChange
End Sub

Private Sub picDN_DblClick()
    RaiseEvent DblClick
End Sub


'=====================================================
' Button Left
'=====================================================
Private Sub picLeft_Click()
    Value = m_Value - m_SmallChange
    MouseDnLeft = False
End Sub

Private Sub picLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    tmrLeft.Enabled = True
    MouseDnLeft = True
    If m_Style = 0 Then
        DrawButton picLeft, ShiftColors(m_ButtonsBackColor, 170), m_ButtonsBackColor, ArrowLeft, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 1 Then
        DrawButton picLeft, ShiftColors(m_ButtonsBackColor, -20), ShiftColors(m_ButtonsBackColor, -20), ArrowLeft, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 2 Then
       If Not m_PicButtonLeft_DOWN Is Nothing Then
           picLeft.PaintPicture m_PicButtonLeft_DOWN, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
       Else
           picLeft.PaintPicture m_PicButtonLeft_UP, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
       End If
    End If
End Sub

Private Sub picLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    MouseDnLeft = False
    If m_Style = 0 Or m_Style = 1 Then
        If CheckMouseOver(picLeft.hWnd) = True Then
            DrawButton picLeft, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowLeft, m_ButtonsArrowColor, m_ButtonsBorderColor
        Else
            DrawButton picLeft, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowLeft, m_ButtonsArrowColor, m_ButtonsBorderColor
        End If
    Else
        If CheckMouseOver(picDN.hWnd) = True Then
            If Not m_PicButtonLeft_HOOVER Is Nothing Then
                picLeft.PaintPicture m_PicButtonLeft_HOOVER, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
            Else
                If Not m_PicButtonLeft_UP Is Nothing Then
                    picLeft.PaintPicture m_PicButtonLeft_UP, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
                Else
                    picLeft.Picture = Nothing
                End If
            End If
        Else
            If Not m_PicButtonLeft_UP Is Nothing Then
                picLeft.PaintPicture m_PicButtonLeft_UP, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
            Else
                picLeft.Picture = Nothing
            End If
        End If
    End If
    tmrLeft.Enabled = False
End Sub

Private Sub picLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If MouseDnLeft = True Then Exit Sub
    If CheckMouseOver(picLeft.hWnd) = True Then
        If x > 60 And x < picLeft.Width - 60 And y > 60 And y < picLeft.Height - 60 Then
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picLeft, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowLeft, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonLeft_HOOVER Is Nothing Then
                    picLeft.PaintPicture m_PicButtonLeft_HOOVER, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
                Else
                    If Not m_PicButtonLeft_UP Is Nothing Then
                        picLeft.PaintPicture m_PicButtonLeft_UP, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
                    Else
                        picLeft.Picture = Nothing
                    End If
                End If
           End If
        Else
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picLeft, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowLeft, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonLeft_UP Is Nothing Then
                    picLeft.PaintPicture m_PicButtonLeft_UP, 0, 0, picLeft.ScaleWidth, picLeft.ScaleHeight
                Else
                    picLeft.Picture = Nothing
                End If
            End If
        End If
    End If

End Sub
Private Sub tmrLeft_Timer()
    Value = m_Value - m_SmallChange
End Sub

Private Sub picLeft_DblClick()
    RaiseEvent DblClick
End Sub

'=====================================================
' Button Right
'=====================================================
Private Sub picRight_Click()
    Value = m_Value + m_SmallChange
    MouseDnRight = False
End Sub

Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    tmrRight.Enabled = True
    MouseDnRight = True
    If m_Style = 0 Then
        DrawButton picRight, ShiftColors(m_ButtonsBackColor, 170), m_ButtonsBackColor, ArrowRight, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 1 Then
        DrawButton picRight, ShiftColors(m_ButtonsBackColor, -20), ShiftColors(m_ButtonsBackColor, -20), ArrowRight, m_ButtonsArrowColor, m_ButtonsBorderColor
    ElseIf m_Style = 2 Then
       If Not m_PicButtonRight_DOWN Is Nothing Then
           Set picRight.Picture = m_PicButtonRight_DOWN
           picRight.PaintPicture m_PicButtonRight_DOWN, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
       Else
           picRight.PaintPicture m_PicButtonRight_UP, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
       End If
    End If
End Sub

Private Sub picRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    MouseDnRight = False
    If m_Style = 0 Or m_Style = 1 Then
        If CheckMouseOver(picRight.hWnd) = True Then
            DrawButton picRight, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowRight, m_ButtonsArrowColor, m_ButtonsBorderColor
        Else
            DrawButton picRight, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowRight, m_ButtonsArrowColor, m_ButtonsBorderColor
        End If
    Else
        If CheckMouseOver(picDN.hWnd) = True Then
            If Not m_PicButtonRight_HOOVER Is Nothing Then
                picRight.PaintPicture m_PicButtonRight_HOOVER, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
            Else
                If Not m_PicButtonRight_UP Is Nothing Then
                    picRight.PaintPicture m_PicButtonRight_UP, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
                Else
                    picRight.Picture = Nothing
                End If
            End If
        Else
            If Not m_PicButtonRight_UP Is Nothing Then
                picRight.PaintPicture m_PicButtonRight_UP, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
            Else
                picRight.Picture = Nothing
            End If
        End If
    End If
    tmrRight.Enabled = False
End Sub

Private Sub picRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If MouseDnRight = True Then Exit Sub
    If CheckMouseOver(picRight.hWnd) = True Then
        If x > 60 And x < picRight.Width - 60 And y > 60 And y < picRight.Height - 60 Then
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picRight, ShiftColors(m_ButtonsBackColor, 70), ShiftColors(m_ButtonsBackColor, 180), ArrowRight, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonRight_HOOVER Is Nothing Then
                    picRight.PaintPicture m_PicButtonRight_HOOVER, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
                Else
                    If Not m_PicButtonRight_UP Is Nothing Then
                        picRight.PaintPicture m_PicButtonRight_UP, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
                    Else
                        picRight.Picture = Nothing
                    End If
                End If
           End If
        Else
            If m_Style = 0 Or m_Style = 1 Then
                DrawButton picRight, m_ButtonsBackColor, ShiftColors(m_ButtonsBackColor, 170), ArrowRight, m_ButtonsArrowColor, m_ButtonsBorderColor
            Else
                If Not m_PicButtonRight_UP Is Nothing Then
                    picRight.PaintPicture m_PicButtonRight_UP, 0, 0, picRight.ScaleWidth, picRight.ScaleHeight
                Else
                    picRight.Picture = Nothing
                End If
            End If
        End If
    End If
End Sub
Private Sub tmrRight_Timer()
    Value = m_Value + m_SmallChange
End Sub

Private Sub picRight_DblClick()
    RaiseEvent DblClick
End Sub

'=====================================================
' Scroller
'=====================================================
Private Sub picScroller_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ' Set colors to darker color or set picture to downpicture
    If m_Style = 0 Then
       DrawScroller ShiftColors(m_ScrollerBackColor, 170), m_ScrollerBackColor, m_ScrollerBorderColor, m_ScrollerGripColor
    ElseIf m_Style = 1 Then
       DrawScroller ShiftColors(m_ScrollerBackColor, -20), ShiftColors(m_ScrollerBackColor, -20), m_ScrollerBorderColor, m_ScrollerGripColor
    ElseIf m_Style = 2 Then
       If Not m_PicScroller_DOWN Is Nothing Then
           picScroller.PaintPicture m_PicScroller_DOWN, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
       Else
           If Not m_PicScroller_UP Is Nothing Then
               picScroller.PaintPicture m_PicScroller_UP, 0, 0, picScroller.ScaleHeight, picScroller.ScaleWidth
           Else
               picScroller.Picture = Nothing
           End If
       End If
    End If
    ' Set coordinates
    If m_Orientation = Vertical Then
      m_MouseY = y
      m_Sliding = True
      m_OldPosn = picScroller.Top + y
    Else
      m_MouseX = x
      m_Sliding = True
      m_OldPosn = picScroller.Left + x
    End If
End Sub

Private Sub picScroller_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If m_Style = 0 Or m_Style = 1 Then
        If CheckMouseOver(picScroller.hWnd) = True Then
            DrawScroller ShiftColors(m_ScrollerBackColor, 70), ShiftColors(m_ScrollerBackColor, 180), m_ScrollerBorderColor, m_ScrollerGripColor
        Else
            DrawScroller m_ScrollerBackColor, ShiftColors(m_ScrollerBackColor, 170), m_ScrollerBorderColor, m_ScrollerGripColor
        End If
    Else
        If CheckMouseOver(picScroller.hWnd) = True Then
            If Not m_PicScroller_HOOVER Is Nothing Then
                picScroller.PaintPicture m_PicScroller_HOOVER, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
            Else
                If Not m_PicScroller_UP Is Nothing Then
                    picScroller.PaintPicture m_PicScroller_UP, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
                Else
                    picScroller.Picture = Nothing
                End If
            End If
        Else
            If Not m_PicScroller_UP Is Nothing Then
                picScroller.PaintPicture m_PicScroller_UP, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
            Else
                picScroller.Picture = Nothing
            End If
        End If
    End If
   m_Sliding = False
   If m_ValueChanged Then RaiseEvent Change
End Sub

Private Sub picScroller_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim NewPosn As Long
    Dim MaxY As Long
    Dim MinY As Long
    Dim Maxx As Long
    Dim MinX As Long
    ' Only redraw when not sliding the scroller
    If Not m_Sliding Then
        ' Only redraw when the mouse is exactly over the scroller (inside the control)
        If CheckMouseOver(picScroller.hWnd) = True Then
            ' redraw to hooverpic when between borders
            If x > 60 And x < picScroller.Width - 60 And y > 60 And y < picScroller.Height - 60 Then
                If m_Style = 0 Or m_Style = 1 Then
                    DrawScroller ShiftColors(m_ScrollerBackColor, 70), ShiftColors(m_ScrollerBackColor, 180), m_ScrollerBorderColor, m_ScrollerGripColor
                Else
                    If Not m_PicScroller_HOOVER Is Nothing Then
                        picScroller.PaintPicture m_PicScroller_HOOVER, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
                    Else
                        If Not m_PicScroller_UP Is Nothing Then
                            picScroller.PaintPicture m_PicScroller_UP, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
                        Else
                            picScroller.Picture = Nothing
                        End If
                    End If
                End If
            Else
                ' Redraw to normal pic when outside borders
                If m_Style = 0 Or m_Style = 1 Then
                    DrawScroller m_ScrollerBackColor, ShiftColors(m_ScrollerBackColor, 170), m_ScrollerBorderColor, m_ScrollerGripColor
                Else
                    If Not m_PicScroller_UP Is Nothing Then
                        picScroller.PaintPicture m_PicScroller_UP, 0, 0, picScroller.ScaleWidth, picScroller.ScaleHeight
                    Else
                        picScroller.Picture = Nothing
                    End If
                End If
            End If
        End If
        Exit Sub
    End If
    If m_Orientation = Vertical Then
       ' Add the .Top value to the Height to take in account the ButtonsVisible property
       MinY = picUP.Top + picUP.Height
       MaxY = picDN.Top - (picScroller.Height)
         ' Determine Position of the Scroller
         NewPosn = picScroller.Top + y - m_MouseY
         If NewPosn >= MaxY Then
             NewPosn = MaxY
         End If
         If NewPosn <= MinY Then
             NewPosn = MinY
         End If
         ' Don't need to do anything if we haven't moved
         If NewPosn <> m_OldPosn Then
             ' Move the Scroller
             picScroller.Move picScroller.Left, NewPosn
             ' Calculate the new Value based on the position of the Scroller between the Up and Down Buttons
             m_Value = ((picScroller.Top - MinY) / (MaxY - MinY)) * (m_MaxValue - m_MinValue) + m_MinValue
             RaiseEvent Scroll
             ' Save position
             m_OldPosn = NewPosn
             ' Set the variable so we know if we should trigger the Change Event on MouseUp
             m_ValueChanged = True
         End If
    Else  ' Horizontal scrolling
       ' Add the .Left value to the Height to take
       ' in account the ButtonsVisible property
       MinX = picLeft.Left + picLeft.Width
       Maxx = picRight.Left - picScroller.Width
         ' Determine  Position of the Scroller
         NewPosn = picScroller.Left + x - m_MouseX
         If NewPosn >= Maxx Then
             NewPosn = Maxx
         End If
         If NewPosn <= MinX Then
             NewPosn = MinX
         End If
         ' Don't need to do anything if we haven't moved
         If NewPosn <> m_OldPosn Then
             ' Move the Scroller
             picScroller.Move NewPosn, picScroller.Top
             ' Calculate the new Value based on the position of the Scroller between the Left and Right Buttons
             m_Value = ((picScroller.Left - MinX) / (Maxx - MinX)) * (m_MaxValue - m_MinValue) + m_MinValue
             RaiseEvent Scroll
             ' Save position
             m_OldPosn = NewPosn
             ' Set the variable so we know if we should trigger the Change Event on MouseUp
             m_ValueChanged = True
         End If
    End If
End Sub

Private Sub picScroller_DblClick()
    RaiseEvent DblClick
End Sub



'=====================================================
' Configure the control
'=====================================================
Private Sub ConfigureControl()
   ' Shut Flag Off - We only need to do this the first time
   bAddedToIDE = False
   ' Automattically Determine if the ScrollBar is Vertical or Horizontal
   ' Based on the Width / Height of the control when the user creates the control on the form.
   m_Orientation = IIf(UserControl.ScaleHeight >= UserControl.ScaleWidth, Vertical, Horizontal)
   picBG.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
   ' Setup Picture Boxes
   Call SetPictureVisability
   Select Case m_Orientation
      Case Vertical
         UserControl.Width = 315
      Case Horizontal
         UserControl.Height = 315
   End Select
End Sub

'=====================================================
' Configure all pictures
'=====================================================
Private Sub SetPictureVisability()
   ' Hide All and Show the ones we need Others
   picRight.Visible = False
   picLeft.Visible = False
   picUP.Visible = False
   picDN.Visible = False
   ' Set the Scroller
   picScroller.Visible = True
   If Not (m_ButtonsVisible) Then Exit Sub
   Select Case m_Orientation
      Case Vertical
         picUP.Visible = True
         picDN.Visible = True
      Case Horizontal
         picRight.Visible = True
         picLeft.Visible = True
   End Select
End Sub

'=====================================================
' Positioning scroller
'=====================================================
Private Function PositionScroller() As Long
    Dim MinY As Single
    Dim MaxY As Single
    Dim MinX As Single
    Dim Maxx As Single
    
    ' Reposition the Scroller based on the
    ' Orientation of the Slider and the Value.
    ' We Get what Percent Value is based on the Min / Max Values
    ' We then multiple this percent by the distance between
    ' the Top / Bottom buttons, taking into account the width/height
    ' of the Scroller.
    With UserControl
       If m_Orientation = Vertical Then
          If ButtonsVisible Then
             MinY = picUP.Height
             MaxY = picDN.Top - picScroller.Height
          Else
             MinY = 0
             MaxY = UserControl.ScaleHeight - picScroller.Height
          End If
          picScroller.Top = (m_Value - m_MinValue) / (m_MaxValue - m_MinValue) * (MaxY - MinY) + MinY
          ' Move the Scroller based on the Alignment
            picScroller.Left = 0
       Else
          If ButtonsVisible Then
             MinX = picLeft.Width
             Maxx = picRight.Left - picScroller.Width
          Else
             MinX = 0
             Maxx = UserControl.ScaleWidth - picScroller.Width
          End If
          picScroller.Left = (m_Value - m_MinValue) / (m_MaxValue - m_MinValue) * (Maxx - MinX) + MinX
          ' Move the Scroller based on the Alignment
          picScroller.Top = 0
       End If
    End With
End Function

'=====================================================
' Draw buttons with arrows
'=====================================================
Public Sub DrawButton(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, ArrowType As ArrowTypes, ArrowColor As OLE_COLOR, BordersColor As OLE_COLOR)    'Horizontal gradient
    On Error Resume Next
    DoEvents
    OldScaleMode = ctlControl.ScaleMode
    ctlControl.ScaleMode = 3
    If m_Style = Flat Or EndColor = &H8000000F Then
        ctlControl.BackColor = StartColor
        GoTo DrawArrows
    End If
    ' Draw background
    If m_Orientation = Vertical Then
        Call InitializeCol(ctlControl, EndColor, StartColor, False)
        For i = 0 To ctlControl.ScaleWidth
            NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
            ctlControl.Line (i, 0)-(i, ctlControl.ScaleHeight), NewColor
        Next
    Else
        Call InitializeCol(ctlControl, EndColor, StartColor, False, True)
        For i = 0 To ctlControl.ScaleHeight
            NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
            ctlControl.Line (0, i)-(ctlControl.ScaleWidth, i), NewColor
        Next
    End If
    DoEvents
DrawArrows:
    ' Draw arrows
    MidX = Round(ctlControl.ScaleWidth / 2)
    MidY = Round(ctlControl.ScaleHeight / 2)
    Dim BordersColor2 As OLE_COLOR
    Dim BordersColor3 As OLE_COLOR
    If m_Style = Flat Then
        BordersColor2 = ArrowColor
        BordersColor3 = ArrowColor
    Else
        BordersColor2 = ShiftColors(ArrowColor, 150)
        BordersColor3 = ShiftColors(ArrowColor, 90)
    End If
    If m_ButtonsArrowStyle = ArrowRaised Then
        If ArrowType = ArrowUp Then
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 1, MidY - 4), BordersColor2
            ctlControl.Line (MidX + 5, MidY + 2)-(MidX - 1, MidY - 4), BordersColor2
            ctlControl.Line (MidX - 4, MidY + 2)-(MidX + 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY + 2)-(MidX - 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY + 2)-(MidX + 1, MidY - 2), ArrowColor
            ctlControl.Line (MidX + 3, MidY + 2)-(MidX - 1, MidY - 2), ArrowColor
        ElseIf ArrowType = ArrowDown Then
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 1, MidY + 4), ArrowColor
            ctlControl.Line (MidX + 5, MidY - 2)-(MidX - 1, MidY + 4), ArrowColor
            ctlControl.Line (MidX - 4, MidY - 2)-(MidX + 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY - 2)-(MidX - 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY - 2)-(MidX + 1, MidY + 2), BordersColor2
            ctlControl.Line (MidX + 3, MidY - 2)-(MidX - 1, MidY + 2), BordersColor2
        ElseIf ArrowType = ArrowLeft Then
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX - 4, MidY + 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY + 5)-(MidX - 4, MidY - 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY - 4)-(MidX - 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY + 4)-(MidX - 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY - 3)-(MidX - 2, MidY + 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY + 3)-(MidX - 2, MidY - 1), ArrowColor
        ElseIf ArrowType = ArrowRight Then
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX + 4, MidY + 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY + 5)-(MidX + 4, MidY - 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY - 4)-(MidX + 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 4)-(MidX + 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY - 3)-(MidX + 2, MidY + 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY + 3)-(MidX + 2, MidY - 1), BordersColor2
        End If
    ElseIf m_ButtonsArrowStyle = ArrowEngraved Then
        BordersColor3 = ShiftColors(ArrowColor, 30)
        If ArrowType = ArrowUp Then
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 1, MidY - 4), ArrowColor
            ctlControl.Line (MidX + 5, MidY + 2)-(MidX - 1, MidY - 4), ArrowColor
            ctlControl.Line (MidX - 4, MidY + 2)-(MidX + 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY + 2)-(MidX - 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY + 2)-(MidX + 1, MidY - 2), BordersColor2
            ctlControl.Line (MidX + 3, MidY + 2)-(MidX - 1, MidY - 2), BordersColor2
        ElseIf ArrowType = ArrowDown Then
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 1, MidY + 4), BordersColor2
            ctlControl.Line (MidX + 5, MidY - 2)-(MidX - 1, MidY + 4), BordersColor2
            ctlControl.Line (MidX - 4, MidY - 2)-(MidX + 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY - 2)-(MidX - 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY - 2)-(MidX + 1, MidY + 2), ArrowColor
            ctlControl.Line (MidX + 3, MidY - 2)-(MidX - 1, MidY + 2), ArrowColor
        ElseIf ArrowType = ArrowLeft Then
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX - 4, MidY + 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY + 5)-(MidX - 4, MidY - 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY - 4)-(MidX - 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY + 4)-(MidX - 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY - 3)-(MidX - 2, MidY + 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY + 3)-(MidX - 2, MidY - 1), BordersColor2
        ElseIf ArrowType = ArrowRight Then
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX + 4, MidY + 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY + 5)-(MidX + 4, MidY - 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY - 4)-(MidX + 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 4)-(MidX + 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY - 3)-(MidX + 2, MidY + 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY + 3)-(MidX + 2, MidY - 1), ArrowColor
        End If
    ElseIf m_ButtonsArrowStyle = TriangleRaised Then
        If ArrowType = ArrowUp Then
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 1, MidY - 4), ArrowColor
            ctlControl.Line (MidX + 5, MidY + 2)-(MidX - 1, MidY - 4), ArrowColor
            ctlControl.Line (MidX - 0, MidY - 2)-(MidX + 1, MidY - 2), BordersColor3
            ctlControl.Line (MidX - 1, MidY - 1)-(MidX + 2, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 0)-(MidX + 3, MidY + 0), BordersColor3
            ctlControl.Line (MidX - 3, MidY + 1)-(MidX + 4, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 5, MidY + 2), ArrowColor
            ctlControl.Line (MidX - 0, MidY - 1)-(MidX + 1, MidY - 1), BordersColor2
            ctlControl.Line (MidX - 1, MidY + 0)-(MidX + 2, MidY + 0), BordersColor2
        ElseIf ArrowType = ArrowDown Then
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 1, MidY + 4), ArrowColor
            ctlControl.Line (MidX + 5, MidY - 2)-(MidX - 1, MidY + 4), ArrowColor
            ctlControl.Line (MidX - 0, MidY + 2)-(MidX + 1, MidY + 2), BordersColor3
            ctlControl.Line (MidX - 1, MidY + 1)-(MidX + 2, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY - 0)-(MidX + 3, MidY - 0), BordersColor3
            ctlControl.Line (MidX - 3, MidY - 1)-(MidX + 4, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 5, MidY - 2), ArrowColor
            ctlControl.Line (MidX - 0, MidY + 1)-(MidX + 1, MidY + 1), BordersColor2
            ctlControl.Line (MidX - 1, MidY - 0)-(MidX + 2, MidY - 0), BordersColor2
        ElseIf ArrowType = ArrowLeft Then
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX - 4, MidY + 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY + 5)-(MidX - 4, MidY - 1), ArrowColor
            ctlControl.Line (MidX + 1, MidY - 3)-(MidX + 1, MidY + 4), BordersColor3
            ctlControl.Line (MidX + 0, MidY - 2)-(MidX + 0, MidY + 3), BordersColor3
            ctlControl.Line (MidX - 1, MidY - 1)-(MidX - 1, MidY + 2), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 0)-(MidX - 2, MidY + 1), BordersColor3
            ctlControl.Line (MidX + 0, MidY - 1)-(MidX + 0, MidY + 2), BordersColor2
            ctlControl.Line (MidX - 1, MidY - 0)-(MidX - 1, MidY + 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX + 2, MidY + 5), ArrowColor
        ElseIf ArrowType = ArrowRight Then
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX + 4, MidY + 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY + 5)-(MidX + 4, MidY - 1), ArrowColor
            ctlControl.Line (MidX - 1, MidY - 3)-(MidX - 1, MidY + 4), BordersColor3
            ctlControl.Line (MidX - 0, MidY - 2)-(MidX - 0, MidY + 3), BordersColor3
            ctlControl.Line (MidX + 1, MidY - 1)-(MidX + 1, MidY + 2), BordersColor3
            ctlControl.Line (MidX + 2, MidY - 0)-(MidX + 2, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 0, MidY - 1)-(MidX - 0, MidY + 2), BordersColor2
            ctlControl.Line (MidX + 1, MidY - 0)-(MidX + 1, MidY + 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX - 2, MidY + 5), ArrowColor
       End If
    ElseIf m_ButtonsArrowStyle = TriangleEngraved Then
       BordersColor3 = ShiftColors(ArrowColor, 40)
       If ArrowType = ArrowUp Then
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 1, MidY - 4), BordersColor2
            ctlControl.Line (MidX + 5, MidY + 2)-(MidX - 1, MidY - 4), BordersColor2
            ctlControl.Line (MidX - 0, MidY - 2)-(MidX + 1, MidY - 2), BordersColor3
            ctlControl.Line (MidX - 1, MidY - 1)-(MidX + 2, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 0)-(MidX + 3, MidY + 0), BordersColor3
            ctlControl.Line (MidX - 3, MidY + 1)-(MidX + 4, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 5, MidY + 2), BordersColor2
            ctlControl.Line (MidX - 0, MidY - 1)-(MidX + 1, MidY - 1), ArrowColor
            ctlControl.Line (MidX - 1, MidY + 0)-(MidX + 2, MidY + 0), ArrowColor
        ElseIf ArrowType = ArrowDown Then
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 1, MidY + 4), BordersColor2
            ctlControl.Line (MidX + 5, MidY - 2)-(MidX - 1, MidY + 4), BordersColor2
            ctlControl.Line (MidX - 0, MidY + 2)-(MidX + 1, MidY + 2), BordersColor3
            ctlControl.Line (MidX - 1, MidY + 1)-(MidX + 2, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY - 0)-(MidX + 3, MidY - 0), BordersColor3
            ctlControl.Line (MidX - 3, MidY - 1)-(MidX + 4, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 5, MidY - 2), BordersColor2
            ctlControl.Line (MidX - 0, MidY + 1)-(MidX + 1, MidY + 1), ArrowColor
            ctlControl.Line (MidX - 1, MidY - 0)-(MidX + 2, MidY - 0), ArrowColor
        ElseIf ArrowType = ArrowLeft Then
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX - 4, MidY + 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY + 5)-(MidX - 4, MidY - 1), BordersColor2
            ctlControl.Line (MidX + 1, MidY - 3)-(MidX + 1, MidY + 4), BordersColor3
            ctlControl.Line (MidX + 0, MidY - 2)-(MidX + 0, MidY + 3), BordersColor3
            ctlControl.Line (MidX - 1, MidY - 1)-(MidX - 1, MidY + 2), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 0)-(MidX - 2, MidY + 1), BordersColor3
            ctlControl.Line (MidX + 0, MidY - 1)-(MidX + 0, MidY + 2), ArrowColor
            ctlControl.Line (MidX - 1, MidY - 0)-(MidX - 1, MidY + 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX + 2, MidY + 5), BordersColor2
        ElseIf ArrowType = ArrowRight Then
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX + 4, MidY + 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY + 5)-(MidX + 4, MidY - 1), BordersColor2
            ctlControl.Line (MidX - 1, MidY - 3)-(MidX - 1, MidY + 4), BordersColor3
            ctlControl.Line (MidX - 0, MidY - 2)-(MidX - 0, MidY + 3), BordersColor3
            ctlControl.Line (MidX + 1, MidY - 1)-(MidX + 1, MidY + 2), BordersColor3
            ctlControl.Line (MidX + 2, MidY - 0)-(MidX + 2, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 0, MidY - 1)-(MidX - 0, MidY + 2), ArrowColor
            ctlControl.Line (MidX + 1, MidY - 0)-(MidX + 1, MidY + 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX - 2, MidY + 5), BordersColor2
       End If
    ElseIf m_ButtonsArrowStyle = CirclesRaised Then
        If ArrowType = ArrowUp Then
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 1, MidY - 4), BordersColor2
            ctlControl.Line (MidX + 5, MidY + 2)-(MidX - 1, MidY - 4), BordersColor2
            ctlControl.Line (MidX - 4, MidY + 2)-(MidX + 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY + 2)-(MidX - 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY + 2)-(MidX + 1, MidY - 2), ArrowColor
            ctlControl.Line (MidX + 3, MidY + 2)-(MidX - 1, MidY - 2), ArrowColor
        ElseIf ArrowType = ArrowDown Then
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 1, MidY + 4), ArrowColor
            ctlControl.Line (MidX + 5, MidY - 2)-(MidX - 1, MidY + 4), ArrowColor
            ctlControl.Line (MidX - 4, MidY - 2)-(MidX + 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY - 2)-(MidX - 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY - 2)-(MidX + 1, MidY + 2), BordersColor2
            ctlControl.Line (MidX + 3, MidY - 2)-(MidX - 1, MidY + 2), BordersColor2
        ElseIf ArrowType = ArrowLeft Then
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX - 4, MidY + 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY + 5)-(MidX - 4, MidY - 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY - 4)-(MidX - 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY + 4)-(MidX - 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY - 3)-(MidX - 2, MidY + 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY + 3)-(MidX - 2, MidY - 1), ArrowColor
        ElseIf ArrowType = ArrowRight Then
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX + 4, MidY + 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY + 5)-(MidX + 4, MidY - 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY - 4)-(MidX + 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 4)-(MidX + 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY - 3)-(MidX + 2, MidY + 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY + 3)-(MidX + 2, MidY - 1), BordersColor2
        End If
        If picScroller.Width < picScroller.Height Then
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleWidth / 2) - 3, ArrowColor
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleWidth / 2) - 4, BordersColor2
        Else
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleHeight / 2) - 3, ArrowColor
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleHeight / 2) - 4, BordersColor2
        End If
    ElseIf m_ButtonsArrowStyle = CirclesEngraved Then
        If ArrowType = ArrowUp Then
            ctlControl.Line (MidX - 5, MidY + 2)-(MidX + 1, MidY - 4), ArrowColor
            ctlControl.Line (MidX + 5, MidY + 2)-(MidX - 1, MidY - 4), ArrowColor
            ctlControl.Line (MidX - 4, MidY + 2)-(MidX + 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY + 2)-(MidX - 1, MidY - 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY + 2)-(MidX + 1, MidY - 2), BordersColor2
            ctlControl.Line (MidX + 3, MidY + 2)-(MidX - 1, MidY - 2), BordersColor2
        ElseIf ArrowType = ArrowDown Then
            ctlControl.Line (MidX - 5, MidY - 2)-(MidX + 1, MidY + 4), BordersColor2
            ctlControl.Line (MidX + 5, MidY - 2)-(MidX - 1, MidY + 4), BordersColor2
            ctlControl.Line (MidX - 4, MidY - 2)-(MidX + 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX + 4, MidY - 2)-(MidX - 1, MidY + 3), BordersColor3
            ctlControl.Line (MidX - 3, MidY - 2)-(MidX + 1, MidY + 2), ArrowColor
            ctlControl.Line (MidX + 3, MidY - 2)-(MidX - 1, MidY + 2), ArrowColor
        ElseIf ArrowType = ArrowLeft Then
            ctlControl.Line (MidX + 2, MidY - 5)-(MidX - 4, MidY + 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY + 5)-(MidX - 4, MidY - 1), ArrowColor
            ctlControl.Line (MidX + 2, MidY - 4)-(MidX - 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY + 4)-(MidX - 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX + 2, MidY - 3)-(MidX - 2, MidY + 1), BordersColor2
            ctlControl.Line (MidX + 2, MidY + 3)-(MidX - 2, MidY - 1), BordersColor2
        ElseIf ArrowType = ArrowRight Then
            ctlControl.Line (MidX - 2, MidY - 5)-(MidX + 4, MidY + 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY + 5)-(MidX + 4, MidY - 1), BordersColor2
            ctlControl.Line (MidX - 2, MidY - 4)-(MidX + 3, MidY + 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY + 4)-(MidX + 3, MidY - 1), BordersColor3
            ctlControl.Line (MidX - 2, MidY - 3)-(MidX + 2, MidY + 1), ArrowColor
            ctlControl.Line (MidX - 2, MidY + 3)-(MidX + 2, MidY - 1), ArrowColor
        End If
        If picScroller.Width < picScroller.Height Then
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleWidth / 2) - 3, BordersColor3
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleWidth / 2) - 4, BordersColor2
        Else
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleHeight / 2) - 3, BordersColor3
            ctlControl.Circle (ctlControl.ScaleWidth / 2, ctlControl.ScaleHeight / 2), (ctlControl.ScaleHeight / 2) - 4, BordersColor2
        End If
    ElseIf m_ButtonsArrowStyle = WafelsEngraved Then
        BordersColor2 = ShiftColors(ArrowColor, 130)
        For iii = 3 To (ctlControl.ScaleHeight - 4) Step 4
            If iii > (ctlControl.ScaleHeight - 4) Then Exit For
            ctlControl.Line (3, iii)-(ctlControl.ScaleWidth - 3, iii), ArrowColor
            If m_Style = Graphical Then ctlControl.Line (4, iii + 1)-(ctlControl.ScaleWidth - 4, iii + 1), BordersColor3
            If m_Style = Graphical Then ctlControl.Line (4, iii + 2)-(ctlControl.ScaleWidth - 4, iii + 2), BordersColor2
        Next iii
        For iii = 3 To (ctlControl.ScaleWidth - 4) Step 4
            If iii > (ctlControl.ScaleWidth - 4) Then Exit For
            ctlControl.Line (iii, 3)-(iii, ctlControl.ScaleHeight - 3), ArrowColor
            If m_Style = Graphical Then ctlControl.Line (iii + 1, 4)-(iii + 1, ctlControl.ScaleHeight - 4), BordersColor3
            If m_Style = Graphical Then ctlControl.Line (iii + 2, 4)-(iii + 2, ctlControl.ScaleHeight - 4), BordersColor2
        Next iii
    ElseIf m_ButtonsArrowStyle = WafelsRaised Then
        BordersColor2 = ShiftColors(ArrowColor, 130)
        For iii = 3 To (ctlControl.ScaleHeight - 4) Step 4
            If iii > (ctlControl.ScaleHeight - 4) Then Exit For
            ctlControl.Line (3, iii)-(ctlControl.ScaleWidth - 3, iii), BordersColor2
            If m_Style = Graphical Then ctlControl.Line (4, iii + 1)-(ctlControl.ScaleWidth - 4, iii + 1), BordersColor3
            If m_Style = Graphical Then ctlControl.Line (4, iii + 2)-(ctlControl.ScaleWidth - 4, iii + 2), ArrowColor
        Next iii
        For iii = 3 To (ctlControl.ScaleWidth - 4) Step 4
            If iii > (ctlControl.ScaleWidth - 4) Then Exit For
            ctlControl.Line (iii, 3)-(iii, ctlControl.ScaleHeight - 3), BordersColor2
            If m_Style = Graphical Then ctlControl.Line (iii + 1, 4)-(iii + 1, ctlControl.ScaleHeight - 4), BordersColor3
            If m_Style = Graphical Then ctlControl.Line (iii + 2, 4)-(iii + 2, ctlControl.ScaleHeight - 4), ArrowColor
        Next iii
    End If
DrawBorders:
    ' Draw borders
    If m_ButtonsBorderStyle <> None Then
        Dim NewBorderColor As OLE_COLOR
        If m_Enabled = True Then
            NewBorderColor = m_BarBorderColor
        Else
            NewBorderColor = m_DisabledBorderColor
        End If
        ' Draw borders based on the type of button (Top,Bottom,Left,Right)
        ' The type of button is recognised by the type of arrow (arrowup is top button,...)
        If ArrowType = ArrowUp Then
            ctlControl.Line (0, 0)-(ctlControl.ScaleWidth, 0), BordersColor
            ctlControl.Line (0, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth, ctlControl.ScaleHeight - 1), NewBorderColor
            ctlControl.Line (1, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight - 1), BordersColor
            ctlControl.Line (0, 0)-(0, ctlControl.ScaleHeight - 1), BordersColor
            ctlControl.Line (ctlControl.ScaleWidth - 1, 0)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight - 1), BordersColor
        ElseIf ArrowType = ArrowDown Then
            ctlControl.Line (0, 0)-(ctlControl.ScaleWidth, 0), NewBorderColor
            ctlControl.Line (1, 0)-(ctlControl.ScaleWidth - 1, 0), BordersColor
            ctlControl.Line (0, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth, ctlControl.ScaleHeight - 1), BordersColor
            ctlControl.Line (0, 1)-(0, ctlControl.ScaleHeight), BordersColor
            ctlControl.Line (ctlControl.ScaleWidth - 1, 1)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight), BordersColor
        ElseIf ArrowType = ArrowLeft Then
            ctlControl.Line (0, 0)-(ctlControl.ScaleWidth, 0), BordersColor
            ctlControl.Line (0, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth, ctlControl.ScaleHeight - 1), BordersColor
            ctlControl.Line (0, 0)-(0, ctlControl.ScaleHeight), BordersColor
            ctlControl.Line (ctlControl.ScaleWidth - 1, 0)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight), NewBorderColor
            ctlControl.Line (ctlControl.ScaleWidth - 1, 1)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight - 1), BordersColor
        ElseIf ArrowType = ArrowRight Then
            ctlControl.Line (0, 0)-(ctlControl.ScaleWidth, 0), BordersColor
            ctlControl.Line (0, ctlControl.ScaleHeight - 1)-(ctlControl.ScaleWidth, ctlControl.ScaleHeight - 1), BordersColor
            ctlControl.Line (ctlControl.ScaleWidth - 1, 0)-(ctlControl.ScaleWidth - 1, ctlControl.ScaleHeight), BordersColor
            ctlControl.Line (0, 0)-(0, ctlControl.ScaleHeight), NewBorderColor
            ctlControl.Line (0, 1)-(0, ctlControl.ScaleHeight - 1), BordersColor
        End If
    End If
    ctlControl.Refresh
    ctlControl.ScaleMode = OldScaleMode
End Sub

'=====================================================
' Draw scroller with grip
'=====================================================
Public Sub DrawScroller(StartColor As OLE_COLOR, EndColor As OLE_COLOR, Bordercolor As OLE_COLOR, GripperColor As OLE_COLOR) 'Horizontal gradient
    On Error Resume Next
    
    DoEvents
    OldScaleMode = picScroller.ScaleMode
    picScroller.ScaleMode = 3
    If m_Style = Flat Or EndColor = &H8000000F Then
       picScroller.BackColor = StartColor
       GoTo DrawGripper
    End If
    ' Draw background
    If m_Orientation = Vertical Then
        Call InitializeCol(picScroller, EndColor, StartColor, False)
        For i = 0 To picScroller.ScaleWidth
            NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
            picScroller.Line (i, 0)-(i, picScroller.ScaleHeight), NewColor
        Next
    Else
        Call InitializeCol(picScroller, EndColor, StartColor, False, True)
        For i = 0 To picScroller.ScaleHeight
            NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
            picScroller.Line (0, i)-(picScroller.ScaleWidth, i), NewColor
        Next
    End If
    DoEvents
DrawGripper:
    ' Draw gripper
    MidX = Round(picScroller.ScaleWidth / 2)
    MidY = Round(picScroller.ScaleHeight / 2)
    Dim GripperColor2 As OLE_COLOR
    Dim GripperColor3 As OLE_COLOR
    If m_Style = Flat Then
        GripperColor2 = GripperColor
        GripperColor3 = GripperColor
    Else
        GripperColor2 = ShiftColors(GripperColor, 130)
        GripperColor3 = ShiftColors(GripperColor, 70)
    End If
    If m_ScrollerGripperStyle = DotRaised Then
        For ii = 3 To (picScroller.ScaleWidth - 4) Step 4
            If ii > (picScroller.ScaleWidth - 4) Then Exit For
            For iii = 3 To (picScroller.ScaleHeight - 4) Step 4
                If iii > (picScroller.ScaleHeight - 4) Then Exit For
                picScroller.Line (ii, iii)-(ii + 2, iii), GripperColor2
                If m_Style = Graphical Then picScroller.Line (ii, iii + 1)-(ii + 1, iii + 1), GripperColor3
                If m_Style = Graphical Then picScroller.Line (ii + 1, iii + 1)-(ii + 1, iii + 2), GripperColor
            Next iii
        Next ii
    ElseIf m_ScrollerGripperStyle = DotEngraved Then
        For ii = 3 To (picScroller.ScaleWidth - 4) Step 4
            If ii > (picScroller.ScaleWidth - 4) Then Exit For
            For iii = 3 To (picScroller.ScaleHeight - 4) Step 4
                If iii > (picScroller.ScaleHeight - 4) Then Exit For
                picScroller.Line (ii, iii)-(ii + 2, iii), GripperColor
                If m_Style = Graphical Then picScroller.Line (ii, iii + 1)-(ii + 1, iii + 1), GripperColor3
                If m_Style = Graphical Then picScroller.Line (ii + 1, iii + 1)-(ii + 1, iii + 2), GripperColor2
            Next iii
        Next ii
    ElseIf m_ScrollerGripperStyle = BoxEngraved Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        picScroller.Line (4, 4)-(picScroller.ScaleWidth - 4, 4), GripperColor
        picScroller.Line (4, 4)-(4, picScroller.ScaleHeight - 4), GripperColor
        picScroller.Line (picScroller.ScaleWidth - 4, 4)-(picScroller.ScaleWidth - 4, picScroller.ScaleHeight - 4), GripperColor2
        picScroller.Line (4, picScroller.ScaleHeight - 4)-(picScroller.ScaleWidth - 3, picScroller.ScaleHeight - 4), GripperColor2
    ElseIf m_ScrollerGripperStyle = BoxRaised Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        picScroller.Line (4, 4)-(picScroller.ScaleWidth - 4, 4), GripperColor2
        picScroller.Line (4, 4)-(4, picScroller.ScaleHeight - 4), GripperColor2
        picScroller.Line (picScroller.ScaleWidth - 4, 4)-(picScroller.ScaleWidth - 4, picScroller.ScaleHeight - 4), GripperColor
        picScroller.Line (4, picScroller.ScaleHeight - 4)-(picScroller.ScaleWidth - 3, picScroller.ScaleHeight - 4), GripperColor
    ElseIf m_ScrollerGripperStyle = DiamondEngraved Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        picScroller.Line (MidX - 5, MidY)-(MidX + 1, MidY - 6), GripperColor
        picScroller.Line (MidX - 5, MidY)-(MidX + 1, MidY + 6), GripperColor2
        picScroller.Line (MidX + 5, MidY)-(MidX - 1, MidY - 6), GripperColor
        picScroller.Line (MidX + 5, MidY)-(MidX - 1, MidY + 6), GripperColor2
    ElseIf m_ScrollerGripperStyle = DiamondRaised Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        picScroller.Line (MidX - 5, MidY)-(MidX + 1, MidY - 6), GripperColor2
        picScroller.Line (MidX - 5, MidY)-(MidX + 1, MidY + 6), GripperColor
        picScroller.Line (MidX + 5, MidY)-(MidX - 1, MidY - 6), GripperColor2
        picScroller.Line (MidX + 5, MidY)-(MidX - 1, MidY + 6), GripperColor
    ElseIf m_ScrollerGripperStyle = CircleEngraved Then
        If picScroller.Width < picScroller.Height Then
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleWidth / 2) - 3, GripperColor3
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleWidth / 2) - 4, GripperColor2
        Else
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleHeight / 2) - 3, GripperColor3
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleHeight / 2) - 4, GripperColor2
        End If
    ElseIf m_ScrollerGripperStyle = CircleRaised Then
        If picScroller.Width < picScroller.Height Then
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleWidth / 2) - 3, GripperColor
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleWidth / 2) - 4, GripperColor2
        Else
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleHeight / 2) - 3, GripperColor
            picScroller.Circle (picScroller.ScaleWidth / 2, picScroller.ScaleHeight / 2), (picScroller.ScaleHeight / 2) - 4, GripperColor2
        End If
    ElseIf m_ScrollerGripperStyle = LineEngraved Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        If m_Orientation = Vertical Then
            For iii = 3 To (picScroller.ScaleHeight - 4) Step 4
                If iii > (picScroller.ScaleHeight - 4) Then Exit For
                picScroller.Line (3, iii)-(picScroller.ScaleWidth - 3, iii), GripperColor
                If m_Style = Graphical Then picScroller.Line (4, iii + 1)-(picScroller.ScaleWidth - 4, iii + 1), GripperColor3
                If m_Style = Graphical Then picScroller.Line (4, iii + 2)-(picScroller.ScaleWidth - 4, iii + 2), GripperColor2
            Next iii
        Else
            For iii = 3 To (picScroller.ScaleWidth - 4) Step 4
                If iii > (picScroller.ScaleWidth - 4) Then Exit For
                picScroller.Line (iii, 3)-(iii, picScroller.ScaleHeight - 3), GripperColor
                If m_Style = Graphical Then picScroller.Line (iii + 1, 4)-(iii + 1, picScroller.ScaleHeight - 4), GripperColor3
                If m_Style = Graphical Then picScroller.Line (iii + 2, 4)-(iii + 2, picScroller.ScaleHeight - 4), GripperColor2
            Next iii
        End If
    ElseIf m_ScrollerGripperStyle = LineRaised Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        If m_Orientation = Vertical Then
            For iii = 3 To (picScroller.ScaleHeight - 4) Step 4
                If iii > (picScroller.ScaleHeight - 4) Then Exit For
                picScroller.Line (3, iii)-(picScroller.ScaleWidth - 3, iii), GripperColor2
                If m_Style = Graphical Then picScroller.Line (4, iii + 1)-(picScroller.ScaleWidth - 4, iii + 1), GripperColor3
                If m_Style = Graphical Then picScroller.Line (4, iii + 2)-(picScroller.ScaleWidth - 4, iii + 2), GripperColor
            Next iii
        Else
            For iii = 3 To (picScroller.ScaleWidth - 4) Step 4
                If iii > (picScroller.ScaleWidth - 4) Then Exit For
                picScroller.Line (iii, 3)-(iii, picScroller.ScaleHeight - 3), GripperColor2
                If m_Style = Graphical Then picScroller.Line (iii + 1, 4)-(iii + 1, picScroller.ScaleHeight - 4), GripperColor3
                If m_Style = Graphical Then picScroller.Line (iii + 2, 4)-(iii + 2, picScroller.ScaleHeight - 4), GripperColor
            Next iii
        End If
    ElseIf m_ScrollerGripperStyle = WafelEngraved Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        For iii = 3 To (picScroller.ScaleHeight - 4) Step 4
            If iii > (picScroller.ScaleHeight - 4) Then Exit For
            picScroller.Line (3, iii)-(picScroller.ScaleWidth - 3, iii), GripperColor
            If m_Style = Graphical Then picScroller.Line (4, iii + 1)-(picScroller.ScaleWidth - 4, iii + 1), GripperColor3
            If m_Style = Graphical Then picScroller.Line (4, iii + 2)-(picScroller.ScaleWidth - 4, iii + 2), GripperColor2
        Next iii
        For iii = 3 To (picScroller.ScaleWidth - 4) Step 4
            If iii > (picScroller.ScaleWidth - 4) Then Exit For
            picScroller.Line (iii, 3)-(iii, picScroller.ScaleHeight - 3), GripperColor
            If m_Style = Graphical Then picScroller.Line (iii + 1, 4)-(iii + 1, picScroller.ScaleHeight - 4), GripperColor3
            If m_Style = Graphical Then picScroller.Line (iii + 2, 4)-(iii + 2, picScroller.ScaleHeight - 4), GripperColor2
        Next iii
    ElseIf m_ScrollerGripperStyle = WafelRaised Then
        GripperColor2 = ShiftColors(GripperColor, 130)
        For iii = 3 To (picScroller.ScaleHeight - 4) Step 4
            If iii > (picScroller.ScaleHeight - 4) Then Exit For
            picScroller.Line (3, iii)-(picScroller.ScaleWidth - 3, iii), GripperColor2
            If m_Style = Graphical Then picScroller.Line (4, iii + 1)-(picScroller.ScaleWidth - 4, iii + 1), GripperColor3
            If m_Style = Graphical Then picScroller.Line (4, iii + 2)-(picScroller.ScaleWidth - 4, iii + 2), GripperColor
        Next iii
        For iii = 3 To (picScroller.ScaleWidth - 4) Step 4
            If iii > (picScroller.ScaleWidth - 4) Then Exit For
            picScroller.Line (iii, 3)-(iii, picScroller.ScaleHeight - 3), GripperColor2
            If m_Style = Graphical Then picScroller.Line (iii + 1, 4)-(iii + 1, picScroller.ScaleHeight - 4), GripperColor3
            If m_Style = Graphical Then picScroller.Line (iii + 2, 4)-(iii + 2, picScroller.ScaleHeight - 4), GripperColor
        Next iii
    End If
DrawBorders:
    ' Borders
    If m_ScrollerBorderStyle <> None Then
        Dim NewBorderColor As OLE_COLOR
        If m_Enabled = True Then
            NewBorderColor = m_BarBorderColor
        Else
            NewBorderColor = m_DisabledBorderColor
        End If
        picScroller.Line (0, 0)-(picScroller.ScaleWidth, 0), NewBorderColor
        picScroller.Line (0, picScroller.ScaleHeight - 1)-(picScroller.ScaleWidth, picScroller.ScaleHeight - 1), NewBorderColor
        picScroller.Line (1, 0)-(picScroller.ScaleWidth - 1, 0), Bordercolor
        picScroller.Line (1, picScroller.ScaleHeight - 1)-(picScroller.ScaleWidth - 1, picScroller.ScaleHeight - 1), Bordercolor
        picScroller.Line (0, 1)-(0, picScroller.ScaleHeight - 1), Bordercolor
        picScroller.Line (picScroller.ScaleWidth - 1, 1)-(picScroller.ScaleWidth - 1, picScroller.ScaleHeight - 1), Bordercolor
    End If
    picScroller.Refresh
    picScroller.ScaleMode = OldScaleMode
End Sub

'=====================================================
' Draw Backgrounds
'=====================================================
Public Sub DrawPicBack(StartColor As OLE_COLOR, EndColor As OLE_COLOR, BordersColor As OLE_COLOR)
    On Error Resume Next
    DoEvents
    OldScaleMode = picBG.ScaleMode
    picBG.ScaleMode = 3
    ' Background
    If m_Style = Flat Or EndColor = &H8000000F Then
       picBG.BackColor = StartColor
       GoTo DrawBorders
    End If
    If m_Orientation = Vertical Then
        Call InitializeCol(picBG, StartColor, EndColor, False)
        For i = 0 To picBG.ScaleWidth
            NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
            picBG.Line (i, 0)-(i, picBG.ScaleHeight), NewColor
        Next
    Else
        Call InitializeCol(picBG, StartColor, EndColor, False, True)
        For i = 0 To picBG.ScaleHeight
            NewColor = RGB(RedStart + i * RedI, GreenStart + i * GreenI, BlueStart + i * BlueI)
            picBG.Line (0, i)-(picBG.ScaleWidth, i), NewColor
        Next
    End If
    DoEvents
DrawBorders:
    ' Borders
    If m_BarBorderStyle <> None Then
        picBG.Line (0, 0)-(picBG.ScaleWidth, 0), BordersColor
        picBG.Line (0, picBG.ScaleHeight - 1)-(picBG.ScaleWidth, picBG.ScaleHeight - 1), BordersColor
        picBG.Line (0, 0)-(0, picBG.ScaleHeight), BordersColor
        picBG.Line (picBG.ScaleWidth - 1, 0)-(picBG.ScaleWidth - 1, picBG.ScaleHeight), BordersColor
    End If
    picBG.Refresh
    picBG.ScaleMode = OldScaleMode
End Sub

'=====================================================
' Resize
'=====================================================
Private Sub UserControl_Resize()
    On Error Resume Next
    Dim xpos As Long
    Dim ypos As Long
    Dim Longest As Long
    ' We have to call the SetPictureVisability sub routine
    ' if we just added the control to the form
    If bAddedToIDE Then
        ConfigureControl
        SetPictureVisability
    End If
    If m_Orientation = Vertical Then
        picBG.Width = UserControl.Width
        picBG.Height = UserControl.Height
        picUP.Width = picBG.Width
        picDN.Width = picBG.Width
        picScroller.Width = picBG.Width
        If m_Style = 0 Or m_Style = 1 Then
            ' Autosize the buttons based on height when height is
            ' smaller than the height of 3 autosized buttons
            If UserControl.Height > (m_ButtonsHeightMax * 3) Then
                picUP.Height = m_ButtonsHeightMax
            Else
                ' Autosize the buttons height
                picUP.Height = UserControl.Height / 3
            End If
            picDN.Height = picUP.Height
            picScroller.Height = picUP.Height
        Else
            ' Autosize buttons based on the largest height of the pictures
            Set picTmp.Picture = m_PicScrollerVertical_UP
            picScroller.Height = picTmp.Height
            If m_PicButtonTop_UP.Height <= m_ButtonsHeightMax Then
                picUP.Height = m_PicButtonTop_UP.Height
            Else
                picUP.Height = m_ButtonsHeightMax
            End If
            If m_PicButtonBottom_UP.Height <= m_ButtonsHeightMax Then
                picDN.Height = m_PicButtonBottom_UP.Height
            Else
                picDN.Height = m_ButtonsHeightMax
            End If
        End If
    Else
        picBG.Width = UserControl.Width
        picBG.Height = UserControl.Height
        picLeft.Height = picBG.Height
        picRight.Height = picBG.Height
        picScroller.Height = picBG.Height
        If m_Style = 0 Or m_Style = 1 Then
            ' Autosize the buttons based on width when width is
            ' smaller than the width of 3 autosized buttons
            If UserControl.Width > (m_ButtonsWidthMax * 3) Then
                picLeft.Width = m_ButtonsWidthMax
            Else
                ' Autosize the buttons width
                picLeft.Width = UserControl.Width / 3
            End If
            picRight.Width = picLeft.Width
            picScroller.Width = picLeft.Width
        Else
            ' Autosize buttons based on the largest width of the pictures
            Set picTmp.Picture = m_PicScrollerHorizontal_UP
            picScroller.Width = picTmp.Width
            If m_PicButtonLeft_UP.Width <= m_ButtonsWidthMax Then
                picLeft.Width = m_PicButtonLeft_UP.Width
            Else
                picLeft.Width = m_ButtonsWidthMax
            End If
            If m_PicButtonRight_UP.Width <= m_ButtonsWidthMax Then
                picRight.Width = m_PicButtonBottom_UP.Width
            Else
                picRight.Width = m_ButtonsWidthMax
            End If
        End If
    End If
    If m_Enabled = True Then
        DrawTheBar
    Else
        DrawTheBarDisabled
    End If
    picBG.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    With UserControl
        If Not ButtonsVisible Then
            ' Buttons are not Enabled - So move them off screen
            If m_Orientation = Vertical Then
                ' Center between the Left and Right Edges
                picUP.Move (.ScaleWidth - picUP.Width) \ 2, -picUP.Height
                picDN.Move (.ScaleWidth - picUP.Width) \ 2, UserControl.ScaleHeight
                PositionScroller
            Else
                ' Between Top and Bottom edges
                picLeft.Move -picLeft.Width, ((.ScaleHeight - picLeft.Height) / 2)
                picRight.Move .ScaleWidth, ((.ScaleHeight - picRight.Height) / 2)
                PositionScroller
            End If
        Else
            ' Buttons are Enabled
            If m_Orientation = Vertical Then
                ' Center between the Left and Right Edges
                picUP.Move 0, 0 '(.ScaleWidth - picUP.Width) \ 2, 0
                picDN.Move 0, (.ScaleHeight - picDN.Height) '(.ScaleWidth - picDN.Width) \ 2, (.ScaleHeight - picDN.Height)
                PositionScroller
            Else
                ' Between Top and Bottom edges
                picLeft.Move 0, 0 '0, ((.ScaleHeight - picLeft.Height) / 2)
                picRight.Move (.ScaleWidth - (picRight.Width)), 0 '(.ScaleWidth - (picRight.Width)), ((.ScaleHeight - picRight.Height) / 2)
                PositionScroller
            End If
        End If
    End With
 
End Sub

'=============================================================================================
' Usercontrol properties
'=============================================================================================

'=====================================================
' InitProperties
'=====================================================
Private Sub UserControl_InitProperties()
    ' Bar
    m_BarBorderColor = m_def_BarBorderColor
    m_BarBackColor = m_def_BarBackColor
    Set m_PicBackVertical = Nothing
    Set m_PicBackVerticalDiasabled = Nothing
    Set m_PicBackHorizontal = LoadPicture("")
    Set m_PicBackHorizontalDisabled = LoadPicture("")
    m_BarBorderStyle = m_def_BarBorderStyle
    ' Buttons
    m_ButtonsBackColor = m_def_ButtonsBackColor
    m_ButtonsBorderColor = m_def_ButtonsBorderColor
    m_ButtonsArrowColor = m_def_ButtonsArrowColor
    m_ButtonsArrowStyle = m_def_ButtonsArrowStyle
    m_ButtonsBorderStyle = m_def_ButtonsBorderStyle
    Set m_PicButtonTop_UP = Nothing
    Set m_PicButtonTop_DOWN = Nothing
    Set m_PicButtonTop_DISABLED = Nothing
    Set m_PicButtonTop_HOOVER = LoadPicture("")
    Set m_PicButtonBottom_UP = Nothing
    Set m_PicButtonBottom_DOWN = Nothing
    Set m_PicButtonBottom_DISABLED = Nothing
    Set m_PicButtonBottom_HOOVER = LoadPicture("")
    Set m_PicButtonLeft_DOWN = Nothing
    Set m_PicButtonLeft_DOWN = Nothing
    Set m_PicButtonLeft_DISABLED = Nothing
    Set m_PicButtonLeft_HOOVER = LoadPicture("")
    Set m_PicButtonRight_UP = Nothing
    Set m_PicButtonRight_UP = Nothing
    Set m_PicButtonRight_DISABLED = Nothing
    Set m_PicButtonRight_HOOVER = LoadPicture("")
    m_ButtonsHeightMax = m_def_ButtonsHeightMax
    m_ButtonsWidthMax = m_def_ButtonsWidthMax
    m_ButtonsVisible = m_def_ButtonsVisible
    ' Scroller
    m_ScrollerBackColor = m_def_ScrollerBackColor
    m_ScrollerBorderColor = m_def_ScrollerBorderColor
    m_ScrollerGripColor = m_def_ScrollerGripColor
    m_ScrollerGripperStyle = m_def_ScrollerGripperStyle
    m_ScrollerBorderStyle = m_def_ScrollerBorderStyle
    Set m_PicScrollerVertical_DOWN = Nothing
    Set m_PicScrollerVertical_UP = Nothing
    Set m_PicScrollerVertical_DISABLED = Nothing
    Set m_PicScrollerVertical_HOOVER = LoadPicture("")
    Set m_PicScrollerHorizontal_UP = LoadPicture("")
    Set m_PicScrollerHorizontal_DOWN = LoadPicture("")
    Set m_PicScrollerHorizontal_DISABLED = LoadPicture("")
    Set m_PicScrollerHorizontal_HOOVER = LoadPicture("")
    m_ScrollerHeightMax = m_def_ScrollerHeightMax
    m_ScrollerWidthMax = m_def_ScrollerWidthMax
    m_ScrollInterval = m_def_ScrollInterval
   ' Disabled
    m_DisabledBackColor = m_def_DisabledBackColor
    m_DisabledBorderColor = m_def_DisabledBorderColor
    m_Locked = m_def_Locked
    ' Style
    m_Style = m_def_Style
    ' Values
    m_Value = m_def_Value
    m_MaxValue = m_def_MaxValue
    m_MinValue = m_def_MinValue
    m_SmallChange = m_def_SmallChange
    m_LargeChange = m_def_LargeChange
    ' Looks
    m_Orientation = m_def_Orientation
    m_Enabled = m_def_Enabled
    m_ToolTipText = m_def_ToolTipText
    ' Set Flag so we know we have to recalculate the sizes of the PictureBoxes
    bAddedToIDE = True
End Sub

'=====================================================
' ReadProperties
'=====================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Bar
    m_BarBorderColor = PropBag.ReadProperty("BarBorderColor", m_def_BarBorderColor)
    m_BarBackColor = PropBag.ReadProperty("BarBackColor", m_def_BarBackColor)
    Set m_PicBackVertical = PropBag.ReadProperty("PicBackVertical", Nothing)
    Set m_PicBackVerticalDiasabled = PropBag.ReadProperty("PicBackVerticalDiasabled", Nothing)
    Set m_PicBackHorizontal = PropBag.ReadProperty("PicBackHorizontal", Nothing)
    Set m_PicBackHorizontalDisabled = PropBag.ReadProperty("PicBackHorizontalDisabled", Nothing)
    m_BarBorderStyle = PropBag.ReadProperty("BarBorderStyle", m_def_BarBorderStyle)
   ' Buttons
    m_ButtonsBackColor = PropBag.ReadProperty("ButtonsBackColor", m_def_ButtonsBackColor)
    m_ButtonsBorderColor = PropBag.ReadProperty("ButtonsBorderColor", m_def_ButtonsBorderColor)
    m_ButtonsArrowColor = PropBag.ReadProperty("ButtonsArrowColor", m_def_ButtonsArrowColor)
    m_ButtonsArrowStyle = PropBag.ReadProperty("ButtonsArrowStyle", m_def_ButtonsArrowStyle)
    m_ButtonsBorderStyle = PropBag.ReadProperty("ButtonsBorderStyle", m_def_ButtonsBorderStyle)
    m_ButtonsHeightMax = PropBag.ReadProperty("ButtonsHeightMax", m_def_ButtonsHeightMax)
    m_ButtonsWidthMax = PropBag.ReadProperty("ButtonsWidthMax", m_def_ButtonsWidthMax)
    Set m_PicButtonTop_UP = PropBag.ReadProperty("PicButtonTop_UP", Nothing)
    Set m_PicButtonTop_DOWN = PropBag.ReadProperty("PicButtonTop_DOWN", Nothing)
    Set m_PicButtonTop_DISABLED = PropBag.ReadProperty("PicButtonTop_DISABLED", Nothing)
    Set m_PicButtonTop_HOOVER = PropBag.ReadProperty("PicButtonTop_HOOVER", Nothing)
    Set m_PicButtonBottom_UP = PropBag.ReadProperty("PicButtonBottom_UP", Nothing)
    Set m_PicButtonBottom_DOWN = PropBag.ReadProperty("PicButtonBottom_DOWN", Nothing)
    Set m_PicButtonBottom_DISABLED = PropBag.ReadProperty("m_PicButtonBottom_DISABLED", Nothing)
    Set m_PicButtonBottom_HOOVER = PropBag.ReadProperty("PicButtonBottom_HOOVER", Nothing)
    Set m_PicButtonRight_UP = PropBag.ReadProperty("PicButtonRight_UP", Nothing)
    Set m_PicButtonRight_DOWN = PropBag.ReadProperty("PicButtonRight_DOWN", Nothing)
    Set m_PicButtonRight_DISABLED = PropBag.ReadProperty("PicButtonRight_DISABLED", Nothing)
    Set m_PicButtonRight_HOOVER = PropBag.ReadProperty("PicButtonRight_HOOVER", Nothing)
    Set m_PicButtonLeft_UP = PropBag.ReadProperty("PicButtonLeft_UP", Nothing)
    Set m_PicButtonLeft_DOWN = PropBag.ReadProperty("PicButtonLeft_DOWN", Nothing)
    Set m_PicButtonLeft_DISABLED = PropBag.ReadProperty("PicButtonLeft_DISABLED", Nothing)
    Set m_PicButtonLeft_HOOVER = PropBag.ReadProperty("PicButtonLeft_HOOVER", Nothing)
    ' Scroller
    m_ScrollerBackColor = PropBag.ReadProperty("ScrollerBackColor", m_def_ScrollerBackColor)
    m_ScrollerBorderColor = PropBag.ReadProperty("ScrollerBorderColor", m_def_ScrollerBorderColor)
    m_ScrollerGripColor = PropBag.ReadProperty("ScrollerGripColor", m_def_ScrollerGripColor)
    m_ScrollerGripperStyle = PropBag.ReadProperty("ScrollerGripperStyle", m_def_ScrollerGripperStyle)
    m_ScrollerBorderStyle = PropBag.ReadProperty("ScrollerBorderStyle", m_def_ScrollerBorderStyle)
    m_ScrollerHeightMax = PropBag.ReadProperty("ScrollerHeightMax", m_def_ScrollerHeightMax)
    m_ScrollerWidthMax = PropBag.ReadProperty("ScrollerWidthMax", m_def_ScrollerWidthMax)
    Set m_PicScrollerVertical_UP = PropBag.ReadProperty("PicScrollerVertical_UP", Nothing)
    Set m_PicScrollerVertical_DOWN = PropBag.ReadProperty("PicScrollerVertical_DOWN", Nothing)
    Set m_PicScrollerVertical_DISABLED = PropBag.ReadProperty("PicScrollerVertical_DISABLED", Nothing)
    Set m_PicScrollerVertical_HOOVER = PropBag.ReadProperty("PicScrollerVertical_HOOVER", Nothing)
    Set m_PicScrollerHorizontal_UP = PropBag.ReadProperty("PicScrollerHorizontal_UP", Nothing)
    Set m_PicScrollerHorizontal_DOWN = PropBag.ReadProperty("PicScrollerHorizontal_DOWN", Nothing)
    Set m_PicScrollerHorizontal_DISABLED = PropBag.ReadProperty("PicScrollerHorizontal_DISABLED", Nothing)
    Set m_PicScrollerHorizontal_HOOVER = PropBag.ReadProperty("PicScrollerHorizontal_HOOVER", Nothing)
    m_ScrollInterval = PropBag.ReadProperty("ScrollInterval", m_def_ScrollInterval)
    ' Disabled
    m_DisabledBackColor = PropBag.ReadProperty("DisabledBackColor", m_def_DisabledBackColor)
    m_DisabledBorderColor = PropBag.ReadProperty("DisabledBorderColor", m_def_DisabledBorderColor)
    m_Locked = PropBag.ReadProperty("Locked", m_def_Locked)
   ' Values
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    m_SmallChange = PropBag.ReadProperty("SmallChange", m_def_SmallChange)
    m_LargeChange = PropBag.ReadProperty("LargeChange", m_def_LargeChange)
    ' looks
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_ButtonsVisible = PropBag.ReadProperty("ButtonsVisible", True)
    If m_Enabled = False Or m_Locked = True Then
        UserControl.Enabled = False
    Else
        UserControl.Enabled = True
    End If
    
    tmrUp.Interval = m_ScrollInterval
    tmrDown.Interval = m_ScrollInterval
    tmrLeft.Interval = m_ScrollInterval
    tmrRight.Interval = m_ScrollInterval
    UserControl_Resize
End Sub

'=====================================================
' WriteProperties
'=====================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     ' Bar
    Call PropBag.WriteProperty("BarBorderColor", m_BarBorderColor, m_def_BarBorderColor)
    Call PropBag.WriteProperty("BarBackColor", m_BarBackColor, m_def_BarBackColor)
    Call PropBag.WriteProperty("PicBackVertical", m_PicBackVertical, Nothing)
    Call PropBag.WriteProperty("PicBackVerticalDiasabled", m_PicBackVerticalDiasabled, Nothing)
    Call PropBag.WriteProperty("BarBorderStyle", m_BarBorderStyle, m_def_BarBorderStyle)
    Call PropBag.WriteProperty("PicBackHorizontal", m_PicBackHorizontal, Nothing)
    Call PropBag.WriteProperty("PicBackHorizontalDisabled", m_PicBackHorizontalDisabled, Nothing)
   ' Buttons
    Call PropBag.WriteProperty("ButtonsBackColor", m_ButtonsBackColor, m_def_ButtonsBackColor)
    Call PropBag.WriteProperty("ButtonsBorderColor", m_ButtonsBorderColor, m_def_ButtonsBorderColor)
    Call PropBag.WriteProperty("ButtonsArrowColor", m_ButtonsArrowColor, m_def_ButtonsArrowColor)
    Call PropBag.WriteProperty("ButtonsArrowStyle", m_ButtonsArrowStyle, m_def_ButtonsArrowStyle)
    Call PropBag.WriteProperty("ButtonsBorderStyle", m_ButtonsBorderStyle, m_def_ButtonsBorderStyle)
    Call PropBag.WriteProperty("ButtonsHeightMax", m_ButtonsHeightMax, m_def_ButtonsHeightMax)
    Call PropBag.WriteProperty("ButtonsWidthMax", m_ButtonsWidthMax, m_def_ButtonsWidthMax)
    Call PropBag.WriteProperty("PicButtonTop_UP", m_PicButtonTop_UP, Nothing)
    Call PropBag.WriteProperty("PicButtonTop_DOWN", m_PicButtonTop_DOWN, Nothing)
    Call PropBag.WriteProperty("PicButtonTop_DISABLED", m_PicButtonTop_DISABLED, Nothing)
    Call PropBag.WriteProperty("PicButtonTop_HOOVER", m_PicButtonTop_HOOVER, Nothing)
    Call PropBag.WriteProperty("PicButtonBottom_UP", m_PicButtonBottom_UP, Nothing)
    Call PropBag.WriteProperty("PicButtonBottom_DOWN", m_PicButtonBottom_DOWN, Nothing)
    Call PropBag.WriteProperty("m_PicButtonBottom_DISABLED", m_PicButtonBottom_DISABLED, Nothing)
    Call PropBag.WriteProperty("PicButtonBottom_HOOVER", m_PicButtonBottom_HOOVER, Nothing)
    Call PropBag.WriteProperty("ButtonsVisible", ButtonsVisible, True)
    Call PropBag.WriteProperty("PicButtonLeft_UP", m_PicButtonLeft_UP, Nothing)
    Call PropBag.WriteProperty("PicButtonLeft_DOWN", m_PicButtonLeft_DOWN, Nothing)
    Call PropBag.WriteProperty("PicButtonLeft_DISABLED", m_PicButtonLeft_DISABLED, Nothing)
    Call PropBag.WriteProperty("PicButtonLeft_HOOVER", m_PicButtonLeft_HOOVER, Nothing)
    Call PropBag.WriteProperty("PicButtonRight_UP", m_PicButtonRight_UP, Nothing)
    Call PropBag.WriteProperty("PicButtonRight_DOWN", m_PicButtonRight_DOWN, Nothing)
    Call PropBag.WriteProperty("PicButtonRight_DISABLED", m_PicButtonRight_DISABLED, Nothing)
    Call PropBag.WriteProperty("PicButtonRight_HOOVER", m_PicButtonRight_HOOVER, Nothing)
    ' Scroller
    Call PropBag.WriteProperty("ScrollerBackColor", m_ScrollerBackColor, m_def_ScrollerBackColor)
    Call PropBag.WriteProperty("ScrollerBorderColor", m_ScrollerBorderColor, m_def_ScrollerBorderColor)
    Call PropBag.WriteProperty("ScrollerGripColor", m_ScrollerGripColor, m_def_ScrollerGripColor)
    Call PropBag.WriteProperty("ScrollerGripperStyle", m_ScrollerGripperStyle, m_def_ScrollerGripperStyle)
    Call PropBag.WriteProperty("ScrollerBorderStyle", m_ScrollerBorderStyle, m_def_ScrollerBorderStyle)
    Call PropBag.WriteProperty("ScrollerHeightMax", m_ScrollerHeightMax, m_def_ScrollerHeightMax)
    Call PropBag.WriteProperty("ScrollerWidthMax", m_ScrollerWidthMax, m_def_ScrollerWidthMax)
    Call PropBag.WriteProperty("PicScrollerVertical_DOWN", m_PicScrollerVertical_DOWN, Nothing)
    Call PropBag.WriteProperty("PicScrollerVertical_UP", m_PicScrollerVertical_UP, Nothing)
    Call PropBag.WriteProperty("PicScrollerVertical_DISABLED", m_PicScrollerVertical_DISABLED, Nothing)
    Call PropBag.WriteProperty("PicScrollerVertical_HOOVER", m_PicScrollerVertical_HOOVER, Nothing)
    Call PropBag.WriteProperty("PicScrollerHorizontal_UP", m_PicScrollerHorizontal_UP, Nothing)
    Call PropBag.WriteProperty("PicScrollerHorizontal_DOWN", m_PicScrollerHorizontal_DOWN, Nothing)
    Call PropBag.WriteProperty("PicScrollerHorizontal_DISABLED", m_PicScrollerHorizontal_DISABLED, Nothing)
    Call PropBag.WriteProperty("PicScrollerHorizontal_HOOVER", m_PicScrollerHorizontal_HOOVER, Nothing)
    Call PropBag.WriteProperty("ScrollInterval", m_ScrollInterval, m_def_ScrollInterval)
    ' Disabled
    Call PropBag.WriteProperty("DisabledBackColor", m_DisabledBackColor, m_def_DisabledBackColor)
    Call PropBag.WriteProperty("DisabledBorderColor", m_DisabledBorderColor, m_def_DisabledBorderColor)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
    ' Style
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    ' Values
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("SmallChange", m_SmallChange, m_def_SmallChange)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange, m_def_LargeChange)
    ' looks
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackColor", BackColor, &HE0E0E0)

End Sub

'=====================================================
' Initialize colors for controls vertical
'=====================================================
Function InitializeCol(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, Clear As Boolean, Optional GridVertical As Boolean)
    StartCol = StartColor
    EndCol = EndColor
    RedStart = StartCol Mod 256
    RedEnd = EndCol Mod 256
    If GridVertical = False Then
        RedI = (RedEnd - RedStart) / (ctlControl.ScaleWidth)
        GreenStart = (StartCol And &HFF00FF00) / 256
        GreenEnd = (EndCol And &HFF00FF00) / 256
        GreenI = (GreenEnd - GreenStart) / (ctlControl.ScaleWidth)
        BlueStart = (StartCol And &HFFFF0000) / (65536)
        BlueEnd = (EndCol And &HFFFF0000) / (65536)
        BlueI = (BlueEnd - BlueStart) / (ctlControl.ScaleWidth)
    Else
        RedI = (RedEnd - RedStart) / (ctlControl.ScaleHeight)
        GreenStart = (StartCol And &HFF00FF00) / 256
        GreenEnd = (EndCol And &HFF00FF00) / 256
        GreenI = (GreenEnd - GreenStart) / (ctlControl.ScaleHeight)
        BlueStart = (StartCol And &HFFFF0000) / (65536)
        BlueEnd = (EndCol And &HFFFF0000) / (65536)
        BlueI = (BlueEnd - BlueStart) / (ctlControl.ScaleHeight)
    End If
    If Clear = True Then ctlControl.Cls
End Function

'=====================================================
' Initialize colors for controls horizontal
'=====================================================
Function InitializeCol2(ctlControl As Control, StartColor As OLE_COLOR, EndColor As OLE_COLOR, Clear As Boolean)
    StartCol = StartColor
    EndCol = EndColor
    RedStart = StartCol Mod 256
    RedEnd = EndCol Mod 256
    RedI = (RedEnd - RedStart) / (ctlControl.ScaleWidth)
    GreenStart = (StartCol And &HFF00FF00) / 256
    GreenEnd = (EndCol And &HFF00FF00) / 256
    GreenI = (GreenEnd - GreenStart) / (ctlControl.ScaleWidth)
    BlueStart = (StartCol And &HFFFF0000) / (65536)
    BlueEnd = (EndCol And &HFFFF0000) / (65536)
    BlueI = (BlueEnd - BlueStart) / (ctlControl.ScaleWidth)
    If Clear = True Then ctlControl.Cls
End Function

'=====================================================
' Shift colors within colorrange
'=====================================================
Private Function ShiftColors(ByVal MyColor As Long, ByVal Base As Long) As Long
    Dim R As Long, G As Long, B As Long, Delta As Long
    R = (MyColor And &HFF)
    G = ((MyColor \ &H100) Mod &H100)
    B = ((MyColor \ &H10000) Mod &H100)
    Delta = &HFF - Base
    B = Base + B * Delta \ &HFF
    G = Base + G * Delta \ &HFF
    R = Base + R * Delta \ &HFF
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255
    ShiftColors = R + 256& * G + 65536 * B
End Function

'check mouseposition
Private Function CheckMouseOver(ctlHwnd As Long) As Boolean
    Dim Pt As POINTAPI
    GetCursorPos Pt
    CheckMouseOver = (WindowFromPoint(Pt.x, Pt.y) = ctlHwnd)
End Function
'===================================================================================================
' Get,Set and Let
'===================================================================================================

'=====================================================
' Show / Hide the buttons
'=====================================================
Public Property Get ButtonsVisible() As Boolean
Attribute ButtonsVisible.VB_ProcData.VB_Invoke_Property = "Stylings"
   ButtonsVisible = m_ButtonsVisible
End Property
Public Property Let ButtonsVisible(vData As Boolean)
   m_ButtonsVisible = vData
   ' Show the nessacary Buttons
   picRight.Visible = False
   picLeft.Visible = False
   picUP.Visible = False
   picDN.Visible = False
   If m_Orientation = Vertical Then
      picUP.Visible = m_ButtonsVisible
      picDN.Visible = m_ButtonsVisible
   Else
      picRight.Visible = m_ButtonsVisible
      picLeft.Visible = m_ButtonsVisible
   End If
   UserControl_Resize
   PropertyChanged "ButtonsVisible"
End Property


'=====================================================
' Value
'=====================================================
Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = "Values"
   Value = m_Value
End Property
Public Property Let Value(nVal As Long)
    '// Make sure we are within the given range
    If nVal >= m_MinValue And nVal <= m_MaxValue Then
       m_Value = nVal
    ElseIf nVal < m_MinValue Then
       m_Value = m_MinValue
    ElseIf nVal > m_MaxValue Then
       m_Value = m_MaxValue
    End If
    '// Move The Scroller
    PositionScroller
    RaiseEvent Change
    PropertyChanged "Value"
End Property

'=====================================================
' Smallchange
'=====================================================
Public Property Get SmallChange() As Long
Attribute SmallChange.VB_ProcData.VB_Invoke_Property = "Values"
   SmallChange = m_SmallChange
End Property
Public Property Let SmallChange(nVal As Long)
   ' Check range
   If nVal >= 1 And nVal <= 32767 Then
      m_SmallChange = nVal
   Else
      MsgBox "Invalid property value", vbCritical Or vbOKOnly, "Error"
      m_SmallChange = 1
   End If
   PropertyChanged "SmallChange"
End Property

'=====================================================
' Largechange
'=====================================================
Public Property Get LargeChange() As Long
Attribute LargeChange.VB_ProcData.VB_Invoke_Property = "Values"
   LargeChange = m_LargeChange
End Property
Public Property Let LargeChange(New_Val As Long)
   ' Check range
   If New_Val >= 1 And New_Val <= 32767 Then
      m_LargeChange = New_Val
   Else
      MsgBox "Invalid property value", vbCritical Or vbOKOnly, "Error"
      m_LargeChange = 1
   End If
   PropertyChanged "LargeChange"
End Property

'=====================================================
' Orientation
'=====================================================
Public Property Get Orientation() As Orientations
    Orientation = m_Orientation
End Property
Public Property Let Orientation(New_Orientation As Orientations)
    Dim NewTop, NewLeft, NewHeight, NewWidth As Long
    If m_Orientation <> New_Orientation Then
        NewHeight = UserControl.Width
        NewWidth = UserControl.Height
        UserControl.Width = NewWidth
        UserControl.Height = NewHeight
    End If
    m_Orientation = New_Orientation
    SetPictureVisability
    UserControl_Resize
    PropertyChanged "Orientation"
End Property

'=====================================================
' MinValue
'=====================================================
Public Property Get MinValue() As Long
Attribute MinValue.VB_ProcData.VB_Invoke_Property = "Values"
    MinValue = m_MinValue
End Property
Public Property Let MinValue(vData As Long)
    Dim nDiff As Double
   nDiff = (CDbl(m_MaxValue) - CDbl(vData))
   ' The difference between Min & Max can't be larger than 2,147,483,647
   If (nDiff > nMaxValue) Then
      MsgBox "Invalid property value. Difference between Min and Max values cannot exceed " & Format$(nMaxValue, "#,###,###,###"), vbCritical Or vbOKOnly, "Error"
      Exit Property
   End If
   m_MinValue = vData
   Value = IIf(m_Value < m_MinValue, m_MinValue, m_Value)
   UserControl_Resize
   PropertyChanged "MinValue"
End Property

'=====================================================
' MaxValue
'=====================================================
Public Property Get MaxValue() As Long
Attribute MaxValue.VB_ProcData.VB_Invoke_Property = "Values"
    MaxValue = m_MaxValue
End Property
Public Property Let MaxValue(vData As Long)
    Dim nDiff As Double
   ' Make sure Max > Min - Can't be Equal
   If (vData <= m_MinValue) Then
      MsgBox "Invalid property value. Max property must be larger than Min property.", vbCritical Or vbOKOnly, "Error"
      Exit Property
   End If
   nDiff = (CDbl(vData) - CDbl(m_MinValue))
   ' The difference between Min & Max can't be larger than 2,147,483,647
   If (nDiff > nMaxValue) Then
      MsgBox "Invalid property value. Difference between Min and Max values cannot exceed " & CStr(nMaxValue), vbCritical Or vbOKOnly, "Error"
      Exit Property
   End If
   m_MaxValue = vData
   Value = IIf(m_Value > m_MaxValue, m_MaxValue, m_Value)
   UserControl_Resize
   PropertyChanged "MaxValue"
End Property


'=====================================================
' Enabled
'=====================================================
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "Stylings"
   Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Enabled = m_Enabled
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Locked() As Boolean
Attribute Locked.VB_ProcData.VB_Invoke_Property = "Stylings"
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    PropertyChanged "Locked"
End Property

'=====================================================
' Style
'=====================================================
Public Property Get Style() As Styles
    Style = m_Style
End Property
Public Property Let Style(ByVal New_Style As Styles)
    m_Style = New_Style
    PropertyChanged "Style"
    SetPictureVisability
    UserControl_Resize
End Property


'=====================================================
' Bar
'=====================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicBackVertical() As Picture
    Set PicBackVertical = m_PicBackVertical
End Property
Public Property Set PicBackVertical(ByVal New_PicBackVertical As Picture)
    Set m_PicBackVertical = New_PicBackVertical
    PropertyChanged "PicBackVertical"
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,1,1,0
Public Property Get PicBackHorizontal() As Picture
    Set PicBackHorizontal = m_PicBackHorizontal
End Property

Public Property Set PicBackHorizontal(ByVal New_PicBackHorizontal As Picture)
    Set m_PicBackHorizontal = New_PicBackHorizontal
    PropertyChanged "PicBackHorizontal"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicBackVerticalDiasabled() As Picture
    Set PicBackVerticalDiasabled = m_PicBackVerticalDiasabled
End Property
Public Property Set PicBackVerticalDiasabled(ByVal New_PicBackVerticalDiasabled As Picture)
    Set m_PicBackVerticalDiasabled = New_PicBackVerticalDiasabled
    PropertyChanged "PicBackVerticalDiasabled"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,1,1,0
Public Property Get PicBackHorizontalDisabled() As Picture
    Set PicBackHorizontalDisabled = m_PicBackHorizontalDisabled
End Property

Public Property Set PicBackHorizontalDisabled(ByVal New_PicBackHorizontalDisabled As Picture)
    On Error Resume Next
    Set m_PicBackHorizontalDisabled = New_PicBackHorizontalDisabled
    PropertyChanged "PicBackHorizontalDisabled"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get BarBorderColor() As OLE_COLOR
    BarBorderColor = m_BarBorderColor
End Property
Public Property Let BarBorderColor(ByVal New_BarBorderColor As OLE_COLOR)
    m_BarBorderColor = New_BarBorderColor
    PropertyChanged "BarBorderColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=25,0,0,0
Public Property Get BarBorderStyle() As BorderStyles
    BarBorderStyle = m_BarBorderStyle
End Property
Public Property Let BarBorderStyle(ByVal New_BarBorderStyle As BorderStyles)
    m_BarBorderStyle = New_BarBorderStyle
    PropertyChanged "BarBorderStyle"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get BarBackColor() As OLE_COLOR
    BarBackColor = m_BarBackColor
End Property
Public Property Let BarBackColor(ByVal New_BarBackColor As OLE_COLOR)
    m_BarBackColor = New_BarBackColor
    PropertyChanged "BarBackColor"
    UserControl_Resize
End Property


'=================================================================
' Scroller
'=================================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ScrollerHeightMax() As Long
Attribute ScrollerHeightMax.VB_ProcData.VB_Invoke_Property = "Values"
    ScrollerHeightMax = m_ScrollerHeightMax
End Property
Public Property Let ScrollerHeightMax(ByVal New_ScrollerHeightMax As Long)
    m_ScrollerHeightMax = New_ScrollerHeightMax
    PropertyChanged "ScrollerHeightMax"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ScrollerWidthMax() As Long
Attribute ScrollerWidthMax.VB_ProcData.VB_Invoke_Property = "Values"
    ScrollerWidthMax = m_ScrollerWidthMax
End Property
Public Property Let ScrollerWidthMax(ByVal New_ScrollerWidthMax As Long)
    m_ScrollerWidthMax = New_ScrollerWidthMax
    PropertyChanged "ScrollerWidthMax"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicScrollerVertical_UP() As Picture
    Set PicScrollerVertical_UP = m_PicScrollerVertical_UP
End Property
Public Property Set PicScrollerVertical_UP(ByVal New_PicScrollerVertical_UP As Picture)
    Set m_PicScrollerVertical_UP = New_PicScrollerVertical_UP
    PropertyChanged "PicScrollerVertical_UP"
    'SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicScrollerVertical_DOWN() As Picture
    Set PicScrollerVertical_DOWN = m_PicScrollerVertical_DOWN
End Property
Public Property Set PicScrollerVertical_DOWN(ByVal New_PicScrollerVertical_DOWN As Picture)
    Set m_PicScrollerVertical_DOWN = New_PicScrollerVertical_DOWN
    PropertyChanged "PicScrollerVertical_DOWN"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicScrollerVertical_DISABLED() As Picture
    Set PicScrollerVertical_DISABLED = m_PicScrollerVertical_DISABLED
End Property
Public Property Set PicScrollerVertical_DISABLED(ByVal New_PicScrollerVertical_DISABLED As Picture)
    Set m_PicScrollerVertical_DISABLED = New_PicScrollerVertical_DISABLED
    PropertyChanged "PicScrollerVertical_DISABLED"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicScrollerVertical_HOOVER() As Picture
    Set PicScrollerVertical_HOOVER = m_PicScrollerVertical_HOOVER
End Property

Public Property Set PicScrollerVertical_HOOVER(ByVal New_PicScrollerVertical_HOOVER As Picture)
    Set m_PicScrollerVertical_HOOVER = New_PicScrollerVertical_HOOVER
    PropertyChanged "PicScrollerVertical_HOOVER"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,1,0,0
Public Property Get PicScrollerHorizontal_UP() As Picture
    Set PicScrollerHorizontal_UP = m_PicScrollerHorizontal_UP
End Property

Public Property Set PicScrollerHorizontal_UP(ByVal New_PicScrollerHorizontal_UP As Picture)
    Set m_PicScrollerHorizontal_UP = New_PicScrollerHorizontal_UP
    PropertyChanged "PicScrollerHorizontal_UP"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,1,1,0
Public Property Get PicScrollerHorizontal_DOWN() As Picture
    Set PicScrollerHorizontal_DOWN = m_PicScrollerHorizontal_DOWN
End Property

Public Property Set PicScrollerHorizontal_DOWN(ByVal New_PicScrollerHorizontal_DOWN As Picture)
    Set m_PicScrollerHorizontal_DOWN = New_PicScrollerHorizontal_DOWN
    PropertyChanged "PicScrollerHorizontal_DOWN"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,1,1,0
Public Property Get PicScrollerHorizontal_DISABLED() As Picture
    Set PicScrollerHorizontal_DISABLED = m_PicScrollerHorizontal_DISABLED
End Property

Public Property Set PicScrollerHorizontal_DISABLED(ByVal New_PicScrollerHorizontal_DISABLED As Picture)
    Set m_PicScrollerHorizontal_DISABLED = New_PicScrollerHorizontal_DISABLED
    PropertyChanged "PicScrollerHorizontal_DISABLED"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,1,1,0
Public Property Get PicScrollerHorizontal_HOOVER() As Picture
    Set PicScrollerHorizontal_HOOVER = m_PicScrollerHorizontal_HOOVER
End Property

Public Property Set PicScrollerHorizontal_HOOVER(ByVal New_PicScrollerHorizontal_HOOVER As Picture)
    Set m_PicScrollerHorizontal_HOOVER = New_PicScrollerHorizontal_HOOVER
    PropertyChanged "PicScrollerHorizontal_HOOVER"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ScrollerGripperStyle() As GripperStyles
    ScrollerGripperStyle = m_ScrollerGripperStyle
End Property
Public Property Let ScrollerGripperStyle(ByVal New_ScrollerGripperStyle As GripperStyles)
    m_ScrollerGripperStyle = New_ScrollerGripperStyle
    PropertyChanged "ScrollerGripperStyle"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ScrollerBackColor() As OLE_COLOR
    ScrollerBackColor = m_ScrollerBackColor
End Property
Public Property Let ScrollerBackColor(ByVal New_ScrollerBackColor As OLE_COLOR)
    m_ScrollerBackColor = New_ScrollerBackColor
    PropertyChanged "ScrollerBackColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ScrollerBorderColor() As OLE_COLOR
    ScrollerBorderColor = m_ScrollerBorderColor
End Property
Public Property Let ScrollerBorderColor(ByVal New_ScrollerBorderColor As OLE_COLOR)
    m_ScrollerBorderColor = New_ScrollerBorderColor
    PropertyChanged "ScrollerBorderColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=25,0,0,0
Public Property Get ScrollerBorderStyle() As BorderStyles
    ScrollerBorderStyle = m_ScrollerBorderStyle
End Property
Public Property Let ScrollerBorderStyle(ByVal New_ScrollerBorderStyle As BorderStyles)
    m_ScrollerBorderStyle = New_ScrollerBorderStyle
    PropertyChanged "ScrollerBorderStyle"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ScrollerGripColor() As OLE_COLOR
    ScrollerGripColor = m_ScrollerGripColor
End Property
Public Property Let ScrollerGripColor(ByVal New_ScrollerGripColor As OLE_COLOR)
    m_ScrollerGripColor = New_ScrollerGripColor
    PropertyChanged "ScrollerGripColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get ScrollInterval() As Integer
Attribute ScrollInterval.VB_ProcData.VB_Invoke_Property = "Values"
    ScrollInterval = m_ScrollInterval
End Property

Public Property Let ScrollInterval(ByVal New_ScrollInterval As Integer)
    If New_ScrollInterval > 10000 Then Exit Property
    m_ScrollInterval = New_ScrollInterval
    PropertyChanged "ScrollInterval"
End Property


'=================================================================
' Top button
'=================================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonTop_UP() As Picture
    Set PicButtonTop_UP = m_PicButtonTop_UP
End Property
Public Property Set PicButtonTop_UP(ByVal New_PicButtonTop_UP As Picture)
    Set m_PicButtonTop_UP = New_PicButtonTop_UP
    PropertyChanged "PicButtonTop_UP"
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonTop_DOWN() As Picture
    Set PicButtonTop_DOWN = m_PicButtonTop_DOWN
End Property
Public Property Set PicButtonTop_DOWN(ByVal New_PicButtonTop_DOWN As Picture)
    Set m_PicButtonTop_DOWN = New_PicButtonTop_DOWN
    PropertyChanged "PicButtonTop_DOWN"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonTop_DISABLED() As Picture
    Set PicButtonTop_DISABLED = m_PicButtonTop_DISABLED
End Property
Public Property Set PicButtonTop_DISABLED(ByVal New_PicButtonTop_DISABLED As Picture)
    Set m_PicButtonTop_DISABLED = New_PicButtonTop_DISABLED
    PropertyChanged "PicButtonTop_DISABLED"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonTop_HOOVER() As Picture
    Set PicButtonTop_HOOVER = m_PicButtonTop_HOOVER
End Property
Public Property Set PicButtonTop_HOOVER(ByVal New_PicButtonTop_HOOVER As Picture)
    Set m_PicButtonTop_HOOVER = New_PicButtonTop_HOOVER
    PropertyChanged "PicButtonTop_HOOVER"
End Property


'=================================================================
' Bottom button
'=================================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonBottom_UP() As Picture
    Set PicButtonBottom_UP = m_PicButtonBottom_UP
End Property
Public Property Set PicButtonBottom_UP(ByVal New_PicButtonBottom_UP As Picture)
    Set m_PicButtonBottom_UP = New_PicButtonBottom_UP
    PropertyChanged "PicButtonBottom_UP"
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonBottom_DOWN() As Picture
    Set PicButtonBottom_DOWN = m_PicButtonBottom_DOWN
End Property
Public Property Set PicButtonBottom_DOWN(ByVal New_PicButtonBottom_DOWN As Picture)
    Set m_PicButtonBottom_DOWN = New_PicButtonBottom_DOWN
    PropertyChanged "PicButtonBottom_DOWN"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonBottom_DISABLED() As Picture
    Set PicButtonBottom_DISABLED = m_PicButtonBottom_DISABLED
End Property
Public Property Set PicButtonBottom_DISABLED(ByVal New_m_PicButtonBottom_DISABLED As Picture)
    Set m_PicButtonBottom_DISABLED = New_m_PicButtonBottom_DISABLED
    PropertyChanged "PicButtonBottom_DISABLED"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonBottom_HOOVER() As Picture
    Set PicButtonBottom_HOOVER = m_PicButtonBottom_HOOVER
End Property
Public Property Set PicButtonBottom_HOOVER(ByVal New_PicButtonBottom_HOOVER As Picture)
    Set m_PicButtonBottom_HOOVER = New_PicButtonBottom_HOOVER
    PropertyChanged "PicButtonBottom_HOOVER"
End Property


'=================================================================
' Left button
'=================================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonLeft_UP() As Picture
    Set PicButtonLeft_UP = m_PicButtonLeft_UP
End Property
Public Property Set PicButtonLeft_UP(ByVal nwPic As Picture)
    Set m_PicButtonLeft_UP = nwPic
    PropertyChanged "PicButtonLeft_UP"
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonLeft_DOWN() As Picture
    Set PicButtonLeft_DOWN = m_PicButtonLeft_DOWN
End Property
Public Property Set PicButtonLeft_DOWN(ByVal nwPic As Picture)
    Set m_PicButtonLeft_DOWN = nwPic
    PropertyChanged "PicButtonLeft_DOWN"
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonLeft_DISABLED() As Picture
    Set PicButtonLeft_DISABLED = m_PicButtonLeft_DISABLED
End Property
Public Property Set PicButtonLeft_DISABLED(ByVal New_PicButtonLeft_DISABLED As Picture)
    Set m_PicButtonLeft_DISABLED = New_PicButtonLeft_DISABLED
    PropertyChanged "PicButtonLeft_DISABLED"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonLeft_HOOVER() As Picture
    Set PicButtonLeft_HOOVER = m_PicButtonLeft_HOOVER
End Property
Public Property Set PicButtonLeft_HOOVER(ByVal New_PicButtonLeft_HOOVER As Picture)
    Set m_PicButtonLeft_HOOVER = New_PicButtonLeft_HOOVER
    PropertyChanged "PicButtonLeft_HOOVER"
End Property


'=================================================================
' Right button
'=================================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonRight_UP() As Picture
    Set PicButtonRight_UP = m_PicButtonRight_UP
End Property
Public Property Set PicButtonRight_UP(ByVal nwPic As Picture)
    Set m_PicButtonRight_UP = nwPic
    PropertyChanged "PicButtonRight_UP"
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonRight_DOWN() As Picture
    Set PicButtonRight_DOWN = m_PicButtonRight_DOWN
End Property
Public Property Set PicButtonRight_DOWN(ByVal nwPic As Picture)
    Set m_PicButtonRight_DOWN = nwPic
    PropertyChanged "PicButtonRight_DOWN"
    SetPictureVisability
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonRight_DISABLED() As Picture
    Set PicButtonRight_DISABLED = m_PicButtonRight_DISABLED
End Property
Public Property Set PicButtonRight_DISABLED(ByVal New_PicButtonRight_DISABLED As Picture)
    Set m_PicButtonRight_DISABLED = New_PicButtonRight_DISABLED
    PropertyChanged "PicButtonRight_DISABLED"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicButtonRight_HOOVER() As Picture
    Set PicButtonRight_HOOVER = m_PicButtonRight_HOOVER
End Property
Public Property Set PicButtonRight_HOOVER(ByVal New_PicButtonRight_HOOVER As Picture)
    Set m_PicButtonRight_HOOVER = New_PicButtonRight_HOOVER
    PropertyChanged "PicButtonRight_HOOVER"
End Property


'=================================================================
' Top, Bottom, Left and Right button
'=================================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonsBackColor() As OLE_COLOR
    ButtonsBackColor = m_ButtonsBackColor
End Property
Public Property Let ButtonsBackColor(ByVal New_ButtonsBackColor As OLE_COLOR)
    m_ButtonsBackColor = New_ButtonsBackColor
    PropertyChanged "ButtonsBackColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonsBorderColor() As OLE_COLOR
    ButtonsBorderColor = m_ButtonsBorderColor
End Property
Public Property Let ButtonsBorderColor(ByVal New_ButtonsBorderColor As OLE_COLOR)
    m_ButtonsBorderColor = New_ButtonsBorderColor
    PropertyChanged "ButtonsBorderColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonsHeightMax() As Long
Attribute ButtonsHeightMax.VB_ProcData.VB_Invoke_Property = "Values"
    ButtonsHeightMax = m_ButtonsHeightMax
End Property
Public Property Let ButtonsHeightMax(ByVal New_ButtonsHeightMax As Long)
    m_ButtonsHeightMax = New_ButtonsHeightMax
    PropertyChanged "ButtonsHeightMax"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonsWidthMax() As Long
Attribute ButtonsWidthMax.VB_ProcData.VB_Invoke_Property = "Values"
    ButtonsWidthMax = m_ButtonsWidthMax
End Property
Public Property Let ButtonsWidthMax(ByVal New_ButtonsWidthMax As Long)
    m_ButtonsWidthMax = New_ButtonsWidthMax
    PropertyChanged "ButtonsWidthMax"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ButtonsArrowStyle() As ArrowStyles
    ButtonsArrowStyle = m_ButtonsArrowStyle
End Property
Public Property Let ButtonsArrowStyle(ByVal New_ButtonsArrowStyle As ArrowStyles)
    m_ButtonsArrowStyle = New_ButtonsArrowStyle
    PropertyChanged "ButtonsArrowStyle"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=25,0,0,0
Public Property Get ButtonsBorderStyle() As BorderStyles
    ButtonsBorderStyle = m_ButtonsBorderStyle
End Property
Public Property Let ButtonsBorderStyle(ByVal New_ButtonsBorderStyle As BorderStyles)
    m_ButtonsBorderStyle = New_ButtonsBorderStyle
    PropertyChanged "ButtonsBorderStyle"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonsArrowColor() As OLE_COLOR
    ButtonsArrowColor = m_ButtonsArrowColor
End Property
Public Property Let ButtonsArrowColor(ByVal New_ButtonsArrowColor As OLE_COLOR)
    m_ButtonsArrowColor = New_ButtonsArrowColor
    PropertyChanged "ButtonsArrowColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get DisabledBackColor() As OLE_COLOR
    DisabledBackColor = m_DisabledBackColor
End Property
Public Property Let DisabledBackColor(ByVal New_DisabledBackColor As OLE_COLOR)
    m_DisabledBackColor = New_DisabledBackColor
    PropertyChanged "DisabledBackColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get DisabledBorderColor() As OLE_COLOR
    DisabledBorderColor = m_DisabledBorderColor
End Property
Public Property Let DisabledBorderColor(ByVal New_DisabledBorderColor As OLE_COLOR)
    m_DisabledBorderColor = New_DisabledBorderColor
    PropertyChanged "DisabledBorderColor"
    UserControl_Resize
End Property




