VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About the programmer..."
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10665
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   711
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "frmAbout.frx":BC6A2
      Left            =   8400
      List            =   "frmAbout.frx":BCC82
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox FontPic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   45
      Picture         =   "frmAbout.frx":BD73A
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   767
      TabIndex        =   1
      Top             =   1845
      Visible         =   0   'False
      Width           =   11505
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9360
      Top             =   600
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   3000
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   3480
      Width           =   4275
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Programmer:        Donckers Frank
'                    DarkManSoft@Gmail.com
'
' Description:       Control Custom Scrollbar

'=====================================================
' API BITBILD
'=====================================================
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'=====================================================
' Declaration
'=====================================================
Dim intSY(499) As Integer
Dim strPhrase3 As String
Dim intXpos As Integer
Dim intCH As Integer
Dim intYpos As Integer
Dim intTT As Integer
Dim intMaxX As Integer
Dim intMaxX2 As Integer
Dim XX As Integer


'=====================================================
' Enable timer for texteffect
'=====================================================
Private Sub Form_Activate()
    Timer1.Enabled = True
End Sub

'=====================================================
' Form Load
'=====================================================
Private Sub Form_Load()
    strPhrase3 = "Darkman ScrollBar"
    strPhrase3 = UCase(strPhrase3)
    ' Fill array with positions
    For XX = 0 To 499
        intSY(XX) = List1.List(XX)
    Next XX
    intTT = 40
End Sub

'=====================================================
' Timer for texteffect
'=====================================================
Private Sub Timer1_Timer()
    ' Copy the letters from FontPic2 to pic1 on position
    For XX = 1 To Len(strPhrase3)
        BitBlt Pic1.hDC, 12 + (XX * 13), intSY(XX + intTT - 1), 13, 22, FontPic2.hDC, 0, 0, vbSrcCopy
    Next XX
    For XX = 1 To Len(strPhrase3)
        intCH = Asc(Mid(strPhrase3, XX, 1))
        intCH = intCH - 32
        BitBlt Pic1.hDC, 12 + (XX * 13), intSY(XX + intTT), 13, 22, FontPic2.hDC, intCH * 13, 0, vbSrcCopy
    Next XX
    intTT = intTT + 1
    If intTT > 310 Then intTT = 1
    Pic1.Refresh
End Sub
