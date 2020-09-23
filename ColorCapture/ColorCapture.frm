VERSION 5.00
Begin VB.Form frmColorCapture 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ColorCapture"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   ControlBox      =   0   'False
   Icon            =   "ColorCapture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pc 
      Height          =   1815
      Left            =   150
      ScaleHeight     =   1755
      ScaleWidth      =   3570
      TabIndex        =   4
      Top             =   1275
      Width           =   3630
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtRed 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   630
      Width           =   1215
   End
   Begin VB.TextBox txtGreen 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1350
      TabIndex        =   1
      Top             =   630
      Width           =   1215
   End
   Begin VB.TextBox txtBlue 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2625
      TabIndex        =   0
      Top             =   630
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmColorCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rgbvalue As Long
Dim pt As POINTAPI



Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpszDriver As String, ByVal lpszDevice As String, ByVal lpszOutput As Long, lpInitData As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long


Private Sub btnExit_Click()
    End
End Sub

Private Sub Timer1_Timer()
Dim DefaultMonitor
Dim HexColor As String
HexColor = "&HFFFFFFFF"
On Error GoTo 10
txtBlue.Text = ""
txtGreen.Text = ""
txtRed.Text = ""

GetCursorPos pt
rgbvalue = GetPixel(GetDC(DefaultMonitor), pt.X, pt.Y)
pc.BackColor = rgbvalue
HexColor = Hex(rgbvalue)
If rgbvalue >= 0 And rgbvalue <= 255 Then
txtBlue.Text = "Blue  = 00"
txtGreen.Text = "Green = 00"
txtRed.Text = "Red   = " + " " + HexColor
End If
If rgbvalue > 255 And rgbvalue <= 65535 Then
txtBlue.Text = "Blue  = 00"
txtGreen.Text = "Green = " + " " + Left(HexColor, 2)
txtRed.Text = "Red   = " + " " + Right(HexColor, 2)
End If
If rgbvalue > 65535 Then
txtBlue.Text = "Blue  = " + " " + Left(HexColor, 2)
txtGreen.Text = "Green = " + " " + Left(Right(HexColor, 4), 2)
txtRed.Text = "Red   = " + " " + Right(HexColor, 2)
End If

Exit Sub
' the error.
10: Exit Sub
End Sub

Private Sub Form_Load()
    Timer1.Interval = 1
End Sub
