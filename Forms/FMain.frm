VERSION 5.00
Begin VB.Form FMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Test FPU x87"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton BtnTestFSqrt 
      Caption         =   "Test Dbl_FSqrt"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://www.youtube.com/watch?v=HekvDjzDZCo
'https://www.youtube.com/watch?v=8R5rf-WggVI

Private Sub BtnTestFSqrt_Click()
    Dim r As Double
    r = GetNextRnd
    PrintF "Dbl_FSqrt(" & r & ") = " & Dbl_FSqrt(r)
End Sub

Function GetNextRnd() As Double
    Randomize Timer
    GetNextRnd = Rnd * 12345678912345.6
End Function

Private Sub PrintF(s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub
