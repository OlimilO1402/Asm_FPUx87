VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Text            =   "10"
      ToolTipText     =   "Interval in ms (Schaltfläche clkDelay)"
      Top             =   120
      Width           =   1530
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Warten (clkDELAY)"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "Warten mit DoEvents"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Warten (clkWAIT)"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Warten ohne DoEvents"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "0.0001"
      ToolTipText     =   "Wartezeit"
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' (softKUS) - X/2003
'---------------------------------------------------------------------------------------
' In diesem Beispiel wird gezeigt, wie der CPU-Takzähler (erst ab
' dem Pentium verfügbar!) genutzt werden kann, um sehr kleine
' Zeitspannen zu messen. Bei der Initialisierung wird sozusagen
' als Nebeneffekt die Taktfrequenz des Rechners ermittelt, was ins-
' besondere für Programme nützlich ist, deren Code auf allen Rechnern
' gleich schnell laufen soll.
' Das Beispiel besteht aus zwei Modulen:
' 1. Form1    - Dient bloß der Demonstration
' 2. Module1  - Enthält die eigentlichen CPU-Clock-Funktionen
' Die ASM-Quellcodes sind in den .ASM-Dateien enthalten.
'---------------------------------------------------------------------------------------
' Ein Formular mit:
' Text1, Text2, Text3 (multilined), Command1(0 bis 1)
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0

Private Sub Command3_Click()
    Form2.Show
End Sub

'Private Const XYcm      As Long = 567       ' größerer Wert => größere Form
'Private Const Ftxt      As String = "asm-Beispiel CPU-Taktzähler"

Private Sub Form_Load()
    'Text3 = clkFMT(-clkInit(1)) & vbCrLf
    Command1.Caption = "Warten (clkWAIT)"
    Command1.ToolTipText = "Warten ohne DoEvents"
    Command2.Caption = "Warten (clkDELAY)"
    Command2.ToolTipText = "Warten mit DoEvents"
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text3.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then
        Text3.Move L, T, W, H
    End If
End Sub

Private Sub Command1_Click()
    Dim txt      As String
    Dim lpcWait  As Currency
    Dim dbl      As Double
    
    On Error Resume Next
    
    dbl = Val(Replace(Text1, ",", "."))
    
    If dbl = 0 Then
        Beep
    Else
        dbl = clkWAIT(dbl, lpcWait)
        
        txt = txt & "Loopcount wait:   " & lpcWait * 10000
        'If Index = 1 Then txt = txt & " (" & clkFMT(lpcWait / clk_sec) & ")"
        txt = txt & vbCrLf & "Exakte Wartezeit: " & Format(dbl, "0.0000000000")
        Text3 = Text3 & txt & vbCrLf & vbCrLf
        Text3.SelStart = Len(Text3)
    End If
    
    Text3.SetFocus
End Sub

Private Sub Command2_Click()
    Dim txt      As String
    Dim lpcDelay As Currency
    Dim lpcWait  As Currency
    Dim dbl      As Double
    
    On Error Resume Next
    
    dbl = Val(Replace(Text1, ",", "."))
    
    If dbl = 0 Then
        Beep
    Else
        dbl = clkDELAY(dbl, Val(Text2), lpcDelay, lpcWait)
        
        txt = "Loopcount delay:  " & lpcDelay & vbCrLf
        txt = txt & "Loopcount wait:   " & lpcWait * 10000
        txt = txt & " (" & clkFMT(lpcWait / clk_sec) & ")"
        txt = txt & vbCrLf & "Exakte Wartezeit: " & Format(dbl, "0.0000000000")
        Text3 = Text3 & txt & vbCrLf & vbCrLf
        Text3.SelStart = Len(Text3)
    End If
    
    Text3.SetFocus
End Sub

