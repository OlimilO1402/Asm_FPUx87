VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Assembler in der IDE"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkASM 
      Caption         =   "Assembler"
      Height          =   240
      Left            =   2925
      TabIndex        =   9
      Top             =   1650
      Width           =   1590
   End
   Begin VB.CommandButton cmdBench 
      Caption         =   "Benchmark"
      Height          =   315
      Left            =   2925
      TabIndex        =   8
      Top             =   1200
      Width           =   1590
   End
   Begin VB.CommandButton cmdSHR 
      Caption         =   "Right Shift"
      Height          =   315
      Left            =   2925
      TabIndex        =   7
      Top             =   675
      Width           =   1590
   End
   Begin VB.CommandButton cmdSHL 
      Caption         =   "Left Shift"
      Height          =   315
      Left            =   2925
      TabIndex        =   6
      Top             =   300
      Width           =   1590
   End
   Begin VB.TextBox txtResult 
      Height          =   285
      Left            =   1050
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   1665
   End
   Begin VB.TextBox txtBits 
      Height          =   285
      Left            =   1050
      TabIndex        =   3
      Text            =   "2"
      Top             =   675
      Width           =   1665
   End
   Begin VB.TextBox txtVal 
      Height          =   285
      Left            =   1050
      MaxLength       =   11
      TabIndex        =   1
      Text            =   "256"
      Top             =   300
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ergebnis:"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   1275
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Bits:"
      Height          =   240
      Left            =   300
      TabIndex        =   2
      Top             =   750
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Wert:"
      Height          =   240
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const BENCH_ITERS   As Long = 3000000

Private Sub chkASM_Click()
    If chkASM.Value Then
        InitFastShifting
    Else
        TermFastShifting
    End If
End Sub

Private Sub cmdBench_Click()
    Dim i               As Long
    Dim blnFSA          As Boolean
    Dim dblTmr          As Double
    Dim dblTimeWOASM    As Double
    Dim dblTimeWASM     As Double
    
    blnFSA = FastShiftingActive
    
    TermFastShifting
    
    dblTmr = Timer
    For i = 1 To BENCH_ITERS
        ShiftLeft &H10000, 5
    Next
    dblTimeWOASM = Timer - dblTmr
    
    InitFastShifting
    
    dblTmr = Timer
    For i = 1 To BENCH_ITERS
        ShiftLeft &H10000, 5
    Next
    dblTimeWASM = Timer - dblTmr
    
    If blnFSA Then
        chkASM.Value = 1
    Else
        TermFastShifting
        chkASM.Value = 0
    End If
    
    MsgBox BENCH_ITERS & " Iterationen" & vbCrLf & _
           "Ohne ASM: " & Round(dblTimeWOASM, 2) & " sec" & vbCrLf & _
           "Mit ASM: " & Round(dblTimeWASM, 2) & " sec", vbInformation
End Sub

Private Sub cmdSHL_Click()
    txtResult.Text = ShiftLeft(CLng(txtVal.Text), CLng(txtBits.Text))
End Sub

Private Sub cmdSHR_Click()
    txtResult.Text = ShiftRight(CLng(txtVal.Text), CLng(txtBits.Text))
End Sub

Private Sub Command1_Click()
    Dim v As Double: v = 2
    Dim r As Double
    Dim i As Long, n As Long: n = 1000000
    ReDim Values(0 To n) As Double
    ReDim Results0(0 To n) As Double
    ReDim Results1(0 To n) As Double
    For i = 0 To n
        Values(i) = Rnd() * 123456.789
        Results0(i) = VBA.Math.Sqr(Values(i))
    Next
    Dim dt As Single: dt = Timer
    For i = 0 To n
        Results1(i) = modShifting.FSqrt(Values(i))
    Next
    dt = Timer - dt
    
    MsgBox dt * 1000 & " ms"
    
    For i = 0 To n
        If Results0(i) <> Results1(i) Then
            MsgBox "Values are different: " & Results0(i) & " " & Results1(i)
            Exit For
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TermFastShifting
End Sub
