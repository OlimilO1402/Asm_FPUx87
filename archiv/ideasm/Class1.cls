VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunctionHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************
' function redirection class
'
' [rm_code]
'***********************************
' Thanks to:
' EBArtSoft's API HOOK Demo II
'***********************************

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Const PAGE_EXECUTE_READWRITE    As Long = &H40&

Private blnHooked       As Boolean
Private lpOldAddr       As Long

Private btOldASM(4)     As Byte
Private btReal(31)      As Byte

Private Sub Class_Initialize()
    '
End Sub

Friend Sub New_(ByVal pOldFunc As Long, ByVal pNewFunc As Long)
    Hook pOldFunc, pNewFunc
End Sub

Private Sub Class_Terminate()
    Unhook
End Sub

Public Function Unhook() As Boolean
    If blnHooked Then
        blnHooked = PutMem(lpOldAddr, VarPtr(btOldASM(0)), UBound(btOldASM) + 1)
        Unhook = blnHooked
        blnHooked = False
    End If
End Function

Public Property Get OldAddress() As Long
    OldAddress = lpOldAddr
End Property

Public Property Get Hooked() As Boolean
    Hooked = blnHooked
End Property

Private Function Hook(ByVal pAddr As Long, ByVal NewAddr As Long, Optional ProxyAddr As Long) As Boolean
    'Dim hModule     As Long
    If Not blnHooked Then
        lpOldAddr = pAddr
        Dim hFnc As Long: hFnc = pAddr
        If Not GetMem(hFnc, VarPtr(btOldASM(0)), UBound(btOldASM) + 1) Then
            Exit Function
        End If
    
        If ProxyAddr <> 0 Then
            CopyMemory btReal(0), btOldASM(0), UBound(btOldASM) + 1
            Redirect VarPtr(btReal(UBound(btOldASM) + 1)), lpOldAddr + UBound(btOldASM) + 1
            Redirect ProxyAddr, VarPtr(btReal(0))
        End If
    
        blnHooked = Redirect(hFnc, NewAddr)
        Hook = blnHooked
    End If
End Function

Public Function Redirect(ByVal OldAddr As Long, ByVal NewAddr As Long) As Boolean
    Dim btAsm(4)    As Byte
    Dim lngNewAddr  As Long

    lngNewAddr = NewAddr - OldAddr - (UBound(btAsm) + 1)

    btAsm(0) = &HE9
    CopyMemory btAsm(1), lngNewAddr, 4

    Redirect = PutMem(OldAddr, VarPtr(btAsm(0)), UBound(btAsm) + 1)

End Function

Private Function GetMem(ByVal lpAddr As Long, ByVal pData As Long, ByVal dlen As Long) As Boolean

    Dim lngOldProtect   As Long

    If VirtualProtect(ByVal lpAddr, dlen, PAGE_EXECUTE_READWRITE, lngOldProtect) = 0 Then
        Exit Function
    End If
    
    CopyMemory ByVal pData, ByVal lpAddr, dlen
    
    VirtualProtect ByVal lpAddr, dlen, lngOldProtect, lngOldProtect

    GetMem = True
End Function

Private Function PutMem(ByVal lpAddr As Long, ByVal pData As Long, ByVal dlen As Long) As Boolean

    Dim lngOldProtect   As Long

    If 0 = VirtualProtect(ByVal lpAddr, dlen, PAGE_EXECUTE_READWRITE, lngOldProtect) Then
        Exit Function
    End If

    CopyMemory ByVal lpAddr, ByVal pData, dlen
    
    VirtualProtect ByVal lpAddr, dlen, lngOldProtect, lngOldProtect

    PutMem = True
End Function
