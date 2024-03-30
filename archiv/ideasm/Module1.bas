Attribute VB_Name = "modShifting"
Option Explicit

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal BytLength As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal module As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal addr As Long) As Long

Private Enum VirtualFreeTypes
    MEM_DECOMMIT = &H4000
    MEM_RELEASE = &H8000
End Enum

Private Enum VirtualAllocTypes
    MEM_COMMIT = &H1000
    MEM_RESERVE = &H2000
    MEM_RESET = &H8000
    MEM_LARGE_PAGES = &H20000000
    MEM_PHYSICAL = &H100000
    MEM_WRITE_WATCH = &H200000
End Enum

Private Enum VirtualAllocPageFlags
    PAGE_EXECUTE = &H10
    PAGE_EXECUTE_READ = &H20
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_EXECUTE_WRITECOPY = &H80
    PAGE_NOACCESS = &H1
    PAGE_READONLY = &H2
    PAGE_READWRITE = &H4
    PAGE_WRITECOPY = &H8
    PAGE_GUARD = &H100
    PAGE_NOCACHE = &H200
    PAGE_WRITECOMBINE = &H400
End Enum

Private Type Memory
    address                     As Long
    bytes                       As Long
End Type

Private Const IDE_ADDROF_REL    As Long = 22

' by Donald, donald@xbeat.net
Private Const SHLCode           As String = "8A4C240833C0F6C1E075068B442404D3E0C20800"
Private Const SHRCode           As String = "8A4C240833C0F6C1E075068B442404D3E8C20800"
Private Const SARCode           As String = "8A4C240833C0F6C1E075068B442404D3F8C20800"
Private Const SqrtCode          As String = "DD442404D9FAC20800"

Private m_memLHS                As Memory
Private m_memRHS                As Memory
Private m_memSAR                As Memory
Private m_memSqrt               As Memory

Private m_HookLHS            As FunctionHook
Private m_HookRHS            As FunctionHook
Private m_HookSAR            As FunctionHook
Private m_HookSqrt           As FunctionHook

Private m_blnInited             As Boolean

Public Property Get FastShiftingActive() As Boolean
    FastShiftingActive = m_blnInited
End Property

Public Sub InitFastShifting()
    If Not m_blnInited Then
        
        m_memLHS = AsmToMem(SHLCode)
        m_memRHS = AsmToMem(SHRCode)
        m_memSAR = AsmToMem(SARCode)
        m_memSqrt = AsmToMem(SqrtCode)
        
        Set m_HookLHS = MNew.FunctionHook(GetFunctionPointer(AddressOf ShiftLeft), m_memLHS.address)
        Set m_HookRHS = MNew.FunctionHook(GetFunctionPointer(AddressOf ShiftRight), m_memRHS.address)
        Set m_HookSAR = MNew.FunctionHook(GetFunctionPointer(AddressOf ShiftRightZ), m_memSAR.address)
        Set m_HookSqrt = MNew.FunctionHook(GetFunctionPointer(AddressOf FSqrt), m_memSqrt.address)
        
        'm_clsHookLHS.Hook GetFunctionPointer(AddressOf ShiftLeft), m_memLHS.address
        'm_clsHookRHS.Hook GetFunctionPointer(AddressOf ShiftRight), m_memRHS.address
        'm_clsHookSAR.Hook GetFunctionPointer(AddressOf ShiftRightZ), m_memSAR.address
        'm_clsHookSqrt.Hook GetFunctionPointer(AddressOf FSqrt), m_memSqrt.address
        
        m_blnInited = True
        
    End If
End Sub

Public Sub TermFastShifting()
    If m_blnInited Then
    
        m_HookLHS.Unhook
        m_HookRHS.Unhook
        m_HookSAR.Unhook
        m_HookSqrt.Unhook
        
        FreeMemory m_memLHS
        FreeMemory m_memRHS
        FreeMemory m_memSAR
        FreeMemory m_memSqrt
        
        Set m_clsHookLHS = Nothing
        Set m_clsHookRHS = Nothing
        Set m_clsHookSAR = Nothing
        Set m_clsHookSqrt = Nothing
        
        m_blnInited = False

    End If
End Sub

' Assembler Hex String in ausführbaren Speicher kopieren
Private Function AsmToMem(ByVal strAsm As String) As Memory
    Dim btAsm() As Byte
    Dim i       As Long
    Dim udtMem  As Memory
    
    ReDim btAsm(Len(strAsm) \ 2 - 1)

    For i = 0 To Len(strAsm) \ 2 - 1
        btAsm(i) = CByte("&H" & Mid$(strAsm, i * 2 + 1, 2))
    Next
    
    udtMem = AllocMemory(UBound(btAsm) + 1, , PAGE_EXECUTE_READWRITE)
    With udtMem
        RtlMoveMemory ByVal .address, btAsm(0), UBound(btAsm) + 1 ', 4
        VirtualProtect .address, .bytes, PAGE_EXECUTE_READ, 0
    End With
    
    AsmToMem = udtMem
End Function

Private Function FncPtr(ByVal addrof As Long) As Long
    Dim pAddr As Long
    If IsRunningInIDE_DirtyTrick() Then
        ' Wird das Programm aus der Entwicklungsumgebung heraus
        ' ausgeführt, befindet sich der eigentliche Zeiger auf
        ' eine Funktion bei (AddressOf X) + 22, AddressOf X
        ' selber zeigt nur auf einen Stub. (getestet mit VB 6)
        RtlMoveMemory pAddr, ByVal addrof + IDE_ADDROF_REL, 4
        If IsBadCodePtr(pAddr) Then pAddr = addrof
    Else
        pAddr = addrof
    End If
    
    FncPtr = pAddr
End Function

Private Function AllocMemory(ByVal bytes As Long, Optional ByVal lpAddr As Long = 0, Optional ByVal PageFlags As VirtualAllocPageFlags = PAGE_READWRITE) As Memory
    With AllocMemory
        .address = VirtualAlloc(lpAddr, bytes, MEM_COMMIT, PageFlags)
        .bytes = bytes
    End With
End Function

Private Function FreeMemory(udtMem As Memory) As Boolean
    VirtualFree udtMem.address, udtMem.bytes, MEM_DECOMMIT
    udtMem.address = 0
    udtMem.bytes = 0
End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

Public Function ShiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Dim mask As Long
    Select Case ShiftCount
    Case 1 To 31
        ' mask out bits that are pushed over the edge anyway
        mask = Pow2(31 - ShiftCount)
        ShiftLeft = Value And (mask - 1)
        ' shift
        ShiftLeft = ShiftLeft * Pow2(ShiftCount)
        ' set sign bit
        If Value And mask Then
            ShiftLeft = ShiftLeft Or &H80000000
        End If
        
    Case 0
        ' ret unchanged
        ShiftLeft = Value
    End Select
End Function

Public Function ShiftRightZ(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20001215
    Select Case ShiftCount
    Case 1 To 31
        If Value And &H80000000 Then
            ShiftRightZ = (Value And Not &H80000000) \ 2
            ShiftRightZ = ShiftRightZ Or &H40000000
            ShiftRightZ = ShiftRightZ \ Pow2(ShiftCount - 1)
        Else
            ShiftRightZ = Value \ Pow2(ShiftCount)
        End If
        
    Case 0
        ' ret unchanged
        ShiftRightZ = Value
    End Select
End Function

Public Static Function ShiftRight(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Donald, donald@xbeat.net, 20011009
    Dim lPow2(0 To 30) As Long
    Dim i As Long
    Select Case ShiftCount
    Case 0
        ShiftRight = Value
        
    Case 1 To 30
        If i = 0 Then
            lPow2(0) = 1
            For i = 1 To 30
                lPow2(i) = 2 * lPow2(i - 1)
            Next
        End If
        
        If Value And &H80000000 Then
            ShiftRight = Value \ lPow2(ShiftCount)
            If ShiftRight * lPow2(ShiftCount) <> Value Then
                ShiftRight = ShiftRight - 1
            End If
        Else
            ShiftRight = Value \ lPow2(ShiftCount)
        End If
        
    Case 31
        If Value And &H80000000 Then
            ShiftRight = -1
        Else
            ShiftRight = 0
        End If
    End Select
End Function

Public Function FSqrt(ByVal Value As Double) As Double
    FSqrt = Value ^ (1 / 2) * 2
End Function

Private Static Function Pow2(ByVal Exponent As Long) As Long
    ' by Donald, donald@xbeat.net, 20001217
    Dim alPow2(0 To 31) As Long
    Dim i As Long
    Select Case Exponent
    Case 0 To 31
        ' initialize lookup table
        If alPow2(0) = 0 Then
            alPow2(0) = 1
            For i = 1 To 30
                alPow2(i) = alPow2(i - 1) * 2
            Next
            alPow2(31) = &H80000000
        End If
        
        ' return
        Pow2 = alPow2(Exponent)
    End Select
End Function

' http://www.activevb.de/tipps/vb6tipps/tipp0347.html
Private Function IsRunningInIDE_DirtyTrick() As Boolean
    On Error GoTo NotCompiled
    Debug.Print 1 / 0
    Exit Function
NotCompiled:
    IsRunningInIDE_DirtyTrick = True
    Exit Function
End Function
