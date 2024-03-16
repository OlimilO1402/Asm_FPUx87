;.386
;.model flat, stdcall
;option casemap:none
;include C:\masm32\include\windows.inc
include C:\masm32\include\masm32rt.inc 
.data?
    hInstance dd ?
	
.code

DllEntry proc instance: dword, reason: dword, unused: dword
    
    .if reason == DLL_PROCESS_ATTACH
        
        mrm hInstance, instance       ; copy local to global
        mov eax, TRUE                 ; return TRUE so DLL will start
        
    .elseif reason == DLL_PROCESS_DETACH
		;
    .elseif reason == DLL_THREAD_ATTACH
		;
    .elseif reason == DLL_THREAD_DETACH
		;
    .endif
	
    ret
	
DllEntry endp

;123456789012345 * 123456789012345
;           = 15241578753238669120562399025
;uint64_max = 18446744073709551615
; int64_max =  9223372036854775807

;The manual to Intel assembler syntax can be found here:
;https://software.intel.com/content/dam/develop/public/us/en/documents/325462-sdm-vol-1-2abcd-3abcd.pdf


OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE

;Align 8

; floatinfpoint square root function
; --------========  dounble precision floating point operations  ========--------
; fsqrt S. 997
; fld   S. 965
Dbl_FSqrt proc 
    
    fld   QWORD ptr[esp+4]   ; load the first double-value from stack to register ST0
    fsqrt                    ; perform the floating-point squareroot operation on ST0
    ret   8                  ; return to caller, remove 8 bytes from stack (->stdcall)
    
Dbl_FSqrt endp

OPTION EPILOGUE:EpilogueDef
OPTION PROLOGUE:PrologueDef

End DllEntry