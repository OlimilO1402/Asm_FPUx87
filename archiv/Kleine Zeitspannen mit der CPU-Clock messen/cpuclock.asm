.586
.MODEL FLAT

.CODE
_run:

cpuclock        PROC    NEAR

init:           rdtsc
                mov     ecx,[esp+4]                     ; PA1
                push    [ecx]
                push    [ecx+4]
                mov     [ecx],eax
                mov     [ecx+4],edx
                sub     eax,[esp+4]
                sbb     edx,[esp]
                mov     [ecx+8],eax
                mov     [ecx+12],edx
                pop     eax
                pop     edx

                stc
                sbb     eax,eax
                ret     16

cpuclock        ENDP

END _run
