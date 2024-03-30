.586
.MODEL FLAT

.CODE
_run:

init:           rdtsc
                mov     ecx,[esp+4]                    ; PA1
                mov     [ecx],eax
                mov     [ecx+4],edx
                xor     eax,eax
                mov     [ecx+24],eax
                mov     [ecx+28],eax
                                                
init_1:         rdtsc
                add     DWord Ptr [ecx+24],1
                adc     DWord Ptr [ecx+28],0
                sub     eax,[ecx]
                sbb     edx,[ecx+4]
                mov     [ecx+8],eax
                mov     [ecx+12],edx
                sub     eax,[ecx+16]
                sbb     edx,[ecx+20]
                jc      init_1
                ret     16

END _run
