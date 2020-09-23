; MMXDigSubAddRR.asm  by  Robert Rayment 4/03/05
; NB Assumes MMX present. cpuid can be used to
; to test for MMX if wanted.

; FlatAssembler syntax

macro movab %1,%2
 {
    push dword %2
    pop dword %1
 }

format binary
Use32

PICW0         equ [ebp-4]
ptrARR0       equ [ebp-8]
PICW1         equ [ebp-12]
PICH1         equ [ebp-16]
ptrARR1       equ [ebp-20]
ptrARRRes     equ [ebp-24]
MODE          equ [ebp-28]    ; 0,1,2,3
BGL           equ [ebp-32]
WDM           equ [ebp-36]
UX            equ [ebp-40]
UY            equ [ebp-44]
ix1           equ [ebp-48]
ix2           equ [ebp-52]
iy1           equ [ebp-56]
iy2           equ [ebp-60]
ALPH          equ [ebp-64]

ixx           equ [ebp-68]
iyy           equ [ebp-72]

; For edging
;Alpha     equ [ebp-76]
;ixs       equ [ebp-80]
;iys       equ [ebp-84]
;StepAlpha equ [ebp-88]
;ix2p1     equ [ebp-92]
;iy2p1     equ [ebp-96]
;PWp1      equ [ebp-100]
;PHp1      equ [ebp-104]
;iup       equ [ebp-108]
;jup       equ [ebp-112]

    emms
    push ebp
    mov ebp,esp
    sub esp,112      ; RR To match lo stack value
    push edi
    push esi
    push ebx
    push edx

    ; Copy structure
    mov ebx,[ebp+8]
    movab PICW0,     [ebx]
    movab ptrARR0,   [ebx+4]
    movab PICW1,     [ebx+8]
    movab PICH1,     [ebx+12]
    movab ptrARR1,   [ebx+16]
    movab ptrARRRes, [ebx+20]
    movab MODE,      [ebx+24]
    movab BGL,       [ebx+28]
    movab WDM,       [ebx+32]    ; Weighting * DiffMul

    movab UX,        [ebx+36]
    movab UY,        [ebx+40]
    movab ix1,       [ebx+44]
    movab ix2,       [ebx+48]
    movab iy1,       [ebx+52]
    movab iy2,       [ebx+56]
    movab ALPH,      [ebx+60]


    ; Place WDM
    mov eax,WDM
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5      ; mm4 = WDM,WDM,WDM,WDM in words
    ; Place BGL Base GreyLevel
    mov eax,BGL
    mov edx,eax
    shl eax,16
    add eax,edx
    movd mm5,eax
    movq mm6,mm5
    punpckldq mm5,mm6      ; mm5 = BGL,BGL,BGL,BGL in words

    pxor mm7,mm7           ; mm7 = 0

; RR Can use a jump table here but this is easier to follow
    mov eax,MODE
    cmp eax,0
    jne Test1
    Call near MODE0      ; Subtract Minus
    jmp near GETOUT
Test1:
    cmp eax,1
    jne Test2
    Call near MODE1      ; Subtract Xor
    jmp near GETOUT
Test2:
    cmp eax,2
    jne Test3
    Call near MODE2      ; Add Alpha
    jmp near GETOUT
Test3:
    cmp eax,3
    jne Test4
    Call near MODE3      ; Add Edge Alpha
    jmp near GETOUT
Test4:

GETOUT:
    emms
    pop edx
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16

;###################################

MODE0:
;   Subtract Minus
;   CulR = (BaseGreyLevel + (CLng(R0) - R1) * Weighting * DiffMul)
;   CulG = (BaseGreyLevel + (CLng(G0) - G1) * Weighting * DiffMul)
;   CulB = (BaseGreyLevel + (CLng(B0) - B1) * Weighting * DiffMul)

    mov iyy,dword 1
    mov edx,dword iy1 ;iy=iy1 to iy2 ; >=1 to <= PICH0
L0:
    mov ixx,dword 1
    mov ecx,dword ix1 ;ix=ix1 to ix2 ; >=1 to <= PICW0  ; Num 4 byte chunks, 1 Long/Pixel at a time
LL0:
    Call near Get_mm1_mm2
    ; mm1 = A0  B0  G0  R0  in 4 words
    ; mm2 = A1  B1  G1  R1  in 4 words
    ; edi -> ptrARRRes offset

    psubsw mm1,mm2    ; mm1 = +/- (ARR0 - ARR1)
    pmullw mm1,mm4    ; mm1 = +/- (ABGR0 - ABRG1) * WDM
    paddsw mm1,mm5    ; mm1 = +/- (BGL + (ABGR0 - ABRG1) * WDM)
    packuswb mm1,mm7      ; mm1 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0
    movd [edi],mm1    ; mm1 to ARRRes

    inc dword ixx     ; ixx+1
    mov eax,UX
    cmp ixx,eax
    jg nexy0

    inc ecx           ; ix+1
    cmp ecx,ix2
    jle LL0
 nexy0:
    inc dword iyy     ; iyy+1
    mov eax,UY
    cmp iyy,eax
    jg outmode0

    inc edx           ; iy+1
    cmp edx,iy2
    jle L0
outmode0:
mov eax,0
RET
;-----------------------------------------------

MODE1:
;    Subtract Xor
;    CulR = (BaseGreyLevel - (R0 Xor R1) * Weighting * DiffMul)
;    CulG = (BaseGreyLevel - (G0 Xor G1) * Weighting * DiffMul)
;    CulB = (BaseGreyLevel - (B0 Xor B1) * Weighting * DiffMul)
    mov iyy,dword 1
    mov edx,dword iy1 ;iy=iy1 to iy2 ; >=1 to <= PICH0
L1:
    mov ixx,dword 1
    mov ecx,dword ix1 ;ix=ix1 to ix2 ; >=1 to <= PICW0  ; Num 4 byte chunks, 1 Long/Pixel at a time
LL1:
    Call near Get_mm1_mm2
    ; mm1 = A0  B0  G0  R0  in 4 words
    ; mm2 = A1  B1  G1  R1  in 4 words
    ; edi -> ptrARRRes offset

    pxor mm1,mm2      ; mm1 = +/- (ARR0 Xor ARR1)
    pmullw mm1,mm4    ; mm1 = +/- (ABGR0 Xor ABRG1) * WDM
    paddsw mm1,mm5    ; mm1 = +/- (BGL + (ABGR0 Xor ABRG1) * WDM)
    packuswb mm1,mm7    ; mm1 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0
    movd [edi],mm1    ; mm1 to ARRRes

    inc dword ixx     ; ixx+1

    mov eax,UX
    cmp ixx,eax
    jg nexy1

    inc ecx           ; ix+1
    cmp ecx,ix2
    jle LL1

 nexy1:
    inc dword iyy     ; iyy+1
    mov eax,UY
    cmp iyy,eax
    jg outmode1

    inc edx           ; iy+1
    cmp edx,iy2
    jle L1
outmode1:
mov eax,1
RET
;-----------------------------------------------
MODE2:
; Add Alpha
    ; Place ALPH
    mov eax,ALPH
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5      ; mm4 = ALPH,ALPH,ALPH,ALPH in words


    mov iyy,dword 1
    mov edx,dword iy1      ;iy=iy1 to iy2 ; >=1 to <= PICH0
L2:
    mov ixx,dword 1
    mov ecx,dword ix1      ;ix=ix1 to ix2 ; >=1 to <= PICW0  ; Num 4 byte chunks, 1 Long/Pixel at a time
LL2:
    Call near Get_mm1_mm2
    ; mm1 = A0  B0  G0  R0  in 4 words
    ; mm2 = A1  B1  G1  R1  in 4 words
    ; edi -> ptrARRRes offset

    psubsw mm2,mm1    ; mm2 = (ARR1 - ARR0)
    pmullw mm2,mm4    ; mm2 = (ABGR1 - ABRG0) * ALPH
    psraw mm2,7       ; mm2 = mm2\128
    paddsw mm2,mm1    ; mm2 = mm2 + ABGR0
    packuswb mm2,mm7  ; mm2 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0
    movd [edi],mm2    ; mm2 to ARRRes

    inc dword ixx     ; ixx+1

    mov eax,UX
    cmp ixx,eax
    jg nexy2

    inc ecx           ; ix+1
    cmp ecx,ix2
    jle LL2

 nexy2:
    inc dword iyy     ; iyy+1
    mov eax,UY
    cmp iyy,eax
    jg outmode2

    inc edx           ; iy+1
    cmp edx,iy2
    jle L2
outmode2:
mov eax,2
RET
;-----------------------------------------------
MODE3:
; Add Edge Alpha
Alpha      equ [ebp-76]
ixs        equ [ebp-80]
iys        equ [ebp-84]
StepAlpha  equ [ebp-88]
ix2p1      equ [ebp-92]
iy2p1      equ [ebp-96]
PWp1       equ [ebp-100]
PHp1       equ [ebp-104]
iup        equ [ebp-108]
jup        equ [ebp-112]

    mov eax,ix2
    inc eax
    mov ix2p1,eax    ; ix2+1

    mov eax,iy2
    inc eax
    mov iy2p1,eax    ; iy2+1

    mov eax,PICW1
    inc eax
    mov PWp1,eax     ; PICW1+1

    mov eax,PICH1
    inc eax
    mov PHp1,eax     ; PICH1+1

    mov eax,PICW1
    shr eax,1
    inc eax
    mov iup,eax      ; iup=(PICW1\2)+1

    mov eax,PICH1
    shr eax,1
    inc eax
    mov jup,eax      ; jup=(PICH1\2)+1

    mov eax,ALPH
    cmp eax,0
    jne SA
    mov eax,128
    jmp SADone
 SA:
    mov eax,128
    mov ebx,ALPH
    xor edx,edx
    div ebx
SADone:
    mov StepAlpha,eax      ; StepAlpha = 128 or 128\AF

    push dword ix1
    push dword ix2
    push dword iy1
    push dword iy2


TopBottom:
    xor eax,eax
    mov Alpha,eax      ; Alpha = 0

    inc eax
    mov ixs,eax        ; ixs = 1
    mov iyy,dword 1    ; iyy = 1

    mov edx,dword iy1  ; iy=iy1 to iy2 ; >=1 to <= PICH0
L3:
    mov eax,ixs
    mov ixx,eax        ; ixx = ixs

    mov ecx,dword ix1  ; ix=ix1 to ix2 ; >=1 to <= PICW0
LL3:
;;;;;;; Calculate TopBottom
    ; TOP
    ; Input:  ptrARR0, PICW0 edx=iy1, ecx=ix1
    ;         ptrARR1, PICW1 iyy, ixx
    Call near Get_mm1_mm2
    ; mm1 = A0  B0  G0  R0  in 4 words
    ; mm2 = A1  B1  G1  R1  in 4 words
    ; edi -> ptrARRRes offset ixx,iyy
    push edx

    ; Place 128-Alpha
    mov eax,128
    sub eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = 128-Alpha,128-Alpha,128-Alpha,128-Alpha in words
    pmullw mm1,mm4     ; mm1 = (128-Alpha)*ARGB0

    ; Place Alpha
    mov eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = Alpha,Alpha,Alpha,Alpha in words
    pmullw mm2,mm4     ; mm2 = Alpha*ARGB1

    paddsw mm2,mm1     ; mm2 = mm2 + mm1
    psraw mm2,7        ; mm2 = mm2\128
    packuswb mm2,mm7   ; mm2 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0
    movd [edi],mm2     ; mm2 to ARRRes
    pop edx
    ; BOTTOM temp edx=iy1,iyy for Get_mm1_mm2
    ; Input:  ptrARR0, PICW0 edx=iy1, ecx=ix1
    ;         ptrARR1, PICW1 iyy, ixx
    push edx
    push dword iyy
    mov eax,iy2p1
    sub eax,iyy
    mov edx,eax        ;edx = iy2p1-iyy

    mov eax,PHp1
    sub eax,iyy
    mov iyy,eax        ;iyy = PHp1-iyy
    Call near Get_mm1_mm2
    ; mm1 = A0  B0  G0  R0  in 4 words
    ; mm2 = A1  B1  G1  R1  in 4 words
    ; edi -> ptrARRRes offset ixx,iyy
    ; Place 128-Alpha
    mov eax,128
    sub eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = 128-Alpha,128-Alpha,128-Alpha,128-Alpha in words
    pmullw mm1,mm4     ; mm1 = (128-Alpha)*ARGB0

    ; Place Alpha
    mov eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = Alpha,Alpha,Alpha,Alpha in words
    pmullw mm2,mm4     ; mm2 = Alpha*ARGB1

    paddsw mm2,mm1     ; mm2 = mm2 + mm1
    psraw mm2,7        ; mm2 = mm2\128
    packuswb mm2,mm7   ; mm2 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0
    movd [edi],mm2     ; mm2 to ARRRes


    pop dword iyy
    pop edx
;;;;;;;
    inc dword ixx      ; ixx+1
    mov eax,UX
    cmp ixx,eax
    jg nexy3

    inc ecx            ; ix+1
    cmp ecx,ix2
    jle LL3

 nexy3:
    inc dword ix1      ; ix1 = ix1+1
    dec dword ix2      ; ix2 = ix2-1
    mov eax,ix2
    cmp eax,ix1
    jl LeftRight       ; ix2<ix1

    inc dword ixs      ; ixs+1
    mov eax,iyy
    inc eax
    cmp eax,jup
    jg LeftRight
    mov iyy,eax        ; iyy+1

    mov eax,Alpha
    add eax,StepAlpha
    cmp eax,128
    jle ny3
    mov eax,128
ny3:
    mov Alpha,eax
    inc edx            ; iy+1
    cmp edx,iy2
    jle L3

LeftRight:

    pop dword iy2
    pop dword iy1
    pop dword ix2
    pop dword ix1

    xor eax,eax
    mov Alpha,eax      ; Alpha = 0

    inc eax
    mov iys,eax        ; iys = 1

    mov ixx,dword 1    ; ixx = 1

    mov ecx,dword ix1  ; ix=ix1 to ix2 ; >=1 to <= PICW0
L4:
    mov eax,iys
    mov iyy,eax        ; iyy = iys

    mov edx,dword iy1  ; iy=iy1 to iy2 ; >=1 to <= PICH0
LL4:
;;;;;;; Calculate LeftRight

    push edx

    ; LEFT
    ; Input:  ptrARR0, PICW0 edx=iy1, ecx=ix1
    ;         ptrARR1, PICW1 iyy, ixx
    Call near Get_mm1_mm2
    ; mm1 = A0  B0  G0  R0  in 4 words
    ; mm2 = A1  B1  G1  R1  in 4 words
    ; edi -> ptrARRRes offset ixx,iyy

    ; Place 128-Alpha
    mov eax,128
    sub eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = 128-Alpha,128-Alpha,128-Alpha,128-Alpha in words
    pmullw mm1,mm4     ; mm1 = (128-Alpha)*ARGB0

    ; Place Alpha
    mov eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = Alpha,Alpha,Alpha,Alpha in words
    pmullw mm2,mm4     ; mm2 = Alpha*ARGB1

    paddsw mm2,mm1     ; mm2 = mm2 + mm1
    psraw mm2,7        ; mm2 = mm2\128
    packuswb mm2,mm7   ; mm2 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0
    movd [edi],mm2     ; mm2 to ARRRes
    pop edx

    ; RIGHT temp ecx=ix1,ixx for Get_mm1_mm2
    ; Input:  ptrARR0, PICW0 edx=iy1, ecx=ix1
    ;         ptrARR1, PICW1 iyy, ixx
    push ecx
    push dword ixx
    mov eax,ix2p1
    sub eax,ixx
    mov ecx,eax        ;ecx = ix2p1-ixx
    mov eax,PWp1
    sub eax,ixx
    mov ixx,eax        ;ixx = PWp1-ixx
    ; Input:  ptrARR0, PICW0 edx=iy1, ecx=ix1
    ;         ptrARR1, PICW1 iyy, ixx
    Call near Get_mm1_mm2
    ; mm1 = A0  B0  G0  R0  in 4 words
    ; mm2 = A1  B1  G1  R1  in 4 words
    ; edi -> ptrARRRes offset ixx,iyy
    push edx
    ; Place 128-Alpha
    mov eax,128
    sub eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = 128-Alpha,128-Alpha,128-Alpha,128-Alpha in words
    pmullw mm1,mm4     ; mm1 = (128-Alpha)*ARGB0

    ; Place Alpha
    mov eax,Alpha
    mov edx,eax
    shl eax,16
    mov ax,dx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5  ; mm4 = Alpha,Alpha,Alpha,Alpha in words
    pmullw mm2,mm4     ; mm2 = Alpha*ARGB1

    paddsw mm2,mm1     ; mm2 = mm2 + mm1
    psraw mm2,7        ; mm2 = mm2\128
    packuswb mm2,mm7   ; mm2 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0
    movd [edi],mm2     ; mm2 to ARRRes
    pop edx
    pop dword ixx
    pop ecx
;;;;;;;
    inc dword iyy      ; iyy+1
    mov eax,UY
    cmp iyy,eax
    jg nexy4

    inc edx            ; iy+1
    cmp edx,iy2
    jle LL4

 nexy4:
    inc dword iy1      ; iy1 = iy1+1
    dec dword iy2      ; iy2 = iy2-1
    mov eax,iy2
    cmp eax,iy1
    jl outmode3        ; iy2<iy1

    inc dword iys      ; iys+1

    mov eax,ixx
    inc eax
    cmp eax,iup
    jg outmode3
    mov ixx,eax        ; ixx+1


    mov eax,Alpha
    add eax,StepAlpha
    cmp eax,128
    jle ny4
    mov eax,128
ny4:
    mov Alpha,eax
    inc ecx            ; ix+1
    cmp ecx,ix2
    jle L4

outmode3:
mov eax,3
RET
;=================================
Get_mm1_mm2:
; Input:  ptrARR0, PICW0 edx=iy1, ecx=ix1
;         ptrARR1, PICW1 iyy, ixx
;         mm7 = 0
; Output: mm1 = A0  B0  G0  R0  in 4 words
;         mm2 = A1  B1  G1  R1  in 4 words
;         edi -> ptrARRRes offset

    ;mov esi,ptrARR0
    ;esi=esi+4*[ PICW0*(iy-1)+(ix-1) ]
    mov esi,ptrARR0
    mov eax,edx    ; iy
    dec eax        ; (iy-1)
    mov ebx,PICW0
push edx
xor edx,edx
    mul ebx        ; PICW0*(iy-1)  NB Res to edx:eax ie edx likely = 0
    mov ebx,ecx    ; ix
    dec ebx        ; (ix-1)
    add eax,ebx
    shl eax,2      ; 4*[ PICW0*(iy-1)+(ix-1) ]
    add esi,eax

    movd mm1,[esi]      ; mm1 = 0 0 0 0 A B R G ARR0 ' ABGR 1 pixel in lo word
    punpcklbw mm1,mm7   ; mm1 = A  B  G  R  in 4 words  mm7 = 0

    ;mov edi,ptrARR1
    ;edi=edi+4*[ (iyy-1)*PICW1+(ixx-1) ]
    mov edi,ptrARR1
    mov eax,iyy     ; iyy
    dec eax         ; (iyy-1)
    mov ebx,PICW1
xor edx,edx
    mul ebx         ; PICW1*(iyy-1)
    mov ebx,ixx     ; ixx
    dec ebx         ; (ixx-1)
    add eax,ebx
    shl eax,2       ; 4*[ (iyy-1)*PICW1+(ixx-1) ]
    add edi,eax

pop edx
    movd mm2,[edi]      ; mm2 = 0 0 0 0 A B R G ARR1 ' ABGR 1 pixel in lo word
    punpcklbw mm2,mm7   ; mm2 = A  B  G  R  ARR1 in words  mm7 = 0

    mov edi,ptrARRRes
    add edi,eax
RET
;####################################
