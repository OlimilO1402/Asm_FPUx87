Attribute VB_Name = "MCPUClock"
Option Explicit
'-------------------------------------------------------------
' CPU-clock zum Messen kleiner Zeiten nutzen
' (ben�tigt einen Pentium-Prozessor!)
'
' (softKUS) - X/2003
'-------------------------------------------------------------

' Aufruf von asm-Funktionen (Deklaration f�r Longs)
Private Declare Function CallWindowProcA Lib "user32" (ByRef adr As Long, ByVal PA1 As Long, ByVal PA2 As Long, ByVal PA3 As Long, ByVal PA4 As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)

Dim cur_txt        As String
Public clk_sec     As Currency
Public clk_dmy     As Currency
Public clk_frq     As Currency
Public clk_run(19) As Currency


' clkDELAY      Sleep-Funktion
'
' AUFRUF:       clkDELAY(D1:Delay, [N1:Interval], [@C1:lpcDelay], [@C2:lpcWait])
'
' EIN:          dbl:D1  Wartezeit in sec
'               lng:N1  =0: entspricht clkWAIT
'                       >0: periodischer Aufruf von sleep/DoEvents
'
' AUS:          cur:C1  Platzhalter f�r loop-counter
'               cur:C2  Platzhalter f�r clkWAIT-loop-counter
'
' R�CKGABE:     cur     Absolut verstrichene Zeit
'
Function clkDELAY(Delay As Double, Optional Interval As Long = 5, Optional lpcDelay As Currency = 0, Optional lpcWait As Currency = 0) As Double
    
    clkInit
    
    Dim cCLK(3) As Currency: clkGET cCLK
    
    lpcDelay = 0
    lpcWait = 0
    
    If Interval <> 0 Then
        If (Delay * clk_frq) > Interval Then
            
            Dim cRUN As Currency: QueryPerformanceCounter cRUN
            
            Dim cEND As Currency: cEND = cRUN + Delay * clk_frq
            Dim cTMP As Currency: cTMP = cEND - 0.02 * clk_frq
            
            Dim cACT As Currency: QueryPerformanceCounter cACT
            
            Do While cACT < cTMP
                
                lpcDelay = lpcDelay + 1
                
                Sleep Interval
                
                DoEvents
                QueryPerformanceCounter cACT
                
            Loop
        End If
    End If
    
    clkGET cCLK, 2
    Dim dbl As Double: dbl = Delay - (cCLK(2) - cCLK(0) + clk_dmy) / clk_sec
    
    If dbl > 0 Then clkWAIT dbl, lpcWait
    
    clkGET cCLK
    clkDELAY = cCLK(1) / clk_sec
End Function

' clkFMT        Formatieren von Zeispannen
'
' AUFRUF:       clkFMT(cur, [cnt])
'
' EIN:          cur:cur Zu formatierender Wert
'               lng:cnt Anzahl Nachkommastellen (Vorgabe=2)
'
' R�CKGABE:     chr     Formatierter Text
'
Function clkFMT(cur As Currency, Optional cnt As Long = 2) As String
    
    Dim Sec As Double
    Dim txt As String
    
    txt = "0." & String(cnt, "0")
    
    If cur < 0 Then
        
        clkFMT = cur_txt
        
    ElseIf cur <> 0 Then
        If clk_sec Then Sec = cur / clk_sec
        Select Case Log(Sec) / Log(10)
        Case Is > 1
            clkFMT = Format(Sec / 60, txt) & " m"
            
        Case Is > -1
            clkFMT = Format(Sec, txt) & " s"
            
        Case Is > -4
            clkFMT = Format(Sec * 1000, txt) & " ms"
            
        Case Else
            clkFMT = Format(Sec * 1000000, txt) & " �s"
        End Select
    End If
End Function

' clkGET        Einlesen des CPU-internen Taktz�hlers
'
' AUFRUF:       clkGET(cur(), [NR])
'
' EIN:          cur:cur()   Array von mind. 2 Currency-Werten
'               lng:NR      Zeiger auf cur()-Element (Vorgabe=0)
'
' R�CKGABE:     cur(NR)     Aktueller CPU-Taktz�hler
'               cur(NR+1)   Differenz zw. aktuellem cur(NR) und
'                           cur(NR) beim Funktionsaufruf
'
' HINWEIS:      Um Zeiten zu messen, sollte clkGET einmal zur Initiali-
'               sierung aufgerufen werden (setzt cur(NR)) und ein weiteres
'               mal zur Berechnung der Zeitspanne (setzt cur(NR+1))
'
Function clkGET(cur() As Currency, Optional NR As Long) As Currency
    
    Static asm(9) As Long
    
    If asm(0) = 0 Then
        asm(0) = &H4C8B310F
        asm(1) = &H31FF0424
        asm(2) = &H890471FF
        asm(3) = &H4518901
        asm(4) = &H424442B
        asm(5) = &H8924141B
        asm(6) = &H51890841
        asm(7) = &HF95A580C
        asm(8) = &H10C2C01B
        asm(9) = &H0
    End If
    
    ' *****************************************************
    
    On Error Resume Next ' Fehler: ung�ltiges Array abfangen
    
    If UBound(cur) >= NR + 1 Then
        CallWindowProcA asm(0), VarPtr(cur(NR)), 0, 0, 0
        clkGET = cur(NR + 1)
    End If
End Function

' clkInit        Initialisieren
'
' AUFRUF:       clkINI([vl])
'
' EIN:          dbl:vl  CheckRate (Vorgabe = 0.5)
'                       Erkl�rung s. unten
'
' R�CKGABE:     cur     Taktfrequenz
'
'
' setzt:        clk_sec Anzahl Takte/Sekunde
'               cur_txt Taktfrequenz als formatierter Text
'               clk_dmy Durchschnittliche Anzahl Takte, die
'                       f�r clkGET() ben�tigt werden
'
' HINWEIS:      CheckRate gibt in 1/sec die Zeitspanne an,
'               die clkINI zwischen dem Lesen des CPU-Taktz�hlers
'               verstreichen l��t. Je h�her der Wert, desto genauer
'               ist das Ergebnis (clk_sec/clk_dmy)
'
Function clkInit(Optional chkRate As Double = 0.5) As Currency
    Dim cur(3) As Currency
    Dim cu1    As Currency
    Dim cu2    As Currency
    
    If clk_sec = 0 Then
        QueryPerformanceFrequency clk_frq
        QueryPerformanceCounter cu1
        cu1 = cu1 + clk_frq * chkRate
        
        clkGET cur, 0
        clkGET cur, 2
        
        Do: QueryPerformanceCounter cu2
            clk_dmy = (clk_dmy + clkGET(cur, 2)) / 2
        Loop Until cu2 >= cu1
        
        clk_sec = clkGET(cur, 0) * (1 / chkRate)
        cu1 = IIf(clk_sec > 100000, 100000, 100)
        cur_txt = "Running at " & Format(clk_sec / cu1, "0.00") & IIf(cu1 = 100, " MHz", " GHz")
    End If
    
    clkInit = clk_sec
End Function

' clkRUN        Zum Testen von Programmen/Funktionen
'
' AUFRUF:       clkRUN([md], [cnt], [txt], [@ret], [prn])
'
' EIN:          bol:md  .F.: Z�hler initialisieren (Vorgabe)
'                       .T.: Zeit messen / Ausgabe
'               lng:cnt Nummer des zu verwendenden Z�hler (Vorgabe = 1)
'               chr:txt Auszugebender Text
'               bol:prn nur mit md=.T.:
'                       .T.: Ergebnis per debug.print ausgeben (Vorgabe)
'                       .F.: Ermittelte Zeit nicht ausgeben
'
' AUS:          chr:txt Textausgabe
'
' R�CKGABE:     chr     (nur mit md=.T.)
'                       txt & clkFMT(Ermittelte Zeit)
'
'
' Hinweise:     Die Funktion eignet sich vor allem, um den Zeitbedarf
'               von Programmteilen zu messen:
'
'               clkRUN
'               ... programm
'               clkRUN True
'
Function clkRUN(Optional MD As Boolean, Optional cnt As Long = 1, Optional txt As String, Optional ret As String, Optional prn As Boolean = True) As String
    
    Dim dsp As String
    Dim tmp As Long
    
    If clk_sec = 0 Then clkINI
    
    tmp = cnt * 2 - 2
    
    If tmp >= 0 And tmp < 20 Then
        If MD Then
            clkGET clk_run, tmp
            clkRUN = clkFMT(clk_run(tmp + 1))
            
            On Error Resume Next
            dsp = Left$(": ", Len(txt) * 2)
            If prn Then Debug.Print txt; dsp; clkRUN
            ret = ret & Left$(vbCrLf, Len(ret) * 2) & txt & dsp & clkRUN
            
        Else
            DoEvents
            clkGET clk_run, tmp
        End If
    End If
End Function

' clkWAIT       Sleep-Funktion
'
' AUFRUF:       clkWAIT(Zeit, [cnt])
'
' EIN:          dbl:Zeit    Abzuwartende Zeit in sec
'               cur:Cnt     Z�hler
'
' AUS:          dbl         Tats�chlich verstrichene Zeit/sec
'               cnt         Anzahl der Loops der ASM-Funktion
'                           > 1: Ergebnis ist verl��lich
'
' HINWEIS:      Die ASM-Funktion wird mit einem Array-Parameter
'               aufgerufen:
'               0: CPU-Taktz�hler beim Funktionsaufruf
'               1: Differenz 0/1 und aktueller CPU-Taktz�hler
'               2: Anzahl Takte, die abgewartet werden soll
'               3: Loop-Z�hler. Ist der Loop-Z�hler nach Ver-
'                  lassen der Funktion >1, wurde die ASM-Schleife
'                  mehr als einmal durchlaufen und ist das Ergeb-
'                  nis verl��lich. Allerdings wird der Aufwand,
'                  der hier f�r den Funktionsaufruf betrieben wird,
'                  nicht ber�cksichtigt!
'
Function clkWAIT(Sec As Double, Optional cnt As Currency) As Double
    Static asm(12) As Long
    
    If asm(0) = 0 Then
        asm(0) = &H4C8B310F:  asm(1) = &H1890424
        asm(2) = &H33045189:  asm(3) = &H184189C0
        asm(4) = &HF1C4189:   asm(5) = &H18418331
        asm(6) = &H1C518301:  asm(7) = &H1B012B00
        asm(8) = &H41890451:  asm(9) = &HC518908
        asm(10) = &H1B10412B: asm(11) = &HE3721451
        asm(12) = &H10C2
    End If
    
    ' *****************************************************

    Dim cur(3) As Currency
    
    cur(2) = Abs(Sec) * clk_sec
    CallWindowProcA asm(0), VarPtr(cur(0)), 0, 0, 0
    cnt = cur(3)
    clkWAIT = cur(1) / clk_sec
End Function
