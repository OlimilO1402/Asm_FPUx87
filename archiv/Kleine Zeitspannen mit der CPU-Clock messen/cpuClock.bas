Attribute VB_Name = "cpuClock"
' Dieser Source stammt von http://www.activevb.de
' und kann frei verwendet werden. Für eventuelle Schäden
' wird nicht gehaftet.
'
' Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
' Ansonsten viel Spaß und Erfolg mit diesem Source!


'-------------------------------------------------------------
' CPU-clock zum Messen kleiner Zeiten nutzen
' (benötigt einen Pentium-Prozessor!)
'
' (softKUS) - X/2003
'-------------------------------------------------------------

Option Explicit

Dim cur_txt        As String
Public clk_sec     As Currency
Public clk_dmy     As Currency
Public clk_frq     As Currency
Public clk_run(19) As Currency


' Aufruf von asm-Funktionen (Deklaration für Longs)
Private Declare Function ASM_cdLong _
    Lib "user32" _
    Alias "CallWindowProcA" _
    (ByRef adr As Long, _
     ByVal PA1 As Long, _
     ByVal PA2 As Long, _
     ByVal PA3 As Long, _
     ByVal PA4 As Long) As Long

Private Declare Function API_PCounter _
    Lib "kernel32" _
    Alias "QueryPerformanceCounter" _
    (lpPerformanceCount As Currency) As Long
        
Private Declare Function API_PFrequency _
    Lib "kernel32" _
    Alias "QueryPerformanceFrequency" _
    (lpFrequency As Currency) As Long

Private Declare Sub API_Sleep _
    Lib "kernel32" Alias "Sleep" _
   (ByVal dwMilliSeconds As Long)

' clkDELAY      Sleep-Funktion
'
' AUFRUF:       clkDELAY(D1:Delay, [N1:Interval], [@C1:lpcDelay], [@C2:lpcWait])
'
' EIN:          dbl:D1  Wartezeit in sec
'               lng:N1  =0: entspricht clkWAIT
'                       >0: periodischer Aufruf von sleep/DoEvents
'
' AUS:          cur:C1  Platzhalter für loop-counter
'               cur:C2  Platzhalter für clkWAIT-loop-counter
'
' RÜCKGABE:     cur     Absolut verstrichene Zeit
'
Function clkDELAY( _
    Delay As Double, _
    Optional Interval As Long = 5, _
    Optional lpcDelay As Currency, _
    Optional lpcWait As Currency) As Double

    Dim cACT    As Currency
    Dim cRUN    As Currency
    Dim cEND    As Currency
    Dim cTMP    As Currency
    Dim cCLK(3) As Currency
    Dim dbl     As Double
    
    clkINI
    clkGET cCLK
    
    lpcDelay = 0
    lpcWait = 0
    
    If Interval = 0 Then
    ElseIf (Delay * clk_frq) > Interval Then
        API_PCounter cRUN
        cEND = cRUN + Delay * clk_frq
        cTMP = cEND - 0.02 * clk_frq
        API_PCounter cACT
        
        Do While cACT < cTMP
            lpcDelay = lpcDelay + 1
            API_Sleep Interval
            DoEvents
            API_PCounter cACT
        Loop
    End If
    
    clkGET cCLK, 2
    dbl = Delay - (cCLK(2) - cCLK(0) + clk_dmy) / clk_sec
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
' RÜCKGABE:     chr     Formatierter Text
'
Function clkFMT( _
    cur As Currency, _
    Optional cnt As Long = 2) As String
    
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
            clkFMT = Format(Sec * 1000000, txt) & " µs"
        End Select
    End If
End Function

' clkGET        Einlesen des CPU-internen Taktzählers
'
' AUFRUF:       clkGET(cur(), [NR])
'
' EIN:          cur:cur()   Array von mind. 2 Currency-Werten
'               lng:NR      Zeiger auf cur()-Element (Vorgabe=0)
'
' RÜCKGABE:     cur(NR)     Aktueller CPU-Taktzähler
'               cur(NR+1)   Differenz zw. aktuellem cur(NR) und
'                           cur(NR) beim Funktionsaufruf
'
' HINWEIS:      Um Zeiten zu messen, sollte clkGET einmal zur Initiali-
'               sierung aufgerufen werden (setzt cur(NR)) und ein weiteres
'               mal zur Berechnung der Zeitspanne (setzt cur(NR+1))
'
Function clkGET( _
    cur() As Currency, _
    Optional NR As Long) As Currency
    
    Static asm(9) As Long

    If asm(0) = 0 Then
        asm(0) = &H4C8B310F:  asm(1) = &H31FF0424
        asm(2) = &H890471FF:  asm(3) = &H4518901
        asm(4) = &H424442B:   asm(5) = &H8924141B
        asm(6) = &H51890841:  asm(7) = &HF95A580C
        asm(8) = &H10C2C01B:  asm(9) = &H0
    End If
    
    ' *****************************************************
    
    On Error Resume Next ' Fehler: ungültiges Array abfangen
    
    If UBound(cur) >= NR + 1 Then
        ASM_cdLong asm(0), VarPtr(cur(NR)), 0, 0, 0
        clkGET = cur(NR + 1)
    End If
End Function

' clkINI        Initialisieren
'
' AUFRUF:       clkINI([vl])
'
' EIN:          dbl:vl  CheckRate (Vorgabe = 0.5)
'                       Erklärung s. unten
'
' RÜCKGABE:     cur     Taktfrequenz
'
'
' setzt:        clk_sec Anzahl Takte/Sekunde
'               cur_txt Taktfrequenz als formatierter Text
'               clk_dmy Durchschnittliche Anzahl Takte, die
'                       für clkGET() benötigt werden
'
' HINWEIS:      CheckRate gibt in 1/sec die Zeitspanne an,
'               die clkINI zwischen dem Lesen des CPU-Taktzählers
'               verstreichen läßt. Je höher der Wert, desto genauer
'               ist das Ergebnis (clk_sec/clk_dmy)
'
Function clkINI(Optional chkRate As Double = 0.5) As Currency
    Dim cur(3) As Currency
    Dim cu1    As Currency
    Dim cu2    As Currency
    
    If clk_sec = 0 Then
        API_PFrequency clk_frq
        API_PCounter cu1
        cu1 = cu1 + clk_frq * chkRate
        
        clkGET cur, 0
        clkGET cur, 2
        
        Do: API_PCounter cu2
            clk_dmy = (clk_dmy + clkGET(cur, 2)) / 2
        Loop Until cu2 >= cu1
        
        clk_sec = clkGET(cur, 0) * (1 / chkRate)
        cu1 = IIf(clk_sec > 100000, 100000, 100)
        cur_txt = "Running at " & _
            Format(clk_sec / cu1, "0.00") & _
            IIf(cu1 = 100, " MHz", " GHz")
    End If
    
    clkINI = clk_sec
End Function

' clkRUN        Zum Testen von Programmen/Funktionen
'
' AUFRUF:       clkRUN([md], [cnt], [txt], [@ret], [prn])
'
' EIN:          bol:md  .F.: Zähler initialisieren (Vorgabe)
'                       .T.: Zeit messen / Ausgabe
'               lng:cnt Nummer des zu verwendenden Zähler (Vorgabe = 1)
'               chr:txt Auszugebender Text
'               bol:prn nur mit md=.T.:
'                       .T.: Ergebnis per debug.print ausgeben (Vorgabe)
'                       .F.: Ermittelte Zeit nicht ausgeben
'
' AUS:          chr:txt Textausgabe
'
' RÜCKGABE:     chr     (nur mit md=.T.)
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
Function clkRUN( _
    Optional MD As Boolean, _
    Optional cnt As Long = 1, _
    Optional txt As String, _
    Optional ret As String, _
    Optional prn As Boolean = True) As String
    
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
'               cur:Cnt     Zähler
'
' AUS:          dbl         Tatsächlich verstrichene Zeit/sec
'               cnt         Anzahl der Loops der ASM-Funktion
'                           > 1: Ergebnis ist verläßlich
'
' HINWEIS:      Die ASM-Funktion wird mit einem Array-Parameter
'               aufgerufen:
'               0: CPU-Taktzähler beim Funktionsaufruf
'               1: Differenz 0/1 und aktueller CPU-Taktzähler
'               2: Anzahl Takte, die abgewartet werden soll
'               3: Loop-Zähler. Ist der Loop-Zähler nach Ver-
'                  lassen der Funktion >1, wurde die ASM-Schleife
'                  mehr als einmal durchlaufen und ist das Ergeb-
'                  nis verläßlich. Allerdings wird der Aufwand,
'                  der hier für den Funktionsaufruf betrieben wird,
'                  nicht berücksichtigt!
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
    ASM_cdLong asm(0), VarPtr(cur(0)), 0, 0, 0
    cnt = cur(3)
    clkWAIT = cur(1) / clk_sec
End Function


