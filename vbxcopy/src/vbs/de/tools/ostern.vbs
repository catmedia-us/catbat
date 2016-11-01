'
' $Id: ostern.vbs,v 1.1 2005/04/06 09:10:01 keilw Exp $
'
'Verwendeter Algorithmus zur Errechnung des Osterdatums (BASIC-Darstellung, alle Variablen integer oder long): 
 

  Select Case J
        Case 1700 To 1799
            M = 23
            N = 3
        Case 1800 To 1899
            M = 23
            N = 4
        Case 1900 To 2099
            M = 24
            N = 5
        Case 2100 To 2199
            M = 24
            N = 6
    End Select
    a = J Mod 19
    b = J Mod 4
    c = J Mod 7
    d = (19 * a + M) Mod 30
    e = (2 * b + 4 * c + 6 * d + N) Mod 7
    OM = 3
    OT = 22 + d + e
    If OT > 31 Then
        OT = OT - 31
        OM = 4
        If OT = 26 Then OT = 19
        If OT = 25 And d = 28 And a > 10 Then OT = 18
    End If
'Wobei 
'J=Jahreszahl (vierstellig) 
'OM= Monat des Ostersonntags 
'OT=Tag des Ostersonntags innerhalb des Monats.

'Der Algorithmus geht auf den Mathematiker und Astronomen Carl Friedrich Gauß (1777-1855) zurück.
