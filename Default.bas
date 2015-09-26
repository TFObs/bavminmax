Attribute VB_Name = "mdlDefault"
Option Explicit

Public Declare Function WritePrivateProfileString Lib _
        "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpString As Any, ByVal _
        lpFileName As String) As Long
        
Public Declare Function GetPrivateProfileString Lib _
        "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal _
        lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize _
        As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileSection Lib _
        "kernel32" Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpString As _
        String, ByVal lpFileName As String) As Long
        
Public Declare Function GetPrivateProfileSection Lib _
        "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpReturnedString _
        As String, ByVal nSize As Long, ByVal lpFileName _
        As String) As Long
        
Public Declare Function InternetGetConnectedState Lib _
  "wininet.dll" (ByRef lpdwFlags As Long, _
  ByVal dwReserved As Long) As Long
  
Public Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" (ByVal hWnd As Long, _
  ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long

  
Public Declare Function InternetDial Lib "wininet.dll" ( _
  ByVal hwndParent As Long, _
  ByVal lpszConiID As String, _
  ByVal dwFlags As Long, _
  ByRef hCon As Long, _
  ByVal dwReserved As Long) As Long

Public Const DIAL_FORCE_ONLINE = 1
Public Const DIAL_FORCE_UNATTENDED = 2
Public dicStbld As Dictionary
    
    Public datei, pfad
    Public pi
    Public x
    Public sSQL
    Public Database
    Public sorter As Boolean
    Public SternName As String, maxSternLen As Byte
    Public coltrigger
    Public minMag_max As Double, minMag_min As Double
    
'Deklaration: Globale Form API-Konstanten
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const HWND_TOPMOST As Long = -1&
Public Const HWND_NOTOPMOST As Long = -2&

'Deklaration: Globale Form API-Funktionen
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function floatwindow(hWnd)
floatwindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Function

Public Function unfloatwindow(hWnd)
unfloatwindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Function
Public Sub INISetValue(ByVal Path$, ByVal Sect$, ByVal Key$, _
                        ByVal Value$)
  Dim result&
    'Wert schreiben
    result = WritePrivateProfileString(Sect, Key, Value, Path)
End Sub
 
Public Function INIGetValue(ByVal Path$, ByVal Sect$, ByVal Key$) _
                             As String
  Dim result&, Buffer$
    'Wert lesen
    Buffer = Space$(32)
    result = GetPrivateProfileString(Sect, Key, vbNullString, _
                                     Buffer, Len(Buffer), Path)
    INIGetValue = Left$(Buffer, result)
End Function
Public Function CheckInetConnection(ByVal hWnd As Long)
Dim result
result = CBool(InternetGetConnectedState(0, 0))
    
    If result = False Then
    'Testen oder Herstellen der Internetverbindung
    result = RASConnect(hWnd)
    End If
    
    If result = False Then
      MsgBox "Internetverbindung konnte nicht aufgebaut werden," & _
      vbCrLf & "bitte überprüfen Sie Ihre Einstellungen", vbCritical, "Keine Verbindung..."
    End If
    CheckInetConnection = result
End Function

' Online-Verbindung starten
Public Function RASConnect(ByVal hWnd As Long, _
  Optional ByVal sDFÜName As String = "", _
  Optional ByVal bAutoStart As Boolean = False) As Boolean

  Dim conID As Long
  Dim nFlags As Long

  nFlags = IIf(bAutoStart, DIAL_FORCE_UNATTENDED, DIAL_FORCE_ONLINE)
  InternetDial hWnd, sDFÜName, nFlags, conID, 0

  RASConnect = (conID <> 0)
End Function

Public Sub DefaultWerte()
Dim result
datei = App.Path & "\prog.ini"
Call INISetValue(datei, "ort", "Breite", 50)
Call INISetValue(datei, "ort", "Länge", 10)
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "bürgerlich")
Call INISetValue(datei, "filter", "höhe", 10)
Call INISetValue(datei, "filter", "Azimut_u", 0)
Call INISetValue(datei, "filter", "Azimut_o", 360)
Call INISetValue(datei, "filter", "BProg", "alle")
Call INISetValue(datei, "filter", "Typ", "alle")
Call INISetValue(datei, "filter", "Monddist", 90)
Call INISetValue(datei, "filter", "Sternbild", "alle")
Call INISetValue(datei, "Standard", "höhe", 10)
Call INISetValue(datei, "Standard", "Azimut_u", 0)
Call INISetValue(datei, "Standard", "Azimut_o", 360)
Call INISetValue(datei, "Standard", "BProg", "alle")
Call INISetValue(datei, "Standard", "Typ", "alle")
Call INISetValue(datei, "Standard", "Monddist", 90)
Call INISetValue(datei, "Standard", "Sternbild", "alle")
Call INISetValue(datei, "filter", "minMag_Max", 18)
Call INISetValue(datei, "filter", "minMag_Min", 18)

End Sub

Public Function StripDuplicates(ByVal Value As Variant, _
Optional ByVal sChar As String = " ") As Variant

If IsNull(Value) Then
    StripDuplicates = Null
Else
    If Value = String$(Len(Value), sChar) Then
        Value = sChar
    Else
    While Len(Value) > 0 And InStr(1, Value, sChar & sChar) > 0
        Value = Replace(Value, sChar & sChar, sChar)
    Wend
    End If
    StripDuplicates = Value
End If
    
        
    
End Function


'==============================================================================================================
'=======================ASTRONOMISCHE BERECHNUNGSFUNKTIONEN====================================================
'==============================================================================================================
'Berechnung des Julianischen Datums
'Uhrzeit wird in Stunden übergeben
Public Function JulDat(Tag, Monat, Jahr, Optional Uhrzeit) As Double
 
    Dim a As Double
    Dim b As Integer
    a = 10000# * Jahr + 100# * Monat + Tag
    If (a < -47120101) Then MsgBox "Warnung: Datum jenseits der Berechnungsgrenze"
    
    If (Monat <= 2) Then
        Monat = Monat + 12
        Jahr = Jahr - 1
    End If
    
        b = Fix(Jahr / 400) - Fix(Jahr / 100) + Fix(Jahr / 4)
    
    a = 365# * Jahr + 1720996.5
    
    JulDat = a + b + Fix(30.6001 * (Monat + 1)) + Tag

    If Not IsMissing(Uhrzeit) Then
        JulDat = JulDat + (Uhrzeit) / 24
    End If
    
End Function

'Berechnung der Gregor. Datums aus dem Julianischen Datum
Public Function JulinDat(JulDat)
Dim a, b, c, D, e, F
Dim Tag, Monat, Jahr
Dim Uhrzeit, Datum, Ergebnis

    a = Int(JulDat + 0.5)
        If a < 2299161 Then
            c = a + 1524
        ElseIf a >= 2299161 Then
            b = Int((a - 1867216.25) / 36524.25)
            c = a + b - Int(b / 4) + 1525
        End If

D = Int((c - 122.1) / 365.25)
e = Int(365.25 * D)
F = Int((c - e) / 30.6001)

Tag = (c - e - Int(30.6001 * F) + (JulDat + 0.5 - a))
Monat = (F - 1 - 12 * Int(F / 14))
Jahr = (D - 4715 - Int((7 + Monat) / 10))
If Jahr > 2100 Then
MsgBox "Datum jenseits der Berechnungsgrenze..", vbCritical, "Jahr > 2100"
Exit Function
End If
Uhrzeit = (Tag) - Int(Tag)
Datum = CDbl(CDate(CStr((CStr(Tag - Uhrzeit) + "." + CStr(Monat) + "." + CStr(Jahr)))))
Ergebnis = Datum + Uhrzeit

' Ausgabe als Double Integer
JulinDat = Ergebnis
End Function
 
'Berechnung der Sternzeit
Public Function STZT(Tag, Monat, Jahr, _
ByVal Uhrzeit As Double, ByVal länge As Double)
Dim JDo As Double, ST As Double

    'Berechnung des Julianischen Datums für 0h Weltzeit
    JDo = JulDat(Tag, Monat, Jahr)

    'Berechnung der Ortssternzeit
    'Uhrzeit muß in Stundenbruchteilen eingegeben werden
    'Sternzeit wird in Tagesbruchteilen berechnet

    ST = 6.66452 + 0.0657098244 * (JDo - 2451544.5) + 1.0027379093 * Uhrzeit
    If ST < 0 Then
        ST = ((24 + ST / 24) - Int(24 + ST / 24))
        Else: ST = ((ST / 24) - Int(ST / 24))
    End If
    
    STZT = ST + (länge / 15 / 24) 'Ausgabe in Tagesbruchteilen
End Function

'Berechnung des Stundenwinkels
Public Function stdw(ByVal RA As Double, ByVal ST As Double)
Dim winkel As Double
    'Umwandlung von RA in Tagesbruchteile
    If RA = 0 Then
        RA = 1
    Else: RA = RA / 24
    End If

winkel = ST - RA
stdw = winkel


'Normieren auf 24h (immer in Tagesbruchteilen!)
If stdw < 0 Then
 stdw = 1 + stdw
End If
 
End Function

'arccos-Funktion muß definiert werden...
Public Function arccos(x)
Dim acsn
    acsn = (Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1))
    arccos = acsn
End Function

'arcsin-Funktion muß definiert werden...
Public Function arcsin(x)
Dim asn
    asn = Atn(x / Sqr(-x * x + 1)) 'achtung!!! normalerweise 1/pi
    arcsin = asn
End Function

'Berechnung der Höhe zum Berechneten Stundenwinkel
Public Function Hoehe(ha As Double, breite As Double, DEC As Double)
Dim Ergebnis
'ha=Stundenwinkel, wird normiert auf 24h übergeben
    Ergebnis = arcsin((Cos(((ha * 360)) * (4 * Atn(1) / 180)) * Cos(breite * (4 * Atn(1) / 180)) * Cos(DEC * (4 * Atn(1) / 180))) + _
    Sin(breite * (4 * Atn(1) / 180)) * Sin(DEC * (4 * Atn(1) / 180))) * 1 / (4 * Atn(1) / 180)
    Hoehe = Ergebnis
End Function

'Berechnung des Azimut
Public Function Azimut(ho As Double, stw As Double, breite As Double, DEC As Double)
Dim Ergebnis
'Übergabe des Stundenwinkels in Tagesbruchteilen
Ergebnis = arccos(((Cos((stw * 360) * (4 * Atn(1) / 180)) * Sin(breite * (4 * Atn(1) / 180)) * Cos(DEC * (4 * Atn(1) / 180))) - _
    Cos(breite * (4 * Atn(1) / 180)) * Sin(DEC * (4 * Atn(1) / 180))) / Cos(ho * (4 * Atn(1) / 180))) * 1 / (4 * Atn(1) / 180)
    
Azimut = Ergebnis

End Function

'Berechnung der Luftmasse
Public Function airmass(ByVal breite As Double, ByVal DEC As Double, ByVal stdw As Double)
Dim zendist, lMass
zendist = 1 / ((Cos((stdw * 360) * (4 * Atn(1) / 180)) * Cos(breite * (4 * Atn(1) / 180)) * Cos(DEC * (4 * Atn(1) / 180))) + _
    Sin(breite * (4 * Atn(1) / 180)) * Sin(DEC * (4 * Atn(1) / 180)))
    
lMass = zendist - 0.0018167 * (zendist - 1) - 0.002875 * (zendist - 1) ^ 2 _
- 0.0008083 * (zendist - 1) ^ 3

airmass = lMass
End Function

'Berechnung des Sonnenauf- und -untergangs
'Formeln gemäß Dr.R. Brodbeck, siehe http://lexikon.astronomie.info/zeitgleichung/
Public Function AufUnter(daten, Jahr) ', zeiger)
Dim breite As Double, länge As Double
Dim b, T, jetzt, ds
Dim sonnenaufgang As Double, sonnenuntergang As Double
Dim zeitdifferenz As Double, zeitgleichung As Double
Dim aufgangUT As Double, untgangUT As Double
Dim AU(2)
'Ermitteln der geogr. Koordinaten aus Registry
länge = INIGetValue(App.Path & "\Prog.ini", "Ort", "Länge") 'Berlin =13.366666
breite = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite") 'Berlin =52.55


'Differenz der Tage zum 1.1.eines Jahres
't=datediff("d","1-1",daten)siehe auch cmdListe_click()!!!

T = DateDiff("d", "1-1", daten)
'Breitengrad im Bogenmass
b = breite * (4 * Atn(1)) / 180

'Deklination der Sonne
ds = 0.40954 * Sin(0.0172 * (T - 79.34974))
'Um eine Höhe von 50' über dem Horizont zu erreichen:

jetzt = INIGetValue(App.Path & "\Prog.ini", "Auf- Untergang", "Dämmerung")

'Berücksichtigung des Filters für Sonnenaufgang
Select Case jetzt
    Case Is = "": sonnenaufgang = -0.0145 '-0.0145 = 50'/
    Case Is = "bürgerlich": sonnenaufgang = -6 * (4 * Atn(1)) / 180
    Case Is = "nautisch": sonnenaufgang = -12 * (4 * Atn(1)) / 180
    Case Is = "astronomisch": sonnenaufgang = -18 * (4 * Atn(1)) / 180
    Case Else: sonnenaufgang = -0.0145
End Select

'Zeitdifferenz
If ((Sin(sonnenaufgang) - Sin(b) * Sin(ds)) / _
(Cos(b) * Cos(ds))) <= -1 Then
 AU(0) = 25: AU(1) = 25
 AufUnter = AU
 Exit Function
 Else
zeitdifferenz = 12 * arccos((Sin(sonnenaufgang) - Sin(b) * Sin(ds)) / _
(Cos(b) * Cos(ds))) / (4 * Atn(1))
End If

sonnenaufgang = 12 - zeitdifferenz
sonnenuntergang = 12 + zeitdifferenz

'Zeitgleichung: WOZ - MOZ
zeitgleichung = -0.1752 * Sin(0.03343 * T + 0.5474) - 0.134 * Sin(0.018234 * T - 0.1939)

'Auf-und Untergang in UT
aufgangUT = sonnenaufgang - zeitgleichung - länge / 15
untgangUT = sonnenuntergang - zeitgleichung - länge / 15

' Ausgabe der Ergebnisse
'Select Case zeiger
'Case 0: AufUnter = aufgangUT / 24
'Case 1: AufUnter = untgangUT / 24
'End Select
AU(0) = aufgangUT / 24
AU(1) = untgangUT / 24
AufUnter = AU
End Function

'==============================================================================================================
'======================= ASTRONOMISCHE BERECHNUNGSFUNKTIONEN ENDE =============================================
'==============================================================================================================



