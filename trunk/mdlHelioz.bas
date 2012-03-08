Attribute VB_Name = "mdlHelioz"
Option Explicit
Public umwand(6), RA, DEC
Dim nalph, ndelt
Dim T, dlp, lnu
Dim g, G2, G4, G5, G6
Dim D, a, U
Dim dl, dl2, dl4, dl5, dl6
Dim dlM, wL
Dim Ro, R2, R4, R5, R6, RM, r
Dim TN, xi, z, theta
Dim alpha, Delta, vgl, se
Dim be, la, erg
Dim hjd, hcorr

Public Function Hkorr(JD, RA, DEC, IsGeoz As Boolean)

'für Umrechnung ins Gradmaß
pi = 4 * Atn(1)
pi = pi / 180

'jul. Jahrhunderte seit 1.1.2000 12h
T = (JD - 2451545) / 36525

'langperperiodische Störungen durch mittlere Anomalie
dlp = (1.866 / 3600 - 0.016 * T / 3600) * Sin((207.51 + 150.27 * T) * pi) + _
6.4 / 3600 * Sin((251.39 + 20.2 * T) * pi) + _
0.266 / 3600 * Sin((150.8 + 119 * T) * pi)


'mittlere Länge der Sonne
lnu = 280.465905 + 36000 * T + 2770.308 / 3600 * T + _
1.089 / 3600 * T ^ 2 + 0.202 / 3600 * Sin((128.9 + 893.3 * T) * pi) + _
dlp
lnu = range360(lnu)

'mittlere Anomalie der Sonne
g = (357.525433 + 35999 * T + 178.02 / 3600 * T - 0.54 / 3600 * T ^ 2 + dlp)
g = range360(g)

'mittlere Anomalie der Venus
G2 = (49.943 + 58517.493 * T)
G2 = range360(G2)
'mittlere Anomalie des Mars
G4 = (19.557 + 19139.977 * T)
G4 = range360(G4)
'mittlere Anomalie des Jupiter
G5 = 19.863 + 3034.583 * T + 1300 / 3600 * Sin((173.58 + 39.8 * T) * pi)
G5 = range360(G5)
'mittlere Anomalie des Saturn
G6 = 317.394 + 1221.794 * T
G6 = range360(G6)
'mittlerer Winkelabstand des Mondes von der Sonne
D = 297.852 + 445267.114 * T
D = range360(D)
'mittlere Anomalie des Mondes
a = 134.954 + 477198.849 * T
a = range360(a)
'mittleres Argument der Breite des Mondes
U = 93.276 + 483202.025 * T
U = range360(U)

'Differenz zw. wahrer und mittlerer Sonnenlänge
dl = (6892.817 - 17.24 * T) * Sin(g * pi) _
     + (71.977 - 0.361 * T) * Sin((2 * g) * pi) _
     + (1.054) * Sin((3 * g) * pi)
dl = range360(dl / 3600)

'dekadischer Logarithmus des Radius in AE
Ro = (0.00003042 - 0.00000015 * T) + (-0.00725598 + 0.00001814 * T) * Cos(g * pi) _
+ (-0.00009092 + 0.00000046 * T) * Cos((2 * g) * pi) _
+ (-0.00000145 * Cos((3 * g) * pi))

'Störungen in Länge
'Längenstörung durch Venus
dl2 = 4.838 * Cos((299.102 + G2 - g) * pi) + 0.116 * Cos((148.9 + 2 * G2 - g) * pi) _
+ 5.526 * Cos((148.313 + 2 * G2 - 2 * g) * pi) + 2.497 * Cos((315.943 + 2 * G2 - 3 * g) * pi) _
+ 0.666 * Cos((177.71 + 3 * G2 - 3 * g) * pi) + 1.559 * Cos((345.253 + 3 * G2 - 4 * g) * pi) _
+ 1.024 * Cos((318.15 + 3 * G2 - 5 * g) * pi) + 0.21 * Cos((206.2 + 4 * G2 - 4 * g) * pi) _
+ 0.144 * Cos((195.4 + 4 * G2 - 5 * g) * pi) + 0.152 * Cos((343.8 + 4 * G2 - 6 * g) * pi) _
+ 0.123 * Cos((195.3 + 5 * G2 - 7 * g) * pi) + 0.154 * Cos((359.6 + 5 * G2 - 8 * g) * pi)

'Längenstörung durch Mars
dl4 = 0.273 * Cos((217.7 - G4 + g) * pi) + 2.043 * Cos((343.888 - 2 * G4 + 2 * g) * pi) _
+ 1.77 * Cos((200.402 - 2 * G4 + g) * pi) + 0.129 * Cos((294.2 - 3 * G4 + 3 * g) * pi) _
+ 0.425 * Cos((338.88 - 3 * G4 + 2 * g) * pi) + 0.5 * Cos((105.18 - 4 * G4 + 3 * g) * pi) _
+ 0.585 * Cos((334.06 - 4 * G4 + 2 * g) * pi) + 0.204 * Cos((100.8 - 5 * G4 + 3 * g) * pi) _
+ 0.154 * Cos((227.4 - 6 * G4 + 4 * g) * pi) + 0.101 * Cos((96.3 - 6 * G4 + 3 * g) * pi) _
+ 0.106 * Cos((222.7 - 7 * G4 + 4 * g) * pi)

'Längenstörung durch Jupiter
dl5 = 0.163 * Cos((198.6 - G5 + 2 * g) * pi) + 7.208 * Cos((179.532 - G5 + g) * pi) _
+ 2.6 * Cos((263.217 - G5) * pi) + 2.731 * Cos((87.145 - 2 * G5 + 2 * g) * pi) _
+ 1.61 * Cos((109.493 - 2 * G5 + g) * pi) + 0.164 * Cos((170.5 - 3 * G5 + 3 * g) * pi) _
+ 0.556 * Cos((82.65 - 3 * G5 + 2 * g) * pi) + 0.21 * Cos((98.5 - 3 * G5 + g) * pi)
dl5 = Format(dl5, "#.000")

'Längenstörung durch Saturn
dl6 = 0.419 * Cos((100.58 - G6 + g) * pi) + 0.32 * Cos((269.46 - G6) * pi) _
+ 0.108 * Cos((290.6 - 2 * G6 + 2 * g) * pi) + 0.112 * Cos((293.6 - 2 * G6 + g) * pi)

'Längenstörung durch Mond

dlM = 6.454 * Sin(D * pi) + 0.177 * Sin((D + a) * pi) _
- 0.424 * Sin((D - a) * pi) + 0.172 * Sin((D - g) * pi)

'wahre Länge der Sonne
wL = lnu + dl + dl2 / 3600 + dl4 / 3600 + dl5 / 3600 + dl6 / 3600 + dlM / 3600
wL = range360(wL)

'Störung in logR durch Venus
R2 = 2359 * Cos((209.08 + G2 - g) * pi) _
+ 160 * Cos((58.4 + 2 * G2 - g) * pi) _
+ 6842 * Cos((58.318 + 2 * G2 - 2 * g) * pi) _
+ 869 * Cos((226.7 + 2 * G2 - 3 * g) * pi) _
+ 1045 * Cos((87.57 + 3 * G2 - 3 * g) * pi) _
+ 1497 * Cos((255.25 + 3 * G2 - 4 * g) * pi) _
+ 194 * Cos((49.5 + 3 * G2 - 5 * g) * pi) _
+ 376 * Cos((116.28 + 4 * G2 - 4 * g) * pi) _
+ 196 * Cos((105.2 + 4 * G2 - 5 * g) * pi) _
+ 163 * Cos((145.4 + 5 * G2 - 5 * g) * pi) _
+ 141 * Cos((105.4 + 5 * G2 - 7 * g) * pi)

'Störung in logR durch Mars
R4 = 150 * Cos((127.7 - G4 + g) * pi) _
+ 2057 * Cos((253.828 - 2 * G4 + 2 * g) * pi) _
+ 151 * Cos((295 - 2 * G4 + g) * pi) _
+ 168 * Cos((203.5 - 3 * G4 + 3 * g) * pi) _
+ 215 * Cos((249 - 3 * G4 + 2 * g) * pi) _
+ 478 * Cos((15.17 - 4 * G4 + 3 * g) * pi) _
+ 105 * Cos((65.9 - 4 * G4 + 2 * g) * pi) _
+ 107 * Cos((324.6 - 5 * G4 + 4 * g) * pi) _
+ 139 * Cos((137.3 - 6 * G4 + 4 * g) * pi)

'Störung in logR durch Jupiter
R5 = 208 * Cos((112 - G5 + 2 * g) * pi) _
+ 7067 * Cos((89.545 - G5 + g) * pi) _
+ 244 * Cos((338.6 - G5) * pi) _
+ 103 * Cos((350.5 - 2 * G5 + 3 * g) * pi) _
+ 4026 * Cos((357.108 - 2 * G5 + 2 * g) * pi) _
+ 1459 * Cos((19.467 - 2 * G5 + g) * pi) _
+ 281 * Cos((81.2 - 3 * G5 + 3 * g) * pi) _
+ 803 * Cos((352.56 - 3 * G5 + 2 * g) * pi) _
+ 174 * Cos((8.6 - 3 * G5 + g) * pi) _
+ 113 * Cos((347.7 - 4 * G5 + 2 * g) * pi)

'Störung in logR durch Saturn
R6 = 429 * Cos((10.6 - G6 + g) * pi) _
+ 162 * Cos((200.6 - 2 * G6 + 2 * g) * pi) _
+ 112 * Cos((203.1 - 2 * G6 + g) * pi)

'Störung in logR durch Mond
RM = 13360 * Cos(D * pi) _
+ 370 * Cos((D + a) * pi) - 1330 * Cos((D - a) * pi) _
- 140 * Cos((D + g) * pi) + 360 * Cos((D - g) * pi)

r = 10 ^ (Ro + ((R2 + R4 + R5 + R6 + RM) / (10 ^ 9)))

'Transformation ekliptikaler koordinaten von J2000 auf aktuelle Epoche
'Nullepoche = J2000
'TN = Jahrhunderte seit 1.Jan 2000 12h
'T = Aquinoktium neu - Äquinoktium alt
TN = (JD - 2451545) / 36525
T = (JD - 2451545) / 36525 '2451545 = JD für Epoche 2000

'Hilfsgrößen
xi = (2306.218 + 1.397 * TN) * T + 0.302 * T ^ 2 + 0.018 * T ^ 3
z = xi + 0.793 * T ^ 2
theta = (2004.311 - 0.853 * TN) * T - 0.427 * T ^ 2 - 0.042 * T ^ 3

'RA alt und Dek alt
alpha = RA / 24 * 360
Delta = DEC

'Umrechnung
theta = theta / 3600
xi = xi / 3600
z = z / 3600

'neue Koordinaten
ndelt = arcsinz((Sin(theta * pi) * Cos(Delta * pi) * Cos((alpha + xi) * pi) + Cos(theta * pi) * Sin(Delta * pi)))
nalph = arcsinz(Cos(Delta * pi) * Sin((alpha + xi) * pi) / Cos(ndelt * pi))
nalph = range360(nalph)
'Näherungswert für die neue RA
vgl = (xi + z) + theta * Tan(Delta * pi) * Sin(alpha * pi) + alpha

If Abs(nalph - vgl) > 0.5 Then
nalph = 180 - (nalph - z)
Else: nalph = nalph + z
End If
nalph = range360(nalph)
'Schiefe der Ekliptik
se = 23.439291 - 0.013004 * T

' ekliptikale Koordinaten
be = arcsinz(Cos(se * pi) * Sin(ndelt * pi) - Sin(se * pi) * Cos(ndelt * pi) * Sin(nalph * pi))
la = arcsinz((Cos(se * pi) * Cos(ndelt * pi) * Sin(nalph * pi) + Sin(se * pi) * Sin(ndelt * pi)) / Cos(be * pi))

If Cos(la * pi) * Cos(be * pi) <> Cos(ndelt * pi) * Cos(nalph * pi) Then
 la = 180 - la
End If

erg = ausg(nalph * 24 / 360, ndelt)
frmHelioz.Label3.Caption = "RA       " & umwand(1) & ":" & umwand(2) & ":" & umwand(3) & vbCrLf & _
"DEC   " & umwand(4) & ":" & umwand(5) & ":" & umwand(6)

'MsgBox lnu & vbCrLf & be & vbCrLf & la

'einfache Formel aus Wischnewski
hjd = (-499 * Cos(be * pi) * Cos((lnu - la) * pi)) / 3600 / 24

'Formel aus http://adsabs.harvard.edu/full/1972PASP...84..784L
'Title: The Calculation of Heliocentric Corrections
'Authors: Landolt, A. U. & Blondeau, K. L.
'Journal: Publications of the Astronomical Society of the Pacific, Vol. 84, No. 502, p.784
'Bibliographic Code: 1972PASP...84..784L

hcorr = (-499 * r * _
((Cos(wL * pi) * Cos(nalph * pi) * Cos(ndelt * pi)) _
+ (Sin(wL * pi) * (Sin(se * pi) * Sin(ndelt * pi) + _
Cos(se * pi) * Cos(ndelt * pi) * Sin(nalph * pi))))) / 3600 / 24

'Correction = KR[{cosAcosBcosC}+{sinA(sinDsinC+cosDcosCsinB)}]
'K = 499.004631696182s
'R = Entfernung Erde-Sonne
'A = lsonne
'B = alpha neu
'c = delta neu
'd = Schiefe Ekliptik

'MsgBox hjd & vbCrLf & hcorr
If IsGeoz = True Then
    Hkorr = Format(JD + hcorr, "#.0000000")
ElseIf IsGeoz = False Then
    Hkorr = Format(JD - hcorr, "#.0000000")
End If

End Function

Public Function arcsinz(x)
 arcsinz = Atn(x / Sqr(-x * x + 1)) * (180 / 4 / Atn(1))
End Function


Public Function range360(x)
' gibt Winkel zwischen 0 und 360° um große
' Werte zu verhindern, die man aus
' den Berechnungen zu mittlerer Länge usw. bekommt
  range360 = x - 360 * Int(x / 360)
End Function

Public Function ausg(rak, deck)
Dim vz

'Umwandlung der RA
umwand(1) = Format(Int(rak), "00")
umwand(2) = Format(Int((rak - Int(rak)) * 60), "00")
umwand(3) = (((rak - Int(rak)) * 60) - Int((rak - Int(rak)) * 60)) * 60

'Falls wg Rundung 59.99999 herauskommt
If umwand(3) >= 60 Then
umwand(2) = umwand(2) + 1
umwand(3) = "00.0"
Else
umwand(3) = Format(umwand(3), "00.0")
End If

'Wenn Dec negativ, nur mit Absolutwerden rechnen!!
If deck < 0 Then
vz = "-"
Else: vz = "+"
End If
deck = Abs(deck)

umwand(4) = vz & Format(Int(deck), "00")
umwand(5) = Format(Int((deck - Int(deck)) * 60), "00")
umwand(6) = (((deck - Int(deck)) * 60) - Int((deck - Int(deck)) * 60)) * 60
umwand(6) = Format(umwand(6), "00.0")

'Wegen Rundungen 59.99999s
If umwand(6) >= 60 Then
umwand(5) = Format(umwand(5) + 1, "00")
umwand(6) = "00.0"
Else
umwand(6) = Format(umwand(6), "00.0")
End If

If vz = "-" Then
deck = -deck
End If

vz = ""
ausg = umwand
End Function

