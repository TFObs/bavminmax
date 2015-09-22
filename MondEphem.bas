Attribute VB_Name = "mdlMondEphem"
Option Explicit
Public Mpi, Mdeg, Mrad
Public sonne, mond
Dim suncoor As koordinaten
Dim MoonCoor As koordinaten
Dim coor As koordinaten

Public Type koordinaten
    lon As Double
    Lat As Double
    DEC As Double
    RA As Double
    AnomalyMean As Double
    diameter As Double
    distance As Double
    parallax As Double
    Sign As String
End Type



Function Sign(lon)
 Dim signs
  signs = Array("Widder", "Stier", "Zwillinge", "Krebs", "Löwe", "Jungfrau", _
    "Waage", "Skorpion", "Schütze", "Steinbock", "Wassermann", "Fische")
  Sign = signs(Int(lon * Mrad / 30))
End Function

' Calculate coordinates for Sun
' Coordinates are accurate to about 10s (right ascension)
' and a few minutes of arc (declination)
Function SunPosition(TDT)
    Dim D, eg, wg, e, a
    Dim diameter0, MSun, nu
    Dim SonnenCoord(8)
    Dim SonneEqu
    
   D = TDT - 2447891.5
  
  eg = 279.403303 * Mdeg
  wg = 282.768422 * Mdeg
  e = 0.016713
  a = 149598500  ' km
  diameter0 = 0.533128 * Mdeg ' angular diameter of Moon at a distance
  
  MSun = 360 * Mdeg / 365.242191 * D + eg - wg
  nu = MSun + 360# * Mdeg / Mpi * e * Sin(MSun)
  
  suncoor.lon = Mod2MPi(nu + wg)
  suncoor.Lat = 0
  suncoor.AnomalyMean = MSun
  
  suncoor.distance = (1 - (e ^ 2)) / (1 + e * Cos(nu)) ' distance in astronomical units
  suncoor.diameter = diameter0 / suncoor.distance ' angular diameter in radians
  suncoor.distance = suncoor.distance * a         ' distance in km
  suncoor.parallax = 6378.137 / suncoor.distance  ' horizonal parallax

  'SonnenCoord = Ecl2Equ(suncoor.lon, suncoor.Lat, TDT)
  SonneEqu = Ecl2Equ(suncoor.lon, suncoor.Lat, TDT)
  
  suncoor.Sign = Sign(suncoor.lon)
  
  'ReDim Preserve SonnenCoord(8)
  SonnenCoord(0) = SonneEqu(0)
  SonnenCoord(1) = SonneEqu(1)
  SonnenCoord(2) = suncoor.lon
  SonnenCoord(3) = suncoor.AnomalyMean
  SonnenCoord(4) = suncoor.Sign
  SonnenCoord(7) = suncoor.parallax
  SonnenCoord(8) = suncoor.diameter
  SunPosition = SonnenCoord
  
  End Function
  
  ' Calculate data and coordinates for the Moon
' Coordinates are accurate to about 1/5 degree (in ecliptic coordinates)
Function MoonPosition(Sunlon, SunAnomalyMean, TDT)
Dim D, l0, P0, N0, i, e, a, diameter0, parallax0, l
Dim phases, mainPhase, p
Dim MMoon, N, c, Ev, Ae, A3
Dim MMoon2, Ec, A4, l2, v, l3, N2
Dim orbitlon, moonAge, phase, moonPhase
Dim MondCoord(8), MondEqu
   D = TDT - 2447891.5
  
  'Mean Moon orbit elements as of 1990.0
   l0 = 318.351648 * Mdeg
   P0 = 36.34041 * Mdeg
   N0 = 318.510107 * Mdeg
   i = 5.145396 * Mdeg
   e = 0.0549
   a = 384401 ' km
   diameter0 = 0.5181 * Mdeg ' angular diameter of Moon at a distance
   parallax0 = 0.9507 * Mdeg ' parallax at distance a
  
   l = 13.1763966 * Mdeg * D + l0
   MMoon = l - 0.1114041 * Mdeg * D - P0 ' Moon's mean anomaly M
   N = N0 - 0.0529539 * Mdeg * D ' Moon's mean ascending node longitude
   c = l - suncoor.lon
   Ev = 1.2739 * Mdeg * Sin(2 * c - MMoon)
   Ae = 0.1858 * Mdeg * Sin(SunAnomalyMean)
   A3 = 0.37 * Mdeg * Sin(SunAnomalyMean)
   MMoon2 = MMoon + Ev - Ae - A3 ' corrected Moon anomaly
   Ec = 6.2886 * Mdeg * Sin(MMoon2) ' equation of centre
   A4 = 0.214 * Mdeg * Sin(2 * MMoon2)
   l2 = l + Ev + Ec - Ae + A4 ' corrected Moon's longitude
   v = 0.6583 * Mdeg * Sin(2 * (l2 - Sunlon))
   l3 = l2 + v ' true orbital longitude

   N2 = N - 0.16 * Mdeg * Sin(SunAnomalyMean)
  
  
  MoonCoor.lon = Mod2MPi(N2 + Atan2(Sin(l3 - N2) * Cos(i), Cos(l3 - N2)))
  MoonCoor.Lat = arcsin(Sin(l3 - N2) * Sin(i))
  orbitlon = l3
  
  'MondCoord = Ecl2Equ(MoonCoor.lon, MoonCoor.Lat, TDT)
  MondEqu = Ecl2Equ(MoonCoor.lon, MoonCoor.Lat, TDT)
  
  'relative distance to semi mayor axis of lunar oribt
  MoonCoor.distance = (1 - (e ^ 2)) / (1 + e * Cos(MMoon2 + Ec))
  MoonCoor.diameter = diameter0 / MoonCoor.distance ' angular diameter in radians
  MoonCoor.parallax = parallax0 / MoonCoor.distance ' horizontal parallax in radians
  MoonCoor.distance = MoonCoor.distance * a 'distance in km

  'Age of Moon in radians since New Moon (0) - Full Moon (pi)
  moonAge = Mod2MPi(l3 - Sunlon)
  phase = 0.5 * (1 - Cos(moonAge)) ' Moon phase, 0-1
  
  phases = Array("Neumond", "Zunehmende Sichel", "Erstes Viertel", "Zunnehmender Mond", _
    "Vollmond", "Abnehmender Mond", "Letztes Viertel", "Abnehmende Sichel", "Neumond")
  mainPhase = 1# / 29.53 * 360 * Mdeg ' show 'Newmoon, 'Quarter' for +/-1 day arond the actual event
  p = Modu(moonAge, 90 * Mdeg)
  
  If (p < mainPhase Or p > 90 * Mdeg - mainPhase) Then
  p = 2 * Round(moonAge / (90# * Mdeg))
  Else: p = 2 * Int(moonAge / (90# * Mdeg)) + 1
  End If
  If p = 8 Then p = 0
  moonPhase = phases(p)
  
  MoonCoor.Sign = Sign(MoonCoor.lon)
  'ReDim Preserve MondCoord(8)
  MondCoord(0) = MondEqu(0)
  MondCoord(1) = MondEqu(1)
  MondCoord(2) = moonPhase
  MondCoord(3) = phase
  MondCoord(4) = moonAge * Mrad / 360 * 29.530588853
  MondCoord(5) = MoonCoor.Sign
  MondCoord(6) = p
  MondCoord(7) = MoonCoor.parallax
  MondCoord(8) = MoonCoor.diameter
MoonPosition = MondCoord
End Function




'========================================================================
'======================Zusatzfunktionen==================================
'========================================================================

'Umwqandlung von ekliptikalen Koordinaten (lon/lat) in äquatoriale Koordinaten (RA/dec)
Function Ecl2Equ(lon, Lat, TDT)
Dim T, eps, coseps, sineps, sinlon
Dim equat(2)
  T = (TDT - 2451545#) / 36525 'Epoche 2000, Januar 1.5
  eps = (23# + (26 + 21.45 / 60) / 60 + T * (-46.815 + T * (-0.0006 + T * 0.00181)) / 3600) * Mdeg
  coseps = Cos(eps)
  sineps = Sin(eps)
  
  sinlon = Sin(lon)
  coor.RA = Mod2MPi(Atan2((sinlon * coseps - Tan(Lat) * sineps), Cos(lon)))
  coor.DEC = arcsin(Sin(Lat) * coseps + Cos(Lat) * sineps * sinlon)
  equat(0) = coor.RA
  equat(1) = coor.DEC
  Ecl2Equ = equat
End Function

'Berechnung des Julianischen Datums
'Uhrzeit wird in Stunden übergeben
Public Function CalcJD(Tag, Monat, Jahr, Optional Uhrzeit) As Double
 
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
    
    CalcJD = a + b + Fix(30.6001 * (Monat + 1)) + Tag

    If Not IsMissing(Uhrzeit) Then
        CalcJD = CalcJD + (Uhrzeit) / 24
    End If
    
End Function

' Find local time of moonrise and moonset
' JD is the Julian Date of 0h local time (midnight)
' Accurate to about 5 minutes or better
' recursive: 1 - calculate rise/set in UTC
' recursive: 0 - find rise/set on the current local day (set could also be first)
' returns '' for moonrise/set does not occur on selected day

'deltaT is normally 65
Function MoonRise(JD, deltaT, lon, Lat, zone, recursive)
Dim jd0UT, suncoor1, coor1, suncoor2, coor2
Dim rise, riseprev, risenext, risetemp
Dim timeinterval ', deltaT

   timeinterval = 0.5
  
   jd0UT = Int(JD - 0.5) + 0.5 ' JD at 0 hours UT
   suncoor1 = SunPosition(jd0UT + 0 * deltaT / 24 / 3600)
   coor1 = MoonPosition(suncoor1(2), suncoor1(3), jd0UT + 0 * deltaT / 24 / 3600)

   suncoor2 = SunPosition(jd0UT + timeinterval + 0 * deltaT / 24 / 3600) ' calculations for noon
  ' calculations for next day's midnight
   coor2 = MoonPosition(suncoor2(2), suncoor2(3), jd0UT + timeinterval + 0 * deltaT / 24 / 3600)
  
    
  ' rise/set time in UTC, time zone corrected later
  rise = RiseSet(jd0UT, coor1, coor2, lon, Lat, timeinterval)
  
  If (recursive = 0) Then ' check and adjust to have rise/set time on local calendar day
    If (zone > 0) Then
      ' recursive call to MoonRise returns events in UTC
      riseprev = MoonRise(JD - 1, deltaT, lon, Lat, zone, 1)
     'End If
      ' recursive call to MoonRise returns events in UTC
      'risenext = MoonRise(JD+1, deltaT, lon, lat, zone, 1)
      'alert("yesterday="+riseprev(1)+"  today="+rise(1)+" tomorrow="+risenext(1))
      'alert("yesterday="+riseprev(0)+"  today="+rise(0)+" tomorrow="+risenext(0))
      'alert("yesterday="+riseprev(2)+"  today="+rise(1)+" tomorrow="+risenext(2))

      If (rise(1) >= 24 - zone Or rise(1) < -zone) Then ' transit time is tomorrow local time
        If (riseprev(1) < 24 - zone) Then rise(1) = "" ' there is no moontransit today
        Else: rise(1) = riseprev(1)
      End If

      If (rise(0) >= 24 - zone Or rise(0) < -zone) Then ' transit time is tomorrow local time
        If (riseprev(0) < 24 - zone) Then rise(0) = "" ' there is no moontransit today
        Else: rise(0) = riseprev(0)
      End If

      If (rise(1) >= 24 - zone Or rise(1) < -zone) Then ' transit time is tomorrow local time
        If (riseprev(2) < 24 - zone) Then rise(1) = "" ' there is no moontransit today
        Else: rise(1) = riseprev(2)
      End If
   End If
    
    ElseIf (zone < 0) Then
      ' rise/set time was tomorrow local time -> calculate rise time for former UTC day
      If (rise(0) < -zone Or rise(1) < -zone Or rise(1) < -zone) Then
        risetemp = MoonRise(JD + 1, deltaT, lon, Lat, zone, 1)
      
       
        If (rise(0) < -zone) Then
          If (risetemp(0) > -zone) Then rise(0) = "" ' there is no moonrise today
          Else: rise(0) = risetemp.rise
        End If
            

        If (rise(1) < -zone) Then
        
          If (risetemp(1) > -zone) Then rise(1) = "" ' there is no moonset today
          Else: rise(1) = risetemp.transit
        End If
       

        If (rise(2) < -zone) Then
        
          If (risetemp(2) > -zone) Then rise(2) = "" ' there is no moonset today
          Else: rise(2) = risetemp(2)
        End If
    End If
   
   End If
        
      
    
    If (rise(0)) Then rise(0) = Modu(rise(0) + zone, 24)    ' correct for time zone, if time is valid
    If (rise(1)) Then rise(1) = Modu(rise(1) + zone, 24) ' correct for time zone, if time is valid
    If (rise(1)) Then rise(1) = Modu(rise(1) + zone, 24)       ' correct for time zone, if time is valid
  
    '==> rise(0): rise
    '==> rise(1): transit
    '==> rise(2): set
 MoonRise = (rise)

End Function


' returns Greenwich sidereal time (hours) of time of rise
' and set of object with coordinates coor.ra/coor.dec
' at geographic position lon/lat (all values in MRadians)
' for mathematical horizon, without refraction and body Radius correction
'coor(0)=ra coor(1)=dec

Function GMSTRiseSet(coor, lon, Lat)
Dim RiseSet(3), tagbogen

 tagbogen = arccos(-Tan(Lat) * Tan(coor(1)))
 
 RiseSet(1) = Mrad / 15 * (coor(0) - lon)
 RiseSet(0) = 24# + Mrad / 15 * (-tagbogen + coor(0) - lon) ' calculate GMST of rise of object
 RiseSet(2) = Mrad / 15 * (tagbogen + coor(0) - lon) ' calculate GMST of set of object

 RiseSet(1) = Modu(RiseSet(1), 24) 'transit
 RiseSet(0) = Modu(RiseSet(0), 24) 'rise
 RiseSet(2) = Modu(RiseSet(2), 24) 'set
 
 GMSTRiseSet = RiseSet
End Function

' JD is the Julian Date of 0h UTC time (midnight)
Function RiseSet(jd0UT, coor1, coor2, lon, Lat, timeinterval)
Dim rise(3), rise1, rise2
Dim T0, T02, decMean, psi, alt, y, dt
   rise1 = GMSTRiseSet(coor1, lon, Lat)
   rise2 = GMSTRiseSet(coor2, lon, Lat)
  
   
  
  'alert( rise1(2)  +"  "+ rise2(2) )
  ' unwrap GMST in case we move across 24h -> 0h
  If (rise1(1) > rise2(1) And Abs(rise1(1) - rise2(1)) > 18) Then rise2(1) = rise2(1) + 24
  If (rise1(0) > rise2(0) And Abs(rise1(0) - rise2(0)) > 18) Then rise2(0) = rise2(0) + 24
  If (rise1(2) > rise2(2) And Abs(rise1(2) - rise2(2)) > 18) Then rise2(2) = rise2(2) + 24
   T0 = gmst(jd0UT)
'   T02 = T0-zone*1.002738 ' Greenwich sidereal time at 0h time zone (zone: hours)

  ' Greenwich sidereal time for 0h at selected longitude
   T02 = T0 - lon * Mrad / 15 * 1.002738
   If (T02 < 0) Then T02 = T02 + 24

  If (rise1(1) < T02) Then
     rise1(1) = rise1(1) + 24
     rise2(1) = rise2(1) + 24
 End If

  If (rise1(0) < T02) Then
    rise1(0) = rise1(0) + 24
    rise2(0) = rise2(0) + 24
 End If
 
  If (rise1(2) < T02) Then
   rise1(2) = rise1(2) + 24
   rise2(2) = rise2(2) + 24
 End If
  
  'alert("after="+ rise1(2)  +"  "+ rise2(2)+ " T0="+ T0 )
 
  ' Refraction and Parallax correction
  
   decMean = 0.5 * (coor1(1) + coor2(1))
   psi = arccos(Sin(Lat) / Cos(decMean))
  ' altitude of sun center: semi-diameter, horizontal parallax and (standard) refraction of 34'
   alt = 0.5 * coor1(7) - coor1(8) + 34# / 60 * Mdeg
   y = arcsin(Sin(alt) / Sin(psi))
   dt = 240 * Mrad * y / Cos(decMean) / 3600 ' time correction due to refraction, parallax

  rise(1) = GMST2UT(jd0UT, InterpolateGMST(T0, rise1(1), rise2(1), timeinterval))
  rise(0) = GMST2UT(jd0UT, InterpolateGMST(T0, rise1(0), rise2(0), timeinterval) - dt)
  rise(2) = GMST2UT(jd0UT, InterpolateGMST(T0, rise1(2), rise2(2), timeinterval) + dt)
  
  'rise(1) = Modu(rise(1), 24.)
  'rise(0)    = Modu(rise(0), 24.)
  'rise(2)     = Modu(rise(2),  24.)
 
  RiseSet = rise
End Function

' Find GMST of rise/set of object from the two calculates
' (start)points (day 1 and 2) and at midnight UT(0)
Function InterpolateGMST(gmst0, gmst1, gmst2, timefactor)

InterpolateGMST = ((timefactor * 24.07 * gmst1 - gmst0 * (gmst2 - gmst1)) / (timefactor * 24.07 + gmst1 - gmst2))

End Function

' Julian Date to Greenwich Mean Sidereal Time
Function gmst(JD)
Dim UT, T, T0
  UT = ((JD - 0.5) - Int(JD - 0.5)) * 24 ' UT in hours
  JD = Int(JD - 0.5) + 0.5 ' JD at 0 hours UT
  T = (JD - 2451545#) / 36525#
  T0 = 6.697374558 + T * (2400.051336 + T * 0.000025862)
  
  gmst = (Modu(T0 + UT * 1.002737909, 24))

End Function



'Convert Greenwich mean sidereal time to UT
Function GMST2UT(JD, gmst)
Dim T, T0, UT

  JD = Int(JD - 0.5) + 0.5 'JD at 0 hours UT
  T = (JD - 2451545#) / 36525#
  T0 = Modu(6.697374558 + T * (2400.051336 + T * 0.000025862), 24)
  'var UT = 0.9972695663*Mod((gmst-T0), 24.)
  UT = 0.9972695663 * ((gmst - T0))
GMST2UT = UT
End Function

'Local Mean Sidereal Time, geographical longitude in MRadians, East is positive
Function GMST2LMST(gmst, lon)
 
   GMST2LMST = Modu(gmst + Mrad * lon / 15, 24#)
End Function

' Find (local) time of sunrise and sunset
' JD is the Julian Date of 0h local time (midnight)
' Accurate to about 1-2 minutes
' recursive: 1 - calculate rise/set in UTC
' recursive: 0 - find rise/set on the current local day (set could also be first)
Function SunRise(JD, deltaT, lon, Lat, zone, recursive)
Dim rise, risetemp
Dim coor1, coor2, jd0UT

  jd0UT = Int(JD - 0.5) + 0.5 ' JD at 0 hours UT
'  alert("jd0UT="+jd0UT+"  JD="+JD)
  coor1 = SunPosition(jd0UT + 0 * deltaT / 24 / 3600)
  coor2 = SunPosition(jd0UT + 1 + 0 * deltaT / 24 / 3600) ' calculations for next day's UTC midnight
  

  rise = RiseSet(jd0UT, coor1, coor2, lon, Lat, 1)  ' rise/set time in UTC
  If (recursive = 1) Then ' check and adjust to have rise/set time on local calendar day
    If (zone > 0) Then
      ' rise time was yesterday local time -> calculate rise time for next UTC day
      If (rise(0) >= 24 - zone Or rise(1) >= 24 - zone Or rise(2) >= 24 - zone) Then
        risetemp = SunRise(JD + 1, deltaT, lon, Lat, zone, 1)
      

        If (rise(0) >= 24 - zone) Then rise(0) = risetemp(0)
        If (rise(1) >= 24 - zone) Then rise(1) = risetemp(1)
        If (rise(2) >= 24 - zone) Then rise(2) = risetemp(2)
      
    End If
    ElseIf (zone < 0) Then
      ' rise time was yesterday local time -> calculate rise time for next UTC day
      If (rise(0) < -zone Or rise(1) < -zone Or rise(2) < -zone) Then
        risetemp = SunRise(JD - 1, deltaT, lon, Lat, zone, 1)
     

        If (rise(0) < -zone) Then rise(0) = risetemp(0)
        If (rise(1) < -zone) Then rise(1) = risetemp(1)
        If (rise(2) < -zone) Then rise(2) = risetemp(2)
      
    End If
   End If
    rise(1) = Modu(rise(1) + zone, 24)
    rise(0) = Modu(rise(0) + zone, 24)
    rise(2) = Modu(rise(2) + zone, 24)
   
   End If
SunRise = rise
  
 End Function


Public Function arccos(x)

arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Public Function arcsin(x)
 arcsin = Atn(x / Sqr(-x * x + 1))
End Function


Function Modu(a, b)
 Modu = (a - Int(a / b) * b)
End Function


'Modulo MPi
Function Mod2MPi(x)
  x = Modu(x, 2 * Mpi)
Mod2MPi = x
End Function



Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
   Dim signy As Integer
   signy = Sgn(y)
   If signy = 0 Then signy = 1 ' removes the problem when Y=0
   If Abs(x) < 0.0000001 Then
        ' (direct comparison with zero doesn't always work)
        Atan2 = Sgn(y) * 1.5707963267949
    ElseIf x < 0 Then
        Atan2 = Atn(y / x) + signy * Mpi
    Else
        Atan2 = Atn(y / x)
    End If
End Function

'http://www.arndt-bruenner.de/mathe/scripts/sphaerischr.htm#rechner
'Berechnung der Monddistanz, Koordinateneingabe in Grad!
Function Moondistance(StarRa, StarDec, MoonRa, MoonDec)
Dim a, b, g, c
    a = (90 - StarDec) * Mdeg
    b = (90 - MoonDec) * Mdeg
    g = Abs(MoonRa - StarRa) * Mdeg
    c = arccos(Cos(a) * Cos(b) + Sin(a) * Sin(b) * Cos(g)) * Mrad
 Moondistance = c
End Function
