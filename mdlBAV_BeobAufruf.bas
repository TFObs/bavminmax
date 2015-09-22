Attribute VB_Name = "mdlBAV_BeobAufruf"
'ACHTUNG!! Nur rückwirkend bis August 2008!!
'Vorher sind die sonstigen infos nicht vorhanden!!!

Option Explicit

Const scUserAgent = "VarEphem"
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_OPEN_TYPE_PRECONFIG = 0

Const INTERNET_FLAG_RELOAD = &H80000000
Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" _
    (ByVal sAgent As String, ByVal lAccessType As Long, _
    ByVal sProxyName As String, ByVal sProxyBypass As String, _
    ByVal lFlags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer

Private Declare Function InternetReadFile Lib "wininet" _
    (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer

Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" _
    (ByVal hInternetSession As Long, ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long
    
 Private Type fehlerloc
    ort As String
 End Type

Public fehler As fehlerloc

Function GetBAVStardata(ByVal Temppath As String, ByVal sURL As String)
    Dim hOpen As Long, hFile As Long, sBuffer As String * 4096, ByteSize As Long
    Dim fs As FileSystemObject, OutStream As TextStream
    Dim DatenInhalt As String
    Dim Textstrom, Werte()
    Dim x As Integer, y As Integer
    Dim lines As Collection
    
    Set lines = New Collection
    Set fs = New FileSystemObject
    
    On Error GoTo errhandler
    
    fehler.ort = "Func_GetBAVStardata: Internet-Verbindung"
    
    If fs.FileExists(Temppath) Then fs.DeleteFile (Temppath)
   
    Set OutStream = fs.CreateTextFile(Temppath)
 
    'Create an internet connection
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    'Open the url
    hFile = InternetOpenUrl(hOpen, sURL, vbNullString, ByVal 0&, INTERNET_FLAG_EXISTING_CONNECT, ByVal 0&)
 
       
    Do
        InternetReadFile hFile, sBuffer, Len(sBuffer), ByteSize
        If ByteSize = 0 Then Exit Do
        DatenInhalt = DatenInhalt & Left(sBuffer, ByteSize)
    Loop

    'clean up
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    
    fehler.ort = "Func_GetBAVStardata: Sourcecode"
    
    'Fehler abfangen wenn Datei nicht vorhanden
    'Es wird eine Seite erzeugt, die verschiedene andere BeobAufrufe auflistet
    'Daher muss der Quellcode nach "not be found" durchsucht werden
    If InStr(1, DatenInhalt, "could not be found on this server") Then
        GetBAVStardata = False
        Exit Function
    End If
    
    'Auflösen in einzelne Zeilen Lf als Kennzeichen für eine neue Zeile
    Textstrom = Split(DatenInhalt, vbLf)

    'Ermitteln der Zeilenanfänge für die Daten
    For x = 0 To UBound(Textstrom) - 1
        OutStream.WriteLine Textstrom(x)
        If InStr(1, Textstrom(x), "new Array") Then lines.Add (OutStream.Line)
    Next
       
    DatenInhalt = ""
    
    fehler.ort = "Func_GetBAVStardata: Formatierung"
    
     On Error GoTo daterror
    
    ReDim Werte(lines.Count - 2)
    
    For x = 1 To lines.Count - 1
        'Sonderzeichen -Wagenrücklauf- wird aus den Zeilen entfernt
        For y = lines.Item(x) - 1 To lines.Item(x + 1) - 3
            DatenInhalt = Trim(DatenInhalt) + Replace(Trim(Textstrom(y)), Chr(13), "")
        Next y
        'Nun noch Löschen von Semikolon und der Klammern und umwandeln ALLER Zahlen ("," entfernen)
        DatenInhalt = Replace((Replace(Replace(DatenInhalt, "(", ""), ")", "")), Chr(34) & ",", Chr(34) & "&")
        DatenInhalt = Replace(Replace(Replace(Replace(DatenInhalt, ",", "."), "&", ","), Chr(34), ""), ";", "")
        DatenInhalt = Replace(Replace(Replace(DatenInhalt, Chr(9), ""), ", ", ","), " ,", ",")
        Werte(x - 1) = DatenInhalt
        'Debug.Print Werte(x - 1)
        DatenInhalt = ""
        
    Next x
    

    GetBAVStardata = Werte

    'cleanup
        
    OutStream.Close
    fehler.ort = ""
    
    Set OutStream = Nothing
    Set fs = Nothing
    Set lines = Nothing
    
    Exit Function
    
errhandler:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf & fehler.ort, vbCritical, "Programmfehler"
    fehler.ort = ""
daterror:
    GetBAVStardata = False
    fehler.ort = ""
End Function


Function CreateBAV_Database_EA(ByVal TemporaryPath As String, ByVal BAV_URL As String)
Dim StarArray, x As Integer, counter As Integer
Dim BAVBeobAuf As Recordset
Dim fs As FileSystemObject
Dim tempVal, realval
'"http://www.bav-astro.de/ea/beob_aufr_08_01.html"
'http://www.bav-astro.de/ea/beob_aufr.php?jahr=8&monat=12
'http://bav-astro.de/rrlyr/beob_aufr_08_03.html
StarArray = (GetBAVStardata(TemporaryPath, BAV_URL))

Set fs = New FileSystemObject
Set BAVBeobAuf = New ADODB.Recordset

On Error Resume Next
'Fehler beim Lesen der Datei, Datei wurde heruntergeladen, aber vermutlich vor 07/2007...
If UBound(StarArray) = 0 Then
    MsgBox "Die Datei ist beschädigt oder nicht vorhanden ." & vbCrLf & _
    "Bitte versuchen Sie den Download erneut." & vbCrLf & vbCrLf & _
    "Es können nur Beobachtungsaufrufe ab Juli 2007 verwendet werden.", vbCritical, _
    "Datenbank kann nicht erzeugt werden.."
    fs.DeleteFile (TemporaryPath)
    CreateBAV_Database_EA = False
    Exit Function
 End If



On Error GoTo errhandler

'Löschen einer bestehenden Datenbank
If fs.FileExists(App.Path & "\BAVBA_EA.dat") Then fs.DeleteFile (App.Path & "\BAVBA_EA.dat")

    fehler.ort = "Sub_CreateBAV_Database_EA: DB erzeugen"
    
   'Erstellen einer leeren Datenbank
   With BAVBeobAuf
        .Fields.Append ("ID"), adInteger
        .Fields.Append ("Kürzel"), adVarChar, 6
        .Fields.Append ("Stbld"), adChar, 3
        .Fields.Append ("BP"), adChar, 5
        .Fields.Append ("LBeob"), adDouble
        .Fields.Append ("Max"), adDouble
        .Fields.Append ("MinI"), adDouble
        .Fields.Append ("MinII"), adDouble
        .Fields.Append ("Spektr"), adVarChar, 11
        .Fields.Append ("D"), adDouble, 4
        .Fields.Append ("kD"), adDouble, 4
        .Fields.Append ("Typ"), adVarChar, 12
        .Fields.Append ("Epoche"), adDouble
        .Fields.Append ("Periode"), adDouble
        .Fields.Append ("for"), adVarChar, 3
        .Fields.Append ("hh"), adInteger, 2
        .Fields.Append ("mm"), adInteger, 2
        .Fields.Append ("ss"), adDouble, 4
        .Fields.Append ("vz"), adChar, 1
        .Fields.Append ("o"), adInteger, 2
        .Fields.Append ("m"), adDouble, 5
        .Open
        .Save App.Path & "\BAVBA_EA.dat"
        .Close
    End With

    BAVBeobAuf.Open (App.Path & "\BAVBA_EA.dat")

    With BAVBeobAuf

        tempVal = Split(StarArray(x), ",")
        
        fehler.ort = "Sub_CreateBAV_Database_EA: Array umwandeln"
        
        'Dimensionieren des Feldes für die Einzeldaten
        ReDim realval(UBound(StarArray), UBound(tempVal))

        'Umwandeln des Arrays(Zeilen aus dem Quellcode)  in das Array(Sterne) für die Ausgabe
        For x = 0 To UBound(StarArray)
            tempVal = Split(StarArray(x), ",")
            For counter = 0 To UBound(tempVal)
                realval(x, counter) = Replace(tempVal(counter), Chr(34), "")
            Next counter
        Next x
        
        fehler.ort = "Sub_CreateBAV_Database_EA: RS erstellen"
        
        On Error Resume Next
    
        'Eintagen der Werte in die Felder des Recordset
        For x = 0 To UBound(Split(StarArray(0), ","))
            If Not Trim(realval(3, x)) = "" And Not Trim(realval(4, x)) = "" Then
                .AddNew
                tempVal = Split(realval(0, x), " ")
                .Fields("ID").Value = x + 1
                .Fields("Kürzel").Value = Trim(tempVal(0))
                .Fields("Stbld").Value = Trim(tempVal(1))
                .Fields("BP").Value = Trim(realval(1, x))
                .Fields("Typ").Value = Trim(realval(3, x))
                .Fields("Epoche").Value = Trim(realval(4, x))
                .Fields("Periode").Value = Trim(realval(5, x))
                .Fields("Max").Value = Trim(realval(6, x))
                .Fields("MinI").Value = Trim(realval(7, x))
                .Fields("MinII").Value = Trim(realval(8, x))
                .Fields("D").Value = Trim(realval(10, x))
                .Fields("kD").Value = Trim(realval(11, x))
        
                realval(12, x) = Replace(Replace(Replace(realval(12, x), "h", ""), "m", ""), "s", "")
                tempVal = Split(realval(12, x), " ")
                .Fields("hh").Value = Trim(tempVal(0))
                .Fields("mm").Value = Trim(tempVal(1))
                .Fields("ss").Value = Trim(tempVal(2))
                
                realval(13, x) = Replace(Replace(realval(13, x), "°", ""), "'", "")
                tempVal = Split(realval(13, x), " ")
                .Fields("vz").Value = Trim(tempVal(0))
                .Fields("o").Value = Trim(tempVal(1))
                .Fields("m").Value = Trim(tempVal(2)) + Trim(tempVal(3)) / 60
                
                .Fields("LBeob") = Trim(realval(15, x))
                .Update
            End If
        Next x
        
        Err.Clear
        
        fehler.ort = ""
    
    .Save App.Path & "\BAVBA_EA.dat"
 
 End With
 
 fs.DeleteFile (TemporaryPath)
 
 Set fs = Nothing
 Set BAVBeobAuf = Nothing
 
 MsgBox "Erzeugung der Datenbank erfolgreich." & vbCrLf & "Die Daten können nun für Berechnungen verwendet werden.", vbInformation, "Download erfolgreich"

 frmHaupt.Form_Load
 frmHaupt.cmdListe.Enabled = True: frmHaupt.VTabs.TabEnabled(1) = True
 frmHaupt.cmbGrundlage.Enabled = True
 
 For x = 1 To frmHaupt.cmbGrundlage.ListCount
 If frmHaupt.cmbGrundlage.List(x) = "BAV-BA_EA" Then
    frmHaupt.cmbGrundlage.ListIndex = x
    Exit For
 End If
 Next
 Exit Function
 
errhandler:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf & fehler.ort, vbCritical, "Programmfehler"
    fehler.ort = ""
End Function


Function CreateBAV_Database_RR(ByVal TemporaryPath As String, ByVal BAV_URL As String)
Dim StarArray, x As Integer, counter As Integer
Dim BAVBeobAuf As Recordset
Dim fs As FileSystemObject
Dim tempVal, realval
'"http://www.bav-astro.de/ea/beob_aufr_08_01.html"
'http://bav-astro.de/rrlyr/beob_aufr_08_03.html
StarArray = (GetBAVStardata(TemporaryPath, BAV_URL))

Set fs = New FileSystemObject
Set BAVBeobAuf = New ADODB.Recordset

On Error Resume Next
'Fehler beim Lesen der Datei, Datei wurde heruntergeladen, aber vermutlich vor 03/2008...
If UBound(StarArray) = 0 Then
    MsgBox "Die Datei ist beschädigt oder nicht vorhanden ." & vbCrLf & _
    "Bitte versuchen Sie den Download erneut." & vbCrLf & vbCrLf & _
    "Es können nur Beobachtungsaufrufe ab März 2008 verwendet werden.", vbCritical, _
    "Datenbank kann nicht erzeugt werden.."
    fs.DeleteFile (TemporaryPath)
    CreateBAV_Database_RR = False
    Exit Function
 End If



On Error GoTo errhandler

'Löschen einer bestehenden Datenbank
If fs.FileExists(App.Path & "\BAVBA_RR.dat") Then fs.DeleteFile (App.Path & "\BAVBA_RR.dat")

    fehler.ort = "Sub_CreateBAV_Database_RR: DB erzeugen"
    
   'Erstellen einer leeren Datenbank
   With BAVBeobAuf
        .Fields.Append ("ID"), adInteger
        .Fields.Append ("Kürzel"), adVarChar, 6
        .Fields.Append ("Stbld"), adChar, 3
        .Fields.Append ("BP"), adChar, 5
        .Fields.Append ("LBeob"), adDouble
        .Fields.Append ("Max"), adDouble
        .Fields.Append ("MinI"), adDouble
        .Fields.Append ("M-m"), adDouble
        .Fields.Append ("Typ"), adVarChar, 12
        .Fields.Append ("Epoche"), adDouble
        .Fields.Append ("Periode"), adDouble
        .Fields.Append ("hh"), adInteger, 2
        .Fields.Append ("mm"), adInteger, 2
        .Fields.Append ("ss"), adDouble, 4
        .Fields.Append ("vz"), adChar, 1
        .Fields.Append ("o"), adInteger, 2
        .Fields.Append ("m"), adDouble, 5
        .Open
        .Save App.Path & "\BAVBA_RR.dat"
        .Close
    End With

    BAVBeobAuf.Open (App.Path & "\BAVBA_RR.dat")

    With BAVBeobAuf

        tempVal = Split(StarArray(x), ",")
        
        fehler.ort = "Sub_CreateBAV_Database_RR: Array umwandeln"
        
        'Dimensionieren des Feldes für die Einzeldaten
        ReDim realval(UBound(StarArray), UBound(tempVal))

        'Umwandeln des Arrays(Zeilen aus dem Quellcode)  in das Array(Sterne) für die Ausgabe
        For x = 0 To UBound(StarArray)
            tempVal = Split(StarArray(x), ",")
            For counter = 0 To UBound(tempVal)
                realval(x, counter) = Replace(tempVal(counter), Chr(34), "")
            Next counter
        Next x
        
        fehler.ort = "Sub_CreateBAV_Database_RR: RS erstellen"
        
        On Error Resume Next
    
        'Eintagen der Werte in die Felder des Recordset
        For x = 0 To UBound(Split(StarArray(0), ","))
            If Not Trim(realval(3, x)) = "" And Not Trim(realval(4, x)) = "" Then
                .AddNew
                tempVal = Split(realval(0, x), " ")
                .Fields("ID").Value = x + 1
                .Fields("Kürzel").Value = Trim(tempVal(0))
                .Fields("Stbld").Value = Trim(tempVal(1))
                '.Fields("BP").Value = Trim(realval(1, x))
                .Fields("Typ").Value = Trim(realval(1, x))
                .Fields("Epoche").Value = Trim(realval(4, x))
                .Fields("Periode").Value = Trim(realval(5, x))
                .Fields("Max").Value = Trim(realval(6, x))
                .Fields("MinI").Value = Trim(realval(7, x))
                .Fields("M-m").Value = Trim(realval(8, x))
        
                realval(12, x) = Replace(Replace(Replace(realval(12, x), "h", ""), "m", ""), "s", "")
                tempVal = Split(realval(12, x), " ")
                .Fields("hh").Value = Trim(tempVal(0))
                .Fields("mm").Value = Trim(tempVal(1))
                .Fields("ss").Value = Trim(tempVal(2))
                
                realval(13, x) = Replace(Replace(realval(13, x), "°", ""), "'", "")
                tempVal = Split(realval(13, x), " ")
                .Fields("vz").Value = Trim(tempVal(0))
                .Fields("o").Value = Trim(tempVal(1))
                .Fields("m").Value = Trim(tempVal(2)) + Trim(tempVal(3)) / 60
                .Update
            End If
        Next x
        
        Err.Clear
        
        fehler.ort = ""
    
    .Save App.Path & "\BAVBA_RR.dat"
 
 End With
 
 fs.DeleteFile (TemporaryPath)
 
 Set fs = Nothing
 Set BAVBeobAuf = Nothing
 
 MsgBox "Erzeugung der Datenbank erfolgreich." & vbCrLf & "Die Daten können nun für Berechnungen verwendet werden.", vbInformation, "Download erfolgreich"

  frmHaupt.Form_Load
  frmHaupt.cmdListe.Enabled = True: frmHaupt.VTabs.TabEnabled(1) = True
  frmHaupt.cmbGrundlage.Enabled = True
 For x = 1 To frmHaupt.cmbGrundlage.ListCount
   If frmHaupt.cmbGrundlage.List(x) = "BAV-BA_RR" Then
    frmHaupt.cmbGrundlage.ListIndex = x
    Exit For
   End If
 Next
 
 Exit Function
 
errhandler:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf & fehler.ort, vbCritical, "Programmfehler"
    fehler.ort = ""
End Function

