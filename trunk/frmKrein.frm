VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKrein 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Aktualisierung der Kreiner-DB"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
      Begin VB.OptionButton optD_save 
         Caption         =   "vorhandene Datenbank sichern"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton OptD_overwrite 
         Caption         =   "Datenbank überschreiben"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdLocKrein 
      Caption         =   "...."
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frmOptions 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optD_umwand 
         Caption         =   "vorhandene Datei Umwandeln"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton optD_Down 
         Caption         =   "aktuelle Datei downloaden"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog cdlgKrein 
      Left            =   3120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSpeichPfad 
      Caption         =   "Bitte Pfad zu der Datei allstars-cat.txt angeben:"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmKrein"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fs As FileSystemObject
Dim rsa As ADODB.Recordset
Dim einstrom As TextStream
Dim ausstrom As TextStream
Dim zeile1 As String
Dim zeile2 As String
Dim zeile As String
Dim zähler As Integer
Dim KreinerThere As Boolean
Dim result


Private Sub cmdDownload_Click()
If optD_umwand = True Then
    cmdLocKrein.Visible = True
    lblSpeichPfad.Visible = True
    cmdDownload.Visible = False
    Exit Sub
End If
Call DownAndChange
End Sub

Private Sub cmdLocKrein_Click()
With cdlgKrein
.InitDir = App.Path
.Filter = "Kreiner DB (allstars-cat.txt)|allstars-cat.txt"
.FileName = ""
.ShowOpen
If .FileName = "" Then Exit Sub
Call DownAndChange
End With

End Sub

Private Sub Form_Load()
Set fs = New FileSystemObject
If fs.FileExists(App.Path & "\kreiner.dat") = False Then
KreinerThere = False
 optD_umwand.Enabled = True
 optD_save.Enabled = False
 OptD_overwrite.Enabled = False
End If

KreinerThere = True
optD_Down = True
optD_save = True

End Sub

Sub DownAndChange()

Set fs = New FileSystemObject
Set rsa = New ADODB.Recordset

On Error GoTo ErrorHandler

Me.MousePointer = 11
DoEvents
If optD_Down = True Then

    cmdDownload.BackColor = &HC0FFFF
    cmdDownload.Caption = "Download in Arbeit..."
    
    result = CBool(InternetGetConnectedState(0, 0))
    
    If result = False Then
    'Testen oder Herstellen der Internetverbindung
    result = RASConnect(Me.hWnd)
    End If
    
    If result = False Then
      MsgBox "Internetverbindung konnte nicht aufgebaut werden," & _
      vbCrLf & "bitte überprüfen Sie Ihre Einstellungen", vbCritical, "Keine Verbindung..."
      Me.MousePointer = 1
      Unload Me
      Exit Sub
    ElseIf result = True Then
    result = URLDownloadToFile(0, "http://www.as.wsp.krakow.pl/ephem/allstars-cat.txt", _
    App.Path & "\kreiner.txt", 0, 0)
    End If
    
End If

If fs.FileExists(App.Path & "\kreiner.txt") = False And Not optD_umwand = True Then
     MsgBox "Es ist ein Fehler beim Download aufgetreten" & vbCrLf & _
     "Bitte versuchen Sie es erneut..", vbCritical, "Download nicht erfolgreich"
     cmdDownload.BackColor = &H8000000F
     cmdDownload.Caption = "Start"
     Me.MousePointer = 1
     Exit Sub
End If
     
     With rsa
     .Fields.Append ("ID"), adInteger
     .Fields.Append ("Kürzel"), adVarChar, 6
     .Fields.Append ("Stbld"), adChar, 3
     .Fields.Append ("BP"), adChar, 3
     .Fields.Append ("LBeob"), adDouble
     .Fields.Append ("Max"), adDouble
     .Fields.Append ("MinI"), adDouble
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
End With

If OptD_overwrite = True Then
    If fs.FileExists(App.Path & "\kreiner.dat") Then fs.DeleteFile (App.Path & "\kreiner.dat")
ElseIf optD_save = True Then
    If fs.FileExists(App.Path & "\kreiner.dat") = True And KreinerThere = True Then
    fs.CopyFile App.Path & "\kreiner.dat", App.Path & "\kreiner" & Format(Date, "ddmmyyyy") & ".dat"
    fs.DeleteFile (App.Path & "\kreiner.dat")
End If
End If

Set fs = New FileSystemObject

If optD_umwand = True Then
    fs.CopyFile cdlgKrein.FileName, App.Path & "\kreiner.txt"
 End If
 
Set einstrom = fs.OpenTextFile(App.Path & "\kreiner.txt")

'ACHTUNG!!!!DAtei wurde geändert!!! Keine Headerinfos mehr---Seit 7.1.08 doch wieder....
'=========================================================
zeile1 = einstrom.ReadLine
While InStr(1, zeile1, "^") = False
 zeile1 = einstrom.ReadLine
Wend
zeile1 = einstrom.ReadLine

'RT    And 2006  8.55 F8V       EA/DW/RS    for  ALL //bis Anfang 2009
'RT    And 2006  8.55  9.47 F8V       EA/DW/RS   ALL  minima elements //ab 03-2009

rsa.Open
zähler = 1

While Not einstrom.AtEndOfStream
zeile1 = einstrom.ReadLine
zeile2 = einstrom.ReadLine
'On Error Resume Next
If Not Trim(Mid(zeile2, 51, 11)) = "" And Not Trim(Left(zeile2, 2)) = "" And _
Not Trim(Mid(zeile2, 39, 10)) = "" Then

  With rsa
     .AddNew
     .Fields("ID") = zähler
     
     If InStr(1, Right(Trim(Mid(zeile1, 49, 10)), 3), "sec", vbTextCompare) <> 0 Then
        .Fields("Kürzel") = Trim(Left(zeile1, 6)) & "*"
     Else
        .Fields("Kürzel") = Trim(Left(zeile1, 6))
     End If
     
     .Fields("Stbld") = Trim(Mid(zeile1, 7, 3))
     .Fields("BP") = "KRE"
     .Fields("LBeob") = Trim(Mid(zeile1, 11, 5))
     .Fields("Max") = IIf(Trim(Mid(zeile1, 16, 5)) = "", 0, Trim(Mid(zeile1, 16, 5)))
     .Fields("MinI") = IIf(Trim(Mid(zeile1, 22, 5)) = "", 0, Trim(Mid(zeile1, 22, 5)))
     .Fields("Spektr") = Trim(Mid(zeile1, 28, 10))
     
'23 11 10.1 +53 01 33 2000.0  0.06 0.0 0.6289286 2452500.3510 0.5
'23 11 10.1 +53 01 33 2000.0   0.107 0.0 0.6289286 2452500.3510 0.5

     .Fields("D") = Trim(Mid(zeile2, 31, 6))
     .Fields("kD") = Trim(Mid(zeile2, 37, 4))
     .Fields("Typ") = Trim(Mid(zeile1, 32, 12))
     .Fields("Epoche") = Trim(Mid(zeile2, 53, 11))
     .Fields("Periode") = Trim(Mid(zeile2, 41, 10))
     .Fields("for") = Right(Trim(Mid(zeile1, 44, 10)), 3)
     .Fields("hh") = Left(zeile2, 2)
     .Fields("mm") = Mid(zeile2, 4, 2)
     .Fields("ss") = Trim(Mid(zeile2, 7, 4))
     .Fields("vz") = Mid(zeile2, 12, 1)
     .Fields("o") = Mid(zeile2, 13, 2)
     .Fields("m") = Round(Mid(zeile2, 16, 2) + CDbl(Mid(zeile2, 19, 2)) / 60, 3)
     .Update
     
  End With
 zähler = zähler + 1
 End If
 
Wend

   einstrom.Close
   
   rsa.Save App.Path & "\Kreiner.dat"
   rsa.Close
   Me.MousePointer = 1
   fs.DeleteFile (App.Path & "\kreiner.txt")
    Set einstrom = Nothing
    Set rsa = Nothing
    Set fs = Nothing
    frmHaupt.Form_Load
    
     For x = 1 To frmHaupt.cmbGrundlage.ListCount
        If frmHaupt.cmbGrundlage.List(x) = "Kreiner" Then
            frmHaupt.cmbGrundlage.ListIndex = x
        Exit For
        End If
    Next

    MsgBox "Die Kreiner-Datenbank kann jetzt für" & vbCrLf & "Berechnungen verwendet werden..", vbInformation, "Implementierung erfolgreich"
    frmHaupt.cmdListe.Enabled = True
    frmHaupt.cmbGrundlage.Enabled = True
    Unload Me
    Exit Sub
    
ErrorHandler:
     MsgBox "Es ist ein Fehler aufgetreten. Bitte versuchen Sie es erneut", vbCritical, "Fehler: Download/Umwandlung nicht erfolgreich"
     cmdDownload.BackColor = &H8000000F
     cmdDownload.Caption = "Start"
     Me.MousePointer = 1
     frmHaupt.Form_Load
     frmHaupt.cmdListe.Enabled = True
     frmHaupt.cmbGrundlage.Enabled = True
     Unload Me
     
End Sub




