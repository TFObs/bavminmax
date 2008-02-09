VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHaupt 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "BAV Min/Max  V1.08"
   ClientHeight    =   8145
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6810
   Icon            =   "frmHaupt.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8145
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   6861.694
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdRecherche 
      Caption         =   "Recherche"
      Height          =   495
      Left            =   4680
      TabIndex        =   36
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdListe 
      BackColor       =   &H00C0C000&
      Height          =   615
      Left            =   1080
      Picture         =   "frmHaupt.frx":08CA
      Style           =   1  'Grafisch
      TabIndex        =   35
      ToolTipText     =   "Berechnung starten"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdInternet 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6120
      Picture         =   "frmHaupt.frx":0BD4
      Style           =   1  'Grafisch
      TabIndex        =   34
      ToolTipText     =   "Internet Informationen"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdÖffnen 
      Height          =   495
      Left            =   3360
      Picture         =   "frmHaupt.frx":0EDE
      Style           =   1  'Grafisch
      TabIndex        =   33
      ToolTipText     =   "bestehende Abfrage öffnen"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdInfo 
      Height          =   495
      Left            =   4080
      Style           =   1  'Grafisch
      TabIndex        =   31
      ToolTipText     =   "Informationsfenster anzeigen"
      Top             =   120
      Width           =   491
   End
   Begin VB.CommandButton cmdGridgross 
      Height          =   495
      Left            =   4680
      Picture         =   "frmHaupt.frx":11E8
      Style           =   1  'Grafisch
      TabIndex        =   30
      ToolTipText     =   "Ergebnisliste in Extrafenster vergrößert darstellen"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdListdrucken 
      Height          =   495
      Left            =   6120
      Picture         =   "frmHaupt.frx":1AB2
      Style           =   1  'Grafisch
      TabIndex        =   29
      ToolTipText     =   "Ausdruck der Tabelle (WYSIWYG)"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbGrundlage 
      Height          =   315
      Left            =   1200
      TabIndex        =   28
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdErgebnis 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   2160
      Picture         =   "frmHaupt.frx":1DBC
      Style           =   1  'Grafisch
      TabIndex        =   26
      ToolTipText     =   "Ereignisse ansehen"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdAbfrag 
      BackColor       =   &H80000000&
      Height          =   615
      Left            =   120
      Picture         =   "frmHaupt.frx":20C6
      Style           =   1  'Grafisch
      TabIndex        =   25
      ToolTipText     =   "Abfragen"
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame frmOrt 
      Caption         =   "Beobachtungsort"
      Height          =   6135
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox gBreite 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   20
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox gLänge 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   19
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdOrtCancel 
         Caption         =   "Abbrechen"
         Height          =   495
         Left            =   4320
         TabIndex        =   14
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton cmdOrtOK 
         Height          =   495
         Left            =   4320
         Picture         =   "frmHaupt.frx":23D0
         Style           =   1  'Grafisch
         TabIndex        =   13
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "übernehmen"
         Height          =   255
         Left            =   4920
         TabIndex        =   32
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "geografische Länge: (Osten = positiv)"
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "geografische Breite:"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Bitte Koordinaten des Beobachtungsortes (in Grad) eingeben:"
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdListspeichern 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      Picture         =   "frmHaupt.frx":26DA
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Speichern in eine Textdatei"
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Rahmen 
      Caption         =   "Berechnungsergebnisse"
      DragMode        =   1  'Automatisch
      Height          =   6135
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   6615
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H000080FF&
         Caption         =   "Filter und Anzeige  >>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         Style           =   1  'Grafisch
         TabIndex        =   22
         ToolTipText     =   "Filtermöglichkeiten einblenden.."
         Top             =   240
         Width           =   1935
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdergebnis 
         Height          =   5175
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   9128
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblHinwZeit 
         Caption         =   "bitte beachten: alle Zeitangaben sind UT!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         Caption         =   "Noch keine Berechnungsergebnisse vorhanden..."
         Height          =   975
         Left            =   1200
         TabIndex        =   9
         Top             =   2160
         Width           =   2535
      End
   End
   Begin VB.Frame Rahmen 
      Caption         =   "Berechnungszeitraum"
      Height          =   6135
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6615
      Begin VB.Timer Timer1 
         Interval        =   20
         Left            =   2160
         Top             =   3120
      End
      Begin MSComctlLib.ProgressBar Balken 
         Height          =   375
         Left            =   360
         TabIndex        =   23
         ToolTipText     =   "Fortschrittsanzeige"
         Top             =   3720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSACAL.Calendar ende 
         Height          =   1455
         Left            =   4200
         TabIndex        =   10
         Top             =   4440
         Visible         =   0   'False
         Width           =   1695
         _Version        =   524288
         _ExtentX        =   2990
         _ExtentY        =   2566
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2006
         Month           =   2
         Day             =   19
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtStartdat 
         Alignment       =   2  'Zentriert
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtende 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox cmbDauer 
         Height          =   315
         Left            =   4920
         TabIndex        =   1
         ToolTipText     =   "Bitte Anzahl der zu berechnenden Tage auswählen"
         Top             =   1200
         Width           =   735
      End
      Begin MSACAL.Calendar Kalender 
         Height          =   2535
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "Mit einem Klick das Datum auswählen"
         Top             =   360
         Width           =   3735
         _Version        =   524288
         _ExtentX        =   6588
         _ExtentY        =   4471
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2005
         Month           =   12
         Day             =   20
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   -1  'True
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog cdlSpeichern 
         Left            =   3600
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         InitDir         =   "pfad"
      End
      Begin VB.Label lblfertig 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label lblZeitr 
         Caption         =   "Zeitraum [Tage] :"
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblStart 
         Caption         =   "Startdatum :"
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblEnd 
         Caption         =   "Enddatum :"
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmHaupt.frx":29E4
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3360
      Picture         =   "frmHaupt.frx":2CEE
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Datenbank:"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu proggi 
      Caption         =   "Programm"
      Begin VB.Menu Beenden 
         Caption         =   " beenden"
      End
   End
   Begin VB.Menu setup 
      Caption         =   "Einstellungen.."
      Index           =   1
      Begin VB.Menu ort 
         Caption         =   "Beobachtungsort"
      End
      Begin VB.Menu Dämmer 
         Caption         =   "Filter für Dämmerung"
         Begin VB.Menu astro 
            Caption         =   "astronomisch"
            Checked         =   -1  'True
         End
         Begin VB.Menu naut 
            Caption         =   "nautisch"
         End
         Begin VB.Menu burger 
            Caption         =   "bürgerlich"
         End
         Begin VB.Menu SaH 
            Caption         =   "Sonne am Horizont"
         End
      End
   End
   Begin VB.Menu helKorr 
      Caption         =   "heliozentrische Korrektur"
      Begin VB.Menu hKorrberech 
         Caption         =   "berechnen"
      End
   End
   Begin VB.Menu DB 
      Caption         =   "Datenbanken"
      Begin VB.Menu DBKrein 
         Caption         =   "Kreiner"
         Begin VB.Menu DBKrein_aktual 
            Caption         =   "laden/aktualisieren"
         End
      End
      Begin VB.Menu DBGcvs 
         Caption         =   "GCVS"
         Begin VB.Menu DBGcvs_aktual 
            Caption         =   "laden/aktualisieren"
         End
      End
   End
   Begin VB.Menu hilfe 
      Caption         =   "Hilfe"
   End
End
Attribute VB_Name = "frmHaupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API-Aufrufe um Trennzeichen und Datumsformat zu ermitteln
Private Declare Function GetSystemDefaultLCID Lib _
        "kernel32" () As Long

Private Declare Function GetLocaleInfo Lib "kernel32" _
        Alias "GetLocaleInfoA" (ByVal Locale As Long, _
        ByVal LCType As Long, ByVal lpLCData As String, _
        ByVal cchData As Long) As Long

Private Declare Function SetLocaleInfo Lib "kernel32.dll" _
                 Alias "SetLocaleInfoA" ( _
                 ByVal Locale As Long, _
                 ByVal LCType As Long, _
                 ByVal lpLCData As String) As Long
                 
'Api für die Hilfedatei
Private Declare Function HtmlHelp Lib "hhctrl.ocx" _
            Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
            ByVal pszFile As String, ByVal uCommand As _
            Long, ByVal dwData As Long) As Long

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_CLOSE_ALL As Long = &H12

'Konstanten für die Trennzeichen und das Datumsformat
Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF
Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Const LOCALE_STIMEFORMAT = &H1003      '  time format string

Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SYSTEM_DEFAULT As Long = &H400


Dim dbsBAVSterne As ADODB.Connection 'Grunddaten
Dim rstAbfrage As ADODB.Recordset    'Abfrage auf Grunddaten
Dim rsergebnis As ADODB.Recordset    'temporäre Ergebnisdatei ergebnisse.dat
Dim fs As New FileSystemObject
Dim result



Private Sub Berechnungsfilter_Click()
 frmBerechnungsfilter.Show
End Sub

Private Sub Beenden_Click()
 Call Form_Unload(0)
End Sub





Private Sub cmdinfo_Click()

    'Ändern des Icon bei Klick
    If cmdInfo.Picture = Image2 Then
        frmSterninfo.Show
        cmdInfo.Picture = Image1
        cmdInfo.ToolTipText = "Informationsfenster ausblenden"
        cmdInternet.Enabled = True
    Else
        cmdInfo.Picture = Image2
        cmdInternet.Enabled = False
        cmdInfo.ToolTipText = "Informationsfenster öffnen"
        Unload frmInternet
        Unload frmSterninfo
    End If
    
    If grdergebnis.col = 1 Then Call Infofüllen(grdergebnis)
    
    
End Sub

Private Sub cmbGrundlage_Click()

 If fs.FileExists(pfad & "\ergebnisse.dat") Then
    cmdErgebnis.Enabled = False
    cmdListspeichern.Enabled = False
    cmdGridgross.Enabled = False
    cmdListdrucken.Enabled = False
    cmdInfo.Enabled = False
 End If
 
End Sub

Private Sub cmdErgebnis_Click()

    gridfüllen
    Me.Rahmen(2).Visible = True
    'Me.Width = 11040
    Me.Rahmen(1).Visible = False
    frmOrt.Visible = False
    cmbGrundlage.Enabled = False
    cmdListdrucken.Enabled = True
    cmdGridgross.Enabled = True
    lblHinwZeit.FontBold = True
    cmdInfo.Enabled = True
    cmdAbfrag.Enabled = True
    cmdListe.Enabled = False
    cmdinfo_Click
    cmdFilter_Click
End Sub

Private Sub cmdAbfrag_Click()

 If Rahmen(1).Visible Then Exit Sub
 
    Rahmen(1).Visible = True
    Rahmen(2).Visible = False
    Me.Width = 7050
    frmOrt.Visible = False
    
  Unload frmInternet
  Unload frmBerechnungsfilter
  Unload frmSterninfo
  
  lblfertig.Caption = ""
  
  ' Aktivieren wenn sonstige Datei vorhanden!
  cmbGrundlage.Enabled = True
  cmdInfo.Enabled = False
  
  cmdAbfrag.Enabled = False
  cmdInfo.Picture = Image2
  cmdInfo.ToolTipText = "Informationsfenster ausblenden"
  cmdInternet.Enabled = False
  cmdListe.Enabled = True
End Sub

Private Sub cmdFilter_Click()
 frmBerechnungsfilter.Show
End Sub

Private Sub cmdGridgross_Click()

If frmGridGross.Visible Then Unload frmGridGross
 frmGridGross.Show
 
End Sub

Private Sub cmdInternet_Click()
result = CheckInetConnection(Me.hwnd)
If result = False Then
Unload frmInternet
Exit Sub
End If

frmInternet.Show
End Sub

'Öffnen einer bestehenden Abfrage
Private Sub cmdÖffnen_Click()
Set dbsBAVSterne = New ADODB.Connection
Set rstAbfrage = New ADODB.Recordset
Dim vstrfile, Y

Me.Rahmen(2).Visible = True
'Me.Width = 11040
Me.Rahmen(1).Visible = False

'infodatei vorhanden?
If fs.FileExists(App.Path & "\info.dat") Then
  fs.DeleteFile (App.Path & "\info.dat")
End If

 'Ermitteln des Dateinamens
        With cdlSpeichern
            .Filter = "Abfragen (*.abf)|*.abf"
            .InitDir = pfad
            .MaxFileSize = 2000
            .DialogTitle = "bestehende Abfrage laden"
            .ShowOpen
        vstrfile = .FileName
        End With
        
        If vstrfile = "" Then
          If cmdErgebnis.Enabled Then cmdErgebnis.Enabled = True
            Call cmdAbfrag_Click
            Exit Sub
        End If
        
 cmdErgebnis.Enabled = False
 
 'Grid aus Datei füllen
 DoEvents
 
 grdergebnis.MousePointer = 11
 
 Call LoadGridData(grdergebnis, vstrfile, ";")
 
grdergebnis.Visible = True

    'Einstellen der Spaltenbreiten
    With grdergebnis
 
        .ColWidth(0) = 200
        
    If Not .ColWidth(1) = 0 Then
        .ColWidth(1) = 800
    End If
        
    If Not .ColWidth(5) = 0 Then
        .ColWidth(5) = 1200
    End If
    
    If Not .ColWidth(6) = 0 Then
        .ColWidth(6) = 600
    End If
    
    If Not .ColWidth(7) = 0 Then
        .ColWidth(7) = 600
    End If
    
    'BAV_Sterne oder BAv_sonstige?
    If .TextMatrix(1, 8) = "" Then
        .ColWidth(8) = 0
        .ColWidth(9) = 800
        Database = 0
    ElseIf Trim(.TextMatrix(1, 8)) = "KRE" Then
        .ColWidth(8) = 0
        .ColWidth(9) = 1300
        Database = 2
    ElseIf Trim(.TextMatrix(1, 8)) = "GCVS" Then
        .ColWidth(8) = 0
        .ColWidth(9) = 1300
        Database = 3
    Else: .ColWidth(9) = 0
        .ColWidth(8) = 1300
        Database = 1
    End If
    
    
    If Not .ColWidth(10) = 0 Then
        .ColWidth(10) = 1200
    End If

    For x = 0 To .Cols - 1
        .ColHeaderCaption(0, x) = .TextMatrix(0, x)
        .ColAlignmentHeader(0, x) = 3
     Next x
     
 .ColAlignmentFixed = flexAlignCenterCenter
 .ColAlignment = 4
 .Cols = .Cols - 1
 
 End With
 
 cmdListspeichern.Enabled = True
 cmdGridgross.Enabled = True
 cmdListdrucken.Enabled = True
 cmdAbfrag.Enabled = True
 cmdFilter.Enabled = True
 cmdInfo.Enabled = True
 
 grdergebnis.MousePointer = 1
 
 Set rsergebnis = New ADODB.Recordset

'Neue Ergebnis-DAtei muß erzeugt werden, da sonst Abfrage nicht funktioniert!!
With rsergebnis
    .Open pfad & "\ergebnisse.dat", , , adLockOptimistic

            'Löschen einer alten Abfrage
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete (adAffectCurrent)
                    .MoveNext
                Wend
            End If
            .Save pfad & "\ergebnisse.dat"
         
End With

'Erstellen der Spalten im Recordset
With rsergebnis
For Y = 1 To grdergebnis.Rows - 1
 
  .AddNew
  .Fields("Stbld") = grdergebnis.TextMatrix(Y, 1)
  .Fields("stern") = grdergebnis.TextMatrix(Y, 2)
  .Fields("Datum") = grdergebnis.TextMatrix(Y, 3)
  .Fields("Uhrzeit") = grdergebnis.TextMatrix(Y, 4)
  
  .Fields("stundenwinkel") = grdergebnis.TextMatrix(Y, 5)
  .Fields("Höhe") = grdergebnis.TextMatrix(Y, 6)
  .Fields("Azimut") = grdergebnis.TextMatrix(Y, 7)
  .Fields("BProg") = grdergebnis.TextMatrix(Y, 8)
  .Fields("Typ") = grdergebnis.TextMatrix(Y, 9)
  .Fields("Epochenzahl") = grdergebnis.TextMatrix(Y, 10)
  .Fields("Monddist") = grdergebnis.TextMatrix(Y, 11)

Next Y

.Update
.Save pfad & "\ergebnisse.dat"

     
    'Abfrage als info.dat speichern, damit info im Fenster erscheint!
    With rstAbfrage
        If Database < 2 Then
        
            'Verbindung zur Datenbank herstellen
            With dbsBAVSterne
                .Provider = "microsoft.Jet.oledb.4.0"
                If Database = 0 Then
                    .ConnectionString = pfad & "\Bav_sterne.mdb"
                ElseIf Database = 1 Then
                    .ConnectionString = pfad & "\BAV_sonstige.mdb"
                End If
                .Open
             End With
             
            .ActiveConnection = dbsBAVSterne
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .CursorType = adOpenForwardOnly ' Kleinster Verwaltungsaufwand
            .Open "BVundRR"
            
        ElseIf Database = 2 Then .Open pfad & "\Kreiner.dat"
        ElseIf Database = 3 Then .Open pfad & "\GCVS.dat"
        
        End If
        .Save pfad & "\info.dat"
    End With
        
'Zerstören der Objekte
Set rstAbfrage = Nothing
Set rsergebnis = Nothing
Set dbsBAVSterne = Nothing
Set fs = Nothing

cmbGrundlage.Enabled = False

End With

End Sub

Public Sub cmdOrtCancel_Click()
    'Register einblenden
    frmOrt.Visible = False
    Rahmen(1).Visible = True
    Rahmen(2).Visible = False
   Me.Width = 7050
End Sub

Private Sub cmdOrtOK_Click()

    'Abfangen einer Fehleingabe
    If Not IsNumeric(gBreite.text) Or _
    Not IsNumeric(gLänge.text) Or _
    Abs(gBreite.text) > 90 Or Abs(gLänge) > 180 Then
        MsgBox "Bitte überprüfen Sie die Eingabe" & vbCrLf & "Nur numerische Werte zwischen " & vbCrLf & _
        ": +- 90° für die Breite und" & vbCrLf & ": +-180° für die Länge zugelassen", vbExclamation, "Eingabefehler"
        Exit Sub
    End If
    
    'Speichern von geogr. Breite und Länge in der Registry
    Call INISetValue(App.Path & "\Prog.ini", "Ort", "Breite", gBreite.text)
    Call INISetValue(App.Path & "\Prog.ini", "Ort", "Länge", gLänge.text)
        
    cmdOrtCancel_Click  'um Register wieder einzublenden
    Me.Form_Load
    Me.cmbGrundlage.Enabled = True
    Me.cmdListe.Enabled = True
End Sub

Private Sub cmdRecherche_Click()
If frmSterninfo.lblKoord.Caption <> "" Then
result = Split(frmSterninfo.lblKoord.Caption, vbCrLf)
frmAladin.txtObj.text = Trim(Mid(CStr(result(0)), 3, Len(CStr(result(0))) - 2)) & " " & Trim(Mid(CStr(result(1)), 4, Len(CStr(result(1))) - 2))
End If
frmAladin.Show
End Sub

Private Sub DBGcvs_aktual_Click()

result = CheckInetConnection(Me.hwnd)
If result = False Then Exit Sub
frmGCVS.Show
End Sub



Private Sub DBKrein_aktual_Click()

result = CheckInetConnection(Me.hwnd)
If result = False Then Exit Sub
frmKrein.Show
End Sub

Public Sub Form_Load()
If App.PrevInstance Then End
Set rsergebnis = New ADODB.Recordset
Dim result
Dim dezimal$, Tausend$
Dim DatumsFormat$, ZeitFormat$

pi = 4 * Atn(1)
pi = pi / 180

'Ermittlung des Programmpfades
If Right(App.Path, 1) <> "\" Then
        pfad = App.Path & "\"
    Else
        pfad = App.Path
    End If

datei = pfad & "\prog.ini"
If Not fs.FileExists(App.Path & "\prog.ini") Or Err.Number = 13 Then
MsgBox "Beschädigte oder fehlende Konfigurationsdatei," & vbCrLf _
& "... es werden jetzt Standardwerte geladen!", vbCritical, "Fehler der Konfigurationsdatei"
DoEvents
Call DefaultWerte
End If

'grdergebnis.ToolTipText = "Mit Doppelklick auf den Spaltenkopf wird eine Spalte ausgeblendet," & vbCrLf & _
'"mit einem einfachen Klick wird die Spalte sortiert"
sorter = 7


'Einstellungen für Trennzeichen und Datum/Zeit
dezimal = GetTrennzeichen(LOCALE_SDECIMAL)
Tausend = GetTrennzeichen(LOCALE_STHOUSAND)
DatumsFormat = GetTrennzeichen(LOCALE_SSHORTDATE)
ZeitFormat = GetTrennzeichen(LOCALE_STIMEFORMAT)

SaveSetting App.Title, "Trennzeichen", "dezimal", dezimal
SaveSetting App.Title, "Trennzeichen", "tausend", Tausend
SaveSetting App.Title, "DatumsFormat", "Format", DatumsFormat
SaveSetting App.Title, "ZeitFormat", "Format", ZeitFormat


result = SetTrennzeichen(LOCALE_SDECIMAL, ".")
result = SetTrennzeichen(LOCALE_STHOUSAND, ",")
result = SetTrennzeichen(LOCALE_SSHORTDATE, "dd.MM.yyyy")
result = SetTrennzeichen(LOCALE_STIMEFORMAT, "HH:mm:ss")

     
    ende.Visible = False
    Kalender.Value = Date
    ende.Value = Kalender.Value + 1

    'Füllen des Zeitraumsfeldes
    For x = 1 To 30
        cmbDauer.AddItem x
    Next x
    cmbDauer.ListIndex = 0  'Zeiger der combobox auf 1. Eintrag
    txtende = ende.Value - 1

    cmdListspeichern.Enabled = False
    cmdListdrucken.Enabled = False
    cmdGridgross.Enabled = False
    cmdInfo.Enabled = False
       
    'Falls schon vorhanden
    If fs.FileExists(pfad & "\ergebnisse.dat") Then
         fs.DeleteFile (pfad & "\ergebnisse.dat")
    End If
    If fs.FileExists(pfad & "\info.dat") Then
        fs.DeleteFile (pfad & "\info.dat")
    End If
     'Erstellen von Tabelle und Spalten
     With rsergebnis
        
     .Fields.Append "Stbld", adChar, 3
     .Fields.Append "Stern", adVarChar, 255
     .Fields.Append "Datum", adVarChar, 255
     .Fields.Append "Uhrzeit", adVarChar, 6
     .Fields.Append "Stundenwinkel", adVarChar, 8
     .Fields.Append "Azimut", adVarNumeric, 6
     .Fields.Append "Höhe", adVarNumeric, 5
     .Fields.Append "BProg", adChar, 5
     .Fields.Append "Typ", adChar, 12                   'zunächst beide Felder erstellen
     .Fields.Append "Epochenzahl", adInteger, 255
     .Fields.Append "Monddist", adVarNumeric, 10
    .Open
    .Save pfad & "\ergebnisse.dat"
    .Close
    End With
'Zerstören der Objekte
Set fs = Nothing
With cmbGrundlage
.Clear
.AddItem ("BAV-Programmsterne")
.AddItem ("sonstige BAV-Sterne")
If fs.FileExists(App.Path & "\kreiner.dat") = True Then
.AddItem ("Kreiner")
End If
If fs.FileExists(App.Path & "\GCVS.dat") = True Then
.AddItem ("GCVS")
End If
.ListIndex = 0
End With

'Übernahme der Einstellungen

'ort_Click
'cmdOrtOK_Click
Me.Rahmen(2).Visible = False
Me.Width = 7050
Me.Rahmen(1).Visible = True
frmOrt.Visible = False
cmdInfo.Picture = Image2
cmdInfo.ToolTipText = "Informationsfenster ausblenden"
cmdInfo.Enabled = False
cmdErgebnis.Enabled = False
Set rsergebnis = Nothing
cmdAbfrag.Enabled = False
End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim result
    'Löschen der Abfragedateien bei Programmende
    If fs.FileExists(pfad & "\ergebnisse.dat") Then
        fs.DeleteFile (pfad & "\ergebnisse.dat")
        
    End If
      If fs.FileExists(pfad & "\info.dat") Then
        fs.DeleteFile (pfad & "\info.dat")
        
    End If
    
    If fs.FileExists(pfad & "\filter.dat") Then
        fs.DeleteFile (pfad & "\filter.dat")
      End If
    'Zurücksetzen der Trennzeichen und des Datums/Zeitformats
    result = SetTrennzeichen(LOCALE_SDECIMAL, GetSetting(App.Title, "Trennzeichen", "dezimal"))
    result = SetTrennzeichen(LOCALE_STHOUSAND, GetSetting(App.Title, "Trennzeichen", "tausend"))
    result = SetTrennzeichen(LOCALE_SSHORTDATE, GetSetting(App.Title, "DatumsFormat", "Format"))
    result = SetTrennzeichen(LOCALE_STIMEFORMAT, GetSetting(App.Title, "ZeitFormat", "Format"))

    
    Set fs = Nothing
    Unload frmBerechnungsfilter
    Unload frmInternet
    Unload frmSterninfo
    Unload frmGridGross
    Unload frmHelioz
    Unload frmGrafik
    Call HtmlHelp(frmHaupt.hwnd, "", HH_CLOSE_ALL, 0)
    Unload Me

End Sub
Private Sub astro_Click()
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "astronomisch")
naut.Checked = False
astro.Checked = True
burger.Checked = False
SaH.Checked = False
Call UnloadAll
End Sub

Private Sub burger_Click()
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "bürgerlich")
naut.Checked = False
astro.Checked = False
burger.Checked = True
SaH.Checked = False
Call UnloadAll
End Sub



Private Sub hilfe_Click()
Dim HDatei As String
    HDatei = App.Path & "\BAVMinMax.chm"
    Call HtmlHelp(0, HDatei, HH_DISPLAY_TOPIC, ByVal 0&)
End Sub

Private Sub hKorrberech_Click()
frmHelioz.Show
End Sub

Private Sub naut_Click()
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "nautisch")
naut.Checked = True
astro.Checked = False
burger.Checked = False
SaH.Checked = False
Call UnloadAll
End Sub

Private Sub Register_Click(Index As Long)
End Sub


Private Sub SaH_Click()
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "S. am Horizont")
naut.Checked = False
astro.Checked = False
burger.Checked = False
SaH.Checked = True
Call UnloadAll
End Sub


Private Sub ort_Click()
    frmOrt.Visible = True
    'frmSpalten.Visible = False
    If cmdErgebnis.Enabled = True Then
        Unload frmBerechnungsfilter
        Unload frmSterninfo
        Unload frmInternet
    End If
    
    'Register ausblenden
    Rahmen(1).Visible = False
    Rahmen(2).Visible = False
    Me.Width = 7050
    'Ermitteln der Werte aus Registry
    gBreite.text = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite")
    gLänge.text = INIGetValue(App.Path & "\Prog.ini", "Ort", "Länge")

    'Wenn nicht vorhanden, dann Standardwerte
    If gBreite.text = "" Then gBreite.text = Format(CDbl(50#), "#.00")
    If gLänge.text = "" Then gLänge.text = Format(CDbl(10#), "#.00")
    
End Sub



Private Sub Form_Resize()
    Dim i As Byte
    
       For i = 1 To Me.Rahmen.UBound - 1
        Me.Rahmen(i).Width = Me.Rahmen(1).Width
        Me.Rahmen(i).Top = Me.Rahmen(1).Top
        Me.Rahmen(i).Left = Me.Rahmen(1).Left
        Me.Rahmen(i).Height = Me.Rahmen(1).Height
       Next i
    'Me.Rahmen(2).Width = Me.Rahmen(1).Width + (10540 - 6302)
    
End Sub

'Übernahme der Kalenderwerte
Private Sub kalender_click()
    ende.Value = Kalender.Value + CInt(cmbDauer.text)
    txtende = ende.Value - 1
    txtStartdat = Kalender.Value
    cmdListspeichern.Enabled = False
    cmdInfo.Enabled = False
    cmdGridgross.Enabled = False
    cmdListdrucken.Enabled = False
        cmdErgebnis.Enabled = False
End Sub


'Zeitraum wählen
Private Sub cmbDauer_Click()
    kalender_click
    grdergebnis.Clear
    lblfertig.Caption = ""
End Sub


Public Sub cmdListe_click()
Set dbsBAVSterne = New ADODB.Connection
Set rstAbfrage = New ADODB.Recordset
Set rsergebnis = New ADODB.Recordset
Dim APeriode, EPeriode, ereignis
Dim BAnfang, bende, JDEreignis
Dim result, RA As Double, DEC As Double, BPrg As String
Dim Typ As String
Dim gLänge, gBreite
Dim Uhrzeit As Double, höhe As Double
Dim Stundenwinkel As Double, Sternzeit As Double
Dim Tag, Jahr, Sauf, Sunter
Dim aktHoehe As Double, aktAzimut As Double
Dim überw
Dim ephem, monddist
'lblfertig.Caption = ""
lblfertig.Visible = True

'informationsdatei für frmSterninfo
If fs.FileExists(App.Path & "\info.dat") Then
fs.DeleteFile (App.Path & "\info.dat")
End If


'Wenn INI Datei nicht da oder beschädigt
If Not fs.FileExists(App.Path & "\prog.ini") Or Err.Number = 13 Then
DoEvents
Form_Load
Exit Sub
End If
Me.MousePointer = 11

gLänge = INIGetValue(App.Path & "\Prog.ini", "Ort", "Länge")
gBreite = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite")
grdergebnis.Clear
 
         
     'Öffnen der Tabelle "Grundlage" und löschen einer evtl. alten Abfrage
    With rsergebnis
    .Open pfad & "\ergebnisse.dat", , , adLockOptimistic

        

            'Löschen einer alten Abfrage
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete (adAffectCurrent)
                    .MoveNext
                Wend
            End If
            .Save pfad & "\ergebnisse.dat"
        
         
    End With
   
    
    'Verbindung zur Datenbank herstellen
    With dbsBAVSterne
        .Provider = "microsoft.Jet.oledb.4.0"
        
        'Auswahl der Datengrundlage, mit der die Berechnung durchgeführt wird
        If cmbGrundlage.ListIndex = 0 Then
        .ConnectionString = pfad & "\Bav_sterne.mdb"
        .Open
        ElseIf cmbGrundlage.ListIndex = 1 Then
        .ConnectionString = pfad & "\BAV_sonstige.mdb"
        .Open
        End If
        
    End With
   Database = cmbGrundlage.ListIndex
        
    'Öffnen der Tabelle "BVundRR" aus der BAV_Sterne.mdb
    'als temporäres Recordset oder verbinden mit Recordset
    'aus der Kreiner DB
    
    With rstAbfrage
      If cmbGrundlage.ListIndex < 2 Then
        .ActiveConnection = dbsBAVSterne
        .LockType = adLockOptimistic
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly ' Kleinster Verwaltungsaufwand
        .Open "BVundRR"
     ElseIf cmbGrundlage.List(cmbGrundlage.ListIndex) = "Kreiner" Then
        Database = 2
        .Open pfad & "\Kreiner.dat"
     ElseIf cmbGrundlage.List(cmbGrundlage.ListIndex) = "GCVS" Then
        Database = 3
        .Open pfad & "\GCVS.dat"
    End If
        
      'Öffnen der Ergebnisdatei und springen zu ersten Eintrag
      
    
        .MoveFirst
      DoEvents
      
    Balken.Value = 0
    'Berechnungen für alle Sterne in BAV_Sterne.mdb
    Do While Not .EOF
    DoEvents
        lblfertig.Caption = "Berechnungen zu " & Format((Balken.Value / rstAbfrage.RecordCount) * 100, "#") & "% fertiggestellt"
    Balken.Max = rstAbfrage.RecordCount
        If Not .Fields("epoche").ActualSize = 0 And Not .Fields("periode").ActualSize = 0 Then
        'Berechnung der Ereignisse im gewählten Zeitraum
        BAnfang = JulDat(Left(Kalender.Value, 2), Mid(Kalender.Value, 4, 2), Mid(Kalender.Value, 7, 4))
        bende = JulDat(ende.Day, ende.Month, ende.year)
        EPeriode = Fix((bende - (!epoche + 2400000)) / !periode)
        APeriode = Fix((BAnfang - (!epoche + 2400000)) / !periode)
   
        For x = APeriode To EPeriode
             ereignis = (!epoche + 2400000) + x * !periode
   
            If ereignis > BAnfang And ereignis < bende Then
            
                If Not .Fields("hh").ActualSize = 0 Then
                        RA = !hh + !mm / 60 + !ss / 3600
                        DEC = CDbl(!vz & !O + !m / 60)
                    Else: RA = 0
                 End If
                 
                 'Heliozentrische Korrektur der HJD Zeitpunkte!!!
                 ereignis = Hkorr(ereignis, RA, DEC, False)
                 
                JDEreignis = JulinDat(ereignis)
                Tag = CDate(Fix(JDEreignis))
                Jahr = Format(JDEreignis, "yyyy")
                Sauf = AufUnter(Tag, Jahr, 0) * 24
                If Sauf = "600" Then
                
            MsgBox "Für den gewählten Filter für den" & vbCrLf & _
            "Sonnenauf- und -untergang können keine Werte berechnet werden!" & vbCrLf & vbCrLf _
            & "    Es ist die Zeit der weißen oder schwarzen Nächte...!" & Chr(13) & Chr(13) & _
            "Bitte den Filter für die Dämmerung auf 'Sonne am Horizont' stellen.", vbInformation, "Berechnung nicht möglich"
            Me.MousePointer = 1
            Exit Sub
            End If
            
                Sunter = AufUnter(Tag, Jahr, 1) * 24
                Uhrzeit = 24 * ((JDEreignis) - Fix(JDEreignis))
    
                If Sunter < Uhrzeit And Uhrzeit <= 24 Or _
                    Sauf > Uhrzeit And Uhrzeit >= 0 Then
        
                    
                     BPrg = !BP
                     If Database >= 2 Then Typ = !Typ
                     
                     
                    Sternzeit = STZT(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit, CDbl(gLänge))
                    
                    Stundenwinkel = CDbl(stdw(RA, Sternzeit))
                    aktHoehe = Hoehe(Stundenwinkel, CDbl(gBreite), DEC)
                    aktAzimut = Azimut(aktHoehe, Stundenwinkel, CDbl(gBreite), DEC)
                    
                    'Berechnung der Monddistanz
                    Mpi = 4 * Atn(1)
                    Mdeg = (4 * Atn(1)) / 180
                    Mrad = 180 / (4 * Atn(1))
                    sonne = SunPosition(JulDat(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit))
                    
                    mond = MoonPosition(sonne(2), sonne(3), JulDat(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit))
                    
                    ephem = MoonRise(JulDat(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit), 65, CDbl(gLänge) * Mdeg, CDbl(gBreite) * Mdeg, 0, 1)
                    
                    monddist = Moondistance(RA * 360 / 24, DEC, (mond(0) * Mrad / 15) / 24 * 360, mond(1) * Mrad)

                    If aktHoehe > 0 Then 'Filter der Höhe über Horizont
                        
                       With rsergebnis
                        
                        .AddNew
                        .Fields("Stbld") = rstAbfrage.Fields("Stbld")
                        .Fields("stern") = rstAbfrage.Fields("Kürzel") & " " & rstAbfrage.Fields("Stbld")
                        .Fields("Datum") = Format(JDEreignis, "dd.mm.yyyy")
                        .Fields("Uhrzeit") = Left(Format(JDEreignis, "hh:mm:ss"), 5)
                        .Fields("Monddist") = (Format(Round(monddist), "0"))
                                                 
                         If Stundenwinkel > 0.5 Then
                         .Fields("stundenwinkel") = "E " & Format(1 - Stundenwinkel, "hh:mm")
                          aktAzimut = 360 - aktAzimut
                          
                            ElseIf Stundenwinkel < 0.5 Then
                               .Fields("stundenwinkel") = "W " & Format(Stundenwinkel, "hh:mm")
                               
                            Else: .Fields("Stundenwinkel") = Format(Stundenwinkel, "hh:mm")
                         
                        End If
                        
                        .Fields("Höhe") = Format(aktHoehe, "0") '"#.00")
                        .Fields("Azimut") = Format(aktAzimut, "0") '"#.0")
                        .Fields("Epochenzahl") = x
                        
                        If Database = 0 Then
                          .Fields("BProg") = BPrg
                        ElseIf Database = 1 Then
                          .Fields("Typ") = BPrg
                         ElseIf Database >= 2 Then
                          .Fields("BProg") = BPrg
                          .Fields("Typ") = Typ
                        End If

                        .Update
                        
                        
     
                        End With
                    End If
                
                End If
                
            End If
        Next x

        End If

    .MoveNext
    

    Balken.Value = Balken.Value + 1
    
    Loop
Balken.Value = 0
lblfertig.Caption = "Berechnungen beendet!"
'Übernahme der Spalten in Recordset
rsergebnis.Save pfad & "\ergebnisse.dat"
cmdFilter.Enabled = True
cmdErgebnis.Enabled = True
'Schließen beider *.mdb's  und
'des Recordsets
'rsergebnis.Close
rstAbfrage.Save pfad & "\info.dat"
rstAbfrage.Close
If Database < 2 Then dbsBAVSterne.Close


End With
Me.MousePointer = 1

'Zerstören der Objekte
'Set rsergebnis = Nothing
Set rstAbfrage = Nothing
Set rsergebnis = Nothing
Set dbsBAVSterne = Nothing
Set fs = Nothing
'cmbGrundlage.Enabled = False

End Sub


Public Sub gridfüllen(Optional ByVal sSQL)
   Dim rsa As ADODB.Recordset
    Dim Abf As ADODB.Recordset
    Dim filt As ADODB.Recordset
    Dim zähler As Integer
        Set dicStbld = New Dictionary
'Fehlerbehandlung
    If Not fs.FileExists(pfad & "\ergebnisse.dat") Then
        Exit Sub
    End If
    If fs.FileExists(pfad & "\filter.dat") Then fs.DeleteFile (pfad & "\filter.dat")
      On Error GoTo Abbruch
      
      Set rsa = New ADODB.Recordset
    If rsa.State = adStateOpen Then rsa.Close
      'Recordset laden
      With rsa
         .CursorType = adOpenKeyset
         .LockType = adLockReadOnly
         .Open pfad & "\ergebnisse.dat", , , adLockOptimistic
       End With
   If rsa.RecordCount = 0 Then Exit Sub
   
   rsa.MoveFirst
   Do While Not rsa.EOF
    On Error Resume Next
      dicStbld.Add rsa.Fields("Stbld").Value, "  "
      rsa.MoveNext
   Loop

      
Set Abf = rsa
If sSQL <> "" Then
For zähler = 0 To 3
Abf.Filter = sSQL(zähler)
Abf.Save pfad & "\filter.dat"
Abf.Close
Abf.Open pfad & "\filter.dat"
Next zähler
End If



If Abf.RecordCount > 0 Then
Set grdergebnis.DataSource = Abf
Else:
MsgBox "Für diese Filterkombination konnten keine Ereignisse" & vbCrLf & _
"gefunden werden...", vbInformation, "Keine Ergebnisse für diesen Filter"
Exit Sub
End If


grdergebnis.ColAlignment = 4

grdergebnis.ColWidth(0) = 200
If Not grdergebnis.ColWidth(1) = 0 Then
grdergebnis.ColWidth(1) = 800
End If
If Not grdergebnis.ColWidth(5) = 0 Then
grdergebnis.ColWidth(5) = 1200
End If
If Not grdergebnis.ColWidth(6) = 0 Then
grdergebnis.ColWidth(6) = 600
End If
If Not grdergebnis.ColWidth(7) = 0 Then
grdergebnis.ColWidth(7) = 600
End If

If Database = 0 Then
        grdergebnis.ColWidth(8) = 600
        grdergebnis.ColWidth(9) = 0
End If

If Database = 1 Then

        grdergebnis.ColWidth(9) = 1300
        
        grdergebnis.ColWidth(8) = 0
        
End If

If Database >= 2 Then
        grdergebnis.ColWidth(8) = 0
        grdergebnis.ColWidth(9) = 1300
End If


If Not grdergebnis.ColWidth(10) = 0 Then
grdergebnis.ColWidth(10) = 1200
End If
grdergebnis.ColAlignmentFixed = flexAlignCenterCenter
grdergebnis.Sort = 5
grdergebnis.ColAlignment = 4
    
      'Recordset schliessen, Speicher freigeben
      rsa.Close
      Set rsa = Nothing
      Abf.Close
      Set Abf = Nothing
      Err.Clear
      cmdListspeichern.Enabled = True
      cmdGridgross.Enabled = True
      cmdListdrucken.Enabled = True
      grdergebnis.Visible = True
      Exit Sub
  If fs.FileExists(pfad & "\filter.dat") Then fs.DeleteFile (pfad & "\filter.dat")
Abbruch:
      MsgBox "Fehler: " & Err.Number & vbCrLf & _
             Err.Description, vbCritical
      Err.Clear
End Sub


'Liste in gewünschte Textdatei speichern
Private Sub cmdListspeichern_Click()
Dim datei As String, zeile As String
Dim spalten As Long, zeilen As Long
Dim x As Long, Y As Long
Dim ikanal As Integer
        
        'Ermitteln des Dateinamens
        With cdlSpeichern
            .Filter = "Abfragen (*.abf)|*.abf"
            .InitDir = pfad
            .MaxFileSize = 2000
            .DialogTitle = "Abfrage speichern"
            .ShowSave
        datei = .FileName
        End With

    If datei = "" Then
        Exit Sub
    End If
    
    'Ausgabe der Gridmatrix in eine Textdatei
    ikanal = FreeFile()

    Open datei For Output As ikanal
        spalten = grdergebnis.Cols
        zeilen = grdergebnis.Rows
        zeile = ";"
        For Y = 0 To zeilen - 1
            For x = 1 To spalten - 1
                zeile = zeile & grdergebnis.TextMatrix(Y, x) & ";"
            Next x
  
            Print #ikanal, zeile
            zeile = ";" 'Löschen des Zeileninhalts
        Next Y
        
    Close #ikanal
  
End Sub

'Trennzeichen-Einstellung ermitteln
Private Function GetTrennzeichen(ID&) As String
Dim lcid&, result&, Buffer$, Length&
    
    lcid = GetSystemDefaultLCID()
    Length = GetLocaleInfo(lcid, ID, Buffer, 0) - 1
    Buffer = Space(Length + 1)
    result = GetLocaleInfo(lcid, ID, Buffer, Length)
    GetTrennzeichen = Left$(Buffer, Length)
End Function

'Trennzeichen Setzen
Private Function SetTrennzeichen(ID&, wert$) As Long
Dim lcid&
    
    lcid = GetSystemDefaultLCID()
    SetTrennzeichen = SetLocaleInfo(lcid, ID, wert)
      
End Function

Public Sub grdergebnis_sortieren(ByVal ColIndex As Integer)
    'Sortieren Datagrid
    grdergebnis.MousePointer = 11
    frmGridGross.grdGross.col = ColIndex
    If ColIndex < 6 Or ColIndex = 9 Then
    If sorter = 7 Then
         grdergebnis.Sort = 8
         sorter = 8
        Else
        grdergebnis.Sort = 7
        sorter = 7
        End If
    End If
    
    If ColIndex >= 6 And Not ColIndex = 9 Then
    If sorter = 7 Then
        grdergebnis.Sort = 4
        sorter = 8
     Else
     grdergebnis.Sort = 3
     sorter = 7
     End If
     End If
            
    frmGridGross.grossGrid_füllen
    grdergebnis.MousePointer = 1
End Sub


Private Sub grdergebnis_DblClick()

'BAV_Sterne/Kreiner oder BAV_Sonstige?
If Database = 1 Then
 frmBerechnungsfilter.chkSpalte(8).Visible = False
 ElseIf Database = 0 Or Database >= 2 Then
 frmBerechnungsfilter.chkSpalte(9).Visible = False
End If
 
       grdergebnis.ColWidth(grdergebnis.col) = 0
       frmBerechnungsfilter.chkSpalte(grdergebnis.col) = 0
       frmBerechnungsfilter.chkSpalte(grdergebnis.col).Value = 0
         
End Sub

Private Sub grdergebnis_MouseUp(Button As Integer, _
  Shift As Integer, x As Single, Y As Single)
  
  ' Rechtsklick?
  If Button = vbRightButton Then
    Dim nRow As Long
    Dim nCol As Long
    Dim nRowCur As Long
    Dim nColCur As Long
    Dim data As String
    
    With grdergebnis
      ' aktuelle Zelle "merken"
      nRowCur = .Row
      nColCur = .col
      
      ' Zelle ermitteln
      nRow = .MouseRow
      nCol = .MouseCol
      
      ' Erfolgte der Mausklick auf eine existierende
      ' Zelle im Grid?
      If nRow >= 0 And nCol >= 0 Then
        ' Zelle zur aktiven Zelle machen (selektieren)
        .Row = nRow
        .col = nCol
        
        If x >= .CellLeft And x <= .CellLeft + .CellWidth And _
          Y >= .CellTop And Y <= .CellTop + .CellHeight Then
        
          ' Beispiel: PopUp-Menü anzeigen
         
          For x = 1 To .Cols - 1
           data = data & .TextMatrix(nRow, x) & " ; "
           Next x
           
        Else
          ' ursprüngliche Zelle wieder selektieren
          .Row = nRowCur
          .col = nColCur
        End If
      End If
    End With
  End If
  
  If Button = vbLeftButton Then
  
  If grdergebnis.MouseRow = 0 Then
    Call grdergebnis_sortieren(grdergebnis.col)
      End If
  End If
 If grdergebnis.col = 1 Then
  Call Infofüllen(grdergebnis)
  End If
End Sub

Private Sub cmdListdrucken_Click()
  Call PrintGrid(grdergebnis, 20, 25, 20, 20, _
                 "BAV MinMax - aktuelle Abfrage" & _
                 "", "")
                 MsgBox "Druckauftrag wurde an den" & Chr(13) & "Standarddrucker gesendet!", vbInformation, "Druckauftrag gesendet"
End Sub

Private Sub UnloadAll()
  Unload frmInternet
  cmdInternet.Enabled = False
  Unload frmBerechnungsfilter
  Unload frmSterninfo
  Me.Form_Load
  cmbGrundlage.Enabled = True
End Sub
