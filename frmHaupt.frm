VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmHaupt 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "VarEphem"
   ClientHeight    =   8130
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6855
   Icon            =   "frmHaupt.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8130
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   6907.036
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdStarChart 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4200
      Picture         =   "frmHaupt.frx":08CA
      Style           =   1  'Grafisch
      TabIndex        =   42
      ToolTipText     =   "Horizont-Ansicht erstellen"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab VTabs 
      Height          =   6615
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Beobachtungsort"
      TabPicture(0)   =   "frmHaupt.frx":397F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmOrt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Abfrage"
      TabPicture(1)   =   "frmHaupt.frx":399B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Rahmen(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ergebnisse"
      TabPicture(2)   =   "frmHaupt.frx":39B7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Rahmen(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "einzelner Stern"
      TabPicture(3)   =   "frmHaupt.frx":39D3
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Suche in Datenbanken"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74880
         TabIndex        =   32
         Top             =   380
         Width           =   6315
         Begin VB.ListBox ListRecherche 
            Height          =   1035
            Left            =   480
            MultiSelect     =   1  '1 -Einfach
            TabIndex        =   36
            Top             =   2400
            Width           =   5055
         End
         Begin VB.CommandButton cmdSingleAusw 
            Caption         =   "ausw�hlen"
            Enabled         =   0   'False
            Height          =   615
            Left            =   2280
            TabIndex        =   35
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdSingleSuch 
            Caption         =   "Suchen"
            Height          =   615
            Left            =   3960
            TabIndex        =   34
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtSingleStar 
            Alignment       =   2  'Zentriert
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   480
            TabIndex        =   33
            Top             =   1440
            Width           =   3375
         End
         Begin VB.Label Label5 
            Caption         =   "Stern     Stbld        Datenbank       Epoche                   Periode"
            Height          =   255
            Left            =   480
            TabIndex        =   37
            Top             =   2160
            Width           =   5055
         End
      End
      Begin VB.Frame Rahmen 
         Caption         =   "Berechnungszeitraum"
         Height          =   6135
         Index           =   1
         Left            =   -74880
         TabIndex        =   26
         Top             =   380
         Width           =   6315
         Begin VB.TextBox Text1 
            Alignment       =   2  'Zentriert
            Height          =   375
            Left            =   4440
            TabIndex        =   41
            Text            =   "1"
            Top             =   1560
            Width           =   360
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   4800
            TabIndex        =   40
            Top             =   1560
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "Text1"
            BuddyDispid     =   196617
            OrigLeft        =   2880
            OrigTop         =   5160
            OrigRight       =   3135
            OrigBottom      =   5535
            Max             =   365
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1680
            TabIndex        =   39
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777088
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            Format          =   73924611
            CurrentDate     =   3
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1680
            TabIndex        =   38
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            Format          =   73924611
            CurrentDate     =   2
         End
         Begin MSComctlLib.ProgressBar Balken 
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            ToolTipText     =   "Fortschrittsanzeige"
            Top             =   3600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComDlg.CommonDialog cdlSpeichern 
            Left            =   5520
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            InitDir         =   "pfad"
         End
         Begin VB.Label lblEnd 
            Caption         =   "Enddatum :"
            Height          =   255
            Left            =   1680
            TabIndex        =   31
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblStart 
            Caption         =   "Startdatum :"
            Height          =   255
            Left            =   1680
            TabIndex        =   30
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblZeitr 
            Caption         =   "Zeitraum [Tage] :"
            Height          =   495
            Left            =   3720
            TabIndex        =   29
            Top             =   1440
            Width           =   735
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
            Height          =   855
            Left            =   360
            TabIndex        =   28
            Top             =   4200
            Width           =   5415
         End
      End
      Begin VB.Frame Rahmen 
         Caption         =   "Berechnungsergebnisse"
         DragMode        =   1  'Automatisch
         Height          =   6135
         Index           =   2
         Left            =   -74880
         TabIndex        =   21
         Top             =   380
         Width           =   6315
         Begin VB.CommandButton cmdFilter 
            BackColor       =   &H000080FF&
            Caption         =   "Filter und Anzeige  >>"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4200
            Style           =   1  'Grafisch
            TabIndex        =   22
            ToolTipText     =   "Filterm�glichkeiten einblenden.."
            Top             =   240
            Width           =   1935
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdergebnis 
            Height          =   5175
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Visible         =   0   'False
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   9128
            _Version        =   393216
            BackColor       =   16777215
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblDat 
            Caption         =   "Datum"
            Height          =   255
            Left            =   3600
            TabIndex        =   46
            Top             =   3480
            Width           =   1575
         End
         Begin VB.Label lblAzimut 
            Caption         =   "Azimut"
            Height          =   375
            Left            =   3600
            TabIndex        =   45
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label lblZeit 
            Caption         =   "Uhrzeit"
            Height          =   255
            Left            =   3600
            TabIndex        =   44
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label lblStern 
            Caption         =   "Stern"
            Height          =   375
            Left            =   3600
            TabIndex        =   43
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Zentriert
            AutoSize        =   -1  'True
            Caption         =   "Keine Berechnungsergebnisse vorhanden..."
            Height          =   195
            Left            =   900
            TabIndex        =   25
            Top             =   2160
            Width           =   3135
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
            TabIndex        =   24
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame frmOrt 
         Caption         =   "Beobachtungsort"
         Height          =   6135
         Left            =   120
         TabIndex        =   12
         Top             =   380
         Visible         =   0   'False
         Width           =   6315
         Begin VB.CommandButton cmdOrtOK 
            Height          =   495
            Left            =   4320
            Picture         =   "frmHaupt.frx":39EF
            Style           =   1  'Grafisch
            TabIndex        =   16
            Top             =   3960
            Width           =   495
         End
         Begin VB.CommandButton cmdOrtCancel 
            Caption         =   "Abbrechen"
            Height          =   495
            Left            =   4320
            TabIndex        =   15
            Top             =   4560
            Width           =   1695
         End
         Begin VB.TextBox gL�nge 
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
            TabIndex        =   14
            Top             =   1800
            Width           =   1935
         End
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
            TabIndex        =   13
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Bitte Koordinaten des Beobachtungsortes (in Grad) eingeben:"
            Height          =   375
            Left            =   480
            TabIndex        =   20
            Top             =   480
            Width           =   4455
         End
         Begin VB.Label Label2 
            Caption         =   "geografische Breite:"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   19
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Zentriert
            Caption         =   "geografische L�nge: (Osten = positiv)"
            Height          =   495
            Index           =   1
            Left            =   360
            TabIndex        =   18
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "�bernehmen"
            Height          =   255
            Left            =   4920
            TabIndex        =   17
            Top             =   4080
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdListe 
      BackColor       =   &H00C0C000&
      Height          =   615
      Left            =   1080
      Picture         =   "frmHaupt.frx":3CF9
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Berechnung starten"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdInternet 
      Height          =   495
      Left            =   5400
      Picture         =   "frmHaupt.frx":4003
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Internet-Recherche"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmd�ffnen 
      Height          =   495
      Left            =   3360
      Picture         =   "frmHaupt.frx":430D
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "bestehende Abfrage �ffnen"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdInfo 
      Height          =   495
      Left            =   4080
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Informationsfenster anzeigen"
      Top             =   120
      Width           =   491
   End
   Begin VB.CommandButton cmdGridgross 
      Height          =   495
      Left            =   4680
      Picture         =   "frmHaupt.frx":4617
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Ergebnisliste in Extrafenster vergr��ert darstellen"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdListdrucken 
      Height          =   495
      Left            =   6120
      Picture         =   "frmHaupt.frx":4EE1
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Ausdruck der Tabelle (WYSIWYG)"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbGrundlage 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdErgebnis 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   2160
      Picture         =   "frmHaupt.frx":51EB
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Ereignisse ansehen"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdAbfrag 
      BackColor       =   &H80000000&
      Height          =   615
      Left            =   120
      Picture         =   "frmHaupt.frx":54F5
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Abfragen"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdListspeichern 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      Picture         =   "frmHaupt.frx":57FF
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Speichern in eine Textdatei"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmHaupt.frx":5B09
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3360
      Picture         =   "frmHaupt.frx":5E13
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Datenbank:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
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
      Begin VB.Menu D�mmer 
         Caption         =   "Filter f�r D�mmerung"
         Begin VB.Menu astro 
            Caption         =   "astronomisch"
            Checked         =   -1  'True
         End
         Begin VB.Menu naut 
            Caption         =   "nautisch"
         End
         Begin VB.Menu burger 
            Caption         =   "b�rgerlich"
         End
         Begin VB.Menu SaH 
            Caption         =   "Sonne am Horizont"
         End
      End
      Begin VB.Menu mnuMag 
         Caption         =   "Filter f�r Helligkeit"
      End
   End
   Begin VB.Menu Berech 
      Caption         =   "Berechnungen"
      Begin VB.Menu mnuEinzel 
         Caption         =   "einzelner Stern"
      End
      Begin VB.Menu hKorrberech 
         Caption         =   "heliozentrische Korrektur"
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
      Begin VB.Menu DBBAVAuf 
         Caption         =   "BAV-Beobachtungsaufrufe"
         Begin VB.Menu DB_BAVEA_aktual 
            Caption         =   "Bedeckungsver�nderliche"
         End
         Begin VB.Menu DB_BAVRR_aktual 
            Caption         =   "kurzper Pulsationssterne"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu DBEigen 
         Caption         =   "Eigene laden"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Hilfe"
      Begin VB.Menu hilfe 
         Caption         =   "Hilfedatei"
      End
      Begin VB.Menu Ueber 
         Caption         =   "�ber.."
      End
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
                 
'Api f�r die Hilfedatei
Private Declare Function HtmlHelp Lib "hhctrl.ocx" _
            Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
            ByVal pszFile As String, ByVal uCommand As _
            Long, ByVal dwData As Long) As Long

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_CLOSE_ALL As Long = &H12

'Konstanten f�r die Trennzeichen und das Datumsformat
Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF
Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Const LOCALE_STIMEFORMAT = &H1003      '  time format string

Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SYSTEM_DEFAULT As Long = &H400


Dim dbsbavsterne As ADODB.Connection 'Grunddaten
Dim rstAbfrage As ADODB.Recordset    'Abfrage auf Grunddaten
Dim rsergebnis As ADODB.Recordset    'tempor�re Ergebnisdatei ergebnisse.dat
Dim fs As New FileSystemObject
Dim result
Dim TOPROW As Integer
Dim lngResult As Long
Dim settings_changed As Boolean


 Dim rssourcerecord  As ADODB.Recordset
 Dim rssingleabfrage  As ADODB.Recordset
 Dim feld As Field
 Dim gew�hlt As Collection
 



Private Sub Berechnungsfilter_Click()
 frmBerechnungsfilter.show
End Sub

Private Sub Beenden_Click()
 Call Form_Unload(0)
End Sub


Private Sub cmdinfo_Click()

    '�ndern des Icon bei Klick
    If cmdInfo.Picture = Image2 Then
        frmSterninfo.show
        cmdInfo.Picture = Image1
        cmdInfo.ToolTipText = "Informationsfenster ausblenden"
        'frmHaupt.cmdStarChart.Visible = True
        'cmdInternet.Enabled = True
    Else
        cmdInfo.Picture = Image2
        'cmdInternet.Enabled = False
        cmdInfo.ToolTipText = "Informationsfenster �ffnen"
        
        Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
    End If
    
    If grdergebnis.col = 1 Then
        Call Infof�llen(grdergebnis)
        Call Mondinfo(grdergebnis)
    End If
    
End Sub

Private Sub cmbGrundlage_Click()

 If fs.FileExists(pfad & "\ergebnisse.dat") Then
    cmdErgebnis.Enabled = False: VTabs.TabEnabled(2) = False
    'If Not frmHaupt.cmbGrundlage.text = "Einzeln" Then
        'VTabs.TabVisible(3) = False
    'End If
    cmdListspeichern.Enabled = False
    cmdGridgross.Enabled = False
    cmdListdrucken.Enabled = False
    cmdInfo.Enabled = False
 End If
 
End Sub

Private Sub cmdErgebnis_Click()

    'gridf�llen
    Me.VTabs.TabEnabled(2) = True
    'Me.Width = 11040
    'Me.VTabs.TabEnabled(1) = False
    VTabs.TabEnabled(0) = False
    cmbGrundlage.Enabled = False
    cmdListdrucken.Enabled = True
    cmdGridgross.Enabled = True
    lblHinwZeit.FontBold = True
    cmdInfo.Enabled = True
    cmdAbfrag.Enabled = True
    cmdListe.Enabled = False: VTabs.TabEnabled(1) = False: VTabs.Tab = 2
    'cmdinfo_Click
    cmdFilter_Click
End Sub

Public Sub cmdAbfrag_Click()
VTabs.Tab = 1
cmbGrundlage.Enabled = True
 If cmdListe.Enabled Then Exit Sub
 
    VTabs.TabEnabled(1) = True
    'VTabs.TabEnabled(2) = False
    Me.Width = 7050
    VTabs.TabEnabled(0) = False
    
  Unload frmAladin
  Unload frmBerechnungsfilter
  Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
  
  lblfertig.Caption = ""
  
  ' Aktivieren wenn sonstige Datei vorhanden!
  cmbGrundlage.Enabled = True
  cmdInfo.Enabled = False
  
  'cmdAbfrag.Enabled = False
  cmdInfo.Picture = Image2
  cmdInfo.ToolTipText = "Informationsfenster ausblenden"
  'cmdInternet.Enabled = False
  cmdListe.Enabled = True: VTabs.TabEnabled(1) = True: VTabs.Tab = 1: cmbGrundlage.Enabled = True
  
   ' VTabs.TabVisible(3) = False
 
End Sub

Private Sub cmdFilter_Click()
 If Not grdergebnis.ColHeaderCaption(0, 1) = "" Then frmBerechnungsfilter.show
End Sub

Private Sub cmdGridgross_Click()
    If frmGridGross.Visible Then Unload frmGridGross
    frmGridGross.show
End Sub

Private Sub cmdInternet_Click()
    result = CheckInetConnection(Me.hWnd)
    If result = False Then
        MsgBox "Internetverbindung konnte nicht aufgebaut werden," & _
      vbCrLf & "dieses Feature steht daher nicht zur Vef�gung", vbCritical, "Keine Verbindung..."
        Unload frmAladin
    Exit Sub
End If

If frmSterninfo.lblKoord.Caption <> "" Then
    result = Split(frmSterninfo.lblKoord.Caption, vbCrLf)
    frmAladin.txtObj.text = Trim(Mid(CStr(result(0)), 3, Len(CStr(result(0))) - 2)) & " " & Trim(Mid(CStr(result(1)), 4, Len(CStr(result(1))) - 2))
End If

frmAladin.show
End Sub

'�ffnen einer bestehenden Abfrage
Private Sub cmd�ffnen_Click()
Set dbsbavsterne = New ADODB.Connection
Set rstAbfrage = New ADODB.Recordset
Dim vstrfile, y

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
          If cmdErgebnis.Enabled Then VTabs.TabEnabled(2) = True
            Call cmdAbfrag_Click
            Exit Sub
        End If
        
 cmdErgebnis.Enabled = False: VTabs.TabEnabled(2) = False
 
 'infodatei vorhanden?
If fs.FileExists(App.Path & "\info.dat") Then
  fs.DeleteFile (App.Path & "\info.dat")
End If

 'Grid aus Datei f�llen
 DoEvents
 
 grdergebnis.MousePointer = 11
 
 Call LoadGridData(grdergebnis, vstrfile, ";")
 
grdergebnis.Visible = True
frmHaupt.cmdStarChart.Visible = True

    'Einstellen der Spaltenbreiten
    With grdergebnis
        .ColWidth(12) = 0
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
        .ColWidth(8) = 600
        .ColWidth(9) = 1300
        Database = 0
    ElseIf Trim(.TextMatrix(1, 8)) = "KRE" Then
        .ColWidth(8) = 0
        .ColWidth(9) = 1300
        Database = 2
    ElseIf Trim(.TextMatrix(1, 8)) = "GCVS" Then
        .ColWidth(8) = 0
        .ColWidth(9) = 1300
        Database = 3
    ElseIf Trim(.TextMatrix(1, 8)) = "EIG" Then
        .ColWidth(8) = 0
        .ColWidth(9) = 1300
        Database = 7
   Else: .ColWidth(9) = 1300
        .ColWidth(8) = 600
        Database = 0
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

'Neue Ergebnis-DAtei mu� erzeugt werden, da sonst Abfrage nicht funktioniert!!
With rsergebnis
    .Open pfad & "\ergebnisse.dat", , , adLockOptimistic

            'L�schen einer alten Abfrage
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
For y = 1 To grdergebnis.Rows - 1
 
  .AddNew
  .Fields("Stbld") = grdergebnis.TextMatrix(y, 1)
  .Fields("stern") = grdergebnis.TextMatrix(y, 2)
  .Fields("Datum") = grdergebnis.TextMatrix(y, 3)
  .Fields("Uhrzeit") = grdergebnis.TextMatrix(y, 4)
  
  .Fields("stundenwinkel") = grdergebnis.TextMatrix(y, 5)
  .Fields("H�he") = grdergebnis.TextMatrix(y, 6)
  .Fields("Azimut") = grdergebnis.TextMatrix(y, 7)
  .Fields("BProg") = grdergebnis.TextMatrix(y, 8)
  .Fields("Typ") = grdergebnis.TextMatrix(y, 9)
  .Fields("Epochenzahl") = grdergebnis.TextMatrix(y, 10)
  .Fields("Monddist") = grdergebnis.TextMatrix(y, 11)
  .Fields("bc") = grdergebnis.TextMatrix(y, 12)
  .Fields("JDEreignis") = grdergebnis.TextMatrix(y, 13)

Next y

.Update
.Save pfad & "\ergebnisse.dat"

     
    'Abfrage als info.dat speichern, damit info im Fenster erscheint!
    With rstAbfrage
        If Database < 2 Then
        
            'Verbindung zur Datenbank herstellen
            With dbsbavsterne
                .Provider = "microsoft.Jet.oledb.4.0"
                If Database = 0 Then
                    .ConnectionString = pfad & "\Bav_sterne.mdb"
                'ElseIf Database = 1 Then
                 '   .ConnectionString = pfad & "\BAV_sterne.mdb" 'onstige.mdb"
                End If
                .Open
             End With
             
            .ActiveConnection = dbsbavsterne
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .CursorType = adOpenForwardOnly ' Kleinster Verwaltungsaufwand
            .Open "BVundRR"
            
        ElseIf Database = 2 Then .Open pfad & "\Kreiner.dat"
        ElseIf Database = 3 Then .Open pfad & "\GCVS.dat"
        ElseIf Database = 4 Then .Open pfad & "\Einzel.dat"
        End If
        .Save pfad & "\info.dat"
    End With
        
'Zerst�ren der Objekte
Set rstAbfrage = Nothing
Set rsergebnis = Nothing
Set dbsbavsterne = Nothing
Set fs = Nothing

Me.VTabs.TabEnabled(2) = True
'Me.Width = 11040
Me.VTabs.TabEnabled(1) = False: Me.VTabs.Tab = 2
cmdErgebnis.Enabled = True: VTabs.TabEnabled(2) = True
cmbGrundlage.Enabled = False: cmdListe.Enabled = False

End With

End Sub

Public Sub cmdOrtCancel_Click()
    'Register einblenden
    VTabs.TabVisible(0) = False
    VTabs.TabEnabled(1) = True
    VTabs.TabEnabled(2) = False
   Me.Width = 7050
End Sub

Private Sub cmdOrtOK_Click()

    'Abfangen einer Fehleingabe
    If Not IsNumeric(gBreite.text) Or _
    Not IsNumeric(gL�nge.text) Or _
    Abs(gBreite.text) > 90 Or Abs(gL�nge) > 180 Then
        MsgBox "Bitte �berpr�fen Sie die Eingabe" & vbCrLf & "Nur numerische Werte zwischen " & vbCrLf & _
        ": +- 90� f�r die Breite und" & vbCrLf & ": +-180� f�r die L�nge zugelassen", vbExclamation, "Eingabefehler"
        Exit Sub
    End If
    
    'Speichern von geogr. Breite und L�nge in der Registry
    Call INISetValue(App.Path & "\Prog.ini", "Ort", "Breite", gBreite.text)
    Call INISetValue(App.Path & "\Prog.ini", "Ort", "L�nge", gL�nge.text)
        
    cmdOrtCancel_Click  'um Register wieder einzublenden
    Me.Form_Load
    Me.cmbGrundlage.Enabled = True
    Me.cmdListe.Enabled = True: VTabs.TabEnabled(1) = True: VTabs.TabEnabled(3) = True
    VTabs.TabVisible(0) = False
End Sub



Private Sub cmdStarChart_Click()

Dim charturl As String
Dim azim As String
Dim plandat, planzeit As String
Dim ew, no
Dim Lat, lon

result = CheckInetConnection(Me.hWnd)
    If result = False Then
        MsgBox "Internetverbindung konnte nicht aufgebaut werden," & _
        vbCrLf & "dieses Feature steht daher nicht zur Vef�gung", vbCritical, "Keine Verbindung..."
        Exit Sub
    End If
    
On Error GoTo errhandler

lon = INIGetValue(App.Path & "\Prog.ini", "Ort", "L�nge")
Lat = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite")
no = IIf(Lat > 0, "North", "South"): ew = IIf(lon > 0, "East", "West")

With grdergebnis

 azim = Trim(IIf(.TextMatrix(.Row, 6) + 180 >= 360, .TextMatrix(.Row, 6) - 180, .TextMatrix(.Row, 6) + 180))
 '"utc=1998%2F02%2F14+20%3A55%3A01&" &
 planzeit = .TextMatrix(.Row, 4)
 plandat = .TextMatrix(.Row, 3)
 planzeit = "+" & Left(planzeit, 2) & "%3A" & Right(planzeit, 2) & "%3A00&"
 plandat = "utc=" & Right(plandat, 4) & "%2F" & Mid(plandat, 4, 2) & "%2F" & Left(plandat, 2)
 
End With

charturl = "http://www.fourmilab.ch/cgi-bin/Yourhorizon?date=1&" & plandat + planzeit & _
"&azimuth=" & azim & "&azideg=" & azim & "&fov=270%B0&lat=" & Lat & "%B0&ns=" & no & _
"&lon=" & lon & "%B0&ew=" & ew & "&coords=on&moonp=on&deepm=3.0&consto=on&constn=on&consts=on&constb=on&limag=5&starn=on&" & _
"starnm=1.5&starbm=3.5&showmb=-1.5&showmd=6.0&terrain=on&terrough=0.7&scenery=on&imgsize=480&" & _
"dynimg=y&fontscale=1.0&scheme=0&elements="


result = URLDownloadToFile(0, charturl, App.Path & "\skychart.gif", 0, 0)

If result = 0 Then
 picShowPicture frmPlanetarium.Picture1, App.Path & "\skychart.gif", True
 With grdergebnis
 frmPlanetarium.Caption = .TextMatrix(.Row, 2) & " am " & .TextMatrix(.Row, 3) & " " & _
 .TextMatrix(.Row, 4) & "UT     Azimut: " & .TextMatrix(.Row, 6) & Chr(176) & ", H�he: " & .TextMatrix(.Row, 7) & Chr(176)
End With
frmPlanetarium.show
Else
 MsgBox "Es ist ein Fehler beim Download des Charts aufgetreten" & vbCrLf & _
     "Bitte versuchen Sie es erneut..", vbCritical, "Download nicht erfolgreich"
End If
Exit Sub
errhandler:
MsgBox "Es ist ein Fehler beim Erstellen des Charts aufgetreten" & vbCrLf & _
     "�berpr�fen Sie Ihre Internetverbindung und ggf. versuchen Sie es " & vbCrLf & _
     "zu einem sp�teren Zeitpunkt noch einmal..", vbCritical, "Fehler beim Erstellen"
     'Unload frmPlanetarium
End Sub

Public Sub picShowPicture(oPictureBox As Object, _
  ByVal sFile As String, _
  Optional ByVal bStretch As Boolean = True)
 
  With oPictureBox
    If bStretch Then
      ' Bild an Gr��e der PictureBox anpassen
      .AutoRedraw = True
      Set .Picture = Nothing
      .PaintPicture LoadPicture(sFile), 0, 0, .ScaleWidth, .ScaleHeight
      .AutoRedraw = False
    Else
      ' PictureBox an Bildgr��e anpassen
      Set .Picture = Nothing
      .Picture = LoadPicture(sFile)
      .AutoSize = True
    End If
  End With
End Sub



Private Sub DB_BAVEA_aktual_Click()
Dim BAVUrl As String

BAVUrl = "http://www.bav-astro.de/ea/beob_aufr.php?jahr=" & Format(Date, "YY") & "&monat=" & Format(Date, "m") '"http://www.bav-astro.de/ea/beob_aufr_" & _
Format(Date, "yy") & "_" & Format(Date, "mm") & ".html"

result = CheckInetConnection(Me.hWnd)
If result = False Then Exit Sub
Call CreateBAV_Database_EA(App.Path & "\testbav_EA.txt", BAVUrl)

End Sub

Private Sub DB_BAVRR_aktual_Click()
Dim BAVUrl As String

BAVUrl = "http://www.bav-astro.de/ea/beob_aufr.php?jahr=9&monat=1" '"http://www.bav-astro.de/rrlyr/beob_aufr_" & _
Format(Date, "yy") & "_" & Format(Date, "mm") & ".html"

result = CheckInetConnection(Me.hWnd)
If result = False Then Exit Sub
Call CreateBAV_Database_RR(App.Path & "\testbav_RR.txt", BAVUrl)

End Sub

Private Sub DBASAS_Click()
Call frmGCVS.createasasas
End Sub

Private Sub DBEigen_Click()
 frmEigene.show
End Sub

Private Sub DBGcvs_aktual_Click()

'result = CheckInetConnection(Me.hWnd)
'If result = False Then Exit Sub
frmGCVS.show
End Sub



Private Sub DBKrein_aktual_Click()

'result = CheckInetConnection(Me.hWnd)
'If result = False Then Exit Sub
frmKrein.show
End Sub

Private Sub DTPicker1_Change()
If DTPicker1.Value >= DTPicker2.Value Then DTPicker2.Value = DTPicker1.Value + 1
Text1.text = DTPicker2.Value - DTPicker1.Value
DTPicker2.Value = DTPicker1.Value + Text1.text
End Sub

Private Sub DTPicker2_change()
If DTPicker2.Value <= DTPicker1.Value Then DTPicker2.Value = DTPicker1.Value + 1
If DTPicker2.Value - DTPicker1.Value > 365 Then DTPicker2.Value = DTPicker1.Value + 365
Text1.text = DTPicker2.Value - DTPicker1.Value
End Sub



Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
Text1.text = DTPicker2.Value - DTPicker1.Value
End Sub

Public Sub Form_Load()
If App.PrevInstance Then End
Set rsergebnis = New ADODB.Recordset
Dim result
Dim daemmfil As String
Dim dezimal$, Tausend$
Dim DatumsFormat$, ZeitFormat$
Err.Clear

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
MsgBox "Besch�digte oder fehlende Konfigurationsdatei," & vbCrLf _
& "... es werden jetzt Standardwerte geladen!", vbCritical, "Fehler der Konfigurationsdatei"
DoEvents
Call DefaultWerte
End If

'F�r den Fall, dass �ltere Ini-Version vorhanden ist, Standardwerte erg�nzen
 If INIGetValue(datei, "Standard", "minMag_Max") = "" Then
    Call INISetValue(datei, "Standard", "minMag_Max", 18)
    Call INISetValue(datei, "Standard", "minMag_Min", 18)
    Call INISetValue(datei, "filter", "minMag_Max", 18)
    Call INISetValue(datei, "filter", "minMag_Min", 18)
 End If
 
Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
Unload frmBerechnungsfilter
Unload frmAladin
Unload frmSingleBerech

'grdergebnis.ToolTipText = "Mit Doppelklick auf den Spaltenkopf wird eine Spalte ausgeblendet," & vbCrLf & _
'"mit einem einfachen Klick wird die Spalte sortiert"
'sorter = 7
sorter = True
'If hook = 0 Then MInit Me

TOPROW = 1

If settings_changed = False Then
    'Einstellungen f�r Trennzeichen und Datum/Zeit
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
settings_changed = True
End If

DTPicker1.CustomFormat = "dd.MM.yyyy"
DTPicker2.CustomFormat = "dd.MM.yyyy"
    
   If DTPicker1.Value = "01.01.1900" Then
    DTPicker1.Value = Date
    DTPicker2.Value = DTPicker1.Value + Text1.text
    ReDim coltrigger(1 To 11)
    coltrigger = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
  End If
  
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
     .Fields.Append "H�he", adVarNumeric, 5
     .Fields.Append "BProg", adChar, 5
     .Fields.Append "Typ", adChar, 12                   'zun�chst beide Felder erstellen
     .Fields.Append "Epochenzahl", adInteger, 255
     .Fields.Append "Monddist", adVarNumeric, 10
     .Fields.Append "bc", adInteger, 1
     .Fields.Append "JDEreignis", adDouble, 32
    .Open
    .Save pfad & "\ergebnisse.dat"
    .Close
    End With
'Zerst�ren der Objekte
Set fs = Nothing
With cmbGrundlage
.Clear
.AddItem ("BAV-Programmsterne")
'.AddItem ("sonstige BAV-Sterne")
If fs.FileExists(App.Path & "\kreiner.dat") = True Then
.AddItem ("Kreiner")
End If
If fs.FileExists(App.Path & "\GCVS.dat") = True Then
.AddItem ("GCVS")
End If
If fs.FileExists(App.Path & "\acvs1.1.dat") = True Then
    .AddItem ("ASAS")
End If
If fs.FileExists(App.Path & "\Einzel.dat") = True Then
.AddItem ("Einzeln")
End If
If fs.FileExists(App.Path & "\BAVBA_EA.dat") = True Then
.AddItem ("BAV-BA_EA")
End If
If fs.FileExists(App.Path & "\BAVBA_RR.dat") = True Then
'.AddItem ("BAV-BA_RR")
End If
If fs.FileExists(App.Path & "\Eigene.dat") = True Then
 .AddItem ("Eigene")
End If

.ListIndex = 0
End With

'�bernahme der Einstellungen
daemmfil = INIGetValue(App.Path & "\Prog.ini", "Auf- Untergang", "D�mmerung")

'Ber�cksichtigung des Filters f�r Sonnenaufgang
Select Case daemmfil
    Case Is = "b�rgerlich": burger.Checked = True: astro.Checked = False: naut.Checked = False: SaH.Checked = False
    Case Is = "nautisch": burger.Checked = False: astro.Checked = False: naut.Checked = True: SaH.Checked = False
    Case Is = "astronomisch": burger.Checked = False: astro.Checked = True: naut.Checked = False: SaH.Checked = False
    Case Is = "S. am Horizont": burger.Checked = False: astro.Checked = False: naut.Checked = False: SaH.Checked = True
    Case Else:: burger.Checked = False: astro.Checked = True: naut.Checked = False: SaH.Checked = False
End Select

'ort_Click
'cmdOrtOK_Click
Me.VTabs.TabEnabled(2) = False
Me.Width = 7050
Me.VTabs.TabEnabled(1) = True
Me.VTabs.Tab = 1
VTabs.TabVisible(0) = False
cmdInfo.Picture = Image2
cmdInfo.ToolTipText = "Informationsfenster ausblenden"
cmdInfo.Enabled = False
cmdErgebnis.Enabled = False: VTabs.TabEnabled(2) = False
'VTabs.TabVisible(3) = IIf(fs.FileExists(App.Path & "/recordsets.dat"), True, False)
Set rsergebnis = Nothing
'cmdAbfrag.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Mende
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim result
Mende
    'L�schen der Abfragedateien bei Programmende
    If fs.FileExists(pfad & "\ergebnisse.dat") Then
        fs.DeleteFile (pfad & "\ergebnisse.dat")
        
    End If
      If fs.FileExists(pfad & "\info.dat") Then
        fs.DeleteFile (pfad & "\info.dat")
        
    End If
    
    If fs.FileExists(pfad & "\filter.dat") Then
        fs.DeleteFile (pfad & "\filter.dat")
      End If
      
If settings_changed Then
    'Zur�cksetzen der Trennzeichen und des Datums/Zeitformats
    result = SetTrennzeichen(LOCALE_SDECIMAL, GetSetting(App.Title, "Trennzeichen", "dezimal"))
    result = SetTrennzeichen(LOCALE_STHOUSAND, GetSetting(App.Title, "Trennzeichen", "tausend"))
    result = SetTrennzeichen(LOCALE_SSHORTDATE, GetSetting(App.Title, "DatumsFormat", "Format"))
    result = SetTrennzeichen(LOCALE_STIMEFORMAT, GetSetting(App.Title, "ZeitFormat", "Format"))
    settings_changed = False
End If
    
    Set fs = Nothing
    Unload frmBerechnungsfilter
    Unload frmAladin
    Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
    Unload frmGridGross
    Unload frmHelioz
    Unload frmGrafik
    Unload frmSingleBerech
    Unload frmKrein
    Unload frmGridGross
    Unload frmGCVS
    Unload frmPlanetarium
    Unload frmEigene
    Call HtmlHelp(frmHaupt.hWnd, "", HH_CLOSE_ALL, 0)
    Unload Me

End Sub

Private Sub astro_Click()
Call INISetValue(datei, "Auf- Untergang", "D�mmerung", "astronomisch")
naut.Checked = False
astro.Checked = True
burger.Checked = False
SaH.Checked = False
Call UnloadAll
cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
End Sub

Private Sub burger_Click()
Call INISetValue(datei, "Auf- Untergang", "D�mmerung", "b�rgerlich")
naut.Checked = False
astro.Checked = False
burger.Checked = True
SaH.Checked = False
Call UnloadAll
cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
End Sub



Private Sub hilfe_Click()
Dim HDatei As String
    HDatei = App.Path & "\VarEphem.chm"
    Call HtmlHelp(0, HDatei, HH_DISPLAY_TOPIC, ByVal 0&)
End Sub

Private Sub hKorrberech_Click()
frmHelioz.show
End Sub

Private Sub ListRecherche_Click()
If Not ListRecherche.ListCount = 0 Then cmdSingleAusw.Enabled = True
End Sub

Private Sub mnuEinzel_Click()

Dim x As Integer
For x = 1 To frmHaupt.cmbGrundlage.ListCount
 If frmHaupt.cmbGrundlage.List(x) = "Einzeln" Then
    frmHaupt.cmbGrundlage.ListIndex = x
    Exit For
 End If
 Next
VTabs.TabVisible(3) = True
VTabs.Tab = 3
End Sub

Private Sub mnuMag_Click()
  
 If cmdErgebnis.Enabled = True Then
        Unload frmBerechnungsfilter
        Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
        Unload frmAladin
    End If
    
    'Register ausblenden
    VTabs.TabEnabled(1) = True
    VTabs.TabEnabled(2) = False
    VTabs.TabEnabled(3) = False
    Me.Width = 7050
    VTabs.Tab = 1
    frmHaupt.Enabled = False
    frmHelligkeit.show
    
End Sub

Private Sub naut_Click()
Call INISetValue(datei, "Auf- Untergang", "D�mmerung", "nautisch")
naut.Checked = True
astro.Checked = False
burger.Checked = False
SaH.Checked = False
Call UnloadAll
cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
End Sub

Private Sub Register_Click(Index As Long)
End Sub


Private Sub SaH_Click()
Call INISetValue(datei, "Auf- Untergang", "D�mmerung", "S. am Horizont")
naut.Checked = False
astro.Checked = False
burger.Checked = False
SaH.Checked = True
Call UnloadAll
cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
End Sub


Private Sub ort_Click()
    VTabs.TabVisible(0) = True
    frmOrt.Visible = True: VTabs.TabEnabled(0) = True
    'frmSpalten.Visible = False
    If cmdErgebnis.Enabled = True Then
        Unload frmBerechnungsfilter
        Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
        Unload frmAladin
    End If
    
    'Register ausblenden
    VTabs.TabEnabled(1) = False
    VTabs.TabEnabled(2) = False
    VTabs.TabEnabled(3) = False
    Me.Width = 7050
    'Ermitteln der Werte aus Registry
    gBreite.text = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite")
    gL�nge.text = INIGetValue(App.Path & "\Prog.ini", "Ort", "L�nge")

    'Wenn nicht vorhanden, dann Standardwerte
    If gBreite.text = "" Then gBreite.text = Format(CDbl(50#), "#.00")
    If gL�nge.text = "" Then gL�nge.text = Format(CDbl(10#), "#.00")
    VTabs.Tab = 0
End Sub



Private Sub Form_Resize()
    Dim i As Byte
    
       For i = 1 To Me.Rahmen.UBound - 1
        Me.Rahmen(i).Width = Me.Rahmen(1).Width
        Me.Rahmen(i).Top = Me.Rahmen(1).Top
        Me.Rahmen(i).Left = Me.Rahmen(1).Left
        Me.Rahmen(i).Height = Me.Rahmen(1).Height
       Next i
        
End Sub


Public Sub cmdListe_click()
Set dbsbavsterne = New ADODB.Connection
Set rstAbfrage = New ADODB.Recordset
Set rsergebnis = New ADODB.Recordset
Dim APeriode, EPeriode, ereignis
Dim BAnfang, bende, JDEreignis
Dim result, RA As Double, DEC As Double, BPrg As String
Dim Typ As String
Dim gL�nge, gBreite
Dim Uhrzeit As Double, h�he As Double
Dim Stundenwinkel As Double, Sternzeit As Double
Dim Tag, Jahr, Sauf, Sunter
Dim aktHoehe As Double, aktAzimut As Double
Dim �berw
Dim ephem, monddist
Dim sPeriode
Dim SonneAU
Dim numereignis

On Error GoTo errhandler

'lblfertig.Caption = ""
lblfertig.Visible = True

'informationsdatei f�r frmSterninfo
If fs.FileExists(App.Path & "\info.dat") Then
fs.DeleteFile (App.Path & "\info.dat")
End If


'Wenn INI Datei nicht da oder besch�digt
If Not fs.FileExists(App.Path & "\prog.ini") Or Err.Number = 13 Then
    DoEvents
    Form_Load
    Exit Sub
End If

Me.MousePointer = 11

gL�nge = INIGetValue(App.Path & "\Prog.ini", "Ort", "L�nge")
gBreite = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite")
minMag_max = INIGetValue(App.Path & "\Prog.ini", "filter", "MinMag_max")
minMag_min = INIGetValue(App.Path & "\Prog.ini", "filter", "MinMag_min")

grdergebnis.Clear
 
 fehler.ort = "with rsergebnis"
         
     '�ffnen der Tabelle "Grundlage" und l�schen einer evtl. alten Abfrage
    With rsergebnis
    .Open pfad & "\ergebnisse.dat", , , adLockOptimistic

        

            'L�schen einer alten Abfrage
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete (adAffectCurrent)
                    .MoveNext
                Wend
            End If
            .Save pfad & "\ergebnisse.dat"
        
         
    End With
   
   fehler.ort = "With dbsbavsterne"
    
    'Verbindung zur Datenbank herstellen
    With dbsbavsterne
        .Provider = "microsoft.Jet.oledb.4.0"
        
        'Auswahl der Datengrundlage, mit der die Berechnung durchgef�hrt wird
        If cmbGrundlage.ListIndex = 0 Then
        .ConnectionString = pfad & "\Bav_sterne.mdb"
        .Open
        'ElseIf cmbGrundlage.ListIndex = 1 Then
        '.ConnectionString = pfad & "\BAV_sonstige.mdb"
        '.Open
        End If
        
    End With
    
   Database = IIf(cmbGrundlage.ListIndex > 0, cmbGrundlage.ListIndex + 1, cmbGrundlage.ListIndex)
        
    '�ffnen der Tabelle "BVundRR" aus der BAV_Sterne.mdb
    'als tempor�res Recordset oder verbinden mit Recordset
    'aus der Kreiner DB
    
    fehler.ort = "With rstAbfrage"
    
    With rstAbfrage
      If cmbGrundlage.ListIndex = 0 Then
        .ActiveConnection = dbsbavsterne
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
    ElseIf cmbGrundlage.List(cmbGrundlage.ListIndex) = "Einzeln" Then
        Database = 4
        .Open pfad & "\Einzel.dat"
    ElseIf cmbGrundlage.List(cmbGrundlage.ListIndex) = "BAV-BA_EA" Then
        Database = 5
        .Open pfad & "\BAVBA_EA.dat"
    ElseIf cmbGrundlage.List(cmbGrundlage.ListIndex) = "ASAS" Then
        Database = 6
        .Open pfad & "\acvs1.1.dat"
    ElseIf cmbGrundlage.List(cmbGrundlage.ListIndex) = "Eigene" Then
        Database = 7
        .Open pfad & "\Eigene.dat"
    End If
        
      '�ffnen der Ergebnisdatei und springen zu ersten Eintrag
      
    
        .MoveFirst
      DoEvents
      
    Balken.Value = 0
    Balken.Max = rstAbfrage.RecordCount
    maxSternLen = 0
    'Berechnungen f�r alle Sterne in BAV_Sterne.mdb
    numereignis = 0
    
    fehler.ort = "Do While"
    Do While Not .EOF
        fehler.ort = "BeginLoop"
        DoEvents
        lblfertig.Caption = "Berechnungen zu " & Format((Balken.Value / rstAbfrage.RecordCount) * 100, "#") & "% fertiggestellt"
    
        If Not .Fields("epoche").ActualSize = 0 And Not .Fields("periode").ActualSize = 0 _
        And CDbl(aussortieren(!Max)) < minMag_max And CDbl(aussortieren(!mini)) < minMag_min Then
            'Berechnung der Ereignisse im gew�hlten Zeitraum
            BAnfang = JulDat(Left(DTPicker1.Value, 2), Mid(DTPicker1.Value, 4, 2), Mid(DTPicker1.Value, 7, 4))
            bende = JulDat(Left(DTPicker2.Value, 2), Mid(DTPicker2.Value, 4, 2), Mid(DTPicker2.Value, 7, 4))
        
            sPeriode = Split(!periode, " ")
        
            EPeriode = Fix((bende - (!Epoche + 2400000)) / sPeriode(0))
            APeriode = Fix((BAnfang - (!Epoche + 2400000)) / sPeriode(0))
        
            fehler.ort = "For..Next Epochenzahl"
            
            For x = APeriode To EPeriode
            
             If UBound(sPeriode) = 2 Then
                ereignis = (!Epoche + 2400000) + x * sPeriode(0) + (x ^ 2 * (sPeriode(1) * 10 ^ sPeriode(2)))
             ElseIf UBound(sPeriode) = 4 Then
                ereignis = (!Epoche + 2400000) + x * sPeriode(0) + (x ^ 2 * (sPeriode(1) * 10 ^ sPeriode(2))) + (x ^ 3 * (sPeriode(3) * 10 ^ sPeriode(4)))
             Else
                ereignis = (!Epoche + 2400000) + x * sPeriode(0)
             End If
   
             If ereignis >= BAnfang And ereignis <= bende Then
            
                If Not .Fields("hh").ActualSize = 0 Then
                        RA = !hh + !mm / 60 + !ss / 3600
                        DEC = CDbl(!vz & !o + !m / 60)
                    Else: RA = 0
                End If
                 
                 fehler.ort = "HKorr"
                 'Heliozentrische Korrektur der HJD Zeitpunkte!!!
                ereignis = Hkorr(ereignis, RA, DEC, False)
                 
                JDEreignis = JulinDat(ereignis)
                Tag = CDate(Fix(JDEreignis))
                Jahr = Format(JDEreignis, "yyyy")
                SonneAU = AufUnter(Tag, Jahr)
                
                fehler.ort = "AufUnter"
                
                'Sauf = AufUnter(Tag, Jahr, 0) * 24
                If SonneAU(0) = 25 Then
                
                    MsgBox "F�r den gew�hlten Filter f�r den" & vbCrLf & _
                    "Sonnenauf- und -untergang k�nnen keine Werte berechnet werden!" & vbCrLf & vbCrLf _
                    & "    Es ist die Zeit der wei�en oder schwarzen N�chte...!" & Chr(13) & Chr(13) & _
                    "Bitte den Filter f�r die D�mmerung auf 'Sonne am Horizont' stellen.", vbInformation, "Berechnung nicht m�glich"
                    Me.MousePointer = 1
                Exit Sub
                End If
                
                fehler.ort = "Sunter"
                
                Sauf = SonneAU(0) * 24
                'Sunter = AufUnter(Tag, Jahr, 1) * 24
                Sunter = SonneAU(1) * 24
                Uhrzeit = 24 * ((JDEreignis) - Fix(JDEreignis))
    
                If Sunter < Uhrzeit And Uhrzeit <= 24 Or Sauf > Uhrzeit And Uhrzeit >= 0 Then
                    BPrg = !BP
                    If Database <> 1 Then Typ = Trim(!Typ)
                                   
                    Sternzeit = STZT(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit, CDbl(gL�nge))
                    
                    Stundenwinkel = CDbl(stdw(RA, Sternzeit))
                    aktHoehe = Hoehe(Stundenwinkel, CDbl(gBreite), DEC)
                    aktAzimut = Azimut(aktHoehe, Stundenwinkel, CDbl(gBreite), DEC)
                    
                    fehler.ort = "Monddist"
                    
                    'Berechnung der Monddistanz
                    Mpi = 4 * Atn(1)
                    Mdeg = (4 * Atn(1)) / 180
                    Mrad = 180 / (4 * Atn(1))
                    sonne = SunPosition(JulDat(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit))
                    
                    mond = MoonPosition(sonne(2), sonne(3), JulDat(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit))
                    
                    ephem = MoonRise(JulDat(CByte(Format(JDEreignis, "dd")), _
                    CByte(Format(JDEreignis, "mm")), CInt(Format(JDEreignis, "yyyy")), Uhrzeit), 65, CDbl(gL�nge) * Mdeg, CDbl(gBreite) * Mdeg, 0, 1)
                    
                    monddist = Moondistance(RA * 360 / 24, DEC, (mond(0) * Mrad / 15) / 24 * 360, mond(1) * Mrad)

                    fehler.ort = "Hoehe"
                    
                    If aktHoehe > 0 Then 'Filter der H�he �ber Horizont
                        numereignis = numereignis + 1
                       With rsergebnis
                        
                        .AddNew
                        .Fields("Stbld") = rstAbfrage.Fields("Stbld")
                        'If Len(rstAbfrage.Fields("K�rzel")) > 10 Then
                            'SternName = rstAbfrage.Fields("K�rzel")
                        'Else
                            SternName = rstAbfrage.Fields("K�rzel") & " " & rstAbfrage.Fields("Stbld")
                        'End If
                        maxSternLen = IIf(Len(SternName) > maxSternLen, Len(SternName), maxSternLen)
                        .Fields("stern") = SternName
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
                        
                        .Fields("H�he") = Format(aktHoehe, "0") '"#.00")
                        .Fields("Azimut") = Format(aktAzimut, "0") '"#.0")
                        .Fields("Epochenzahl") = x
                        
                        
                        If Database = 0 Then
                          .Fields("BProg") = BPrg
                          .Fields("Typ") = Typ
                        ElseIf Database = 1 Then
                          .Fields("Typ") = BPrg
                         ElseIf Database >= 2 Then
                          .Fields("BProg") = BPrg
                         .Fields("Typ") = Typ
                        End If
                        
                        If BPrg = "KRE" Then
                            .Fields("bc").Value = 2
                        ElseIf BPrg = "GCVS" Then
                            .Fields("bc").Value = 3
                        ElseIf BPrg = "ASAS" Then
                            .Fields("bc").Value = 4
                        ElseIf BPrg = "EIG" Then
                            .Fields("bc").Value = 7
                        End If
                            .Fields("JDEreignis").Value = CDbl(JDEreignis)
                        .Update
     
                        End With
                        
                    End If 'Hoehe
                
                End If 'Auf/Untergang
                
            End If 'Innerhalb Zeitraum
        Next x

        End If ' Helligkeit & Test ob Periode 0

    .MoveNext
    
    fehler.ort = "EndLoop"

    Balken.Value = Balken.Value + 1
    
    Loop
    
Balken.Value = 0
lblfertig.Caption = "Berechnungen beendet!" & vbCrLf & numereignis & " Ereignisse ermittelt"

fehler.ort = "EndeErgebnis"

  If rsergebnis.RecordCount > 0 Then
    '�bernahme der Spalten in Recordset
    rsergebnis.Save pfad & "\ergebnisse.dat"
    cmdFilter.Enabled = True
    cmdErgebnis.Enabled = True: VTabs.TabEnabled(2) = True
    'Schlie�en beider *.mdb's  und
    'des Recordsets
    'rsergebnis.Close
    rstAbfrage.Save pfad & "\info.dat"
    rstAbfrage.Close
    If Database < 2 Then dbsbavsterne.Close
  Else
    MsgBox "Es konnte kein Ereignis berechnet werden." & vbCrLf & "Bitte versuchen sie ein anderes " & vbCrLf & _
    "Berechnungsintervall.", vbInformation, "Kein Ereignis gefunden"
    cmdAbfrag.Enabled = True: cmdErgebnis.Enabled = False: VTabs.TabEnabled(1) = True
  End If

End With
Me.MousePointer = 1

'Zerst�ren der Objekte
'Set rsergebnis = Nothing
Set rstAbfrage = Nothing
Set rsergebnis = Nothing
Set dbsbavsterne = Nothing
Set fs = Nothing
'cmbGrundlage.Enabled = False
fehler.ort = ""

Exit Sub

errhandler:
MsgBox "Fehler in cmdListe_click() :" & vbCrLf & fehler.ort & vbCrLf & Err.Number & vbCrLf & Err.Description

End Sub
Public Sub gridsortieren(cols2sort)
Dim Abf As ADODB.Recordset
Dim i As Integer
Dim abfsource As String, sortstring As String
    
  Set Abf = New ADODB.Recordset
  
    If Not fs.FileExists(pfad & "\ergebnisse.dat") Then Exit Sub
    
    abfsource = IIf(fs.FileExists(pfad & "\filter.dat"), pfad & "\filter.dat", pfad & "\ergebnisse.dat")
    
    If Abf.State = adStateOpen Then Abf.Close
    
    'Recordset laden
    With Abf
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open abfsource, , , adLockOptimistic
    End With
    
    If Abf.RecordCount = 0 Then Exit Sub
    
    sortstring = ""
    If UBound(cols2sort) = 1 Then
        sortstring = IIf(sorter = True, cols2sort(0) & " ASC", cols2sort(0) & " DESC")
    Else
        For i = 0 To UBound(cols2sort) - 2
            sortstring = IIf(sorter = True, sortstring & cols2sort(i) & " ASC,", sortstring & cols2sort(i) & " DESC,")
        Next i
    
        sortstring = IIf(sorter = True, sortstring & cols2sort(i) & " ASC", sortstring & cols2sort(i) & " DESC")
        
    End If
    
    sorter = IIf(sorter = True, False, True)
    
    Abf.Sort = sortstring
    
    Set grdergebnis.DataSource = Abf
    grdergebnis.ColAlignmentFixed = flexAlignCenterCenter
    grdergebnis.ColAlignment = 4
    Abf.Close
    Set Abf = Nothing
    
    If Database = 4 Then Call zellfarbe(grdergebnis.Row)
End Sub


Public Sub gridf�llen(Optional sSQL)
   Dim rsa As ADODB.Recordset
    Dim Abf As ADODB.Recordset
    Dim filt As ADODB.Recordset
    Dim z�hler As Integer
    
    Me.MousePointer = 11
    
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
    For z�hler = 0 To 3
        Abf.Filter = sSQL(z�hler)
        Abf.Save pfad & "\filter.dat"
        Abf.Close
        Abf.Open pfad & "\filter.dat"
    Next z�hler
End If



    If Abf.RecordCount > 0 Then
        Set grdergebnis.DataSource = Abf
        If Database = 4 Then Call zellfarbe(grdergebnis.Row)
    Else:
        MsgBox "F�r diese Filterkombination konnten keine Ereignisse" & vbCrLf & _
        "gefunden werden...", vbInformation, "Keine Ergebnisse f�r diesen Filter"
        Me.MousePointer = 1
        Exit Sub
    End If
    
    grdergebnis.ColWidth(2) = maxSternLen * 105
    grdergebnis.ColWidth(12) = 0
    grdergebnis.ColWidth(13) = 0
    grdergebnis.ColAlignment = 4

    grdergebnis.ColWidth(0) = 200
    
    If coltrigger(0) = 1 Then grdergebnis.ColWidth(1) = 800
    If coltrigger(4) = 1 Then grdergebnis.ColWidth(5) = 1200
    If coltrigger(5) = 1 Then grdergebnis.ColWidth(6) = 600
    If coltrigger(6) = 1 Then grdergebnis.ColWidth(7) = 600

    If Database = 0 Or Database = 5 Then
        If coltrigger(7) = 1 Then grdergebnis.ColWidth(8) = 600
        
        If coltrigger(8) = 1 Then grdergebnis.ColWidth(9) = 1300
                
    End If

    If Database >= 2 And Database <> 5 Then
        grdergebnis.ColWidth(8) = IIf(coltrigger(7) = 0, 0, 600)
        grdergebnis.ColWidth(9) = IIf(coltrigger(8) = 0, 0, 1300)
        
    End If

    If coltrigger(9) = 1 Then grdergebnis.ColWidth(10) = 1500
    If coltrigger(9) = 0 Then grdergebnis.ColWidth(10) = 0

    grdergebnis.ColAlignmentFixed = flexAlignCenterCenter

    'grdergebnis.col = 13
    'grdergebnis.Sort = 5
    grdergebnis.ColAlignment = 4

    'Form Berechnungsfilter anpassen

    'For z�hler = 1 To 11
        
     '     frmHaupt.grdergebnis.ColWidth(z�hler) = coltrigger(z�hler - 1)
    'Next

        'frmHaupt.grdergebnis.ColWidth(12) = 0 'Zellenfarbe..
        
      'Recordset schliessen, Speicher freigeben
      rsa.Close: Set rsa = Nothing
      Abf.Close: Set Abf = Nothing
      Err.Clear
      
      cmdListspeichern.Enabled = True
      cmdGridgross.Enabled = True
      cmdListdrucken.Enabled = True
      grdergebnis.Visible = True
          frmHaupt.cmdStarChart.Visible = True
    Me.MousePointer = 1
    Exit Sub
    
  If fs.FileExists(pfad & "\filter.dat") Then fs.DeleteFile (pfad & "\filter.dat")
  frmGridGross.grossGrid_f�llen

Abbruch:
      MsgBox "Fehler: " & Err.Number & vbCrLf & _
             Err.Description, vbCritical
      Err.Clear
      Me.MousePointer = 1
End Sub


'Liste in gew�nschte Textdatei speichern
Private Sub cmdListspeichern_Click()
Dim datei As String, zeile As String
Dim spalten As Long, zeilen As Long
Dim x As Long, y As Long
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
        For y = 0 To zeilen - 1
            For x = 1 To spalten - 1
                zeile = zeile & grdergebnis.TextMatrix(y, x) & ";"
            Next x
  
            Print #ikanal, zeile
            zeile = ";" 'L�schen des Zeileninhalts
        Next y
        
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


Private Sub grdergebnis_DblClick()

'BAV_Sterne/Kreiner oder BAV_Sonstige?
If Database = 1 Then
 frmBerechnungsfilter.chkSpalte(8).Visible = False
 ElseIf Database = 0 Or Database >= 2 Then
 'frmBerechnungsfilter.chkSpalte(9).Visible = False
End If
 
       grdergebnis.ColWidth(grdergebnis.col) = 0
       frmBerechnungsfilter.chkSpalte(grdergebnis.col) = 0
       coltrigger(grdergebnis.col - 1) = 0
         
End Sub

Private Sub grdergebnis_MouseUp(Button As Integer, _
  Shift As Integer, x As Single, y As Single)
  
  ' Rechtsklick?
  If Button = vbRightButton Then
    Dim nRow As Long
    Dim nCol As Long
    Dim nRowCur As Long
    Dim nColCur As Long
    Dim data As String
    Dim ColCounter As Integer, c2s, i As Integer
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
          y >= .CellTop And y <= .CellTop + .CellHeight Then
        
          ' Beispiel: PopUp-Men� anzeigen
         
          For x = 1 To .Cols - 1
           data = data & .TextMatrix(nRow, x) & " ; "
          Next x
           
        Else
          ' urspr�ngliche Zelle wieder selektieren
          .Row = nRowCur
          .col = nColCur
        End If
      End If
    End With
  End If
  
  If Button = vbLeftButton Then
  
    If grdergebnis.MouseRow = 0 Then
        ColCounter = Abs(grdergebnis.col - grdergebnis.ColSel) + 1
        ReDim c2s(ColCounter)
        For i = 1 To ColCounter
            If i + grdergebnis.col - 1 = 3 Or i + grdergebnis.col - 1 = 4 Then
                c2s(i - 1) = "JDEreignis"
            Else
                c2s(i - 1) = grdergebnis.TextMatrix(0, i + grdergebnis.col - 1)
            End If
        Next i
            Call gridsortieren(c2s)
    End If
  
    If grdergebnis.col = 1 Then
        Call Infof�llen(grdergebnis)
        Call Mondinfo(grdergebnis)
    End If
  End If
  
  If Database = 4 Then Call zellfarbe(grdergebnis.Row)
End Sub

Private Sub cmdListdrucken_Click()
  Call PrintGrid(grdergebnis, 15, 25, 10, 20, _
                 "VarEphem - aktuelle Abfrage" & _
                 "", "")
                 MsgBox "Druckauftrag wurde an den" & Chr(13) & "Standarddrucker gesendet!", vbInformation, "Druckauftrag gesendet"
End Sub

Private Sub UnloadAll()
  Unload frmAladin
  Unload frmBerechnungsfilter
  Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
  Unload frmSternauswahl
  
  
  Me.Form_Load
  cmbGrundlage.Enabled = True
End Sub


'//--[ScrollUp]---------------------------//
'
'  called from the MainModule WndProc sub
'  when a up-scrolling mouse message is
'  received
'
Public Sub ScrollUp()
    ' scroll up..
    If TOPROW > 1 Then
        TOPROW = TOPROW - 1
        frmHaupt.grdergebnis.TOPROW = TOPROW
        frmGridGross.grdGross.TOPROW = TOPROW
    End If
End Sub

'//--[ScrollDown]---------------------------//
'
'  called from the MainModule WndProc sub
'  when a down-scrolling mouse message is
'  received
'
Public Sub ScrollDown()
    ' scroll down..
    If TOPROW < frmHaupt.grdergebnis.Rows - 1 Then
        TOPROW = TOPROW + 1
        frmHaupt.grdergebnis.TOPROW = TOPROW
        frmGridGross.grdGross.TOPROW = TOPROW
    End If
End Sub



Private Sub Text1_Change()
If Text1.text = "" Then Text1.text = 1
If CInt(Text1.text) > 365 Then Text1.text = 365

DTPicker2.Value = DTPicker1.Value + CInt(Text1.text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub





Private Sub Ueber_Click()
MsgBox "VarEphem, Version 1.1.0.6, Stand: 26.09.2015" & vbCrLf & vbCrLf & _
"geschrieben von: J�rg Hanisch, Gescher" & vbCrLf & _
"Alle Rechte vorbehalen" & vbCrLf & vbCrLf & _
"Fragen, Anregungen, (Spenden ;)) bitte per E-Mail an: " & vbCrLf & _
"hanisch.joerg@gmx.de", vbInformation, "Informationen �ber das Programm"
End Sub

Private Sub VTabs_Click(PreviousTab As Integer)
If VTabs.Tab = 1 And VTabs.TabEnabled(1) Then cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
If VTabs.Tab = 2 Then
    VTabs.TabEnabled(3) = True
    gridf�llen
    'cmdErgebnis_Click
    cmdinfo_Click
End If
If VTabs.Tab = 3 Then
    cmdAbfrag.Enabled = True: cmbGrundlage.Enabled = IIf(ListRecherche.ListCount > 0, True, False)
    Dim x As Integer
    For x = 1 To frmHaupt.cmbGrundlage.ListCount
        If frmHaupt.cmbGrundlage.List(x) = "Einzeln" Then
            frmHaupt.cmbGrundlage.ListIndex = x
            Exit For
        End If
        Unload frmBerechnungsfilter
        Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
        Unload frmAladin
    Next
 'VTabs.TabEnabled(1) = IIf(cmdAbfrag.Enabled, True, False)
End If
End Sub


Sub zellfarbe(ByVal reihe)
Dim res, x, y, bcol
For x = 1 To grdergebnis.Rows - 1
    res = grdergebnis.TextMatrix(x, 12)
    Select Case res
        Case 0, 1, 5
        bcol = vbWhite
        Case 2
        bcol = &HC0FFC0
        Case 3
        bcol = &HC0FFFF
        Case 4
        bcol = &HFFFFC0
        Case 7
        bcol = &HE0E0E0
    End Select
    grdergebnis.Row = x
    For y = 1 To 11
    grdergebnis.col = y
    grdergebnis.CellBackColor = bcol
    Next y
Next x
 grdergebnis.Row = reihe: grdergebnis.col = 1
End Sub
 
Private Sub cmdSingleAusw_Click()
Set gew�hlt = New Collection
Dim sSQL As String
Dim Auswahl
Set fs = New FileSystemObject
Dim x As Integer
If rssourcerecord Is Nothing Then
    If fs.FileExists(App.Path & "\recordsets.dat") Then
     Set rssourcerecord = New ADODB.Recordset
     rssourcerecord.Open App.Path & "\recordsets.dat"
    End If
End If

sSQL = ""

For x = 0 To ListRecherche.ListCount - 1

   If ListRecherche.Selected(x) Then
        Auswahl = Split(ListRecherche.List(x), vbTab)
        gew�hlt.Add "(BP = '" & Auswahl(2) & "')"
        'AND Epoche = '" & Auswahl (3) & "' AND Periode = '" & Auswahl(4) & "') "
   End If

Next x

If gew�hlt.Count > 0 Then
    For x = 1 To gew�hlt.Count - 1
        sSQL = sSQL & gew�hlt.Item(x) & " OR "
    Next x
        sSQL = sSQL & gew�hlt.Item(gew�hlt.Count)
        
Else: result = MsgBox("Es ist kein Datensatz ausgew�hlt..." & vbCrLf _
& "Alle angezeigten Elemente werde �bernommen." & vbCrLf & vbCrLf & "Fortfahren ?", vbExclamation + vbYesNo, "keine Auswahl getroffen...")

    If result = vbYes Then
        cmdSingleAusw.Enabled = False
        Else: Exit Sub
    End If

End If

If Not rssourcerecord Is Nothing Then
    rssourcerecord.Filter = sSQL
    If fs.FileExists(App.Path & "\Einzel.dat") Then fs.DeleteFile (App.Path & "\Einzel.dat")
    rssourcerecord.Save (App.Path & "\Einzel.dat")
    rssourcerecord.Close
Else
    MsgBox "Abfragedatei nicht vorhanden oder besch�digt." & vbCrLf & _
        "Bitte Abfrage erneut durchf�hren", vbCritical, "Fehler der Abfragedatei"
End If



 frmHaupt.Form_Load
 frmHaupt.cmdListe.Enabled = True: frmHaupt.VTabs.TabEnabled(1) = True
 frmHaupt.cmbGrundlage.Enabled = True
 
 For x = 1 To frmHaupt.cmbGrundlage.ListCount
 If frmHaupt.cmbGrundlage.List(x) = "Einzeln" Then
    frmHaupt.cmbGrundlage.ListIndex = x
    Exit For
 End If
 Next
 
Unload frmSterninfo: frmHaupt.cmdStarChart.Visible = False
Unload frmAladin
Unload frmBerechnungsfilter

Set gew�hlt = Nothing
Set rssourcerecord = Nothing
Set fs = Nothing

End Sub
Private Sub cmdSingleSuch_Click()
Dim searchstar
If Not txtSingleStar.text = "" Then
  cmdSingleAusw.Enabled = False
  ListRecherche.Clear
  searchstar = Split(txtSingleStar.text, " ")
  FillList searchstar, App.Path
End If
End Sub

Sub FillList(ByRef StarName, ByVal pfad As String)
Dim Listentext As String
Dim x As Integer
 Set dbsbavsterne = New ADODB.Connection
 Set rssourcerecord = New ADODB.Recordset
 Set fs = New FileSystemObject
On Error GoTo errhandler
 If fs.FileExists(pfad & "\recordsets.dat") Then fs.DeleteFile (pfad & "\recordsets.dat")

   'Neues Recordset zur Aufnahme der Rechercheergebnisse erzeugen
   With rssourcerecord
       .Fields.Append ("ID"), adInteger
       .Open
       .Save pfad & "\recordsets.dat"
   End With
 
For x = 0 To 7

Set rssingleabfrage = New ADODB.Recordset

 With rssingleabfrage
 If x = 1 Then x = 2
        If x < 2 Then
        
            'Verbindung zur Datenbank herstellen
            With dbsbavsterne
                .Provider = "microsoft.Jet.oledb.4.0"
                If x = 0 Then
                    .ConnectionString = pfad & "\Bav_sterne.mdb"
                'ElseIf x = 1 Then
                 '   .ConnectionString = pfad & "\BAV_sterne.mdb" 'onstige.mdb"
                End If
                .Open
             End With
             
            .ActiveConnection = dbsbavsterne
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .CursorType = adOpenForwardOnly ' Kleinster Verwaltungsaufwand
            .Open "SELECT * FROM BVundRR Where K�rzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
            
        ElseIf x = 2 Then
       
        If fs.FileExists(pfad & "\Kreiner.dat") Then
          .Open pfad & "\Kreiner.dat"
          .Filter = "K�rzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If

        ElseIf x = 3 Then
        
        If fs.FileExists(pfad & "\GCVS.dat") Then
          .Open pfad & "\GCVS.dat"
          .Filter = "K�rzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If

        ElseIf x = 4 Then
        
        If fs.FileExists(pfad & "\BAVBA_EA.dat") Then
          .Open pfad & "\BAVBA_EA.dat"
          .Filter = "K�rzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        
        ElseIf x = 5 Then
        
        If fs.FileExists(pfad & "\BAVBA_RR.dat") Then
          .Open pfad & "\BAVBA_RR.dat"
          .Filter = "K�rzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        
        ElseIf x = 6 Then
        If fs.FileExists(pfad & "\acvs1.1.dat") Then
        .Open pfad & "\acvs1.1.dat"
        .Filter = "K�rzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        
        'F�r Sp�tere Erweiterungen: Berechnungen aus eigener Datenbank
        ElseIf x = 7 Then
        If fs.FileExists(pfad & "\Eigene.dat") Then
          .Open pfad & "\Eigene.dat"
          .Filter = "K�rzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If

        End If
        
    If rssingleabfrage.State = adStateOpen Then
    
        If rssingleabfrage.RecordCount > 0 Then
        
           If rssourcerecord.State = adStateClosed Then rssourcerecord.Open pfad & "\recordsets.dat"
        
                With rssourcerecord
                 ' Nur, wenn Daten vorhanden
                    If .RecordCount > 0 Then
                        If Not .BOF Then .MoveLast
                         
                        ' Neuer Datensatz erzeugen
                        If Not rssingleabfrage.Fields.Count = 0 Then _
                        .AddNew
      
                        ' Felder aus Clone-Recordset lesen und in
                        ' Original-Recordset speichern

                        On Error Resume Next
                                            
                        For Each feld In rssingleabfrage.Fields
                            If FieldExists(rssourcerecord, feld.Name) Then _
                            rssourcerecord.Fields(feld.Name).Value = feld.Value
                           
                        Next feld
                              
                        ' Datensatz speichern
                        .Update
                    .Save pfad & "\recordsets.dat"
                    End If
                 End With
                
            'Wenn noch keine Daten vorhanden sind, muss die
        'rsSingleAbfrage "gecloned" werden. Allerdings muss der
        'Filter erhalten bleiben...
            If rssourcerecord.RecordCount = 0 Then
                fs.DeleteFile (pfad & "\recordsets.dat")
                rssingleabfrage.Save pfad & "\recordsets.dat"
                rssourcerecord.Close
                rssourcerecord.Open pfad & "\recordsets.dat"
                fs.DeleteFile (pfad & "\recordsets.dat")
                rssourcerecord.Save pfad & "\recordsets.dat"
             End If
             

         End If
                
     End If

    'Schliessen der Abfrage und Beenden der Verbindung zur Access-Datenbank
    If rssingleabfrage.State = adStateOpen Then rssingleabfrage.Close
    If dbsbavsterne.State = adStateOpen And x < 2 Then dbsbavsterne.Close

    End With

Next x
 

  'Anzeigen der Ergebnisse in einer Liste
  With rssourcerecord

    If .RecordCount > 0 Then
        cmdSingleAusw.Enabled = True
        .MoveFirst
        Do While Not .EOF
            Listentext = .Fields("K�rzel").Value & vbTab & .Fields("Stbld").Value & vbTab & _
          .Fields("BP").Value & vbTab & "   " & Format(.Fields("Epoche").Value, "#.0000") & vbTab & _
          Format(.Fields("Periode").Value, "#.00000000")
          
          If Not IsInList(Listentext) Then ListRecherche.AddItem Listentext
          .MoveNext
        Loop
     ElseIf .RecordCount = 0 Then
        MsgBox "Es konnte kein Eintrag in den Datenbanken " & vbCrLf & "gefunden werden. Bitte " & _
        "�ndern Sie die Abfrage.", vbInformation, "Kein Eintrag vorhanden"
        Exit Sub
             
    End If
    
  End With
    
'If fs.FileExists(pfad & "\recordsets.dat") Then fs.DeleteFile (pfad & "\recordsets.dat")

Set fs = Nothing
Set rssingleabfrage = Nothing
Exit Sub

errhandler:

MsgBox Err.Number & " " & Err.Description & vbCrLf & _
     "Form: Haupt, Sub: FillList" & vbCrLf & vbCrLf & _
     "Bitte �berpr�fen Sie die Eingabe.", vbCritical, "Unzul�ssige Eingabe"


End Sub



Private Function IsInList(ByVal Listentext As String) As Boolean
Dim x As Integer
For x = 0 To ListRecherche.ListCount - 1
If Listentext = ListRecherche.List(x) Then
    IsInList = True
    Exit Function
End If
Next x

IsInList = False
End Function

Private Function aussortieren(feld)
    Dim arra
    arra = Array(":", " ", "(")
    For x = 0 To UBound(arra)
       feld = Replace(feld, arra(x), "")
    Next x
    If feld = "" Then feld = 0
    aussortieren = feld
End Function
