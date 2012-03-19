VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmHaupt 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "VarEphem"
   ClientHeight    =   8145
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6855
   Icon            =   "frmHaupt.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8145
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   6907.036
   StartUpPosition =   3  'Windows-Standard
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
      Tab             =   3
      TabHeight       =   520
      TabCaption(0)   =   "Beobachtungsort"
      TabPicture(0)   =   "frmHaupt.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmOrt"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Abfrage"
      TabPicture(1)   =   "frmHaupt.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Rahmen(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Berechnungsergebnisse"
      TabPicture(2)   =   "frmHaupt.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Rahmen(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Einzelextrema"
      TabPicture(3)   =   "frmHaupt.frx":091E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Suche in Datenbanken"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   6315
         Begin VB.ListBox ListRecherche 
            Height          =   1035
            Left            =   240
            MultiSelect     =   1  '1 -Einfach
            TabIndex        =   41
            Top             =   1320
            Width           =   5055
         End
         Begin VB.CommandButton cmdSingleAusw 
            Caption         =   "auswählen"
            Enabled         =   0   'False
            Height          =   615
            Left            =   2040
            TabIndex        =   40
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CommandButton cmdSingleSuch 
            Caption         =   "Suchen"
            Height          =   615
            Left            =   2280
            TabIndex        =   39
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtSingleStar 
            Alignment       =   2  'Zentriert
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Stern     Stbld        Datenbank       Epoche                   Periode"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   5055
         End
      End
      Begin VB.Frame Rahmen 
         Caption         =   "Berechnungszeitraum"
         Height          =   6135
         Index           =   1
         Left            =   -74880
         TabIndex        =   26
         ToolTipText     =   "Internet-Recherche"
         Top             =   380
         Width           =   6315
         Begin VB.TextBox Text1 
            Alignment       =   1  'Rechts
            Height          =   375
            Left            =   2280
            TabIndex        =   46
            Text            =   "1"
            Top             =   5160
            Width           =   360
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   2641
            TabIndex        =   45
            Top             =   5160
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "Text1"
            BuddyDispid     =   196616
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
            Left            =   360
            TabIndex        =   44
            Top             =   5400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51642371
            CurrentDate     =   40979
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   360
            TabIndex        =   43
            Top             =   4800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   51642371
            CurrentDate     =   40979
         End
         Begin VB.ComboBox cmbDauer 
            Height          =   315
            Left            =   4920
            TabIndex        =   31
            ToolTipText     =   "Bitte Anzahl der zu berechnenden Tage auswählen"
            Top             =   1200
            Width           =   735
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
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1095
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
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.Timer Timer1 
            Interval        =   20
            Left            =   2160
            Top             =   3120
         End
         Begin MSComctlLib.ProgressBar Balken 
            Height          =   375
            Left            =   360
            TabIndex        =   27
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
            TabIndex        =   28
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
         Begin MSACAL.Calendar Kalender 
            Height          =   2535
            Left            =   360
            TabIndex        =   32
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
         Begin VB.Label lblEnd 
            Caption         =   "Enddatum :"
            Height          =   255
            Left            =   4560
            TabIndex        =   36
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label lblStart 
            Caption         =   "Startdatum :"
            Height          =   255
            Left            =   4560
            TabIndex        =   35
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblZeitr 
            Caption         =   "Zeitraum [Tage] :"
            Height          =   495
            Left            =   4200
            TabIndex        =   34
            Top             =   1080
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
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   4200
            Width           =   3615
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
            ToolTipText     =   "Filtermöglichkeiten einblenden.."
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
         Left            =   -74880
         TabIndex        =   12
         Top             =   380
         Visible         =   0   'False
         Width           =   6315
         Begin VB.CommandButton cmdOrtOK 
            Height          =   495
            Left            =   4320
            Picture         =   "frmHaupt.frx":093A
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
            Caption         =   "geografische Länge: (Osten = positiv)"
            Height          =   495
            Index           =   1
            Left            =   360
            TabIndex        =   18
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "übernehmen"
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
      Picture         =   "frmHaupt.frx":0C44
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Berechnung starten"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdInternet 
      Height          =   495
      Left            =   5400
      Picture         =   "frmHaupt.frx":0F4E
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Internet-Recherche"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdÖffnen 
      Height          =   495
      Left            =   3360
      Picture         =   "frmHaupt.frx":1258
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "bestehende Abfrage öffnen"
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
      Picture         =   "frmHaupt.frx":1562
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Ergebnisliste in Extrafenster vergrößert darstellen"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdListdrucken 
      Height          =   495
      Left            =   6120
      Picture         =   "frmHaupt.frx":1E2C
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
      Picture         =   "frmHaupt.frx":2136
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
      Picture         =   "frmHaupt.frx":2440
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
      Picture         =   "frmHaupt.frx":274A
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Speichern in eine Textdatei"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmHaupt.frx":2A54
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3360
      Picture         =   "frmHaupt.frx":2D5E
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
   Begin VB.Menu Berech 
      Caption         =   "Berechnungen"
      Begin VB.Menu mnuEinzel 
         Caption         =   "Einzelextrema"
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
            Caption         =   "Bedeckungsveränderliche"
         End
         Begin VB.Menu DB_BAVRR_aktual 
            Caption         =   "kurzper Pulsationssterne"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "Hilfe"
      Begin VB.Menu hilfe 
         Caption         =   "Hilfedatei"
      End
      Begin VB.Menu Ueber 
         Caption         =   "über.."
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


Dim dbsbavsterne As ADODB.Connection 'Grunddaten
Dim rstAbfrage As ADODB.Recordset    'Abfrage auf Grunddaten
Dim rsergebnis As ADODB.Recordset    'temporäre Ergebnisdatei ergebnisse.dat
Dim fs As New FileSystemObject
Dim result
Dim TOPROW As Integer
Dim lngResult As Long

 Dim rssourcerecord  As ADODB.Recordset
 Dim rssingleabfrage  As ADODB.Recordset
 Dim feld As Field
 Dim gewählt As Collection
 



Private Sub Berechnungsfilter_Click()
 frmBerechnungsfilter.show
End Sub

Private Sub Beenden_Click()
 Call Form_Unload(0)
End Sub






Private Sub cmdinfo_Click()

    'Ändern des Icon bei Klick
    If cmdInfo.Picture = Image2 Then
        frmSterninfo.show
        cmdInfo.Picture = Image1
        cmdInfo.ToolTipText = "Informationsfenster ausblenden"
        'cmdInternet.Enabled = True
    Else
        cmdInfo.Picture = Image2
        'cmdInternet.Enabled = False
        cmdInfo.ToolTipText = "Informationsfenster öffnen"
        
        Unload frmSterninfo
    End If
    
    If grdergebnis.col = 1 Then Call Infofüllen(grdergebnis)
    
    
End Sub

Private Sub cmbGrundlage_Click()

 If fs.FileExists(pfad & "\ergebnisse.dat") Then
    cmdErgebnis.Enabled = False: VTabs.TabEnabled(2) = False
    If Not frmHaupt.cmbGrundlage.text = "Einzeln" Then
        VTabs.TabVisible(3) = False
    End If
    cmdListspeichern.Enabled = False
    cmdGridgross.Enabled = False
    cmdListdrucken.Enabled = False
    cmdInfo.Enabled = False
 End If
 
End Sub

Private Sub cmdErgebnis_Click()

    gridfüllen
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
    cmdinfo_Click
    cmdFilter_Click
End Sub

Public Sub cmdAbfrag_Click()

 If cmdListe.Enabled Then Exit Sub
 
    VTabs.TabEnabled(1) = True
    VTabs.TabEnabled(2) = False
    Me.Width = 7050
    VTabs.TabEnabled(0) = False
    
  Unload frmAladin
  Unload frmBerechnungsfilter
  Unload frmSterninfo
  
  lblfertig.Caption = ""
  
  ' Aktivieren wenn sonstige Datei vorhanden!
  cmbGrundlage.Enabled = True
  cmdInfo.Enabled = False
  
  cmdAbfrag.Enabled = False
  cmdInfo.Picture = Image2
  cmdInfo.ToolTipText = "Informationsfenster ausblenden"
  'cmdInternet.Enabled = False
  cmdListe.Enabled = True: VTabs.TabEnabled(1) = True: VTabs.Tab = 1
  
    VTabs.TabVisible(3) = False
 
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
Unload frmAladin
Exit Sub
End If

If frmSterninfo.lblKoord.Caption <> "" Then
result = Split(frmSterninfo.lblKoord.Caption, vbCrLf)
frmAladin.txtObj.text = Trim(Mid(CStr(result(0)), 3, Len(CStr(result(0))) - 2)) & " " & Trim(Mid(CStr(result(1)), 4, Len(CStr(result(1))) - 2))
End If

frmAladin.show
End Sub

'Öffnen einer bestehenden Abfrage
Private Sub cmdÖffnen_Click()
Set dbsbavsterne = New ADODB.Connection
Set rstAbfrage = New ADODB.Recordset
Dim vstrfile, y

Me.VTabs.TabEnabled(2) = True
'Me.Width = 11040
Me.VTabs.TabEnabled(1) = False



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
        
 cmdErgebnis.Enabled = False: VTabs.TabEnabled(2) = False
 
 'infodatei vorhanden?
If fs.FileExists(App.Path & "\info.dat") Then
  fs.DeleteFile (App.Path & "\info.dat")
End If

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
    ElseIf Trim(.TextMatrix(1, 8)) = "EIGEN" Then
        .ColWidth(8) = 0
        .ColWidth(9) = 1300
        Database = 4
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
For y = 1 To grdergebnis.Rows - 1
 
  .AddNew
  .Fields("Stbld") = grdergebnis.TextMatrix(y, 1)
  .Fields("stern") = grdergebnis.TextMatrix(y, 2)
  .Fields("Datum") = grdergebnis.TextMatrix(y, 3)
  .Fields("Uhrzeit") = grdergebnis.TextMatrix(y, 4)
  
  .Fields("stundenwinkel") = grdergebnis.TextMatrix(y, 5)
  .Fields("Höhe") = grdergebnis.TextMatrix(y, 6)
  .Fields("Azimut") = grdergebnis.TextMatrix(y, 7)
  .Fields("BProg") = grdergebnis.TextMatrix(y, 8)
  .Fields("Typ") = grdergebnis.TextMatrix(y, 9)
  .Fields("Epochenzahl") = grdergebnis.TextMatrix(y, 10)
  .Fields("Monddist") = grdergebnis.TextMatrix(y, 11)

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
                ElseIf Database = 1 Then
                    .ConnectionString = pfad & "\BAV_sonstige.mdb"
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
        
'Zerstören der Objekte
Set rstAbfrage = Nothing
Set rsergebnis = Nothing
Set dbsbavsterne = Nothing
Set fs = Nothing

cmbGrundlage.Enabled = False

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
    Me.cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
    VTabs.TabVisible(0) = False
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

Private Sub DBGcvs_aktual_Click()

result = CheckInetConnection(Me.hWnd)
If result = False Then Exit Sub
frmGCVS.show
End Sub



Private Sub DBKrein_aktual_Click()

result = CheckInetConnection(Me.hWnd)
If result = False Then Exit Sub
frmKrein.show
End Sub

Private Sub DTPicker1_Change()
DTPicker2.Value = DTPicker1.Value + cmbDauer.List(cmbDauer.ListIndex)
End Sub

Private Sub DTPicker2_change()
If DTPicker2.Value <= DTPicker1.Value Then DTPicker2.Value = DTPicker1.Value + 1
If DTPicker2.Value - DTPicker1.Value > 365 Then DTPicker2.Value = DTPicker1.Value + 365
Text1.text = DTPicker2.Value - DTPicker1.Value
End Sub



Public Sub Form_Load()
If App.PrevInstance Then End
Set rsergebnis = New ADODB.Recordset
Dim result
Dim daemmfil As String
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

Unload frmSterninfo
Unload frmBerechnungsfilter
Unload frmAladin
Unload frmSingleBerech

'grdergebnis.ToolTipText = "Mit Doppelklick auf den Spaltenkopf wird eine Spalte ausgeblendet," & vbCrLf & _
'"mit einem einfachen Klick wird die Spalte sortiert"
sorter = 7

'If hook = 0 Then MInit Me

TOPROW = 1

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

DTPicker1.CustomFormat = "dd.MM.yyyy"
DTPicker2.CustomFormat = "dd.MM.yyyy"
    ende.Visible = False
    Kalender.Value = Date
    DTPicker1.Value = Date
    
    ende.Value = Kalender.Value + 1

    'Füllen des Zeitraumsfeldes
    For x = 1 To 30
        cmbDauer.AddItem x
    Next x
    cmbDauer.ListIndex = 0  'Zeiger der combobox auf 1. Eintrag
    txtende = ende.Value - 1
DTPicker2.Value = DTPicker1.Value + cmbDauer.List(cmbDauer.ListIndex)
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
     .Fields.Append "bc", adInteger, 1
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

.ListIndex = 0
End With

'Übernahme der Einstellungen
daemmfil = INIGetValue(App.Path & "\Prog.ini", "Auf- Untergang", "Dämmerung")

'Berücksichtigung des Filters für Sonnenaufgang
Select Case daemmfil
    Case Is = "bürgerlich": burger.Checked = True: astro.Checked = False: naut.Checked = False: SaH.Checked = False
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
VTabs.TabVisible(3) = IIf(fs.FileExists(App.Path & "/recordsets.dat"), True, False)
Set rsergebnis = Nothing
cmdAbfrag.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Mende
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim result
Mende
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
    Unload frmAladin
    Unload frmSterninfo
    Unload frmGridGross
    Unload frmHelioz
    Unload frmGrafik
    Unload frmSingleBerech
    Unload frmKrein
    Unload frmGridGross
    Unload frmGCVS
  
    Call HtmlHelp(frmHaupt.hWnd, "", HH_CLOSE_ALL, 0)
    Unload Me

End Sub
Private Sub astro_Click()
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "astronomisch")
naut.Checked = False
astro.Checked = True
burger.Checked = False
SaH.Checked = False
Call UnloadAll
cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
End Sub

Private Sub burger_Click()
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "bürgerlich")
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
'frmSingleBerech.show
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

Private Sub naut_Click()
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "nautisch")
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
Call INISetValue(datei, "Auf- Untergang", "Dämmerung", "S. am Horizont")
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
        Unload frmSterninfo
        Unload frmAladin
    End If
    
    'Register ausblenden
    VTabs.TabEnabled(1) = False
    VTabs.TabEnabled(2) = False
    Me.Width = 7050
    'Ermitteln der Werte aus Registry
    gBreite.text = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite")
    gLänge.text = INIGetValue(App.Path & "\Prog.ini", "Ort", "Länge")

    'Wenn nicht vorhanden, dann Standardwerte
    If gBreite.text = "" Then gBreite.text = Format(CDbl(50#), "#.00")
    If gLänge.text = "" Then gLänge.text = Format(CDbl(10#), "#.00")
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
        cmdErgebnis.Enabled = False: VTabs.TabEnabled(2) = False
End Sub


'Zeitraum wählen
Private Sub cmbDauer_Click()
    kalender_click
    grdergebnis.Clear
    lblfertig.Caption = ""
End Sub


Public Sub cmdListe_click()
Set dbsbavsterne = New ADODB.Connection
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
Dim sPeriode
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
    With dbsbavsterne
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
    End If
        
      'Öffnen der Ergebnisdatei und springen zu ersten Eintrag
      
    
        .MoveFirst
      DoEvents
      
    Balken.Value = 0
    Balken.Max = rstAbfrage.RecordCount
    'Berechnungen für alle Sterne in BAV_Sterne.mdb
    Do While Not .EOF
    DoEvents
        lblfertig.Caption = "Berechnungen zu " & Format((Balken.Value / rstAbfrage.RecordCount) * 100, "#") & "% fertiggestellt"
    
        If Not .Fields("epoche").ActualSize = 0 And Not .Fields("periode").ActualSize = 0 Then
        'Berechnung der Ereignisse im gewählten Zeitraum
        BAnfang = JulDat(Left(Kalender.Value, 2), Mid(Kalender.Value, 4, 2), Mid(Kalender.Value, 7, 4))
        bende = JulDat(ende.Day, ende.Month, ende.year)
        
        sPeriode = Split(!periode, " ")
        
        EPeriode = Fix((bende - (!epoche + 2400000)) / sPeriode(0))
            APeriode = Fix((BAnfang - (!epoche + 2400000)) / sPeriode(0))
        
        
        For x = APeriode To EPeriode
             If UBound(sPeriode) = 1 Then
             ereignis = (!epoche + 2400000) + x * sPeriode(0) + x ^ 2 * sPeriode(1)
             Else
             ereignis = (!epoche + 2400000) + x * sPeriode(0)
             End If
   
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
                        
                        If BPrg = "KRE" Then
                        .Fields("bc").Value = 2
                        ElseIf BPrg = "GCVS" Then
                        .Fields("bc").Value = 3
                        ElseIf BPrg = "ASAS" Then
                        .Fields("bc").Value = 4
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
  If rsergebnis.RecordCount > 0 Then
    'Übernahme der Spalten in Recordset
    rsergebnis.Save pfad & "\ergebnisse.dat"
    cmdFilter.Enabled = True
    cmdErgebnis.Enabled = True
    'Schließen beider *.mdb's  und
    'des Recordsets
    'rsergebnis.Close
    rstAbfrage.Save pfad & "\info.dat"
    rstAbfrage.Close
    If Database < 2 Then dbsbavsterne.Close
    Else
    MsgBox "Es konnte kein Ereignis berechnet werden." & vbCrLf & "Bitte versuchen sie ein anderes " & vbCrLf & _
    "Berechnungsintervall.", vbInformation, "Kein Ereignis gefunden"
    End If

End With
Me.MousePointer = 1

'Zerstören der Objekte
'Set rsergebnis = Nothing
Set rstAbfrage = Nothing
Set rsergebnis = Nothing
Set dbsbavsterne = Nothing
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
If Database = 4 Then Call zellfarbe
Else:
MsgBox "Für diese Filterkombination konnten keine Ereignisse" & vbCrLf & _
"gefunden werden...", vbInformation, "Keine Ergebnisse für diesen Filter"
Exit Sub
End If
grdergebnis.ColWidth(12) = 0

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

'Form Berechnungsfilter anpassen

For zähler = 1 To 11
    If frmHaupt.grdergebnis.ColWidth(zähler) = 0 Then
          frmBerechnungsfilter.chkSpalte(zähler).Value = 0
    End If
Next

        'frmHaupt.grdergebnis.ColWidth(12) = 0 'Zellenfarbe..
        
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
  frmGridGross.grossGrid_füllen
Abbruch:
      MsgBox "Fehler: " & Err.Number & vbCrLf & _
             Err.Description, vbCritical
      Err.Clear
End Sub


'Liste in gewünschte Textdatei speichern
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
            zeile = ";" 'Löschen des Zeileninhalts
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
    If Database = 4 Then Call zellfarbe
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
  Shift As Integer, x As Single, y As Single)
  
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
          y >= .CellTop And y <= .CellTop + .CellHeight Then
        
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
  If Database = 4 Then Call zellfarbe
End Sub

Private Sub cmdListdrucken_Click()
  Call PrintGrid(grdergebnis, 20, 25, 20, 20, _
                 "VarEphem - aktuelle Abfrage" & _
                 "", "")
                 MsgBox "Druckauftrag wurde an den" & Chr(13) & "Standarddrucker gesendet!", vbInformation, "Druckauftrag gesendet"
End Sub

Private Sub UnloadAll()
  Unload frmAladin
  'cmdInternet.Enabled = False
  Unload frmBerechnungsfilter
  Unload frmSterninfo
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
MsgBox "VarEphem, Version 1.0.9, Stand: 05.03.2012" & vbCrLf & vbCrLf & _
"geschrieben von: Jörg Hanisch, Gescher" & vbCrLf & _
"Alle Rechte vorbehalen" & vbCrLf & vbCrLf & _
"Fragen, Anregungen, (Spenden ;)) bitte per E-Mail an: " & vbCrLf & _
"hanisch.joerg@gmx.de", vbInformation, "Informationen über das Programm"
End Sub

Private Sub VTabs_Click(PreviousTab As Integer)
If VTabs.Tab = 1 And VTabs.TabEnabled(1) Then cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
If VTabs.Tab = 3 Then
Dim x As Integer
For x = 1 To frmHaupt.cmbGrundlage.ListCount
 If frmHaupt.cmbGrundlage.List(x) = "Einzeln" Then
    frmHaupt.cmbGrundlage.ListIndex = x
    Exit For
 End If
 Next
End If
End Sub


Sub zellfarbe()
Dim res, x, y, bcol
For x = 1 To grdergebnis.Rows - 1
    res = grdergebnis.TextMatrix(x, 12)
    Select Case res
        Case 0, 1
        bcol = vbWhite
        Case 2
        bcol = &HC0FFC0
        Case 3
        bcol = &HC0FFFF
        Case 4
        bcol = &HFFFFC0
        Case 5
    End Select
    grdergebnis.Row = x
    For y = 1 To 11
    grdergebnis.col = y
    grdergebnis.CellBackColor = bcol
    Next y
Next x
 grdergebnis.Row = 1: grdergebnis.col = 1
End Sub
 
Private Sub cmdSingleAusw_Click()
Set gewählt = New Collection
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
        gewählt.Add "(BP = '" & Auswahl(2) & "')"
        'AND Epoche = '" & Auswahl (3) & "' AND Periode = '" & Auswahl(4) & "') "
   End If

Next x

If gewählt.Count > 0 Then
    For x = 1 To gewählt.Count - 1
        sSQL = sSQL & gewählt.Item(x) & " OR "
    Next x
        sSQL = sSQL & gewählt.Item(gewählt.Count)
        
Else: result = MsgBox("Es ist kein Datensatz ausgewählt..." & vbCrLf _
& "Alle angezeigten Elemente werde übernommen." & vbCrLf & vbCrLf & "Fortfahren ?", vbExclamation + vbYesNo, "keine Auswahl getroffen...")

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
    MsgBox "Abfragedatei nicht vorhanden oder beschädigt." & vbCrLf & _
        "Bitte Abfrage erneut durchführen", vbCritical, "Fehler der Abfragedatei"
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
 
Unload frmSterninfo
Unload frmAladin
Unload frmBerechnungsfilter

Set gewählt = Nothing
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
 
For x = 0 To 6

Set rssingleabfrage = New ADODB.Recordset

 With rssingleabfrage

        If x < 2 Then
        
            'Verbindung zur Datenbank herstellen
            With dbsbavsterne
                .Provider = "microsoft.Jet.oledb.4.0"
                If x = 0 Then
                    .ConnectionString = pfad & "\Bav_sterne.mdb"
                ElseIf x = 1 Then
                    .ConnectionString = pfad & "\BAV_sonstige.mdb"
                End If
                .Open
             End With
             
            .ActiveConnection = dbsbavsterne
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .CursorType = adOpenForwardOnly ' Kleinster Verwaltungsaufwand
            .Open "SELECT * FROM BVundRR Where Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
            
        ElseIf x = 2 Then
       
        If fs.FileExists(pfad & "\Kreiner.dat") Then
          .Open pfad & "\Kreiner.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If

        ElseIf x = 3 Then
        
        If fs.FileExists(pfad & "\GCVS.dat") Then
          .Open pfad & "\GCVS.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If

        ElseIf x = 4 Then
        
        If fs.FileExists(pfad & "\BAVBA_EA.dat") Then
          .Open pfad & "\BAVBA_EA.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        
        ElseIf x = 5 Then
        
        If fs.FileExists(pfad & "\BAVBA_RR.dat") Then
          .Open pfad & "\BAVBA_RR.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        
        ElseIf x = 6 Then
        If fs.FileExists(pfad & "\acvs1.1.dat") Then
        .Open pfad & "\acvs1.1.dat"
        .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        'Für Spätere Erweiterungen: Berechnungen aus eigener Datenbank
        'ElseIf x = 6 Then
        'If fs.FileExists(pfad & "\Einzel.dat") Then
         ' .Open pfad & "\Einzel.dat"
         ' .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        'End If

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
                            .Fields(feld.Name).Value = feld.Value
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
            Listentext = .Fields("Kürzel").Value & vbTab & .Fields("Stbld").Value & vbTab & _
          .Fields("BP").Value & vbTab & "   " & Format(.Fields("Epoche").Value, "#.0000") & vbTab & _
          Format(.Fields("Periode").Value, "#.00000000")
          
          If Not IsInList(Listentext) Then ListRecherche.AddItem Listentext
          .MoveNext
        Loop
     ElseIf .RecordCount = 0 Then
        MsgBox "Es konnte kein Eintrag in den Datenbanken " & vbCrLf & "gefunden werden. Bitte " & _
        "ändern Sie die Abfrage.", vbInformation, "Kein Eintrag vorhanden"
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
     "Bitte überprüfen Sie die Eingabe.", vbCritical, "Unzulässige Eingabe"


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
