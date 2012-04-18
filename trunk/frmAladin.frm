VERSION 5.00
Begin VB.Form frmAladin 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Internet-Recherche"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmAladin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkAladDirekt 
      Caption         =   "für Aladin übernehmen"
      Height          =   255
      Left            =   5280
      TabIndex        =   34
      Top             =   3480
      Value           =   1  'Aktiviert
      Width           =   2175
   End
   Begin VB.CommandButton cmdHinterVor 
      Caption         =   ">>"
      Height          =   375
      Left            =   6840
      TabIndex        =   33
      ToolTipText     =   "Fenster in den Hintergrund"
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtStern 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   32
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame fraInfo 
      Caption         =   "sonst. Informationen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   5040
      TabIndex        =   27
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton cmdNSVS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "NSVS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   36
         ToolTipText     =   "NSVS-Abfrage"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdGCVS 
         BackColor       =   &H00C0C0FF&
         Caption         =   "GCVS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   30
         ToolTipText     =   "Info's des GCVS"
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimbad 
         BackColor       =   &H00C0C0FF&
         Caption         =   "SIMBAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   29
         ToolTipText     =   "SIMBAD-Query"
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdOCSearch 
         BackColor       =   &H00C0C0FF&
         Caption         =   "General Search Gateway"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   28
         ToolTipText     =   "General Search Gateway Abfrage"
         Top             =   1920
         Width           =   1815
      End
   End
   Begin VB.Frame fraOC 
      Caption         =   "B-R und O-C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2640
      TabIndex        =   20
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton cmdCRTS 
         BackColor       =   &H0080C0FF&
         Caption         =   "CSDR1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   35
         ToolTipText     =   "CRTS-Datenbank"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdGEOS 
         BackColor       =   &H0080C0FF&
         Caption         =   "GEOS RR-Lyr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   24
         ToolTipText     =   "Ansicht in GEOS RR-Lyrae DB"
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdKreiner 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kreiner/ Kim/ Nha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   23
         ToolTipText     =   "Kreiner DB"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdBAV 
         BackColor       =   &H0080C0FF&
         Caption         =   "Lichtenknecker Database of the BAV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   22
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton cmdOCGate 
         BackColor       =   &H0080C0FF&
         Caption         =   "OC - Gateway"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   21
         ToolTipText     =   "O-C Diagramme"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "RR Lyrae:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Eclipsing Binaries:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Frame fraKArten 
      Caption         =   "Aufsuchkarten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton cmdDSS 
         BackColor       =   &H00C0C000&
         Caption         =   "DSS Bild"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   17
         ToolTipText     =   "DSS-Bild ansehen"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdAAVSO 
         BackColor       =   &H00C0C000&
         Caption         =   "Karte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   16
         ToolTipText     =   "Karte der AAVSO laden"
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdAAVSO_D 
         BackColor       =   &H00C0C000&
         Caption         =   "Karte/ DSS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Style           =   1  'Grafisch
         TabIndex        =   15
         ToolTipText     =   "Karte der AAVSO mit DSS-Bild laden"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbPicMeas 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Bildgröße:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "AAVSO:"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdOPen 
      BackColor       =   &H000080FF&
      Caption         =   "Aladin Sky Atlas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame fraCats 
      Caption         =   "VizieR Surveys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2400
      TabIndex        =   2
      Top             =   4440
      Width           =   4935
      Begin VB.CheckBox chkESA 
         Caption         =   "Hipparcos + Tycho (ESA '97)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CheckBox chkHog 
         Caption         =   "TYCHO 2 (Hog+ 2000)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox chk2MASS 
         Caption         =   "2 MASS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkGSC 
         Caption         =   "GSC 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox chkUSNOA 
         Caption         =   "USNO A2.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkUSNOB 
         Caption         =   "USNO - B1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraVizier 
      Caption         =   "Haupt-Daten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
      Begin VB.CheckBox chkSIM 
         Caption         =   "SIMBAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkNED 
         Caption         =   "NED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkPOSS 
         Caption         =   "POSS - Bild"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txtObj 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
      TabIndex        =   0
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   120
      X2              =   7440
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label5 
      Caption         =   "Stern:"
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
      Left            =   5400
      TabIndex        =   31
      Top             =   2880
      Width           =   615
   End
End
Attribute VB_Name = "frmAladin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fs As New FileSystemObject
Dim result
Dim Connstr As String
Dim searchstar

Private Sub chkAladDirekt_Click()
If chkAladDirekt.Value = 1 Then txtObj.text = txtStern.text
End Sub

Private Sub cmdHinterVor_Click()

If cmdHinterVor.Caption = ">>" Then
 unfloatwindow Me.hWnd
 cmdHinterVor.Caption = "<<"
 cmdHinterVor.ToolTipText = "Fenster immer im Vordergrund"
 Exit Sub
 Else
 floatwindow Me.hWnd
 cmdHinterVor.Caption = ">>"
 cmdHinterVor.ToolTipText = "Fenster in den Hintergrund"
 End If
 
End Sub





Public Sub Form_Load()
'result = CheckInetConnection(Me.hWnd)
'If result = False Then Exit Sub

cmdDSS.Visible = False

chkPOSS.Value = 1
chkSIM.Value = 1
chkUSNOA.Value = 1
chkESA.Value = 1
chkHog.Value = 1
chk2MASS.Value = 1

    'Füllen der ComboBox für die Bildgröße
    With cmbPicMeas
    .AddItem "15' x 15'"
    .AddItem "30' x 30'"
    .AddItem "45' x 45'"
    .AddItem "60' x 60'"
    End With
    cmbPicMeas.ListIndex = 1

If frmSterninfo.Visible Then
    cmdDSS.Visible = True
    txtStern.text = frmSterninfo.lblStern
End If

floatwindow Me.hWnd
End Sub

Private Function GetAladinURL(ByVal Koord As String)
 Dim KoordZeile, koordzeile1, koordzeile2
 Dim x As Integer
 Dim Plates As Collection, AddPlates As Collection
 Dim GetString As String
 Dim URLString As String

Set Plates = New Collection
Set AddPlates = New Collection
If chkPOSS Then Plates.Add "aladin"
If chkSIM Then Plates.Add "simbad"
If chkNED Then Plates.Add "NED"
If chkUSNOB Then AddPlates.Add "USNOB"
If chkUSNOA Then AddPlates.Add "USNO2"
If chkGSC Then AddPlates.Add "GSC2.2"
If chk2MASS Then AddPlates.Add "2MASS"
If chkESA Then AddPlates.Add "I/239"
If chkHog Then AddPlates.Add "I/259"
On Error GoTo errhandler

 If IsNumeric(Left(Koord, 2)) Then
 If InStr(1, Koord, ":") Then
    KoordZeile = Split(Koord, " ")
    koordzeile1 = Split(KoordZeile(0), ":")
    koordzeile2 = Split(KoordZeile(1), ":")
    ReDim KoordZeile(5)
    For x = 0 To 2
     KoordZeile(x) = koordzeile1(x)
     KoordZeile(x + 3) = koordzeile2(x)
    Next x
 Else
    KoordZeile = Split(Koord, " ")
End If
    
    GetString = KoordZeile(0) & "%20" & KoordZeile(1) & "%20" & KoordZeile(2)

    If Left(KoordZeile(3), 1) = "+" Then
       GetString = GetString & "%20%2b" & KoordZeile(3)
    ElseIf Left(KoordZeile(3), 1) = "-" Then
       GetString = GetString & "%20" & KoordZeile(3)
    End If

    GetString = GetString & "%20" & KoordZeile(4) & "%20" & KoordZeile(5) & "%3b"
 
 Else
 KoordZeile = Split(Koord, " ")
    GetString = KoordZeile(0) & "%20" & KoordZeile(1) & "%3b"
 
 End If

 
 
 'Plates = Array("aladin", "simbad", "NED")
 URLString = "http://aladin.u-strasbg.fr/java/nph-aladin.pl?script=sync%3b"
 'http://aladin.u-strasbg.fr/java/nph-aladin.pl?script=sync%3bGP%20And%3b
If Plates.Count >= 1 Then
 For x = 1 To Plates.Count
    URLString = URLString & "get%20" & Plates(x) & "%20" & GetString
 Next x
 End If
 
 If AddPlates.Count >= 1 Then
 'AddPlates = Array("USNO2", "GSC2.2", "2MASS", "I/239", "I/259")
 For x = 1 To AddPlates.Count
        URLString = URLString & "get%20VizieR%28" & AddPlates(x) & "%29%20" & GetString
 Next x
End If

 URLString = URLString & "sync%3b" & GetString

GetAladinURL = URLString
'Set Plates = Nothing
'#Set AddPlates = Nothing
If Err.Number = 0 Then
Exit Function
Else: GoTo errhandler
End If
Exit Function
errhandler:
unfloatwindow Me.hWnd
MsgBox "Fehler: " & Err.Number & " : " & Err.Description & vbCrLf & "Bitte Eingabe überprüfen!"
GetAladinURL = "-"
floatwindow Me.hWnd
End Function

Private Sub cmdopen_Click()
result = CheckInetConnection(Me.hWnd)

If result = False Then Exit Sub
If txtObj.text = "" Then

    If txtStern.text = "" Then
        Exit Sub
    Else
    txtObj.text = txtStern.text
    End If
    
End If

result = GetAladinURL(txtObj.text)
If Not result = "-" Then URLGoTo Me.hWnd, result
End Sub

Private Sub cmdAAVSO_Click()

  If GetSearchStar <> "-" Then
    Connstr = "http://www.aavso.org/cgi-bin/vsp.pl?action=render&name=" & _
    Trim(searchstar(0)) & "+" & Trim(searchstar(1)) & "&ra=&dec=&charttitle=&chartcomment=&aavsoscale=C&" & _
    "fov=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&resolution=75&maglimit=12&ccdbox=0&north=up&east=left&Submit=Plot+Chart"
    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdAAVSO_D_Click()

  If GetSearchStar <> "-" Then
    Connstr = "http://www.aavso.org/cgi-bin/vsp.pl?action=render&name=" & _
    Trim(searchstar(0)) & "+" & Trim(searchstar(1)) & "&ra=&dec=&charttitle=&chartcomment=&aavsoscale=C&" & _
    "fov=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&resolution=75&maglimit=12&ccdbox=0&north=up&east=left&" & _
    "dss=on&Submit=Plot+Chart"
    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdBAV_Click()

  If GetSearchStar <> "-" Then
    Connstr = "http://www.bavdata-astro.de/~tl/cgi-bin/vs_html?stern=" & _
    Trim(searchstar(0)) & "+" & Trim(searchstar(1)) & _
    "&datum0=&datum1=&perioden=0&bmr=bmr&lkdb=Lichtenknecker+Database&minfoto=nein&" & _
    "filter=&quality=&lang=de&listobs=html&epoche=&periode="
    URLGoTo Me.hWnd, Connstr
 End If

End Sub

Private Sub cmdDSS_Click()
If IsNumeric(Left(txtObj.text, 2)) Then
'http://archive.stsci.edu/cgi-bin/dss_search?v=poss2ukstu_red&r=17+33+59.38&d=-01+04+51.6&e=J2000&h=15.0&w=15.0&f=gif&c=none&fov=NONE&v3=
Connstr = "http://archive.stsci.edu/cgi-bin/dss_search?v=poss2ukstu_red&r=" & frmSterninfo.lblStarRA & "&d=" & frmSterninfo.lblStarDec & "&e=J2000&" _
& "h=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&w=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&f=gif&c=none&fov=NONE&v3="
URLGoTo Me.hWnd, Connstr
Else: cmdDSS.Visible = False
End If
End Sub

Private Sub cmdGCVS_Click()

  If GetSearchStar <> "-" Then
    Connstr = "http://www.sai.msu.su/groups/cluster/gcvs/cgi-bin/search.cgi?search=" & _
    Trim(searchstar(0)) & "+" & Trim(searchstar(1))
    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdGEOS_Click()

  If GetSearchStar <> "-" Then
    'Connstr = "http://dbrr.ast.obs-mip.fr/listostar.html#" & Trim(searchstar(1))
    Connstr = "http://rr-lyr.ast.obs-mip.fr/dbrr/dbrr-V1.0_08.php?" & Trim(searchstar(0)) & " " & UCase(Left(searchstar(1), 1)) & LCase(Mid(searchstar(1), 2, 2))
    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdCRTS_Click()

  If GetSearchStar <> "-" Then
    Connstr = "http://nesssi.cacr.caltech.edu/cgi-bin/getcssconedbid_release.cgi?Name=" & Trim(searchstar(0)) & " " & Trim(searchstar(1)) & _
    "&DB=photcat&Rad=0.5&OUT=csv&SHORT=short&PLOT=plot"

    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdKreiner_Click()
Dim kreistar As String
  If GetSearchStar <> "-" Then
    Connstr = "http://www.as.up.krakow.pl/o-c/data/getdata.php3?" & Trim(searchstar(0)) & "%20" & Trim(searchstar(1))
    'http://www.as.up.krakow.pl/o-c/diagram_html/and_rt_small.html
    'kreistar = Trim(searchstar(1))
    'StrToUpper....
    'kreistar = Left(kreistar, 1) & Chr(Asc(Mid(kreistar, 2, 1)) - 32) & Chr(Asc(Right(kreistar, 1)) - 32)
    'Connstr = "http://www.as.up.krakow.pl/minicalc/" & UCase(searchstar(1)) & Trim(searchstar(0)) & ".HTM"
    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdOCGate_Click()

  If GetSearchStar <> "-" Then
    Connstr = "http://var.astro.cz/ocgate/ocgate.php?star=" & _
    Trim(searchstar(0)) & "+" & Trim(searchstar(1)) & "&submit=Submit&lang=en"
    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdOCSearch_Click()

  If GetSearchStar <> "-" Then
    Connstr = "http://var.astro.cz/gsg/vsgateway.php?star=" & _
    Trim(searchstar(0)) & "+" & Trim(searchstar(1)) & "&all=yes&alldata=yes&oejv=yes&gcvs=yes" & _
    "&nsv=yes&brka=yes&meka=yes&czev=yes&bcvs=yes&dssplate=yes&usecoords=GCVS&rezim=search_now"
    URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdSimbad_Click()

  If GetSearchStar <> "-" Then
   Connstr = "http://simbad.u-strasbg.fr/simbad/sim-id?protocol=html&Ident=" & _
   Trim(searchstar(0)) & "+" & Trim(searchstar(1)) & "&NbIdent=1&Radius=2&Radius.unit=arcmin&submit=submit+id"
   URLGoTo Me.hWnd, Connstr
  End If

End Sub

Private Sub cmdNSVS_Click()
'http://skydot.lanl.gov/nsvs/cone_search.php?&ra=23:11:32.1&dec=+36:53:35.0&rad=0.5&saturated=on&nocorr=on&lonpts=on&hiscat=on&hicorr=on&hisigcorr=on&radecflip=on
Dim radec

If IsNumeric(Left(txtObj.text, 2)) Then
radec = Split(txtObj.text, " ")
Connstr = "http://skydot.lanl.gov/nsvs/cone_search.php?&ra=" & radec(0) & "&dec=" & radec(1) & _
"&rad=1&saturated=on&nocorr=on&lonpts=on&hiscat=on&hicorr=on&hisigcorr=on&radecflip=on"
URLGoTo Me.hWnd, Connstr
Else: cmdNSVS.Visible = False
End If
End Sub

Private Function GetSearchStar() As String
 If frmSterninfo.Visible And txtStern.text = "" Then
   searchstar = Split(frmSterninfo.lblStern.Caption, " ")
Else
   searchstar = Split(txtStern.text, " ")
End If

If UBound(searchstar) <> 1 Then
GetSearchStar = "-"
Else
GetSearchStar = Trim(searchstar(0)) & "+" & Trim(searchstar(1))
End If
End Function


Private Sub URLGoTo(ByVal hWnd As Long, ByVal URL As String)

  ' hWnd: Das Fensterhandle des
  ' aufrufenden Formulars

  Screen.MousePointer = 11
  Call ShellExecute(hWnd, "Open", URL, "", "", 2)
  Screen.MousePointer = 0
End Sub


Private Sub txtObj_Change()
 If Not Len(txtObj.text) > 1 Then Exit Sub
    cmdDSS.Enabled = IIf(IsNumeric(Left(txtObj.text, 2)), True, False)
    cmdNSVS.Enabled = IIf(IsNumeric(Left(txtObj.text, 2)), True, False)
End Sub

Private Sub txtStern_Change()
If chkAladDirekt.Value = 1 Then txtObj.text = txtStern.text
GetSearchStar
End Sub
