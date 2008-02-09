VERSION 5.00
Begin VB.Form frmAladin 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Internet-Recherche"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmAladin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
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
      Height          =   1935
      Left            =   5040
      TabIndex        =   27
      Top             =   120
      Width           =   2295
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
         ToolTipText     =   "SIMBAD-Query"
         Top             =   1320
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
      Height          =   3255
      Left            =   2640
      TabIndex        =   20
      Top             =   120
      Width           =   2295
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
         Top             =   1320
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
         Top             =   1800
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
         Top             =   2520
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
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame fraKArten 
      BackColor       =   &H8000000A&
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
      Caption         =   "Aladin Previewer"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   5760
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
      Top             =   3720
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
      Top             =   3720
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
      Top             =   5760
      Width           =   3135
   End
End
Attribute VB_Name = "frmAladin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim result

Public Sub Form_Load()
result = CheckInetConnection(Me.hwnd)
If result = False Then Exit Sub
chkPOSS.Value = 1
chkSIM.Value = 1
chkUSNOA.Value = 1
chkESA.Value = 1
chkHog.Value = 1
chk2MASS.Value = 1
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
MsgBox Err.Number & Err.Description & vbCrLf & "Bitte Eingabe überprüfen!"
GetAladinURL = "-"
End Function

Private Sub cmdopen_Click()
result = CheckInetConnection(Me.hwnd)

If result = False Then Exit Sub
result = GetAladinURL(txtObj.text)
If Not result = "-" Then frmInternet.URLGoTo Me.hwnd, result
End Sub



