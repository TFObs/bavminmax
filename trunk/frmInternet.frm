VERSION 5.00
Begin VB.Form frmInternet 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "                     Internet"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   2895
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
      Left            =   240
      TabIndex        =   13
      Top             =   6000
      Width           =   2295
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
         TabIndex        =   18
         ToolTipText     =   "SIMBAD-Query"
         Top             =   1320
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
         TabIndex        =   15
         ToolTipText     =   "SIMBAD-Query"
         Top             =   840
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
         TabIndex        =   14
         ToolTipText     =   "Info's des GCVS"
         Top             =   360
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
      Height          =   3135
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
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
         TabIndex        =   17
         ToolTipText     =   "O-C Diagramme"
         Top             =   2520
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
         TabIndex        =   10
         Top             =   1800
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
         TabIndex        =   7
         ToolTipText     =   "Kreiner DB"
         Top             =   1320
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
         TabIndex        =   6
         ToolTipText     =   "Ansicht in GEOS RR-Lyrae DB"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Eclipsing Binaries:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "RR Lyrae:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1215
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      Begin VB.ComboBox cmbPicMeas 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1800
         Width           =   975
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
         TabIndex        =   4
         ToolTipText     =   "Karte der AAVSO mit DSS-Bild laden"
         Top             =   1200
         Width           =   1215
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
         TabIndex        =   3
         ToolTipText     =   "Karte der AAVSO laden"
         Top             =   1200
         Width           =   735
      End
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
         TabIndex        =   2
         ToolTipText     =   "DSS-Bild ansehen"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "AAVSO:"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   960
         Width           =   615
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
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "<<"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Fenster schließen, Filter aufheben"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowPlacement Lib "user32" _
        (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As _
        Long
        
Private Declare Function SetWindowPos Lib "user32" (ByVal _
        hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x _
        As Long, ByVal Y As Long, ByVal cx As Long, ByVal _
        cy As Long, ByVal wFlags As Long) As Long


' Die nachfolgende Prozedur aktiviert den im
' System registrierten Standard-Browser und lädt
' die durch URL angegebene Internetadresse

Private Type POINTAPI
  x As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SW_HIDE = 0
Const SW_NORMAL = 1
Const SW_MINIMIZED = 2
Const SW_MAXIMIZE = 3

Dim fs As New FileSystemObject
Dim result
Dim Connstr As String
Dim SearchStar

Private Sub cmdAAVSO_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
  Connstr = "http://www.aavso.org/cgi-bin/vsp.pl?action=render&name=" & _
  Trim(SearchStar(0)) & "+" & Trim(SearchStar(1)) & "&ra=&dec=&charttitle=&chartcomment=&aavsoscale=C&" & _
  "fov=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&resolution=75&maglimit=12&ccdbox=0&north=up&east=left&Submit=Plot+Chart"
  URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdAAVSO_D_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
  Connstr = "http://www.aavso.org/cgi-bin/vsp.pl?action=render&name=" & _
  Trim(SearchStar(0)) & "+" & Trim(SearchStar(1)) & "&ra=&dec=&charttitle=&chartcomment=&aavsoscale=C&" & _
  "fov=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&resolution=75&maglimit=12&ccdbox=0&north=up&east=left&" & _
  "dss=on&Submit=Plot+Chart"
 URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdBAV_Click()
'SearchStar = frmSterninfo.lblStern.Caption
'Me.MousePointer = 11
'frmLKDB.Show
 SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
 Connstr = "http://www.bavdata-astro.de/~tl/cgi-bin/vs_html?stern=" & _
 Trim(SearchStar(0)) & "+" & Trim(SearchStar(1)) & _
"&datum0=&datum1=&perioden=0&bmr=bmr&lkdb=Lichtenknecker+Database&minfoto=nein&" & _
"filter=&quality=&lang=de&listobs=html&epoche=&periode="
  URLGoTo Me.hwnd, Connstr
 'Me.MousePointer = 1
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdDSS_Click()
'http://archive.stsci.edu/cgi-bin/dss_search?v=poss2ukstu_red&r=17+33+59.38&d=-01+04+51.6&e=J2000&h=15.0&w=15.0&f=gif&c=none&fov=NONE&v3=
Connstr = "http://archive.stsci.edu/cgi-bin/dss_search?v=poss2ukstu_red&r=" & frmSterninfo.lblStarRA & "&d=" & frmSterninfo.lblStarDec & "&e=J2000&" _
& "h=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&w=" & (cmbPicMeas.ListIndex) * 15 + 15 & "&f=gif&c=none&fov=NONE&v3="
URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdGCVS_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
 Connstr = "http://www.sai.msu.su/groups/cluster/gcvs/cgi-bin/search.cgi?search=" & _
 Trim(SearchStar(0)) & "+" & Trim(SearchStar(1))
 URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdGEOS_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
 Connstr = "http://dbrr.ast.obs-mip.fr/listostar.html#" & Trim(SearchStar(1))
 URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdKreiner_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
    Connstr = "http://www.as.wsp.krakow.pl/o-c/data/getdata.php3?" & _
    Trim(SearchStar(0)) & "%20" & Trim(SearchStar(1))
    URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdOCGate_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
    Connstr = "http://var.astro.cz/ocgate/ocgate.php?star=" & _
    Trim(SearchStar(0)) & "+" & Trim(SearchStar(1)) & "&submit=Submit&lang=en"
    URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdOCSearch_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
Connstr = "http://var.astro.cz/gsg/vsgateway.php?star=" & _
Trim(SearchStar(0)) & "+" & Trim(SearchStar(1)) & "&all=yes&alldata=yes&oejv=yes&gcvs=yes" & _
"&nsv=yes&brka=yes&meka=yes&czev=yes&bcvs=yes&dssplate=yes&usecoords=GCVS&rezim=search_now"
URLGoTo Me.hwnd, Connstr
End Sub

Private Sub cmdSimbad_Click()
SearchStar = Split(frmSterninfo.lblStern.Caption, " ")
   Connstr = "http://simbad.u-strasbg.fr/simbad/sim-id?protocol=html&Ident=" & _
   Trim(SearchStar(0)) & "+" & Trim(SearchStar(1)) & "&NbIdent=1&Radius=2&Radius.unit=arcmin&submit=submit+id"
   URLGoTo Me.hwnd, Connstr
End Sub

Private Sub Form_Load()
result = CheckInetConnection(Me.hwnd)

If result = False Then Exit Sub
   
    With cmbPicMeas
    .AddItem "15' x 15'"
    .AddItem "30' x 30'"
    .AddItem "45' x 45'"
    .AddItem "60' x 60'"
    End With
    cmbPicMeas.ListIndex = 1
     
End Sub

'Fenster an Hauptfenster "andocken"
Private Sub Timer1_Timer()
  Dim WPM As WINDOWPLACEMENT
  Dim Lft&, Tp&, Hgh&, TwpX&, TwpY&
  Static OnTop As Boolean
  
    Timer1.Enabled = False
    TwpX = Screen.TwipsPerPixelX
    TwpY = Screen.TwipsPerPixelY
    
    WPM.Length = Len(WPM)
    If GetWindowPlacement(frmHaupt.hwnd, WPM) = 0 Then Exit Sub
      
    Select Case WPM.showCmd
      Case SW_HIDE:      Me.Visible = False
      
      Case SW_NORMAL:    Me.WindowState = vbNormal
                         If OnTop Then
                           Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, _
                                             0, 0, 0, 0, SWP_NOSIZE Or _
                                             SWP_NOMOVE)
                           OnTop = False
                         End If
                         
                         Lft = (WPM.rcNormalPosition.Right * TwpX) _
                                
                         If Lft < 0 Then
                           Lft = WPM.rcNormalPosition.Left * TwpX
                         End If
                         Tp = WPM.rcNormalPosition.Top * TwpX '+ 1500
                         
                         Hgh = (WPM.rcNormalPosition.Bottom - _
                                WPM.rcNormalPosition.Top) * TwpY '- 1500
                         Me.Move Lft, Tp, Me.Width, frmHaupt.Height 'hgh
      
      Case SW_MINIMIZED: WindowState = vbMinimized
      
      Case SW_MAXIMIZE:
                         If Not OnTop Then
                           OnTop = True
                           Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, _
                                             0, 0, 0, SWP_NOSIZE Or _
                                             SWP_NOMOVE)
                         End If
    End Select
    Timer1.Enabled = True
    
   
End Sub

Public Sub URLGoTo(ByVal hwnd As Long, ByVal URL As String)

  ' hWnd: Das Fensterhandle des
  ' aufrufenden Formulars

  Screen.MousePointer = 11
  Call ShellExecute(hwnd, "Open", URL, "", "", 2)
  Screen.MousePointer = 0
End Sub
