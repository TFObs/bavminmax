VERSION 5.00
Begin VB.Form frmSterninfo 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Informationsfenster"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5288.961
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMondEphem 
      Caption         =   "Mondephemeriden"
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   2055
      Begin VB.Label lblDist 
         Caption         =   "-"
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
         Left            =   120
         TabIndex        =   40
         Top             =   1610
         Width           =   1815
      End
      Begin VB.Image imgMoonPhase 
         Height          =   615
         Left            =   1370
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblPhase 
         Caption         =   "-"
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
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblPhaseText 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblUnter 
         Alignment       =   2  'Zentriert
         Caption         =   "--:--"
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblTrans 
         Alignment       =   2  'Zentriert
         Caption         =   "--:--"
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblAuf 
         Alignment       =   2  'Zentriert
         Caption         =   "--:--"
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Untergg:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Transit:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Aufgang:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datenbank-Infos"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
      Begin VB.Label lblLBeo 
         Alignment       =   2  'Zentriert
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fest Einfach
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
         Left            =   600
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "L.B.:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Quelle/BP:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblQuelle 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fest Einfach
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
         Left            =   960
         TabIndex        =   28
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Mo:"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblEpoche 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   600
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblD 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   720
         TabIndex        =   25
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblkd 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   720
         TabIndex        =   24
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label lblMM 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label lblMinI 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblMinII 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   720
         TabIndex        =   21
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblMax 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblPeriode 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblTyp 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fest Einfach
         Height          =   285
         Left            =   600
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblStern 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblKoord 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fest Einfach
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
         Left            =   120
         TabIndex        =   13
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "M-m:"
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "d:"
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "D:"
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Min II:"
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Min I:"
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Max:"
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Periode:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         Caption         =   "Koord [J2000]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         Caption         =   "Lichtkurve"
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
         Left            =   240
         TabIndex        =   4
         Top             =   3900
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         Caption         =   "Helligkeiten"
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
         Left            =   240
         TabIndex        =   3
         Top             =   2565
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Typ:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Stern:"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   3000
   End
   Begin VB.Image imgPhase 
      Height          =   615
      Index           =   7
      Left            =   840
      Picture         =   "frmSterninfo.frx":0000
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image imgPhase 
      Height          =   615
      Index           =   6
      Left            =   240
      Picture         =   "frmSterninfo.frx":0391
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image imgPhase 
      Height          =   615
      Index           =   5
      Left            =   840
      Picture         =   "frmSterninfo.frx":072F
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   615
   End
   Begin VB.Image imgPhase 
      Height          =   615
      Index           =   4
      Left            =   240
      Picture         =   "frmSterninfo.frx":0ABC
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   615
   End
   Begin VB.Image imgPhase 
      Height          =   615
      Index           =   3
      Left            =   840
      Picture         =   "frmSterninfo.frx":0D81
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image imgPhase 
      Height          =   615
      Index           =   2
      Left            =   240
      Picture         =   "frmSterninfo.frx":1151
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image imgPhase 
      Height          =   615
      Index           =   1
      Left            =   840
      Picture         =   "frmSterninfo.frx":14EF
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   615
   End
   Begin VB.Image imgPhase 
      Height          =   600
      Index           =   0
      Left            =   240
      Picture         =   "frmSterninfo.frx":187C
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label lblStarDec 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblStarRA 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmSterninfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowPlacement Lib "user32" _
        (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As _
        Long
        
Private Declare Function SetWindowPos Lib "user32" (ByVal _
        hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x _
        As Long, ByVal y As Long, ByVal cx As Long, ByVal _
        cy As Long, ByVal wFlags As Long) As Long

Private Type POINTAPI
  x As Long
  y As Long
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

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
End Sub

'Andocken an Hauptfenster
Private Sub Timer1_Timer()
  Dim WPM As WINDOWPLACEMENT
  Dim Lft&, Tp&, Hgh&, TwpX&, TwpY&
  Static OnTop As Boolean
  
    Timer1.Enabled = False
    TwpX = Screen.TwipsPerPixelX
    TwpY = Screen.TwipsPerPixelY
    
    WPM.Length = Len(WPM)
    If GetWindowPlacement(frmHaupt.hWnd, WPM) = 0 Then Exit Sub
      
    Select Case WPM.showCmd
      Case SW_HIDE:      Me.Visible = False
      
      Case SW_NORMAL:    Me.WindowState = vbNormal
                         If OnTop Then
                           Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, _
                                             0, 0, 0, 0, SWP_NOSIZE Or _
                                             SWP_NOMOVE)
                           OnTop = False
                         End If
                         
                         Lft = ((WPM.rcNormalPosition.Left * TwpX) - 2430) _
                                
                        ' If Lft < 0 Then
                        '   Lft = WPM.rcNormalPosition.Left * TwpX
                        'End If
                         Tp = WPM.rcNormalPosition.Top * TwpX
                         
                         Hgh = (WPM.rcNormalPosition.Bottom - _
                                WPM.rcNormalPosition.Top) * TwpY - 1500
                         Me.Move Lft, Tp, Me.Width, frmHaupt.Height 'Hgh
      
      Case SW_MINIMIZED: WindowState = vbMinimized
      
      Case SW_MAXIMIZE:
                         If Not OnTop Then
                           OnTop = True
                           Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, _
                                             0, 0, 0, SWP_NOSIZE Or _
                                             SWP_NOMOVE)
                         End If
    End Select
    Timer1.Enabled = True
    
   
End Sub




