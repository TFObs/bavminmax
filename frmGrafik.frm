VERSION 5.00
Begin VB.Form frmGrafik 
   BackColor       =   &H80000000&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "grafische Darstellung helioz. Korrektur"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8400
   Icon            =   "frmGrafik.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5520
      Top             =   0
   End
   Begin VB.PictureBox PBox1 
      BackColor       =   &H80000000&
      Height          =   2535
      Left            =   5400
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.PictureBox PBox 
      BackColor       =   &H80000000&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "RA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   2
         Top             =   2760
         Width           =   270
      End
   End
   Begin VB.Label lblstz 
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmGrafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowPlacement Lib "user32" _
        (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As _
        Long
        
Private Declare Function SetWindowPos Lib "user32" (ByVal _
        hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x _
        As Long, ByVal Y As Long, ByVal cx As Long, ByVal _
        cy As Long, ByVal wFlags As Long) As Long

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


Private Function ycoord(ByVal xcoord As Double)
ycoord = 0.40954 * Sin(0.0172 * (xcoord - 79.34974)) * 180 / (4 * Atn(1))
End Function


Public Sub zeichnen(Optional RA, Optional DEC)
Dim xkoord
With PBox
.BackColor = QBColor(8)
.ForeColor = QBColor(0)

PBox.Scale (-70, 110)-(400, -130)
.AutoRedraw = True

.DrawWidth = 1
.DrawStyle = 0
PBox.Line (0, 90)-(365, 90), QBColor(0)
PBox.Line (365, 90)-(365, -90), QBColor(0)
PBox.Line (365, -90)-(0, -90), QBColor(0)
PBox.Line (0, -90)-(0, 90), QBColor(0)
PBox.Line (102, -80)-(102, -90), QBColor(0) '12
PBox.Line (240, -85)-(240, -90), QBColor(0)
PBox.Line (194, -80)-(194, -90), QBColor(0) '6
PBox.Line (148, -85)-(148, -90), QBColor(0)
PBox.Line (286, -80)-(286, -90), QBColor(0) '0
PBox.Line (56.5, -85)-(56.5, -90), QBColor(0)
PBox.Line (11, -80)-(11, -90), QBColor(0) '18
PBox.Line (331, -85)-(331, -90), QBColor(0)
For x = -90 To 90 Step 30
PBox.Line (-10, x)-(0, x)
Next x
.CurrentX = 91
.CurrentY = -90
PBox.Print "12 h"
.CurrentX = 183
.CurrentY = -90
PBox.Print "6 h"
.CurrentX = 275
.CurrentY = -90
PBox.Print "0 h"
.CurrentX = 0
.CurrentY = -90
PBox.Print "18 h"

.CurrentX = -37
.CurrentY = -82
PBox.Print "-90°"
.CurrentX = -37
.CurrentY = 8
PBox.Print "0°"
.CurrentX = -37
.CurrentY = 98
PBox.Print "90°"
.ForeColor = QBColor(12)

.DrawWidth = 1
.DrawStyle = 1
PBox.Line (0, 0)-(365, 0), QBColor(0)
.DrawWidth = 3

For x = 365 To 0 Step -1
 PBox.PSet (x, ycoord(365 - x))
Next x

.FillStyle = 0
.FillColor = QBColor(14)

Datum = Format(JulinDat(frmHelioz.txtJD.text), "dd,mm.yyyy")
tage = DateDiff("d", "1-1", CStr(Format(Datum, "dd-mm")))

PBox.Circle ((365 - tage), ycoord(tage)), 7, QBColor(14)
.FillColor = QBColor(15)
xkoord = 286 - (RA / 24 * 365)

If xkoord < 0 Then xkoord = 365 + xkoord

PBox.Circle (xkoord, DEC), 4, QBColor(15)
End With

With PBox1
.BackColor = QBColor(0)
PBox1.Scale (-60, 60)-(60, -60)
.AutoRedraw = True
.ForeColor = QBColor(7)
.CurrentX = 45
.CurrentY = 4
PBox1.Print "12 h"
.CurrentX = -3
.CurrentY = 55
PBox1.Print "18 h"
.CurrentX = -55
.CurrentY = 2
PBox1.Print "0 h"
.CurrentX = -2
.CurrentY = -45
PBox1.Print " 6 h"

.FillStyle = 0
.FillColor = QBColor(14)
PBox1.Circle (0, 0), 6, QBColor(14)
.FillStyle = 0
.FillColor = QBColor(9)

tage = DateDiff("d", "21-03", CStr(Format(Datum, "dd-mm")))
If tage < 0 Then tage = 365 + tage
For x = 0 To tage
winkel = (x / 365) * 2 * (4 * Atn(1)) '2 * pi - (((365 - x) / 365) * 2 * pi)
PBox1.PSet (20 * Cos(winkel), 20 * Sin(winkel)), QBColor(9)
Next x
PBox1.Circle (20 * Cos(winkel), 20 * (Sin(winkel))), 3, QBColor(9)
xkoord = ((RA / 24) - 0.5) * 2 * (4 * Atn(1))
.FillColor = QBColor(15)
PBox1.Circle (40 * Cos(xkoord), 40 * Sin(xkoord)), 3, QBColor(15)

End With
With frmHelioz
zeit = CDbl(.txtstunde.text + .txtminute.text / 60 + .txtsekunde.text / 3600)
stez = STZT(.txtTag.text, .txtmonat.text, .txtjahr.text, zeit, 10)
lblstz.Caption = "lokale Sternzeit : " & Format(stez, "hh:mm:ss") _
& " [UT]"
End With

PBox.DrawWidth = 1
PBox.DrawStyle = 1
stez = 286 - (365 * stez)
If stez < 0 Then stez = 365 + stez

PBox.Line (stez, -80)-(stez, 90), QBColor(15)
End Sub


Private Sub Form_Load()
Call zeichnen(RA, DEC)
frmHelioz.Show
End Sub

Private Sub Timer1_Timer()
  Dim WPM As WINDOWPLACEMENT
  Dim Lft&, Tp&, Hgh&, TwpX&, TwpY&
  Static OnTop As Boolean
  
    Timer1.Enabled = False
    TwpX = Screen.TwipsPerPixelX
    TwpY = Screen.TwipsPerPixelY
    
    WPM.Length = Len(WPM)
    If GetWindowPlacement(frmHelioz.hwnd, WPM) = 0 Then Exit Sub
      
    Select Case WPM.showCmd
      Case SW_HIDE:      Me.Visible = False
      
      Case SW_NORMAL:    Me.WindowState = vbNormal
                         If OnTop Then
                           Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, _
                                             0, 0, 0, 0, SWP_NOSIZE Or _
                                             SWP_NOMOVE)
                           OnTop = False
                         End If
                         
                         Lft = (WPM.rcNormalPosition.Left * TwpX) _
                                
                         If Lft < 0 Then
                           Lft = WPM.rcNormalPosition.Left * TwpY
                         End If
                         Tp = WPM.rcNormalPosition.Bottom * TwpY
                         
                          Hgh = (WPM.rcNormalPosition.Bottom - WPM.rcNormalPosition.Top) * TwpY / 1.5 '- 1500 ' - _
                                'WPM.rcNormalPosition.Top) * TwpY
                         Me.Move Lft, Tp, frmHelioz.Width, Hgh
      
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

