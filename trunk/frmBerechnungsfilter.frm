VERSION 5.00
Begin VB.Form frmBerechnungsfilter 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "     Filter bei der Berechnung"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSpalten 
      Caption         =   "Spalten ein- und ausblenden"
      Height          =   3735
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   2535
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   11
         Left            =   1680
         TabIndex        =   34
         Top             =   3360
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   33
         Top             =   3015
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   32
         Top             =   2670
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   31
         Top             =   2670
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   30
         Top             =   2325
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   29
         Top             =   1980
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   28
         Top             =   1620
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   27
         Top             =   1275
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   26
         Top             =   930
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   25
         Top             =   585
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.CheckBox chkSpalte 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   255
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   23
         Top             =   3360
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   22
         Top             =   3015
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   21
         Top             =   2670
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   20
         Top             =   2325
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   19
         Top             =   1980
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   18
         Top             =   1620
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   1275
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   930
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   585
         Width           =   1110
      End
      Begin VB.Label lblSpalte 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label5"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdStandard 
      BackColor       =   &H00C0C000&
      Caption         =   "Standardwerte"
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
      Left            =   720
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Standardwerte laden"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame frmFilter 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2655
      Begin VB.ComboBox cmbStbld 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   35
         ToolTipText     =   "BAV-Beobachtungsprogramm bzw. Typfilter auswählen"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMonddist 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbBpro 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         ToolTipText     =   "BAV-Beobachtungsprogramm bzw. Typfilter auswählen"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtAzi_o 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtAzi_u 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txthoe 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Sternbild :"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Monddistanz : >"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Beob Prog. :"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "Azimut [°] :  >                 <"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         Caption         =   "Höhe [°] :    >"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdaktual 
      BackColor       =   &H000080FF&
      Caption         =   "Filter anwenden"
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
      Left            =   600
      MaskColor       =   &H8000000F&
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Filter anwenden"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "<<"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Fenster schließen, Filter aufheben"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   600
   End
End
Attribute VB_Name = "frmBerechnungsfilter"
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

Private Sub chkSpalte_Click(Index As Integer)
If chkSpalte(Index).Value = 0 Then
    frmHaupt.grdergebnis.ColWidth(Index) = 0
    
ElseIf chkSpalte(Index).Value = 1 Then
    Select Case Index
    Case Is = 1: frmHaupt.grdergebnis.ColWidth(1) = 800
    Case Is = 5: frmHaupt.grdergebnis.ColWidth(5) = 1200
    Case Is = 6: frmHaupt.grdergebnis.ColWidth(6) = 600
    Case Is = 7: frmHaupt.grdergebnis.ColWidth(7) = 600
    Case Is = 8: frmHaupt.grdergebnis.ColWidth(8) = 600
    Case Is = 9: frmHaupt.grdergebnis.ColWidth(9) = 800
    Case Is = 10: frmHaupt.grdergebnis.ColWidth(10) = 1200
    Case Else: frmHaupt.grdergebnis.ColWidth(Index) = 945
End Select

    If Index = 9 And Database >= 2 Then
     frmHaupt.grdergebnis.ColWidth(9) = 1300
    End If
     
End If
End Sub

Private Sub cmdaktual_Click()

'Speichern der Werte in INI_Datei
If IsNumeric(txthoe.text) And IsNumeric(txtAzi_u.text) And _
    IsNumeric(txtAzi_o.text) And IsNumeric(txtMonddist.text) Then
    Call INISetValue(datei, "filter", "höhe", txthoe.text)
    Call INISetValue(datei, "filter", "Azimut_u", txtAzi_u.text)
    Call INISetValue(datei, "filter", "Azimut_o", txtAzi_o.text)
    Call INISetValue(datei, "filter", "Monddist", txtMonddist.text)
    Call INISetValue(datei, "filter", "Sternbild", cmbStbld.text)
    'BAV_Sterne oder BAV_sonstige?
    If Database = 0 Then
        Call INISetValue(datei, "filter", "BProg", cmbBpro.text)
    Else
        Call INISetValue(datei, "filter", "Typ", cmbBpro.text)
    End If

    Else: MsgBox "Bitte überprüfen Sie die Eingabe," & vbCrLf _
    & "es sind nur numerische Werte erlaubt.", vbExclamation, "Fehleingabe!"
        Exit Sub
End If
ReDim sSQL(4)
'Aufstellen des Abfragefilters
If Not cmbStbld.text = "alle" Then
     sSQL(0) = "Stbld = '" & cmbStbld.text & "'"
     Else: sSQL(0) = ""
 End If
 
 If CDbl(txtAzi_u.text) > CDbl(txtAzi_o.text) Then
  sSQL(1) = "(Azimut >= " & txtAzi_u.text & " OR Azimut <= " & txtAzi_o.text & ")"
 Else
  sSQL(1) = "Azimut >= " & txtAzi_u.text & " AND Azimut <= " & txtAzi_o.text
 End If
 
 sSQL(2) = "Höhe >= " & txthoe.text & " AND Monddist >= " & txtMonddist.text
 
 ' sSQL = sSQL & "Höhe >= " & txthoe.text & " AND Azimut BETWEEN " & txtAzi_u.text & _
'" AND  360 AND Azimut BETWEEN 0 AND " & txtAzi_o.text & " AND Monddist >= " & txtMonddist.text
'Else
'sSQL = sSQL & "Höhe >= " & txthoe.text & " AND Azimut >= " & txtAzi_u.text & _
'" AND Azimut <= " & txtAzi_o.text & " AND Monddist >= " & txtMonddist.text
'End If

If Database = 0 Or Database = 1 Then
    'Filter für Typ oder Bprog
    If Not cmbBpro.text = "alle" Then
        If cmbBpro.text = "alle E" Then
            sSQL(3) = "Typ  Like 'E%' "
        ElseIf cmbBpro.text = "alle RR" Then
            sSQL(3) = "Typ  Like 'RR%' "
        ElseIf Database = 0 Then
            sSQL(3) = "BProg = '" & cmbBpro.text & "'"
        ElseIf Database = 1 Then
            sSQL(3) = "Typ = '" & cmbBpro.text & "'"
        End If
    End If
End If

If Database > 1 Then
If Not cmbBpro.text = "alle" Then
    If cmbBpro.text = "alle E" Then
        sSQL(3) = "Typ  Like 'E%' "
    ElseIf cmbBpro.text = "alle RR" Then
        sSQL(3) = "Typ  Like 'RR%' "
    ElseIf cmbBpro.text = "alle CEP" Then
        sSQL(3) = "(Typ  Like '%CEP%' OR Typ LIKE 'BLBOO%')"
    ElseIf cmbBpro.text = "alle DScuti" Then
       sSQL(3) = "(Typ  Like 'DSC%'OR Typ LIKE 'SXPHE%')"
    Else
       sSQL(3) = "Typ LIKE '" & cmbBpro.text & "%'"
    End If
End If
End If
 


frmHaupt.gridfüllen (sSQL)
frmGridGross.grossGrid_füllen
If frmGridGross.Visible Then
Unload frmGridGross
 frmGridGross.Show
 End If
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub



Private Sub cmdStandard_Click()
 'Standardwerte Laden
 cmbStbld.text = INIGetValue(datei, "Standard", "Sternbild")
 txthoe.text = INIGetValue(datei, "Standard", "höhe")
 txtAzi_u.text = INIGetValue(datei, "Standard", "Azimut_u")
 txtAzi_o.text = INIGetValue(datei, "Standard", "Azimut_o")
 txtMonddist.text = INIGetValue(datei, "Standard", "Monddist")
 
 'BAV_Sterne oder BAV_sonstige?
 If Database = 0 Then
    cmbBpro.text = INIGetValue(datei, "Standard", "BProg")
  Else
    cmbBpro.text = INIGetValue(datei, "Standard", "Typ")
 End If


End Sub

Private Sub Form_Load()
Dim stbldwerte
Dim x As Integer

'Ergebnisdatei vorhanden?
If Not fs.FileExists(datei) Then
 fs.CreateTextFile datei
 DefaultWerte
End If

 frmHaupt.Show
 cmbBpro.Clear
 
 'BAV_Sterne oder BAV_sonstige?
 'Füllen der Combobox
 If Database = "" Then
  If frmHaupt.grdergebnis.TextMatrix(1, 8) <> "" Then
  Database = 0
  'Else: Database = 1
  End If
End If

stbldwerte = dicStbld.Keys
cmbStbld.AddItem "alle"
 For x = 0 To UBound(stbldwerte)
    cmbStbld.AddItem stbldwerte(x)
 Next x
 cmbStbld.ListIndex = 0
   
If Database = 0 Then
 With cmbBpro
  .AddItem "20"
  .AddItem "90"
  .AddItem "RR"
  .AddItem "ST"
  .AddItem "DS"
  .AddItem "alle"
  .ListIndex = 0
  End With
  Label4.Caption = "Beob Prog. :"
  lblSpalte(8) = frmHaupt.grdergebnis.ColHeaderCaption(0, 8)
  chkSpalte(9).Visible = False
  
ElseIf Database = 1 Then
  With cmbBpro
  .AddItem "E"
  .AddItem "EA"
  .AddItem "EB"
  .AddItem "EW"
  .AddItem "DS"
  .AddItem "RR"
  .AddItem "RRAB"
  .AddItem "RRC"
  .AddItem "Son"
  .AddItem "alle E"
  .AddItem "alle RR"
  .AddItem "alle"
  .ListIndex = 0
  End With
  Label4.Caption = "      Typ. :"
  lblSpalte(8) = frmHaupt.grdergebnis.ColHeaderCaption(0, 9)
  chkSpalte(8).Visible = False
  
  ElseIf Database >= 2 Then
  With cmbBpro
  .AddItem "EA"
  .AddItem "EB"
  .AddItem "EW"
  .AddItem "RRAB"
  .AddItem "RRC"
  .AddItem "alle CEP"
  .AddItem "alle DScuti"
  .AddItem "alle E"
  .AddItem "alle RR"
  .AddItem "alle"
  .ListIndex = 0
  End With
  Label4.Caption = "      Typ. :"
  lblSpalte(8) = frmHaupt.grdergebnis.ColHeaderCaption(0, 9)
  chkSpalte(8).Visible = False
  chkSpalte(9).Visible = True
End If

If INIGetValue(datei, "Standard", "Sternbild") = "" Then
    Call INISetValue(datei, "Standard", "Sternbild", "alle")
    Call INISetValue(datei, "filter", "Sternbild", "alle")
End If

 'Für den FAll, dass ältere Ini-Version vorhanden ist, Standardwerte ergänzen
 If INIGetValue(datei, "Standard", "Monddist") = "" Then
    Call INISetValue(datei, "Standard", "Typ", "alle")
    Call INISetValue(datei, "Standard", "Monddist", 30)
    Call INISetValue(datei, "filter", "Typ", "alle")
    Call INISetValue(datei, "filter", "Monddist", 30)
 End If

 txthoe.text = INIGetValue(datei, "filter", "höhe")
 txtAzi_u.text = INIGetValue(datei, "filter", "Azimut_u")
 txtAzi_o.text = INIGetValue(datei, "filter", "Azimut_o")
 txtMonddist.text = INIGetValue(datei, "filter", "Monddist")
 cmbStbld.text = INIGetValue(datei, "filter", "Sternbild")
 
 'BAV_Sterne oder BAV_sonstige?
 If Database = 0 Then
    cmbBpro.text = INIGetValue(datei, "filter", "BProg")
  Else 'If Database = 1 Then
    cmbBpro.text = INIGetValue(datei, "filter", "Typ")
 End If
 
 
 cmdStandard.Enabled = True
 
 'Beschädigte INI-Datei?
 If txthoe.text = "" Then
  cmdStandard.Enabled = False
  MsgBox "Die Konfigurationsdatei ist beschädigt," & vbCrLf _
  & "es werden Standardwerte geladen...", vbCritical, "beschädigte Konfigurationsdatei"
  DefaultWerte
  Call Form_Load
 End If
 
'Spaltenüberschriften, auch bei Laden einer Datei festlegen
    For x = 1 To 7
        lblSpalte(x) = frmHaupt.grdergebnis.ColHeaderCaption(0, x)
    Next x
    
'Spalten Epochenzahl und Monddist, Spalten 8 und 9 schon bei
'Füllen der Combobox erledigt
lblSpalte(10) = frmHaupt.grdergebnis.ColHeaderCaption(0, 10)
lblSpalte(11) = frmHaupt.grdergebnis.ColHeaderCaption(0, 11)

End Sub



Private Sub Form_Unload(Cancel As Integer)
If frmAladin.Visible = False Then
sSQL = ""
frmHaupt.gridfüllen
End If

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




