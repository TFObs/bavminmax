VERSION 5.00
Begin VB.Form frmHelioz 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Berechnung der heliozentrischen Korrektur"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   ForeColor       =   &H8000000C&
   Icon            =   "frmHelioz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame3 
      Caption         =   "heliozentrische Korrektur"
      Enabled         =   0   'False
      Height          =   2535
      Left            =   3480
      TabIndex        =   21
      Top             =   2760
      Width           =   4455
      Begin VB.CommandButton cmdClipboard 
         BackColor       =   &H80000007&
         Height          =   615
         Left            =   3360
         Picture         =   "frmHelioz.frx":08CA
         Style           =   1  'Grafisch
         TabIndex        =   43
         ToolTipText     =   "Speichern in Zwischenablage"
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdGraph 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   3360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmHelioz.frx":0DBB
         Style           =   1  'Grafisch
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdHcorr 
         Caption         =   "Korrektur berechnen"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000080FF&
         Caption         =   "3"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   38
         Top             =   650
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   375
         Index           =   2
         Left            =   1680
         Shape           =   3  'Kreis
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Zentriert
         Caption         =   "Grafik"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label label11 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Korrektur :"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Julianisches Datum [helioz.]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   600
         TabIndex        =   23
         Top             =   1800
         Width           =   1740
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stern auswählen"
      Height          =   2535
      Left            =   3480
      TabIndex        =   17
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Opt_Liste 
         Caption         =   "Liste"
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Opt_Eingabe 
         Caption         =   "Eingabe"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtStern 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdSuch 
         Caption         =   "auswählen"
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Auswahl per:"
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblAirmass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Koordinaten (aktuell)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   27
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "_ _ _"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   26
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Koordinaten (J2000)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   20
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2400
         TabIndex        =   19
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "JD - Kalkulator"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtJD 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         TabIndex        =   7
         Top             =   4440
         Width           =   1740
      End
      Begin VB.CommandButton cmdJDTi 
         Height          =   735
         Left            =   480
         Picture         =   "frmHelioz.frx":1685
         Style           =   1  'Grafisch
         TabIndex        =   10
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdTiJD 
         Height          =   735
         Left            =   1440
         Picture         =   "frmHelioz.frx":1927
         Style           =   1  'Grafisch
         TabIndex        =   9
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtstunde 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         TabIndex        =   4
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtminute 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtsekunde 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   6
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtTag 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtmonat 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtjahr 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         Caption         =   "Zeit bitte in UT eingeben"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Zentriert
         Caption         =   "JD in Zeitpunkt umrechnen"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Zentriert
         Caption         =   "Zeitpunkt in JD umrechnen"
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   33
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Julianisches Datum [geoz.]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         Caption         =   "TT"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         Caption         =   "MM"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         Caption         =   "JJJJ"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         Caption         =   "ss"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         Caption         =   "mm"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         Caption         =   "hh"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000080FF&
         Caption         =   "1"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   37
         Top             =   4440
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   375
         Index           =   0
         Left            =   2280
         Shape           =   3  'Kreis
         Top             =   4380
         Width           =   495
      End
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Zentriert
      BackColor       =   &H000080FF&
      Caption         =   "2"
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
      Index           =   1
      Left            =   3120
      TabIndex        =   42
      Top             =   540
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   3240
      X2              =   3360
      Y1              =   2400
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   3240
      X2              =   3120
      Y1              =   2400
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Index           =   1
      X1              =   3000
      X2              =   3240
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Index           =   0
      X1              =   3240
      X2              =   3480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   3240
      X2              =   3240
      Y1              =   4800
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   375
      Index           =   1
      Left            =   3000
      Shape           =   3  'Kreis
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmHelioz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim result
Dim ClipboardText As String

Private Sub cmdClipboard_Click()
Clipboard.Clear
Clipboard.SetText txtStern.text & vbTab & txtTag.text & "." & txtmonat.text & "." & txtjahr.text & vbTab & _
txtstunde.text & ":" & txtminute.text & ":" & txtsekunde.text & vbTab & txtJD.text & vbTab & _
Trim(Left(Label10.Caption, InStr(1, Label10.Caption, "d") - 1)) & vbTab & Label2.Caption
MsgBox "Ergebnisse in Zwischenablage" & vbCrLf & "gespeichert.", vbInformation, "Ergebnisse gespeichert"
End Sub

Private Sub cmdGraph_Click()
If Not frmGrafik.Visible Then
frmGrafik.show
Call frmGrafik.zeichnen(RA, DEC)
Else: Unload frmGrafik
End If
End Sub

Private Sub cmdHcorr_Click()
Dim korrektur, gBreite, gLänge
Dim nRA, nDEC, stez, zeit
Dim result
Call cmdJDTi_Click
korrektur = Hkorr(CDbl(txtJD.text), RA, DEC, True)
Label2.Caption = Format(korrektur, "#.00000")
Label10.Caption = Format((korrektur - txtJD.text), "#.000000") & " d" & vbCrLf ' * 3600 * 24, "#.0") & " s" & vbCrLf
label11.Caption = Format((korrektur - txtJD.text) / 0.0001, "#.0")
gBreite = INIGetValue(App.Path & "\Prog.ini", "Ort", "Breite")
gLänge = INIGetValue(App.Path & "\Prog.ini", "Ort", "Länge")

nRA = umwand(1) + umwand(2) / 60 + umwand(3) / 3600
nDEC = umwand(4) + umwand(5) / 60 + umwand(6) / 3600
zeit = CDbl(txtstunde.text + txtminute.text / 60 + txtsekunde.text / 3600)
stez = STZT(txtTag.text, txtmonat.text, txtjahr.text, zeit, gLänge)
lblAirmass.Caption = "Luftmasse: " & Format(airmass(gBreite, nDEC, stdw(nRA, stez)), "#.00")
cmdGraph.Visible = True
cmdClipboard.Visible = True
Label13.Visible = True
Call frmGrafik.zeichnen(nRA, nDEC)
End Sub


Private Sub cmdJDTi_Click()
Dim wert
Dim zeit As Date, Datum As Date
If Not IsNumeric(txtJD.text) Then
txtJD.text = ""
Exit Sub
End If
wert = JulinDat(txtJD.text)
Datum = Format(CDate(Fix(wert)), "dd.mm.yyyy")
zeit = Format(wert - Fix(wert), "hh:mm:ss")

txtTag.text = Format(Datum, "dd")
txtmonat.text = Format(Datum, "mm")
txtjahr.text = Format(Datum, "yyyy")
txtstunde.text = Mid(zeit, 1, 2)
txtminute.text = Mid(zeit, 4, 2)
txtsekunde = Mid(zeit, 7, 2)

End Sub

Public Sub cmdSuch_Click()
If Opt_Eingabe = True Then
    Dim ikanal As Integer
    Close #ikanal
    Dim datei$, zeile$
    Dim Koord, erg, stern, vz
    Label3.Caption = "_ _ _"
    Label10.Caption = ""
    Dim Werte As Collection
    Set Werte = New Collection
    ikanal = FreeFile
    datei = App.Path & "\sternkoord.txt"
    Open datei For Input As ikanal
    Do Until EOF(ikanal)
    Input #ikanal, zeile
    stern = Left(zeile, InStr(1, zeile, ";") - 1)
    Koord = Right(zeile, Len(zeile) - Len(stern) - 1)
    Werte.Add Koord, stern
    Loop
    Close #ikanal

    On Error GoTo ndatbank
    Koord = Werte.Item(txtStern.text)
    RA = Left(Koord, InStr(1, Koord, ";") - 1)
    DEC = Right(Koord, Len(Koord) - Len(RA) - 1)

    erg = ausg(RA, DEC)
    Label1(0).Caption = "RA       " & umwand(1) & ":" & umwand(2) & ":" & umwand(3) & vbCrLf & _
    "DEC   " & umwand(4) & ":" & umwand(5) & ":" & umwand(6)
    Frame3.Enabled = True
    result = Hkorr(CDbl(txtJD.text), RA, DEC, True)
    
ElseIf Opt_Liste = True Then frmSternauswahl.show
End If
Exit Sub

ndatbank:
    If txtStern.text = "" Or Err.Number = 5 Then
    MsgBox "Stern nicht gefunden..", vbExclamation, "Stern nicht in der Datenbank"
    txtStern.text = ""
    End If
End Sub

Private Sub cmdTiJD_Click()
Dim JD
Dim zeit
If Not IsNumeric(txtTag.text) Or Not IsNumeric(txtmonat.text) Or Not IsNumeric(txtjahr.text) _
Or Not IsNumeric(txtstunde.text) Or Not IsNumeric(txtminute.text) Or Not IsNumeric(txtsekunde.text) Then
txtTag.text = Format(Date, "dd")
txtmonat.text = Format(Date, "mm")
txtjahr.text = Format(Date, "jjjj")
txtstunde.text = Format(Time, "hh")
txtminute.text = Format(Time, "mm")
txtsekunde.text = Format(Time, "ss")
End If
zeit = txtstunde.text + txtminute.text / 60 + txtsekunde.text / 3600
JD = JulDat(txtTag.text, txtmonat.text, txtjahr.text, zeit)
txtJD.text = Format(JD, "#.00000")


End Sub

Private Sub Form_Load()
txtJD.text = Format(JulDat(Format(Date, "dd"), Format(Date, "mm"), Format(Date, "yyyy"), CDbl(Time * 24)), "#.00000")
txtTag.text = Format(Date, "dd")
txtmonat.text = Format(Date, "MM")
txtjahr.text = Format(Date, "yyyy")
txtstunde.text = Format(Time, "hh")
txtminute.text = Mid(Time, 4, 2)
txtsekunde = Format(Time, "ss")
Opt_Eingabe = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload frmGrafik
Unload frmSternauswahl
Unload Me
End Sub





Private Sub txtJahr_LostFocus()
If Not IsNumeric(txtjahr.text) Then
txtjahr.text = Format(Date, "yyyy")
End If
End Sub





Private Sub txtJD_Change()
Label3.Caption = "_ _ _"
Label2.Caption = ""
End Sub

Private Sub txtminute_LostFocus()
If Not IsNumeric(txtminute.text) Then
txtminute.text = Format(Time, "mm")
ElseIf CInt(txtminute.text) > 60 Then txtminute.text = 0
ElseIf CInt(txtminute.text) < 1 Then txtminute.text = 0
End If

End Sub

Private Sub txtmonat_GotFocus()
txtmonat.SelStart = 0
End Sub

Private Sub txtStern_Change()
Label3.Caption = "_ _ _"
Label2.Caption = ""
Frame3.Enabled = False
Unload frmGrafik
End Sub

Private Sub txtTag_GotFocus()
txtTag.SelStart = 0
End Sub

Private Sub txtJahr_GotFocus()
txtjahr.SelStart = 0
End Sub
Private Sub txtstunde_GotFocus()
txtstunde.SelStart = 0
End Sub
Private Sub txtminute_GotFocus()
txtminute.SelStart = 0
End Sub
Private Sub txtsekunde_GotFocus()
txtsekunde.SelStart = 0
End Sub


Private Sub txtMonat_LostFocus()
If Not IsNumeric(txtmonat.text) Then
txtmonat.text = Format(Date, "mm")
ElseIf CInt(txtmonat.text) > 12 Then
txtmonat.text = 12
ElseIf CInt(txtmonat.text) < 1 Then
txtmonat.text = 1
End If
End Sub



Private Sub txtsekunde_lostfocus()
If Not IsNumeric(txtsekunde.text) Then
txtsekunde.text = Format(Time, "ss")
ElseIf CInt(txtsekunde.text) > 60 Then txtsekunde.text = 0
ElseIf CInt(txtsekunde.text) < 1 Then txtsekunde.text = 0
End If

End Sub

Private Sub txtstunde_lostfocus()
If Not IsNumeric(txtstunde.text) Then
txtstunde.text = Format(Time, "hh")
ElseIf CInt(txtstunde.text) > 24 Then txtstunde.text = 0
ElseIf CInt(txtstunde.text) < 1 Then txtstunde.text = 0
End If

End Sub


Private Sub txtTag_LostFocus()
If Not IsNumeric(txtTag.text) Then
txtTag.text = Format(Date, "dd")
ElseIf CInt(txtTag.text) > 31 Then txtTag.text = 31
ElseIf CInt(txtTag.text) < 1 Then txtTag.text = 1
End If
End Sub
