VERSION 5.00
Begin VB.Form frmHelligkeit 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Filter für die Helligkeit"
   ClientHeight    =   1260
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtminMag_min 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtminMag_max 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Übernehmen"
      Height          =   375
      Left            =   4200
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "[ mag ]"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "[ mag ]"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   285
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Helligkeit im Minimum:     >"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Helligkeit im Maximum:    >"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmHelligkeit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload frmHelligkeit
End Sub

Private Sub Form_Load()
    txtminMag_max.text = FormatNumber(CDbl(INIGetValue(App.Path & "\Prog.ini", "filter", "MinMag_max")), 1)
    txtminMag_min.text = FormatNumber(CDbl(INIGetValue(App.Path & "\Prog.ini", "filter", "MinMag_min")), 1)
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmHaupt
        .Enabled = True
        .cmbGrundlage.Enabled = True
        .cmdListe.Enabled = True: .VTabs.TabEnabled(1) = True: .VTabs.TabEnabled(3) = True
        .VTabs.TabVisible(0) = False
    End With
End Sub

Private Sub OKButton_Click()
    If IsNumeric(txtminMag_max.text) And IsNumeric(txtminMag_min.text) Then
        minMag_max = CDbl(txtminMag_max.text)
        minMag_min = CDbl(txtminMag_min.text)
        Call INISetValue(datei, "filter", "minMag_Max", txtminMag_max.text)
        Call INISetValue(datei, "filter", "minMag_Min", txtminMag_min.text)
        
     Else: MsgBox "Bitte überprüfen Sie die Eingabe," & vbCrLf _
    & "es sind nur numerische Werte erlaubt.", vbExclamation, "Fehleingabe!"
        Exit Sub
    End If
    Unload Me
End Sub
