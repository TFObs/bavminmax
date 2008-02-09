VERSION 5.00
Begin VB.Form frmSternauswahl 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Auswahl des Sterns"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "frmSternauswahl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox lstStern 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox cmbStbld 
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblAuswahl 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "ausgewählt:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "Stern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Sternbild"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSternauswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As FileSystemObject
Dim StarDic As Dictionary
Dim RecStar As ADODB.Recordset
Dim ConstDic As Dictionary
Dim einstrom As TextStream
Dim zeile() As String
Dim col As Collection
Dim x As Integer

Dim werte

Private Sub cmbStbld_Click()
RecStar.Filter = ""
lstStern.Clear
RecStar.Filter = "Stbld = '" & cmbStbld.List(cmbStbld.ListIndex) & "'"
 While Not RecStar.EOF
lstStern.AddItem RecStar.Fields(1).Value
RecStar.MoveNext
Wend

End Sub

Private Sub cmdOK_Click()
If lblAuswahl.Caption = "" Then Exit Sub

frmHelioz.Opt_Eingabe = True
frmHelioz.txtStern.text = CStr(lblAuswahl.Caption)
Call frmHelioz.cmdSuch_Click
Unload Me
End Sub

Private Sub Form_Load()
Set fso = New FileSystemObject
Set StarDic = New Dictionary
Set RecStar = New ADODB.Recordset

With RecStar
 .Fields.Append ("Stbld"), adVarChar, 3
 .Fields.Append ("Stern"), adVarChar, 15
 End With
 
 RecStar.Open
 
Set einstrom = fso.OpenTextFile(App.Path & "\sternkoord.txt")
While Not einstrom.AtEndOfStream
 zeile = Split(einstrom.ReadLine, ";")
 With RecStar
   .AddNew
    .Fields("stbld") = Right(zeile(0), 3)
    If Not StarDic.Exists(Right(zeile(0), 3)) Then StarDic.Add (Right(zeile(0), 3)), " "
    zeile(0) = Left(zeile(0), InStr(1, zeile(0), " "))
    .Fields("Stern") = Trim(zeile(0))
    .Update
  End With
  Wend
  
  werte = StarDic.Keys
  For x = LBound(werte) To UBound(werte)
  cmbStbld.AddItem werte(x)
  Next x
 cmbStbld.ListIndex = 0
 
End Sub

Private Sub lstStern_Click()
lblAuswahl.Caption = lstStern.List(lstStern.ListIndex) & " " & cmbStbld.List(cmbStbld.ListIndex)
End Sub
