VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEigene 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Importieren eigener Dateien"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdLocEigen 
      Caption         =   "..."
      Height          =   495
      Left            =   3600
      TabIndex        =   29
      Top             =   600
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "optional"
      Height          =   1335
      Left            =   240
      TabIndex        =   20
      Top             =   5640
      Width           =   3855
      Begin VB.CheckBox chkopt 
         Caption         =   "vorhanden"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   7
         Left            =   2880
         TabIndex        =   25
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   6
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox chkopt 
         Caption         =   "vorhanden"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Spalte Nr:"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   26
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Typ:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "zuletzt Beob.:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Spaltenzuordnung"
      Height          =   2760
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   3855
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   5
         Left            =   2880
         TabIndex        =   19
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   4
         Left            =   2880
         TabIndex        =   18
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   3
         Left            =   2880
         TabIndex        =   17
         Top             =   1560
         Width           =   615
      End
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   2
         Left            =   2880
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox cmbCol 
         Height          =   315
         Index           =   0
         Left            =   2880
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "DEC(J2000):"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   14
         Top             =   2265
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "RA (J2000):"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   13
         Top             =   1905
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Spalte Nr:"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Periode:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nullepoche:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Sternbild:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Bezeichnung:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmbEigenCreate 
      Caption         =   "Datenbankdatei erstellen"
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dateieigenschaften"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      Begin VB.ComboBox cmbColNum 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cmbTrenner 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Trennzeichen:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Anzahl der Spalten:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog cdlgEigen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEigPath 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   495
      Left            =   240
      TabIndex        =   30
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Eingabedatei:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmEigene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkopt_Click(Index As Integer)
    cmbCol(Index + 6).Enabled = IIf(chkopt(Index).Value, True, False)
    cmbCol(Index + 6).BackColor = IIf(chkopt(Index).Value, &H80000005, &H8000000F)
    cmbCol(Index + 6).ListIndex = 1
End Sub

Private Sub cmbColNum_Click()
Dim x As Integer, y As Integer

For y = 0 To 7
    cmbCol(y).Clear
Next y
 
 For x = 1 To cmbColNum.List(cmbColNum.ListIndex)
 
  For y = 0 To 7
    cmbCol(y).AddItem x
   Next
   
 
 Next x


End Sub

Private Sub cmbEigenCreate_Click()
Dim x As Byte
Me.MousePointer = 11

If checkEingabe = False Then
 MsgBox "Eingaben fehlerhaft, bitte alle Eingaben überfrüfen!...", vbCritical, "Kann nicht fortfahren.."
 Me.MousePointer = 1
Exit Sub
End If

If createEigenDatabase = True Then
    If Not IsInList("Eigene") Then frmHaupt.cmbGrundlage.AddItem "Eigene"
    
    For x = 1 To frmHaupt.cmbGrundlage.ListCount
        If frmHaupt.cmbGrundlage.List(x) = "Eigene" Then
            frmHaupt.cmbGrundlage.ListIndex = x
        Exit For
        End If
    Next x
    frmHaupt.cmdAbfrag_Click
    Unload Me
End If
Me.MousePointer = 1
End Sub

Private Sub cmdLocEigen_Click()
lblEigPath.Caption = ""
With cdlgEigen
.InitDir = App.Path
.Filter = "ASCII-Daten (*.txt,*.csv,*.dat)|*.txt; *.csv; *.dat|Alle Dateien| *.*"
.FileName = ""
.ShowOpen
If .FileName = "" Then Exit Sub
lblEigPath.Caption = .FileName
End With


End Sub

Private Sub Form_Load()
 Dim x As Integer
 For x = 6 To 15
    cmbColNum.AddItem x
 Next x
 
 With cmbTrenner
  .AddItem "<TAB>"
  .AddItem ";"
  .AddItem ","
  .AddItem "|"
  .AddItem ":"
  End With
 
 For x = 6 To 7
  cmbCol(x).Enabled = False
  cmbCol(x).BackColor = &H8000000F
 Next
 cmbColNum.ListIndex = 0
 cmbTrenner.ListIndex = 0
 floatwindow Me.hWnd
End Sub

Private Function checkEingabe()
Dim x As Integer, y As Integer
Dim spaltenwert As Integer

checkEingabe = False
If lblEigPath.Caption = "" Then
    MsgBox "Bitte die Eingabedatei angeben...", vbExclamation, "Kann nicht fortfahren.."
    checkEingabe = False
    Exit Function
End If


On Error GoTo errhandler:

For x = 0 To 7
If cmbCol(x).List(cmbCol(x).ListIndex) <> "" Then
spaltenwert = cmbCol(x).List(cmbCol(x).ListIndex)
    For y = 0 To 7
        If y <> x And cmbCol(y).Enabled Then
            If cmbCol(y).List(cmbCol(y).ListIndex) = spaltenwert Then
                MsgBox "Bitte überprüfen Sie die Spaltenzuordnungen...", vbExclamation, "Kann nicht fortfahren.."
                checkEingabe = False
                Exit Function
            Else
                checkEingabe = True
            End If
        End If
    Next y
End If
Next x
Exit Function
errhandler:
checkEingabe = False

End Function

Function createEigenDatabase()
Dim fso As Object, infile As Object
Dim stern, x%
Dim rsEigen As Object
Dim RAWert As Double, DECWert As Double
Dim Epoche As String
Set rsEigen = New ADODB.Recordset
Set fso = New FileSystemObject


Set infile = fso.OpenTextFile(lblEigPath.Caption)

With rsEigen
     .Fields.Append ("ID"), adInteger
     .Fields.Append ("Kürzel"), adVarChar, 128
     .Fields.Append ("Stbld"), adChar, 3
     .Fields.Append ("BP"), adChar, 3
     .Fields.Append ("LBeob"), adDouble
     .Fields.Append ("Max"), adDouble
     .Fields.Append ("MinI"), adDouble
     .Fields.Append ("Typ"), adVarChar, 12
     .Fields.Append ("Epoche"), adDouble
     .Fields.Append ("Periode"), adChar, 128
     .Fields.Append ("hh"), adInteger, 2
     .Fields.Append ("mm"), adInteger, 2
     .Fields.Append ("ss"), adDouble, 4
     .Fields.Append ("vz"), adChar, 1
     .Fields.Append ("o"), adInteger, 2
     .Fields.Append ("m"), adDouble, 6
End With

rsEigen.Open

x = 1
On Error GoTo errhandler
Dim trenner As String
trenner = cmbTrenner.List(cmbTrenner.ListIndex)

While Not infile.AtEndOfStream
If trenner = "<TAB>" Then
    stern = Split(infile.ReadLine, vbTab)
Else
    stern = Split(infile.ReadLine, trenner)
    If UBound(stern) < 5 Then Err.Raise 9, , , "falsche Spaltenanzahl"
    End If
    If UBound(stern) + 1 >= cmbColNum.List(cmbColNum.ListIndex) Then
        With rsEigen
            .AddNew
            .Fields("ID") = x
            .Fields("Kürzel") = stern(cmbCol(0).List(cmbCol(0).ListIndex) - 1)
            .Fields("Stbld") = stern(cmbCol(1).List(cmbCol(1).ListIndex) - 1)
            .Fields("BP") = "EIG"
        
            Epoche = stern(cmbCol(2).List(cmbCol(2).ListIndex) - 1)
            If Epoche > 2400000 Then Epoche = Epoche - 2400000
            .Fields("Epoche") = Epoche
            
            .Fields("Periode") = stern(cmbCol(3).List(cmbCol(3).ListIndex) - 1)
            
            RAWert = CDbl(stern(cmbCol(4).List(cmbCol(4).ListIndex) - 1))
            DECWert = CDbl(stern(cmbCol(5).List(cmbCol(5).ListIndex) - 1))
            
            .Fields("hh") = CInt(Int(RAWert))
            RAWert = CDbl((RAWert - Int(RAWert)) * 60)
                       
            .Fields("mm") = CInt(RAWert)
             RAWert = CDbl((RAWert - Int(RAWert)) * 60)
            .Fields("ss") = CDbl(FormatNumber(RAWert, 1))
            .Fields("vz") = IIf(DECWert < 0, "-", "+")
            
            DECWert = Abs(DECWert)
            
            .Fields("o") = CInt(Int(DECWert))
            .Fields("m") = CDbl(FormatNumber((DECWert - Int(DECWert)) * 60, 3))
            If chkopt(0).Value = 1 Then
                .Fields("Typ") = stern(cmbCol(7).List(cmbCol(7).ListIndex) - 1)
                Else
                .Fields("Typ") = ""
            End If
            
            .Fields("Max") = 0 'IIf(chk2.Value, CDbl(stern(cmbCol(7).List(cmbCol(7).ListIndex) - 1)), 0)
            .Fields("MinI") = 0 ' IIf(chk3.Value, CDbl(stern(cmbCol(8).List(cmbCol(8).ListIndex) - 1)), 0)
            If chkopt(1).Value = 1 Then
                .Fields("LBeob") = stern(cmbCol(6).List(cmbCol(6).ListIndex) - 1)
                Else
                .Fields("LBeob") = 0
            End If
            .Update
        End With
        x = x + 1
    Else
    MsgBox "nicht erfolgreich bei: " & vbCrLf & _
    stern(cmbCol(0).List(cmbCol(0).ListIndex) - 1) & " " & stern(cmbCol(1).List(cmbCol(1).ListIndex) - 1)
    End If
Wend

If fso.FileExists(App.Path & "\\Eigene.dat") Then fso.DeleteFile (App.Path & "\\Eigene.dat")
rsEigen.Save App.Path & "\Eigene.dat"

MsgBox "Die Eigene Datenbank kann jetzt für" & vbCrLf & "Berechnungen verwendet werden..", vbInformation, "Implementierung erfolgreich"
createEigenDatabase = True
Exit Function

errhandler:
    If UBound(stern) <> 0 Then
        MsgBox " Es ist ein Fehler bei der Umwandlung aufgetreten bei:" & vbCrLf & _
        stern(cmbCol(0).List(cmbCol(0).ListIndex) - 1) & " " & stern(cmbCol(1).List(cmbCol(1).ListIndex) - 1) & _
        " Bitte die Eingabedatei korrigieren..." & vbCrLf & "Fehlernummer: " & Err.Number & _
        vbCrLf & Err.Description, vbCritical, "Fehler: Umwandlung nicht erfolgreich"
    Else
        MsgBox " Es ist ein Fehler bei der Umwandlung aufgetreten..." & vbCrLf & _
        " Bitte die Angaben überprüfen (Trennzeichen?) und" & vbCrLf & "ggf. die Eingabedatei korrigieren..." & vbCrLf & "Fehlernummer: " & Err.Number & _
        vbCrLf & Err.Description, vbCritical, "Fehler: Umwandlung nicht erfolgreich"
    End If
    
createEigenDatabase = False
End Function

Private Function IsInList(ByVal Listentext As String) As Boolean
Dim x As Integer
For x = 0 To frmHaupt.cmbGrundlage.ListCount - 1
If Listentext = frmHaupt.cmbGrundlage.List(x) Then
    IsInList = True
    Exit Function
End If
Next x

IsInList = False
End Function
