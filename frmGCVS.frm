VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGCVS 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Aktualisierung der GCVS-DB"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtFilter 
      Alignment       =   2  'Zentriert
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   1800
      TabIndex        =   9
      Top             =   3360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
      Begin VB.OptionButton optD_save 
         Caption         =   "vorhandene Datenbank sichern"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton OptD_overwrite 
         Caption         =   "Datenbank überschreiben"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdDownload 
      BackColor       =   &H8000000A&
      Caption         =   "Start"
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
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdLocGCVS 
      Caption         =   "...."
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frmOptions 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton optD_umwand 
         Caption         =   "vorhandene Datei Umwandeln"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton optD_Down 
         Caption         =   "aktuelle Datei downloaden"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog cdlgGCVS 
      Left            =   3120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "[ ° ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "nur Sterne mit Deklination größer"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Deklinationsfilter:"
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
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblSpeichPfad 
      Caption         =   "Bitte Pfad zu der Datei iii.dat angeben:"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmGCVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fs As FileSystemObject
Dim rsa As ADODB.Recordset
Dim einstrom As TextStream
Dim ausstrom As TextStream
Dim zeile
Dim stern(2)
Dim zähler As Integer
Dim GCVSThere As Boolean
Dim result
Dim DBFilter As String


Private Sub cmdDownload_Click()
If optD_umwand = True Then
    cmdLocGCVS.Visible = True
    lblSpeichPfad.Visible = True
    cmdDownload.Visible = False
    Exit Sub
End If
Call DownAndChange
End Sub

Private Sub cmdLocGCVS_Click()
With cdlgGCVS
.InitDir = App.Path
.Filter = "GCVS DB (iii.dat)|iii.dat"
.FileName = ""
.ShowOpen
If .FileName = "" Then Exit Sub
Call DownAndChange
End With

End Sub

Private Sub Form_Load()

Set fs = New FileSystemObject
If fs.FileExists(App.Path & "\gcvs.dat") = False Then
 GCVSThere = False
 optD_umwand.Enabled = True
 optD_save.Enabled = False
 OptD_overwrite.Enabled = False
 End If
 
GCVSThere = True
optD_Down = True
optD_save = True
txtFilter.text = "-20"

End Sub

Sub DownAndChange()

Set fs = New FileSystemObject
Set rsa = New ADODB.Recordset
On Error GoTo ErrorHandler
Me.MousePointer = 11
DoEvents

If optD_Down = True Then

    cmdDownload.BackColor = &HC0FFFF
    cmdDownload.Caption = "Download in Arbeit..."
    
    result = CBool(InternetGetConnectedState(0, 0))
    
    If result = False Then
    'Testen oder Herstellen der Internetverbindung
    result = RASConnect(Me.hWnd)
    End If
    
    If result = False Then
      MsgBox "Internetverbindung konnte nicht aufgebaut werden," & _
      vbCrLf & "bitte überprüfen Sie Ihre Einstellungen", vbCritical, "Keine Verbindung..."
      Me.MousePointer = 1
      Unload Me
      Exit Sub
    ElseIf result = True Then
    
    result = MsgBox("Achtung! Der Download kann bei langsamer Internetverbindung " & _
    "sehr lange dauern." & vbCrLf & "Der Download benötigt sämtliche Systemresourcen und der PC " & Chr(34) & _
    "hängt" & Chr(34) & " in dieser Zeit scheinbar." & vbCrLf & vbCrLf & _
    "Weiter mit OK..", vbExclamation + vbOKCancel, "Hinweis zum Download")
        If result = 1 Then
        result = URLDownloadToFile(0, "http://www.sai.msu.su/groups/cluster/gcvs/gcvs/iii/iii.dat", _
        App.Path & "\gcvs.txt", 0, 0)
        Else
        Exit Sub
        End If
    End If
    
End If

If fs.FileExists(App.Path & "\gcvs.txt") = False And Not optD_umwand = True Then
     MsgBox "Es ist ein Fehler beim Download aufgetreten" & vbCrLf & _
     "Bitte versuchen Sie es erneut..", vbCritical, "Download nicht erfolgreich"
     cmdDownload.BackColor = &H8000000F
     cmdDownload.Caption = "Start"
     Me.MousePointer = 1
     Exit Sub
End If

With rsa
     .Fields.Append ("ID"), adInteger
     .Fields.Append ("Kürzel"), adVarChar, 6
     .Fields.Append ("Stbld"), adChar, 3
     .Fields.Append ("BP"), adChar, 4
     .Fields.Append ("LBeob"), adDouble
     .Fields.Append ("Max"), adDouble
     .Fields.Append ("MinI"), adVarChar, 5
     .Fields.Append ("Spektr"), adVarChar, 17
     .Fields.Append ("Typ"), adVarChar, 12
     .Fields.Append ("Epoche"), adDouble
     .Fields.Append ("Periode"), adDouble
     .Fields.Append ("hh"), adInteger, 2
     .Fields.Append ("mm"), adInteger, 2
     .Fields.Append ("ss"), adDouble, 4
     .Fields.Append ("vz"), adChar, 1
     .Fields.Append ("o"), adInteger, 2
     .Fields.Append ("m"), adDouble, 5
     .Fields.Append ("DEC"), adDouble
End With

If OptD_overwrite = True Then
    If fs.FileExists(App.Path & "\gcvs.dat") Then fs.DeleteFile (App.Path & "\gcvs.dat")
ElseIf optD_save = True Then
    If fs.FileExists(App.Path & "\gcvs.dat") = True And GCVSThere = True Then
    fs.CopyFile App.Path & "\gcvs.dat", App.Path & "\GCVS" & Format(Date, "ddmmyyyy") & ".dat"
    fs.DeleteFile (App.Path & "\gcvs.dat")
    End If
End If
 
Set fs = New FileSystemObject

If optD_umwand = True Then
    fs.CopyFile cdlgGCVS.FileName, App.Path & "\gcvs.txt"
 End If
 
Set einstrom = fs.OpenTextFile(App.Path & "\gcvs.txt")


While InStr(1, einstrom.ReadLine, "--") = False
Wend



rsa.Open
zähler = 1

While Not einstrom.AtEndOfStream
    zeile = Split(einstrom.ReadLine, "|")
    stern(0) = Trim(Left(zeile(1), 6))
    stern(1) = Mid(zeile(1), 7, 3)

   If Not zeile(2) = "" And Not Trim(zeile(7)) = "" And Not Trim(zeile(9)) = "" Then
    'On Error GoTo errhandler
    With rsa
        .AddNew
        .Fields("ID") = zähler
        .Fields("Kürzel") = stern(0)
        .Fields("Stbld") = stern(1)
        .Fields("BP") = "GCVS"
        .Fields("Max") = Trim(Mid(zeile(4), 2, 5))
        .Fields("MinI") = Trim(Mid(zeile(5), 2, 5))
        .Fields("Spektr") = Trim(zeile(11))
        .Fields("Typ") = Trim(zeile(3))
        .Fields("LBeob") = 0
        If InStr(1, zeile(7), ":") <> 0 Then
           zeile(7) = Trim(Left(zeile(7), InStr(zeile(7), ":") - 1))
           Else
           zeile(7) = Trim(Left(zeile(7), InStr(1, zeile(7), " ")))
        End If
        
        .Fields("Epoche") = zeile(7)
        .Fields("Periode") = Trim(Mid(zeile(9), 2, 16))
        .Fields("hh") = Left(zeile(2), 2)
        .Fields("mm") = Mid(zeile(2), 3, 2)
        .Fields("ss") = Trim(Mid(zeile(2), 5, 4))
        .Fields("vz") = Mid(zeile(2), 9, 1)
        .Fields("o") = Mid(zeile(2), 10, 2)
        .Fields("m") = Round(Mid(zeile(2), 12, 2) + CDbl(Mid(zeile(2), 14, 2)) / 60, 3)
        .Fields("Dec") = FormatNumber(CDbl(.Fields("vz") & .Fields("o") + .Fields("m") / 60), 4)
        .Update
     
    End With
     zähler = zähler + 1

   End If

Wend

   einstrom.Close
   DBFilter = "DEC > " & CInt(txtFilter.text)
   rsa.Filter = DBFilter
   rsa.Save App.Path & "\gcvs.dat"
   Me.MousePointer = 1
   rsa.Close
    fs.DeleteFile (App.Path & "\gcvs.txt")
    Set einstrom = Nothing
    Set rsa = Nothing
    Set fs = Nothing
    frmHaupt.Form_Load
    
     For x = 1 To frmHaupt.cmbGrundlage.ListCount
   If frmHaupt.cmbGrundlage.List(x) = "GCVS" Then
    frmHaupt.cmbGrundlage.ListIndex = x
    Exit For
   End If
 Next

    MsgBox "Die GCVS-Datenbank kann jetzt für" & vbCrLf & "Berechnungen verwendet werden..", vbInformation, "Implementierung erfolgreich"
    frmHaupt.cmdListe.Enabled = True
    frmHaupt.cmbGrundlage.Enabled = True
    Unload Me
    Exit Sub
    
ErrorHandler:
     MsgBox "Es ist ein Fehler aufgetreten. Bitte versuchen Sie es erneut", vbCritical, "Fehler: Download/Umwandlung nicht erfolgreich"
     cmdDownload.BackColor = &H8000000F
     cmdDownload.Caption = "Start"
     Me.MousePointer = 1
     frmHaupt.Form_Load
     frmHaupt.cmdListe.Enabled = True
     frmHaupt.cmbGrundlage.Enabled = True
     Unload Me
End Sub




Private Sub txtFilter_Change()
 If Not IsNumeric(txtFilter.text) Then txtFilter.text = "-20"
End Sub

Private Sub UpDown1_DownClick()
  If CInt(txtFilter.text) - 1 > -36 Then
  txtFilter.text = CInt(txtFilter.text) - 1
  End If
End Sub

Private Sub UpDown1_UpClick()
 If CInt(txtFilter.text) + 1 < 26 Then
  txtFilter.text = CInt(txtFilter.text) + 1
  End If
End Sub


