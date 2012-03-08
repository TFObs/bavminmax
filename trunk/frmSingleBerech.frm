VERSION 5.00
Begin VB.Form frmSingleBerech 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Extrema von Einzelsternen"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton Option2 
      Caption         =   "Neueingabe in eigene Datenbank"
      Height          =   375
      Left            =   480
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sternauswahl über Datenbanksuche"
      Height          =   375
      Left            =   480
      TabIndex        =   27
      Top             =   120
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Suche in Datenbanken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text1 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Suchen"
         Height          =   615
         Left            =   2280
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdAusw 
         Caption         =   "auswählen"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5400
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ListBox ListRecherche 
         Height          =   1035
         Left            =   240
         MultiSelect     =   1  '1 -Einfach
         TabIndex        =   22
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Stern     Stbld        Datenbank       Epoche                   Periode"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Neueingabe"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   6975
      Begin VB.CommandButton cmdNeu 
         Caption         =   "Speichern"
         Height          =   495
         Left            =   5040
         TabIndex        =   20
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox cmbvz 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   17
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbDECd 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   16
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbDECs 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbDECm 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbRAs 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox cmbRAm 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox cmbRAh 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtPeriode 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtEpoche 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtStbld 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtStern 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         Caption         =   "  [+-]         gg          mm         ss       "
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Zentriert
         Caption         =   "hh       mm        ss       "
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Sternbild:"
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
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Epoche:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Periode:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "RA (J2000):"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "DEC (J2000):"
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
         Index           =   1
         Left            =   3000
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Stern:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSingleBerech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Dim dbsbavsterne As ADODB.Connection
 Dim rssourcerecord  As ADODB.Recordset
 Dim rssingleabfrage  As ADODB.Recordset
 Dim feld As Field
 Dim fs As FileSystemObject
 Dim gewählt As Collection
 Dim x As Integer
 Dim result

Sub FillList(ByRef StarName, ByVal pfad As String)
Dim Listentext As String

 Set dbsbavsterne = New ADODB.Connection
 Set rssourcerecord = New ADODB.Recordset
 Set fs = New FileSystemObject
On Error GoTo errhandler
 If fs.FileExists(pfad & "\recordsets.dat") Then fs.DeleteFile (pfad & "\recordsets.dat")

   'Neues Recordset zur Aufnahme der Rechercheergebnisse erzeugen
   With rssourcerecord
       .Fields.Append ("ID"), adInteger
       .Open
       .Save pfad & "\recordsets.dat"
   End With
 
For x = 0 To 5

Set rssingleabfrage = New ADODB.Recordset

 With rssingleabfrage

        If x < 2 Then
        
            'Verbindung zur Datenbank herstellen
            With dbsbavsterne
                .Provider = "microsoft.Jet.oledb.4.0"
                If x = 0 Then
                    .ConnectionString = pfad & "\Bav_sterne.mdb"
                ElseIf x = 1 Then
                    .ConnectionString = pfad & "\BAV_sonstige.mdb"
                End If
                .Open
             End With
             
            .ActiveConnection = dbsbavsterne
            .LockType = adLockOptimistic
            .CursorLocation = adUseClient
            .CursorType = adOpenForwardOnly ' Kleinster Verwaltungsaufwand
            .Open "SELECT * FROM BVundRR Where Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
            
        ElseIf x = 2 Then
       
        If fs.FileExists(pfad & "\Kreiner.dat") Then
          .Open pfad & "\Kreiner.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If

        ElseIf x = 3 Then
        
        If fs.FileExists(pfad & "\GCVS.dat") Then
          .Open pfad & "\GCVS.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If

        ElseIf x = 4 Then
        
        If fs.FileExists(pfad & "\BAVBA_EA.dat") Then
          .Open pfad & "\BAVBA_EA.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        
        ElseIf x = 5 Then
        
        If fs.FileExists(pfad & "\BAVBA_RR.dat") Then
          .Open pfad & "\BAVBA_RR.dat"
          .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        End If
        
        'Für Spätere Erweiterungen: Berechnungen aus eigener Datenbank
        'ElseIf x = 6 Then
        'If fs.FileExists(pfad & "\Einzel.dat") Then
         ' .Open pfad & "\Einzel.dat"
         ' .Filter = "Kürzel = '" & StarName(0) & "' AND Stbld = '" & StarName(1) & "'"
        'End If

        End If
        
    If rssingleabfrage.State = adStateOpen Then
    
        If rssingleabfrage.RecordCount > 0 Then
        
           If rssourcerecord.State = adStateClosed Then rssourcerecord.Open pfad & "\recordsets.dat"
        
                With rssourcerecord
                 ' Nur, wenn Daten vorhanden
                    If .RecordCount > 0 Then
                        If Not .BOF Then .MoveLast
                         
                        ' Neuer Datensatz erzeugen
                        If Not rssingleabfrage.Fields.Count = 0 Then _
                        .AddNew
      
                        ' Felder aus Clone-Recordset lesen und in
                        ' Original-Recordset speichern

                        On Error Resume Next
                        For Each feld In rssingleabfrage.Fields
                            .Fields(feld.Name).Value = feld.Value
                        Next feld
                              
                        ' Datensatz speichern
                        .Update
                    .Save pfad & "\recordsets.dat"
                    End If
                 End With
                
            'Wenn noch keine Daten vorhanden sind, muss die
        'rsSingleAbfrage "gecloned" werden. Allerdings muss der
        'Filter erhalten bleiben...
            If rssourcerecord.RecordCount = 0 Then
                fs.DeleteFile (pfad & "\recordsets.dat")
                rssingleabfrage.Save pfad & "\recordsets.dat"
                rssourcerecord.Close
                rssourcerecord.Open pfad & "\recordsets.dat"
                fs.DeleteFile (pfad & "\recordsets.dat")
                rssourcerecord.Save pfad & "\recordsets.dat"
             End If
             

         End If
                
     End If

    'Schliessen der Abfrage und Beenden der Verbindung zur Access-Datenbank
    If rssingleabfrage.State = adStateOpen Then rssingleabfrage.Close
    If dbsbavsterne.State = adStateOpen And x < 2 Then dbsbavsterne.Close

    End With

Next x
 

  'Anzeigen der Ergebnisse in einer Liste
  With rssourcerecord

    If .RecordCount > 0 Then
        cmdAusw.Enabled = True
        .MoveFirst
        Do While Not .EOF
            Listentext = .Fields("Kürzel").Value & vbTab & .Fields("Stbld").Value & vbTab & _
          .Fields("BP").Value & vbTab & "   " & Format(.Fields("Epoche").Value, "#.0000") & vbTab & _
          Format(.Fields("Periode").Value, "#.00000000")
          
          If Not IsInList(Listentext) Then ListRecherche.AddItem Listentext
          .MoveNext
        Loop
     ElseIf .RecordCount = 0 Then
        MsgBox "Es konnte kein Eintrag in den Datenbanken " & vbCrLf & "gefunden werden. Bitte " & _
        "ändern Sie die Abfrage.", vbInformation, "Kein Eintrag vorhanden"
        Exit Sub
             
    End If
    
  End With
    
If fs.FileExists(pfad & "\recordsets.dat") Then fs.DeleteFile (pfad & "\recordsets.dat")

Set fs = Nothing
Set rssingleabfrage = Nothing
Exit Sub

errhandler:

MsgBox Err.Number & " " & Err.Description & vbCrLf & _
     "Form: SingleBerech, Sub: FillList" & vbCrLf & vbCrLf & _
     "Bitte überprüfen Sie die Eingabe.", vbCritical, "Unzulässige Eingabe"


End Sub
Private Function IsInList(ByVal Listentext As String) As Boolean

For x = 0 To ListRecherche.ListCount - 1
If Listentext = ListRecherche.List(x) Then
    IsInList = True
    Exit Function
End If
Next x

IsInList = False
End Function
Private Sub cmdNeu_Click()
Dim rsneu As ADODB.Recordset

Set fs = New FileSystemObject
Set rsneu = New ADODB.Recordset

  'Anlegen einer Datenbank, wenn nicht vorhanden
  If Not fs.FileExists(App.Path & "\Einzel.dat") Then
    With rsneu
        .Fields.Append ("ID"), adInteger
        .Fields.Append ("Kürzel"), adVarChar, 6
        .Fields.Append ("Stbld"), adChar, 3
        .Fields.Append ("BP"), adChar, 5
        .Fields.Append ("LBeob"), adDouble
        .Fields.Append ("Max"), adDouble
        .Fields.Append ("Spektr"), adVarChar, 11
        .Fields.Append ("D"), adDouble, 4
        .Fields.Append ("kD"), adDouble, 4
        .Fields.Append ("Typ"), adVarChar, 12
        .Fields.Append ("Epoche"), adDouble
        .Fields.Append ("Periode"), adDouble
        .Fields.Append ("for"), adVarChar, 3
        .Fields.Append ("hh"), adInteger, 2
        .Fields.Append ("mm"), adInteger, 2
        .Fields.Append ("ss"), adDouble, 4
        .Fields.Append ("vz"), adChar, 1
        .Fields.Append ("o"), adInteger, 2
        .Fields.Append ("m"), adDouble, 5
        .Open
        .Save App.Path & "\Einzel.dat"
        .Close
      End With
  End If

 'Öffnen der Eigenen Datenbank
 rsneu.Open App.Path & "\Einzel.dat"

  With rsneu

    If .RecordCount >= 0 Then

    If .RecordCount > 0 Then .MoveLast  'Nur wenn schon ein Eintrag vorhanden ist
        'Eintagen der Felder
        .AddNew
        .Fields("ID").Value = .RecordCount + 1
        .Fields("Kürzel").Value = txtStern.text
        .Fields("Stbld").Value = txtStbld.text
        .Fields("BP").Value = "EIGEN"
        .Fields("Epoche").Value = txtEpoche.text
        .Fields("Periode").Value = txtPeriode.text
        .Fields("hh").Value = cmbRAh.List(cmbRAh.ListIndex)
        .Fields("mm").Value = cmbRAm.List(cmbRAm.ListIndex)
        .Fields("ss").Value = cmbRAs.List(cmbRAs.ListIndex)
        .Fields("vz").Value = cmbvz.List(cmbvz.ListIndex)
        .Fields("o").Value = cmbDECd.List(cmbDECd.ListIndex)
        .Fields("m").Value = cmbDECm.List(cmbDECm.ListIndex) + cmbDECs.List(cmbDECs.ListIndex) / 60
        .Update
        .Save App.Path & "\Einzel.dat"
    End If

  End With
  
Unload Me

frmHaupt.Form_Load
frmHaupt.cmbGrundlage.ListIndex = frmHaupt.cmbGrundlage.ListCount - 1
frmHaupt.cmdListe.Enabled = True: VTabs.TabEnabled(1) = True
frmHaupt.cmbGrundlage.Enabled = True
Unload frmSterninfo
Unload frmAladin
Unload frmBerechnungsfilter

Set fs = Nothing
Set rsneu = Nothing

End Sub


Private Sub Command1_Click()
Dim searchstar
If Not Text1.text = "" Then
  cmdAusw.Enabled = False
  ListRecherche.Clear
  searchstar = Split(Text1.text, " ")
  FillList searchstar, App.Path
End If
End Sub

Private Sub Form_Load()
Set fs = New FileSystemObject
    If fs.FileExists(App.Path & "\Einzel.dat") Then fs.DeleteFile (App.Path & "\Einzel.dat")

  'Eintragen von Standardwerten in die Comboboxen
  For x = 0 To 23
    cmbRAh.AddItem x
    cmbRAm.AddItem x
    cmbRAs.AddItem x
    cmbDECd.AddItem x
    cmbDECm.AddItem x
    cmbDECs.AddItem x
  Next x

  For x = 24 To 59
    cmbRAm.AddItem x
    cmbRAs.AddItem x
    cmbDECm.AddItem x
    cmbDECs.AddItem x
    Next x

  For x = 60 To 90
    cmbDECd.AddItem x
  Next x

  cmbvz.AddItem "+"
  cmbvz.AddItem "-"

  cmbRAh.ListIndex = 0
  cmbRAm.ListIndex = 0
  cmbRAs.ListIndex = 0
  cmbDECd.ListIndex = 0
  cmbDECm.ListIndex = 0
  cmbDECs.ListIndex = 0
  cmbvz.ListIndex = 0

  If frmSterninfo.Visible Then Text1.text = frmSterninfo.lblStern.Caption
  
  Call Option1_Click
  Frame1Zeigen False
Set fs = Nothing
End Sub

Private Sub Option2_Click()
If Frame2.Enabled Then
    Frame2.Enabled = False
    Frame1.Enabled = True
    Frame1Zeigen True
End If
End Sub

Private Sub Option1_Click()
If Frame1.Enabled Then
    Frame1.Enabled = False
    Frame1Zeigen False
    Frame2.Enabled = True
End If
End Sub

Private Sub cmdAusw_Click()
Set gewählt = New Collection
Dim sSQL As String
Dim Auswahl
Set fs = New FileSystemObject

sSQL = ""

For x = 0 To ListRecherche.ListCount - 1

   If ListRecherche.Selected(x) Then
        Auswahl = Split(ListRecherche.List(x), vbTab)
        gewählt.Add "(BP = '" & Auswahl(2) & "')"
        'AND Epoche = '" & Auswahl (3) & "' AND Periode = '" & Auswahl(4) & "') "
   End If

Next x

If gewählt.Count > 0 Then
    For x = 1 To gewählt.Count - 1
        sSQL = sSQL & gewählt.Item(x) & " OR "
    Next x
        sSQL = sSQL & gewählt.Item(gewählt.Count)
        
Else: result = MsgBox("Es ist kein Datensatz ausgewählt..." & vbCrLf _
& "Alle angezeigten Elemente werde übernommen." & vbCrLf & vbCrLf & "Fortfahren ?", vbExclamation + vbYesNo, "keine Auswahl getroffen...")

    If result = vbYes Then
        cmdAusw.Enabled = False
        Else: Exit Sub
    End If

End If

If Not rssourcerecord Is Nothing Then
    rssourcerecord.Filter = sSQL
    If fs.FileExists(App.Path & "\Einzel.dat") Then fs.DeleteFile (App.Path & "\Einzel.dat")
    rssourcerecord.Save (App.Path & "\Einzel.dat")
    rssourcerecord.Close
Else
    MsgBox "Abfragedatei nicht vorhanden oder beschädigt." & vbCrLf & _
        "Bitte Abfrage erneut durchführen", vbCritical, "Fehler der Abfragedatei"
End If

Unload Me

 frmHaupt.Form_Load
 frmHaupt.cmdListe.Enabled = True: frmHaupt.VTabs.TabEnabled(1) = True
 frmHaupt.cmbGrundlage.Enabled = True
 
 For x = 1 To frmHaupt.cmbGrundlage.ListCount
 If frmHaupt.cmbGrundlage.List(x) = "Einzeln" Then
    frmHaupt.cmbGrundlage.ListIndex = x
    Exit For
 End If
 Next
 
Unload frmSterninfo
Unload frmAladin
Unload frmBerechnungsfilter

Set gewählt = Nothing
Set rssourcerecord = Nothing
Set fs = Nothing

End Sub

Sub Frame1Zeigen(ByVal show As Boolean)

  cmbvz.Enabled = show
  cmbvz.Enabled = show
  cmbRAh.Enabled = show
  cmbRAm.Enabled = show
  cmbRAs.Enabled = show
  cmbDECd.Enabled = show
  cmbDECm.Enabled = show
  cmbDECs.Enabled = show
  cmbvz.Enabled = show
  For x = 0 To 5
    Label2(x).Enabled = show
  Next x
  Label3.Enabled = show
  Label4.Enabled = show
  cmdNeu.Enabled = show
  
  Label1.Enabled = Not show
  Text1.Enabled = Not show
  ListRecherche.Enabled = Not show
  cmdAusw.Enabled = Not show
  Command1.Enabled = Not show
  
End Sub
