VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGridGross 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Ergebnisse"
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   Icon            =   "frmGridGross.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows-Standard
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdGross 
      Height          =   9975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   17595
      _Version        =   393216
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblTemp 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmGridGross"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
     grossGrid_füllen
End Sub

Public Sub grossGrid_füllen()
Dim y, x As Integer
Dim text As String

       'Spaltenanzahl ermitteln
       grdGross.Cols = frmHaupt.grdergebnis.Cols
       grdGross.Rows = frmHaupt.grdergebnis.Rows
       
        'Füllen der Zellen
        For y = 0 To frmHaupt.grdergebnis.Rows - 1
             For x = 0 To frmHaupt.grdergebnis.Cols - 2
                text = frmHaupt.grdergebnis.TextMatrix(y, x)
                grdGross.TextMatrix(y, x) = text
            Next x
        Next y
        
        For y = 1 To frmHaupt.grdergebnis.Rows - 1
            frmHaupt.grdergebnis.Row = y: grdGross.Row = y
            For x = 1 To 11
            frmHaupt.grdergebnis.col = x: grdGross.col = x
            grdGross.CellBackColor = frmHaupt.grdergebnis.CellBackColor
            Next x
        Next y
        
        frmHaupt.grdergebnis.Row = 1: frmHaupt.grdergebnis.col = 1
        'Spalte mit Sternname in fetter Schriftart
        For x = 1 To grdGross.Rows - 1
             grdGross.Row = x
            grdGross.col = 2
            grdGross.CellFontBold = True
        Next x
        
        Grid_AutoSize grdGross, lblTemp
        
        'Spalten in Haupttabelle aus- oder eingeblendet?
        For x = 0 To grdGross.Cols - 1
           If frmHaupt.grdergebnis.ColWidth(x) = 0 Then grdGross.ColWidth(x) = 0
        Next x
         
         
         grdGross.ColAlignmentFixed = flexAlignCenterCenter
         grdGross.ColAlignment = flexAlignCenterCenter

 
End Sub

' AutoSize für das mshflexgrid-Control
 Function Grid_AutoSize(oGrid As MSHFlexGrid, oLabel As Label)
  Dim nRow As Long
  Dim nCol As Long
  Dim nWidth As Long
  Dim nMaxWidth As Long
  
  ' Setzen der Eigenschaften
  With oLabel
    With .Font
      .Name = oGrid.Font.Name
      .Size = oGrid.Font.Size
      .Bold = oGrid.Font.Bold
      .Italic = oGrid.Font.Italic
      .Strikethrough = oGrid.Font.Strikethrough
      .Underline = oGrid.Font.Underline
    End With

    ' Wichtig!
    .WordWrap = False
    .AutoSize = True
  End With
  
  ' Auswerten und Setzen der Grössen
  With oGrid
    For nCol = 0 To .Cols - 1
      nMaxWidth = 0
      For nRow = 0 To .Rows - 1
        oLabel.Caption = .TextMatrix(nRow, nCol)
        nWidth = oLabel.Width
        If nWidth + 100 > nMaxWidth Then nMaxWidth = nWidth + 250
      Next nRow
      
      .ColWidth(nCol) = nMaxWidth
    Next nCol
  End With
End Function


'Sortieren von Spalten bei Klick
Private Sub grdGross_MouseUp(Button As Integer, _
  Shift As Integer, x As Single, y As Single)
' Rechtsklick?
  If Button = vbRightButton Then
    Dim nRow As Long
    Dim nCol As Long
    Dim nRowCur As Long
    Dim nColCur As Long
    Dim data As String
    
    With grdGross
      ' aktuelle Zelle "merken"
      nRowCur = .Row
      nColCur = .col
      
      ' Zelle ermitteln
      nRow = .MouseRow
      nCol = .MouseCol
      
      ' Erfolgte der Mausklick auf eine existierende
      ' Zelle im Grid?
      If nRow >= 0 And nCol >= 0 Then
        ' Zelle zur aktiven Zelle machen (selektieren)
        .Row = nRow
        .col = nCol
        
        If x >= .CellLeft And x <= .CellLeft + .CellWidth And _
          y >= .CellTop And y <= .CellTop + .CellHeight Then
        
          ' Beispiel: PopUp-Menü anzeigen
         
          For x = 1 To .Cols - 1
           data = data & .TextMatrix(nRow, x) & " ; "
           Next x
           
        Else
          ' ursprüngliche Zelle wieder selektieren
          .Row = nRowCur
          .col = nColCur
        End If
      End If
    End With
  End If
  
  If Button = vbLeftButton Then
  
  If grdGross.MouseRow = 0 Then
    grdgross_sortieren (grdGross.col)
      End If
  End If
 If grdGross.col = 1 Then
  Call Infofüllen(grdGross)
  End If
End Sub

'Sortieren des Grid unabhängig vom Hauptfenster
Public Sub grdgross_sortieren(ByVal ColIndex As Integer)
    'Sortieren Datagrid
    frmGridGross.grdGross.col = ColIndex
    If ColIndex < 6 Or ColIndex = 9 Then
    If sorter = 7 Then
         grdGross.Sort = 8
         sorter = 8
        Else
        grdGross.Sort = 7
        sorter = 7
        End If
    End If
    
    If ColIndex >= 6 And Not ColIndex = 9 Then
    If sorter = 7 Then
        grdGross.Sort = 4
        sorter = 8
     Else
     grdGross.Sort = 3
     sorter = 7
     End If
     End If
            
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmHaupt.WindowState = 0
End Sub
