Attribute VB_Name = "mdlGridladen"
Option Explicit

'Grid aus Textdatei füllen
Public Sub LoadGridData(ByVal vGrid As MSHFlexGrid, ByVal _
             vstrfile As String, ByVal vstrSep As String, _
             Optional ByVal nFixedCols As Long = 1, _
             Optional ByVal nFixedRows As Long = 1)

  Dim Fn As Integer

  Dim astrData As Variant
  Dim intCols As Integer

  Dim strTemp As String

  Dim r As Long
  Dim c As Long
        
  Fn = FreeFile()
  Open vstrfile For Input As #Fn

  With vGrid
    .Redraw = False

    .Rows = 0
    .Cols = 0

    Do While Not EOF(Fn)
      Line Input #Fn, strTemp

      If Len(strTemp) <> 0 Then
        astrData = Split(strTemp, vstrSep)
        intCols = UBound(astrData)

        If intCols + 1 > .Cols Then .Cols = intCols + 1

        .Rows = .Rows + 1
        r = .Rows - 1

        For c = 0 To intCols
        .TextMatrix(r, c) = astrData(c)
        Next c

      Else
        .Rows = .Rows + 1
      End If
    Loop

    .Redraw = True
  End With

  Close #Fn
frmHaupt.grdergebnis.Visible = True
  With vGrid
    If .Rows >= nFixedRows + 1 Then
      .FixedRows = nFixedRows
    Else
      .FixedRows = .Rows - 1
    End If

    If .Cols >= nFixedCols + 1 Then
      .FixedCols = nFixedCols
    Else
      .FixedCols = .Cols - 1
    End If

    If .Cols > 1 Then
      For c = 0 To .Cols - 1
        .ColWidth(c) = 900
      Next c
    Else
      .ColWidth(0) = .Width * 0.99
    End If
  End With
  
  Set dicStbld = New Dictionary
  For x = 1 To frmHaupt.grdergebnis.Rows
   On Error Resume Next
      dicStbld.Add frmHaupt.grdergebnis.TextMatrix(x, 1), "  "
   Next x
End Sub

'Speichern von Griddaten in Textdatei (Nicht benutzt in BAVMinMAx)
Public Sub SaveGridData(ByVal vGrid As MSHFlexGrid, ByVal _
               vstrfile As String, ByVal vstrSep As String)

  Dim Fn As Integer

  Dim nRows As Long
  Dim nCols As Long

  Dim strTemp As String

  Dim r As Long
  Dim c As Long

  With vGrid
    nRows = .Rows - 1
    nCols = .Cols - 1

    Fn = FreeFile()
    Open vstrfile For Output As #Fn
      For r = 0 To nRows
        strTemp = vbNullString
        For c = 0 To nCols - 1
          strTemp = strTemp & (.TextMatrix(r, c) & vstrSep)
        Next c
        strTemp = strTemp & .TextMatrix(r, c)

        Print #Fn, strTemp
      Next r
    Close #Fn
  End With
End Sub

