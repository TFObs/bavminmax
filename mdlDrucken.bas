Attribute VB_Name = "mdlDrucken"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, ByVal lParam As Long) _
        As Long
        
Const WM_USER = &H400
Const VP_FORMATRANGE = WM_USER + 125
Const VP_YESIDO = 456654

Private Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Private Type TFormatRange
  hdc         As Long
  hdcTarget   As Long
  rc          As RECT
  rcPage      As RECT
End Type





Sub PrintGrid(grid As MSHFlexGrid, ByVal LeftMargin As Single, _
              ByVal TopMargin As Single, ByVal RightMargin As _
              Single, ByVal BottomMargin As Single, Titel As _
              String, Datum As String, Optional many As Integer)
              
  Dim tRange As TFormatRange
  Dim lReturn As Long
  Dim DName As String
  Dim DSchacht As Integer
  Dim gbeg As Long
  Dim CopyCW() As Long
  Dim GRef As Boolean
  Dim x%
  
    GRef = False
    If many > 0 Then
      ' Anzahl der zu druckenden Colums festlegen
      ' Alles > many wird auf colwidth = 0 gesetzt
      If grid.Cols > many Then
        gbeg = grid.Cols - many
        ReDim CopyCW(gbeg)
        grid.Redraw = False
        For x = many To grid.Cols - 1
          CopyCW(x - many) = grid.ColWidth(x)
          grid.ColWidth(x) = 0
        Next x
        GRef = True
      End If
    End If
    
    'mit wParam <> 0 kann überprüft werden
    'ob das Control OPP unterstützt, wenn ja wird
    '456654 (VP_YESIDO) zurückgeliefert
    lReturn = SendMessage(grid.hWnd, VP_FORMATRANGE, 1, 0)
    
    If lReturn = VP_YESIDO Then
      
      'Struktur mit Formatierungsinformationen füllen
      Printer.ScaleMode = vbPixels
      
      With tRange
        .hdc = Printer.hdc
        
        'Höhe und Breite einer Seite (in Pixel)
        .rcPage.Right = Printer.ScaleWidth
        .rcPage.Bottom = Printer.ScaleHeight
        
        'Lage und Abmessungen des Bereichs auf den
        'gedruckt werden soll (in Pixel)
        .rc.Left = Printer.ScaleX(LeftMargin, vbMillimeters)
        .rc.Top = Printer.ScaleY(TopMargin, vbMillimeters)
        .rc.Right = .rcPage.Right - Printer.ScaleX(RightMargin, _
                                                   vbMillimeters)
                                                   
        .rc.Bottom = .rcPage.Bottom - Printer.ScaleY(BottomMargin, _
                                                     vbMillimeters)
      End With
  
      'Drucker initialisieren
      Printer.Print vbNullString
      
      'Seite(n) drucken
      x = 1
      Do
        Printer.CurrentX = Printer.ScaleX(LeftMargin, vbMillimeters)
        Printer.CurrentY = Printer.ScaleY(10, vbMillimeters)
        If Titel <> "" Then Printer.Print Titel & ", Seite " & x 'Anzeige der Seitenzahl
  
        Printer.CurrentX = Printer.ScaleX(LeftMargin, vbMillimeters)
        Printer.CurrentY = Printer.ScaleY(16, vbMillimeters)
        
        If Datum <> "" Then
          Printer.Print Datum
        Else
          Printer.Print Format(Date, "DD.MM.YYYY")
        End If
        lReturn = SendMessage(grid.hWnd, VP_FORMATRANGE, 0, _
                              VarPtr(tRange))
        
        If lReturn < 0 Then
          Exit Do
        Else
          Printer.NewPage
          x = x + 1 ' Neue Seite
        End If
      Loop
      Printer.EndDoc
  
      'Reset
      lReturn = SendMessage(grid.hWnd, VP_FORMATRANGE, 0, 0)
    End If
    
    If GRef Then
      'Alle Colums wieder in richtiger Breite darstellen
      For x = many To grid.Cols - 1
        grid.ColWidth(x) = CopyCW(x - many)
      Next x
      grid.Redraw = True
    End If
End Sub

