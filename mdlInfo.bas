Attribute VB_Name = "mdlInfo"
Public Sub Infofüllen(grid As MSHFlexGrid)
   Dim rsa As ADODB.Recordset
    Dim Abf As ADODB.Recordset
     Dim fs As New FileSystemObject
     Dim sSQL() As String
     Dim sSQLi
     Dim infobez As String
     Dim infostbld As String
     Dim decsec, infokoord
     Dim decmin
     
     
'Fehlerbehandlung
    If Not fs.FileExists(pfad & "\info.dat") Then
        Exit Sub
    End If
    
With frmSterninfo
        .lblStern.Caption = ""
        .lblTyp.Caption = ""
        .lblEpoche.Caption = ""
        .lblPeriode.Caption = ""
        .lblMax.Caption = ""
        .lblMinI.Caption = ""
        .lblMinII.Caption = ""
        .lblD.Caption = ""
        .lblkd.Caption = ""
        .lblMM.Caption = ""
        .lblKoord.Caption = ""
 End With
 
      On Error GoTo Abbruch
      
      Set rsa = New ADODB.Recordset
    If rsa.State = adStateOpen Then rsa.Close
    
      'Recordset laden "info.dat" mit allen infos
      With rsa
         .CursorType = adOpenKeyset
         .LockType = adLockReadOnly
         .Open pfad & "\info.dat", , , adLockOptimistic
       End With
       
   If rsa.RecordCount = 0 Then Exit Sub
   
 Set Abf = New ADODB.Recordset
 

infostern = grid.TextMatrix(grid.Row, 2)


infostbld = Right(infostern, 3)
infostern = Trim(Left(infostern, Len(infostern) - 3))
sSQLi = "Kürzel = '" & infostern & "' AND Stbld = '" & infostbld & "'"

If Not grid.TextMatrix(grid.Row, 8) = "" Then _
sSQLi = sSQLi & " AND BP = '" & grid.TextMatrix(grid.Row, 8) & "'"

Set Abf = rsa
DoEvents
Abf.Filter = sSQLi

On Error Resume Next

'Füllen der Labels aus dem Recordset
With frmSterninfo
        .lblStern.Caption = infostern & " " & infostbld
        .lblStern.ToolTipText = .lblStern.Caption
        frmAladin.txtStern.text = .lblStern.Caption
        frmHaupt.txtSingleStar.text = .lblStern.Caption: frmHaupt.ListRecherche.Clear
         If frmAladin.chkAladDirekt.Value = 1 Then frmAladin.txtObj.text = frmAladin.txtStern.text
        .lblTyp.Caption = Abf.Fields("Typ").Value
        .lblEpoche.Caption = Format(Abf.Fields("Epoche").Value, "#.0000")
        .lblPeriode.Caption = Format(Abf.Fields("Periode").Value, "#.000000")
        .lblPeriode.ToolTipText = Abf.Fields("Periode").Value
        .lblQuelle.Caption = Abf.Fields("BP").Value
        If .lblQuelle.Caption = "ASAS" Then
         .Label8.Caption = "Ampl.:"
        Else
         .Label8.Caption = "Min I:"
        End If
        'If Database < 2 Then
            If FieldExists(Abf, "LBeob") Then
                .lblLBeo.Caption = Format(Abf.Fields("LBeob").Value, "#.00")
                .lblLBeo.ToolTipText = Format(JulinDat(.lblLBeo.Caption + 2400000), "dd.mm.yyyy hh:mm:ss")
            End If 'Else
            '.lblLBeo.Caption = "k.A."
        'End If
            
        If FieldExists(Abf, "Max") Then .lblMax.Caption = Format(Abf.Fields("Max").Value, "#.00")
        If FieldExists(Abf, "MinI") Then .lblMinI.Caption = Format(Abf.Fields("MinI").Value, "#.00")
        
        If FieldExists(Abf, "MinII") Then .lblMinII.Caption = Format(Abf.Fields("MinII").Value, "#.00")
        If FieldExists(Abf, "D") Then .lblD.Caption = Format(Abf.Fields("D").Value, "#.00")
        If FieldExists(Abf, "kd") Then .lblkd.Caption = Format(Abf.Fields("kd").Value, "#.00")
        If FieldExists(Abf, "M-m") Then .lblMM.Caption = Format(Abf.Fields("M-m").Value, "#.0")
        
        If InStr(1, Abf.Fields("m").Value, ".") Then
         decmin = Format(Left(Abf.Fields("m").Value, InStr(1, Abf.Fields("m").Value, ".") - 1), "00")
         decsec = Format((Right(Abf.Fields("m").Value, InStrRev(Abf.Fields("m").Value, "."))) / 10 * 60, "00")
         Else
         decmin = Format(Abf.Fields("m").Value, "00")
         decsec = "00"
        End If
         
         frmSterninfo.lblStarRA.Caption = (Abf.Fields("hh").Value + Abf.Fields("mm").Value / 60 + Abf.Fields("ss").Value / 3600) * 360 / 24
         If Abf.Fields("vz") = "-" Then
         frmSterninfo.lblStarDec.Caption = (Abf.Fields("o").Value + Abf.Fields("m").Value / 60) * -1
         Else
         frmSterninfo.lblStarDec.Caption = (Abf.Fields("o").Value + Abf.Fields("m").Value / 60)
         End If
         
        decsec = IIf(InStr(1, Abf.Fields("m").Value, "."), Format((Right(Abf.Fields("m").Value, InStrRev(Abf.Fields("m").Value, "."))) / 10 * 60, "00"), "00")
        infokoord = ausg(.lblStarRA.Caption * 24 / 360, .lblStarDec.Caption)
        If umwand(1) = "24" Then umwand(1) = "00"
        .lblKoord.Caption = "RA       " & umwand(1) & ":" & umwand(2) & ":" & umwand(3) & vbCrLf & _
        "DEC   " & umwand(4) & ":" & umwand(5) & ":" & umwand(6)
        If frmHaupt.cmdInfo.Enabled And frmSterninfo.Visible Then frmHaupt.cmdStarChart.Visible = True
       
 End With
 
        If frmAladin.Visible Then
          If frmSterninfo.lblKoord.Caption <> "" Then
            result = Split(frmSterninfo.lblKoord.Caption, vbCrLf)
            frmAladin.txtObj.text = Trim(Mid(CStr(result(0)), 3, Len(CStr(result(0))) - 2)) & " " & Trim(Mid(CStr(result(1)), 4, Len(CStr(result(1))) - 2))
          End If
        
        frmAladin.txtStern.text = frmSterninfo.lblStern.Caption
        frmHaupt.txtSingleStar.text = frmSterninfo.lblStern.Caption: frmHaupt.ListRecherche.Clear
        If frmAladin.chkAladDirekt.Value = 1 Then frmAladin.txtObj.text = frmAladin.txtStern.text
       frmAladin.cmdDSS.Visible = True
        End If
        
      'Recordset schliessen, Speicher freigeben
      rsa.Close
      Abf.Close
      Set rsa = Nothing
      Set Abf = Nothing
      Err.Clear
      'result = Mondinfo(grid)
      Exit Sub
      
Abbruch:
If Err.Number = 3265 Then Resume Next

      MsgBox "Fehler: " & Err.Number & vbCrLf & _
             Err.Description, vbCritical
      Err.Clear
End Sub

Public Sub Mondinfo(gitter As MSHFlexGrid)
Dim Day, Month, year, Uhrzeit As Double
Dim sonne, mond, ephem, monddist

'Wichtig! Hier Deklaration von PI,RAD,Degree
Mpi = 4 * Atn(1)
Mdeg = (4 * Atn(1)) / 180
Mrad = 180 / (4 * Atn(1))

On Error Resume Next

'Ermitteln der Zeitangaben aus dem Grid
Uhrzeit = CDbl(FormatNumber(gitter.TextMatrix(gitter.Row, 4), 5)) * 24
Day = CDbl(Left(gitter.TextMatrix(gitter.Row, 3), 2))
Month = CDbl(Mid(gitter.TextMatrix(gitter.Row, 3), 4, 2))
year = CDbl(Right(gitter.TextMatrix(gitter.Row, 3), 4))


'Berechnung von Sonnen und Mondephemeriden
sonne = SunPosition(CalcJD(Day, Month, year, Uhrzeit))
mond = MoonPosition(sonne(2), sonne(3), CalcJD(Day, Month, year, Uhrzeit))
ephem = MoonRise(CalcJD(Day, Month, year, Uhrzeit), 65, 10 * Mdeg, 50 * Mdeg, 0, 1)
monddist = gitter.TextMatrix(gitter.Row, 11) 'Moondistance(frmSterninfo.lblStarRA.Caption, frmSterninfo.lblStarDec.Caption, (mond(0) * Mrad / 15) / 24 * 360, mond(1) * Mrad)

With frmSterninfo
'Ausgabe in Labelfeld
.lblAuf.Caption = FormatDateTime(ephem(0) / 24, vbShortTime)
.lblTrans.Caption = FormatDateTime(ephem(1) / 24, vbShortTime)
.lblUnter.Caption = FormatDateTime(ephem(2) / 24, vbShortTime)
.lblPhaseText.Caption = mond(2)
.lblPhase.Caption = "Phase: " & Format(mond(3), "0.0 %")
.lblDist.Caption = "Monddistanz: " & Format(Round(monddist), "#") & " °"
.imgMoonPhase.Picture = .imgPhase(mond(6)).Picture
End With

End Sub
Public Function FieldExists(rs As Recordset, sFieldName As String) As Boolean
    Dim fld As Field
    For Each fld In rs.Fields
        If UCase(fld.Name) = UCase(sFieldName) Then
            FieldExists = True
            Exit Function
        End If
    Next fld
End Function
