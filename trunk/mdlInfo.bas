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
Abf.Filter = sSQLi

On Error Resume Next

'Füllen der Labels aus dem Recordset
With frmSterninfo
        .lblStern.Caption = infostern & " " & infostbld
        frmAladin.txtStern.text = .lblStern.Caption
         If frmAladin.chkAladDirekt.Value = 1 Then frmaladain.txtObj.text = frmAladin.txtStern.text
        .lblTyp.Caption = Abf.Fields("Typ").Value
        .lblEpoche.Caption = Format(Abf.Fields("Epoche").Value, "#.0000")
        .lblPeriode.Caption = Format(Abf.Fields("Periode").Value, "#.000000")
        .lblQuelle.Caption = Abf.Fields("BP").Value
        If .lblQuelle.Caption = "ASAS" Then
         .Label8.Caption = "Ampl.:"
        Else
         .Label8.Caption = "Min I:"
        End If
        'If Database < 2 Then
            .lblLBeo.Caption = Abf.Fields("LBeob").Value
            'Else
            '.lblLBeo.Caption = "k.A."
        'End If
            
        .lblMax.Caption = Format(Abf.Fields("Max").Value, "#.00")
        .lblMinI.Caption = Format(Abf.Fields("MinI").Value, "#.00")
        
        .lblMinII.Caption = Format(Abf.Fields("MinII").Value, "#.00")
        .lblD.Caption = Format(Abf.Fields("D").Value, "#.00")
        .lblkd.Caption = Abf.Fields("kd").Value
        .lblMM.Caption = Abf.Fields("M-m").Value
        If InStr(1, Abf.Fields("m").Value, ".") Then
         decmin = Format(Left(Abf.Fields("m").Value, InStr(1, Abf.Fields("m").Value, ".") - 1), "00")
         decsec = Format((Right(Abf.Fields("m").Value, InStrRev(Abf.Fields("m").Value, ".") - 2)) / 10 * 60, "00")
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
         
        decsec = IIf(InStr(1, Abf.Fields("m").Value, "."), Format((Right(Abf.Fields("m").Value, InStrRev(Abf.Fields("m").Value, ".") - 2)) / 10 * 60, "00"), "00")
        infokoord = ausg(.lblStarRA.Caption * 24 / 360, .lblStarDec.Caption)
        If umwand(1) = "24" Then umwand(1) = "00"
        .lblKoord.Caption = "RA       " & umwand(1) & ":" & umwand(2) & ":" & umwand(3) & vbCrLf & _
        "DEC   " & umwand(4) & ":" & umwand(5) & ":" & umwand(6)

 End With
 
        If frmAladin.Visible Then
          If frmSterninfo.lblKoord.Caption <> "" Then
            result = Split(frmSterninfo.lblKoord.Caption, vbCrLf)
            frmAladin.txtObj.text = Trim(Mid(CStr(result(0)), 3, Len(CStr(result(0))) - 2)) & " " & Trim(Mid(CStr(result(1)), 4, Len(CStr(result(1))) - 2))
          End If
        
        frmAladin.txtStern.text = frmSterninfo.lblStern.Caption
        If frmAladin.chkAladDirekt.Value = 1 Then frmaladain.txtObj.text = frmAladin.txtStern.text
       frmAladin.cmdDSS.Visible = True
        End If
        
      'Recordset schliessen, Speicher freigeben
      rsa.Close
      Set rsa = Nothing
      Err.Clear
      Call frmSterninfo.Mondinfo
      Exit Sub
      
Abbruch:
If Err.Number = 3265 Then Resume Next
      MsgBox "Fehler: " & Err.Number & vbCrLf & _
             Err.Description, vbCritical
      Err.Clear
End Sub

