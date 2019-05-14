Function SheetExists(ByVal SheetNameOrIndex As Variant, _
    Optional ByVal Wb As Workbook = Nothing) As Boolean
  'True if sheet SheetNameOrIndex exists
  On Error Resume Next
  If Wb Is Nothing Then Set Wb = ActiveWorkbook
  SheetExists = Not Wb.Sheets(SheetNameOrIndex) Is Nothing
End Function

Function ZellenGrauAusblenden(HEs As Integer, he As Integer, marker As Integer)
    
HEs = Sheets("Schrank").Range("G" & 53 - he).Value

If (HEs > 1) Then
    If (marker = 1) Then
        Sheets("Schrank").Range("C" & 53 - he + 1 & ":G" & 53 - he + HEs - 1).Interior.Color = RGB(250, 250, 250)
        Sheets("Schrank").Range("C" & 53 - he + 1 & ":G" & 53 - he + HEs - 1).Interior.Pattern = xlPatternCrissCross
    Else
        Sheets("Schrank").Range("C" & 53 - he + 1 & ":G" & 53 - he + HEs - 1).Interior.Color = 0
        Sheets("Schrank").Range("C" & 53 - he + 1 & ":G" & 53 - he + HEs - 1).Interior.Pattern = xlPatternNone
    End If
End If

End Function

Sub TabelleErstellen(he As Integer)
Sheets("HE_Vorlage").Visible = True
Sheets("AKTIV_Vorlage").Visible = True
Dim Tabelle As Worksheet
Dim HEs As Integer
With ActiveWorkbook
 If SheetExists("HE" + Str(he)) Then
  Worksheets("HE" + Str(he)).Activate
 Else
 
  info = Sheets("Schrank").Range("E" & 53 - he).Value
  ports = Sheets("Schrank").Range("C" & 53 - he).Value
  stecker = Sheets("Schrank").Range("D" & 53 - he).Value
 
    If ports = "AKTIV" Then
        Set Tabelle = .Worksheets("AKTIV_Vorlage")
        Tabelle.Copy .Sheets(Sheets.count)
        ActiveSheet.name = "HE" + Str(he)
        ActiveSheet.Range("G4").Value = he
        ActiveSheet.Range("H5").Value = info
    Else
        Set Tabelle = .Worksheets("HE_Vorlage")
        Tabelle.Copy .Sheets(Sheets.count)
        ActiveSheet.name = "HE" + Str(he)
        ActiveSheet.Range("F4").Value = "HE:" & Str(he)
        ActiveSheet.Range("H4").Value = ports
        ActiveSheet.Range("H5").Value = info
        ActiveSheet.Range("J4").Value = stecker
        copyKomp (0)
    End If
  
  Worksheets("Schrank").Activate
  ActiveSheet.Range("B" & 53 - he).Interior.Color = RGB(153, 216, 130)
  ActiveSheet.Range("B" & 53 - he).Value = "[" & he & "][" & ChrW(&H2713) & "]"
 End If
End With
Sheets("HE_Vorlage").Visible = False
Sheets("AKTIV_Vorlage").Visible = False

HEs = Sheets("Schrank").Range("G" & 53 - he).Value

ZellenGrauAusblenden HEs, he, 1

End Sub
Sub TabelleLoschen(he As Integer)
Dim answer As Integer
Dim HEs As Integer
If SheetExists("HE" + Str(he)) Then
    answer = MsgBox("HE " & he & " wirklich löschen?", vbYesNo + vbQuestion)
    If answer = vbYes Then
        Sheets("HE" + Str(he)).Delete
        ActiveSheet.Range("B" & 53 - he).Interior.Color = RGB(242, 242, 242)
        ActiveSheet.Range("B" & 53 - he).Value = "[" & he & "][—]"
    End If
End If

HEs = Sheets("Schrank").Range("G" & 53 - he).Value
ZellenGrauAusblenden HEs, he, 0

End Sub


