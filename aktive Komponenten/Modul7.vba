Sub aktivGenerieren()
Dim modules As Integer
Dim ports As Integer
Dim ende As Integer
Dim answer As Integer
Dim name As String
Dim he As Integer
Dim mode As String


ende = 11
modules = Int(Cells(10, 4).Value)
ports = Int(Cells(10, 6).Value)
answer = MsgBox("Sollen alle Ports neu generiert werden?", vbYesNo + vbQuestion)

If answer = vbYes Then
    With ActiveSheet
        mode = .Range("K6")
        
        If mode = "Manuell" Then
            name = "Manuell"
        Else
            name = .Range("H7")
        End If
        he = Int(.Cells(4, 7).Value)
        Sheets("Schrank").Cells(53 - he, 4).Value = name
        
        .Range("A11:L110").Clear
        
        For i = 1 To modules
            Sheets("Informationen").Range("A40", "L41").Copy
            .Range("A" & ende, "L" & ende + 1).PasteSpecial Paste:=xlPasteAll
            .Cells(ende + 1, 2).Value = "Modul " & i
            ende = ende + 2
            For j = 1 To ports
                Sheets("Informationen").Range("A43", "L44").Copy
                .Range("A" & ende, "L" & ende + 1).PasteSpecial Paste:=xlPasteAll
                .Cells(ende, 2).Value = "Port " & i & "." & j & ":"
                ende = ende + 2
            Next j
        Next i
        Sheets("Informationen").Range("A46", "L46").Copy
        .Range("A" & ende, "L" & ende).PasteSpecial Paste:=xlPasteAll
        
    End With
End If
End Sub

Sub aktivloeschen()
Dim answer As Integer
answer = MsgBox("Sollen alle Ports neu generiert werden?", vbYesNo + vbQuestion)
If answer = vbYes Then
With ActiveSheet
    .Range("A11:L110").Clear
    Sheets("Informationen").Range("A46", "L46").Copy
    .Range("A" & 11, "L" & 11).PasteSpecial Paste:=xlPasteAll
End With
End If
End Sub
