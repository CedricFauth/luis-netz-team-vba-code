Sub copyKomp(ask As Integer)
    Dim answer As Integer
    Dim ports As Integer
    Dim ende As Integer
    
    ende = 13
    ports = Int(Cells(4, 8).Value)
    If ask = 1 Then
        answer = MsgBox("Sollen alle Ports neu generiert werden?", vbYesNo + vbQuestion)
    Else
        answer = vbYes
    End If
    If answer = vbYes Then
        With ActiveSheet
        
            .Range("A16:L110").Clear
            
            Sheets("Informationen").Range("A50", "L52").Copy
            For i = 1 To ports
                .Range("A" & ende, "L" & ende + 2).PasteSpecial Paste:=xlPasteAll
                
                .Cells(ende + 1, 2).Value = i
                ende = ende + 3
                
            Next i
            Sheets("Informationen").Range("A54", "L54").Copy
            .Range("A" & ende, "L" & ende).PasteSpecial Paste:=xlPasteAll
            ende = ende + 1
        End With
    End If
End Sub

Sub deleteKomp()
    Dim answer As Integer
    answer = MsgBox("Soll alle Ports gelöscht werden?", vbYesNo + vbQuestion)
    If answer = vbYes Then
        With ActiveSheet
            .Range("A13:L110").Clear
            Sheets("Informationen").Range("A54", "L54").Copy
            .Range("A13", "L13").PasteSpecial Paste:=xlPasteAll
        End With
    End If
End Sub



