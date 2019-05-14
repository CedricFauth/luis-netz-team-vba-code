Sub newWorkbook()
Dim relativePath As String
'Copy the data
    Sheets("PF_Export").Range("A59:I102").Copy
'Create a new workbook
    Workbooks.Add
'Paste the data
    ActiveSheet.Paste Destination:=Range("A1")
'Save the newly created workbook
    Application.DisplayAlerts = False
    relativePath = ThisWorkbook.Path & "\" & Left(ThisWorkbook.name, Len(ThisWorkbook.name) - 5) & "_pf_export.xlsx"
    ActiveWorkbook.SaveAs Filename:=relativePath
    Application.DisplayAlerts = True
End Sub

Sub generateExportTable()


Dim startRow As Integer
Dim HEs As String
Dim Campusname As String
Dim BuildingName As String
Dim FloorName As String
Dim RoomNumber As String
Dim RackName As String
Dim Rackheight As String

Dim Portzahl As String
Dim Beschreibung As String
Dim Steckertyp As String

'General
Campusname = Sheets("Schrank").Cells(4, 14).Value
BuildingName = Sheets("Schrank").Cells(4, 3).Value
FloorName = Sheets("Schrank").Cells(6, 14).Value
RoomNumber = Sheets("Schrank").Cells(5, 3).Value
RackName = Sheets("Schrank").Cells(5, 7).Value
Rackheight = Sheets("Schrank").Cells(6, 7).Value
startRow = 10


'Rack
Sheets("PF_Export").Cells(60, 1).Value = Campusname
Sheets("PF_Export").Cells(60, 2).Value = BuildingName
Sheets("PF_Export").Cells(60, 3).Value = FloorName
Sheets("PF_Export").Cells(60, 4).Value = RoomNumber
Sheets("PF_Export").Cells(60, 6).Value = RackName

'If Sheets("Schrank").Cells(6, 7).Value = "40" Then
'    Sheets("PF_Export").Cells(3, 1).Copy
'ElseIf Sheets("Schrank").Cells(6, 7).Value = "33" Then
'    Sheets("PF_Export").Cells(4, 1).Copy
'Else
'    Sheets("PF_Export").Cells(2, 1).Copy
'End If

Sheets("PF_Export").Cells(60, 5).Value = "19" & Chr(34) & " Schrank " & Rackheight & "HE"


'All Units
For i = 1 To 42
    Beschreibung = ""
    'Setup general information
    HEs = Sheets("Schrank").Cells(i + startRow, 7).Value
    If Sheets("Schrank").Cells(i + startRow, 3).Value <> "LEER" And _
    Sheets("Schrank").Cells(i + startRow, 3).Value <> "" Then
    
        Beschreibung = Sheets("Schrank").Cells(i + startRow, 5).Value
    
        Sheets("PF_Export").Cells(i + 60, 1).Value = Campusname
        Sheets("PF_Export").Cells(i + 60, 2).Value = BuildingName
        Sheets("PF_Export").Cells(i + 60, 3).Value = FloorName
        Sheets("PF_Export").Cells(i + 60, 4).Value = RoomNumber
        Sheets("PF_Export").Cells(i + 60, 6).Value = Beschreibung
        Sheets("PF_Export").Cells(i + 60, 7).Value = 43 - i
        Sheets("PF_Export").Cells(i + 60, 8).Value = RackName
        
        Portzahl = Sheets("Schrank").Cells(i + startRow, 3).Value
        Steckertyp = Sheets("Schrank").Cells(i + startRow, 4).Value
    
        'BP
        If Portzahl = "BP" Then
            Sheets("PF_Export").Cells(10 + HEs, 2).Copy
            Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
        '6
        ElseIf Portzahl = "6" And Steckertyp = "DSC" Then
            Sheets("PF_Export").Cells(15 + HEs, 3).Copy
            Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
        '10
        ElseIf Portzahl = "10" Then
            If Steckertyp = "LSA+" Then
                Sheets("PF_Export").Cells(40 + HEs, 4).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        '12
        ElseIf Portzahl = "12" Then
            If Steckertyp = "DSC" Then
                Sheets("PF_Export").Cells(15 + HEs, 5).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "SC" Then
                Sheets("PF_Export").Cells(20 + HEs, 5).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "DLC" Then
                Sheets("PF_Export").Cells(25 + HEs, 5).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "DST" Then
                Sheets("PF_Export").Cells(30 + HEs, 5).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "ST" Then
                Sheets("PF_Export").Cells(35 + HEs, 5).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        '20
        ElseIf Portzahl = "20" Then
            If Steckertyp = "LSA+" Then
                Sheets("PF_Export").Cells(40 + HEs, 6).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "RJ45" Then
                Sheets("PF_Export").Cells(45 + HEs, 6).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        '24
        ElseIf Portzahl = "24" Then
            If Steckertyp = "DSC" Then
                Sheets("PF_Export").Cells(15 + HEs, 7).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "DLC" Then
                Sheets("PF_Export").Cells(25 + HEs, 7).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "DST" Then
                Sheets("PF_Export").Cells(30 + HEs, 7).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            ElseIf Steckertyp = "RJ45" Then
                Sheets("PF_Export").Cells(45 + HEs, 7).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        '25
        ElseIf Portzahl = "25" Then
            If Steckertyp = "RJ45" Then
                Sheets("PF_Export").Cells(45 + HEs, 8).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        '48
        ElseIf Portzahl = "48" Then
            If Steckertyp = "RJ45" Then
                Sheets("PF_Export").Cells(45 + HEs, 9).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        '50
        ElseIf Portzahl = "50" Then
            If Steckertyp = "RJ45" Then
                Sheets("PF_Export").Cells(45 + HEs, 10).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        '64
        ElseIf Portzahl = "64" Then
            If Steckertyp = "RJ45" Then
                Sheets("PF_Export").Cells(45 + HEs, 11).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        'aktive Komp.
        ElseIf Portzahl = "AKTIV" Then
            If Steckertyp = "Manuell" Then
                Sheets("PF_Export").Cells(10 + HEs, 2).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
                Beschreibung = "Switch manuell hinzufügen"
            Else
                Sheets("Schrank").Cells(10 + i, 4).Copy
                Sheets("PF_Export").Cells(i + 60, 5).PasteSpecial xlPasteValues
            End If
        Else
            Sheets("PF_Export").Cells(i + 60, 5).Value = "UNBEKANNT"
        End If
        
        'ComponentName
        If Steckertyp = "Manuell" Then
            Sheets("PF_Export").Cells(i + 60, 6).Value = Beschreibung
        Else
            Sheets("PF_Export").Cells(i + 60, 5).Copy
            Sheets("PF_Export").Cells(i + 60, 6).PasteSpecial xlPasteValues
        End If
        
        Sheets("PF_Export").Cells(i + 60, 6).Value = Sheets("PF_Export").Cells(i + 60, 6).Value & "-" & 43 - i
        
        'ComponentDescription
        Sheets("PF_Export").Cells(i + 60, 9).Value = Beschreibung
        
    Else
        Sheets("PF_Export").Range("A" & i + 60, "I" & i + 60).Value = ""
    End If
    
Next i

newWorkbook

End Sub






