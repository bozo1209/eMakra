Attribute VB_Name = "enett"
Option Explicit
Dim nazwaExcelaDocelowego As String
Dim daneDoSortowania As Variant


'wpisywanie porównania z ostatni¹ wartoœci¹

Function znajdowanieWczorajszejDaty()
    Dim wczorajszaData As Date
    Dim i, j As Integer
    wczorajszaData = Date - 1
    
    With Sheets("EUR_VAN - GWTTP")
        For i = .Range("a1").End(xlDown).Row To .Range("a1").Row Step -1
            If .Cells(i, 1) = wczorajszaData Then
                j = i
                Exit For
            End If
        Next i
    End With
    znajdowanieWczorajszejDaty = j
End Function

Sub uzupelnianieCheka()
    Dim i As Integer
    i = znajdowanieWczorajszejDaty
    
    Sheets("EUR_VAN - GWTTP").Range("p" & i).FormulaR1C1 = "=RC[-1]-'Activity_Ledger EUR'!R" & Sheets("Activity_Ledger EUR").Range("h1").End(xlDown).Row & "C8"
    Sheets("USD_VAN - GWTTP").Range("p" & i).FormulaR1C1 = "=RC[-1]-'Activity_Ledger USD'!R" & Sheets("Activity_Ledger USD").Range("h1").End(xlDown).Row & "C8"
    Sheets("GBP_VAN - GWTTP").Range("p" & i).FormulaR1C1 = "=RC[-1]-'Activity_Ledger GBP'!R" & Sheets("Activity_Ledger GBP").Range("h1").End(xlDown).Row & "C8"
    Sheets("PLN_VAN - GWTTP").Range("p" & i).FormulaR1C1 = "=RC[-1]-'Activity_Ledger PLN'!R" & Sheets("Activity_Ledger PLN").Range("h1").End(xlDown).Row & "C8"
    Sheets("HUF_VAN - GWTTP").Range("p" & i).FormulaR1C1 = "=RC[-1]-'Activity_Ledger HUF'!R" & Sheets("Activity_Ledger HUF").Range("h1").End(xlDown).Row & "C8"
    Sheets("RUB_VAN - GWTTP").Range("p" & i).FormulaR1C1 = "=RC[-1]-'Activity_Ledger RUB'!R" & Sheets("Activity_Ledger RUB").Range("h1").End(xlDown).Row & "C8"
    Sheets("HKD_VAN - GWTTP (Asia)").Range("p" & i).FormulaR1C1 = "=RC[-1]-'Activity_Ledger HKD'!R" & Sheets("Activity_Ledger HKD").Range("h1").End(xlDown).Row & "C8"
End Sub

'sortowanie od ostatniej daty do najnowszej

Private Sub sortowanieDoNajnowszej(wb As Workbook)
    Dim temp As Variant
    Dim i, j As Integer
    Dim workbookName As String
    
    With wb.Sheets(1)
        If (.Range("a2").Value <> "" And .Range("a3").Value <> "") Then
            daneDoSortowania = .Range("a2:i" & .Range("a2").End(xlDown).Row).Value
        ElseIf (.Range("a2").Value <> "") Then
            daneDoSortowania = .Range("a2:i2").Value
        End If
    End With
    
    For i = 1 To (UBound(daneDoSortowania, 1) + 1) / 2
        For j = 1 To UBound(daneDoSortowania, 2)
            temp = daneDoSortowania(i, j)
            daneDoSortowania(i, j) = daneDoSortowania(UBound(daneDoSortowania, 1) + 1 - i, j)
            daneDoSortowania(UBound(daneDoSortowania, 1) + 1 - i, j) = temp
        Next j
    Next i

    workbookName = Left(wb.Name, InStr(1, wb.Name, ".", vbTextCompare) - 1)
    With Workbooks(nazwaExcelaDocelowego)
        Select Case workbookName
            Case "EUR"
                Call wklejanie("Activity_Ledger EUR")
                .Sheets("Activity_Ledger EUR").PivotTables("PivotTable4").PivotCache.Refresh
            Case "GBP"
                Call wklejanie("Activity_Ledger GBP")
                .Sheets("Activity_Ledger GBP").PivotTables("PivotTable7").PivotCache.Refresh
            Case "HKD"
                Call wklejanie("Activity_Ledger HKD")
                .Sheets("Activity_Ledger HKD").PivotTables("PivotTable3").PivotCache.Refresh
            Case "HUF"
                Call wklejanie("Activity_Ledger HUF")
                .Sheets("Activity_Ledger HUF").PivotTables("PivotTable9").PivotCache.Refresh
            Case "PLN"
                Call wklejanie("Activity_Ledger PLN")
                .Sheets("Activity_Ledger PLN").PivotTables("PivotTable8").PivotCache.Refresh
            Case "RUB"
                Call wklejanie("Activity_Ledger RUB")
                .Sheets("Activity_Ledger RUB").PivotTables("PivotTable2").PivotCache.Refresh
            Case "USD"
                Call wklejanie("Activity_Ledger USD")
                .Sheets("Activity_Ledger USD").PivotTables("PivotTable6").PivotCache.Refresh
            Case "VAN"
                Call wklejanie("VANS")
                .Sheets("VAN_Pivot").PivotTables("PivotTable2").PivotCache.Refresh
        End Select
    End With
    
End Sub

Private Sub wklejanie(nazwaArkusza As String)
    With Workbooks(nazwaExcelaDocelowego).Sheets(nazwaArkusza)
        If (.Range("a1").Value <> "" And .Range("a2").Value <> "") Then
            .Range("a" & .Range("a1").End(xlDown).Row + 1 & ":i" & .Range("a1").End(xlDown).Row + UBound(daneDoSortowania, 1)) = daneDoSortowania
        Else
            .Range("a2:i" & UBound(daneDoSortowania, 1)) = daneDoSortowania
        End If
    End With
End Sub

Sub wklejenieDoEnetta()
    Dim wb As Workbook
    nazwaExcelaDocelowego = "eNett 02.2021 — kopia.xlsb"
    
    For Each wb In Workbooks
        If wb.Name <> nazwaExcelaDocelowego And wb.Name <> "PERSONAL.XLSB" Then
            Call sortowanieDoNajnowszej(wb)
            wb.Close savechanges:=False
        End If
    Next wb
    Call uzupelnianieCheka
End Sub

