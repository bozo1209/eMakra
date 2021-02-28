Attribute VB_Name = "enett"
Option Explicit

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


