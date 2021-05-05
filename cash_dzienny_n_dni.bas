Attribute VB_Name = "cash_dzienny_n_dni"
Option Explicit
Dim data As String
Dim dzien, Miesiac, Rok As Integer
Dim ileDnii As Integer
Dim kolDlaBanku As Integer


Sub dodWszystko()
    Call grupowanie
    Call dodKol
    Call wykresy
    Call dlaBanku
    Call podsumowanieWPLN
    Call podsumowanieWWalutach
    Call podsumowaniePerBank
    Call dodawanie
End Sub

Private Sub ileDni()
    ileDnii = 31
End Sub

Private Sub dataSub(dzien As Integer)

    data = Format(dzien & "-05-2021", "mm-dd-yyyy")

End Sub


Private Sub grupowanie()
    Dim i As Integer
    For i = 1 To Sheets.Count
        Sheets(i).Outline.ShowLevels RowLevels:=2, ColumnLevels:=2
    Next i
End Sub




Private Sub dodKol()
    Call ileDni
        
    With Sheets("Podsumowanie w PLN")
        .Range(.Cells(4, .Range("c4").End(xlToRight).Column + 2), .Cells(120, .Range("c4").End(xlToRight).Column + ileDnii + 1)).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
    
    With Sheets("dla banku")
        kolDlaBanku = .Range("d8").End(xlToRight).Column
        
        .Range(.Cells(8, .Range("d8").End(xlToRight).Column + 1), .Cells(188, .Range("d8").End(xlToRight).Column + ileDnii)).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
End Sub

Private Sub dodawanie()
    Call ileDni
    Dim x, dzien As Integer
    For x = 1 To ileDnii
        dzien = x
        Call dataSub(dzien)
        dodawanieDetale
        dodawanieDostepnyKredytIZalegle
    Next x
End Sub

Private Sub dodawanieDetale()

    Dim kolumny(3 To 210, 1 To 3) As Variant

    Dim i, j As Integer
    
    For j = 1 To 3
        For i = 3 To UBound(kolumny, 1)

            If (i = 3 Or i = 131 Or i = 152 Or i = 170 Or i = 185 Or i = 207) And j = 3 Then
                kolumny(i, j) = data
            ElseIf i = 4 Or i = 132 Or i = 153 Or i = 171 Or i = 186 Or i = 208 Then

                Select Case j
                    Case 1
                        kolumny(i, j) = "Wartoœæ w walucie"
                    Case 2
                        kolumny(i, j) = "kurs"
                    Case 3
                        kolumny(i, j) = "w PLN"
                End Select

            ElseIf i = 124 Then
                Select Case j
                    Case 1
                        kolumny(i, j) = "=SUM(R[-119]C:R[-1]C)"
                    Case 3
                        kolumny(i, j) = "=SUM(R[-119]C:R[-1]C)"
                End Select
            ElseIf i = 147 Then

                Select Case j
                    Case 1
                        kolumny(i, j) = "=SUM(R[-14]C:R[-1]C)"
                    Case 3
                        kolumny(i, j) = "=SUM(R[-14]C:R[-1]C)"
                End Select

            ElseIf i = 160 Then

                Select Case j
                    Case 1
                        kolumny(i, j) = "=SUM(R[-6]C:R[-1]C)"
                    Case 3
                        kolumny(i, j) = "=SUM(R[-6]C:R[-1]C)"
                End Select

            ElseIf i = 179 Then

                Select Case j
                    Case 1
                        kolumny(i, j) = "=SUM(R[-7]C:R[-1]C)"
                    Case 3
                        kolumny(i, j) = "=SUM(R[-7]C:R[-1]C)"
                End Select

            ElseIf i = 195 Then

                Select Case j
                    Case 1
                        kolumny(i, j) = "=SUM(R[-8]C:R[-1]C)"
                    Case 3
                        kolumny(i, j) = "=SUM(R[-8]C:R[-1]C)"
                End Select

            ElseIf i = 165 And j = 3 Then
                kolumny(i, j) = "=R[-41]C+R[-18]C+R[-5]C"
            ElseIf i = 198 And j = 3 Then
                kolumny(i, j) = "=R[-33]C+R[-19]C+R[-3]C"
            ElseIf i > 4 And i < 124 Then
                Select Case j
                    Case 2
                        Select Case i
                            Case Is = 34, Is = 50, Is = 70, Is = 85

                                kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)/100"
                            Case Else
                                kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)"
                        End Select
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select

            ElseIf i > 132 And i < 147 Then

                Select Case j
                    Case 2
                        kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)"
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select

            ElseIf i > 153 And i < 160 Then

                Select Case j
                    Case 2
                        kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)"
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select

            ElseIf i > 171 And i < 179 Then
                Select Case j
                    Case 2
                        Select Case i
                            Case 176

                                kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)/100"
                            Case Else
                                kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)"
                        End Select
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select

            ElseIf i > 186 And i < 195 Then
                Select Case j
                    Case 2
                        Select Case i
                            Case 191

                                kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)/100"
                            Case Else
                                kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)"
                        End Select
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select

            ElseIf i > 208 Then

                Select Case j
                    Case 2
                        kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R3C[-2],kursy!R1C1:R400C1,1),0),1)"
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select
            End If
        Next i
    Next j
    With Sheets("detale")
        .Range(.Cells(3, .Range("F4").End(xlToRight).Column + 1), .Cells(UBound(kolumny, 1), .Range("F4").End(xlToRight).Column + 3)).FormulaR1C1 = kolumny
    End With
End Sub

Private Sub dodawanieDostepnyKredytIZalegle()
    Dim kolumny(6 To 22, 1 To 3) As Variant
    Dim i, j As Integer
    
    For j = 1 To 3
        For i = 6 To UBound(kolumny, 1)
            If (i = 6 Or i = 17) And j = 3 Then
                kolumny(i, j) = data
            ElseIf i = 7 Or i = 18 Then
                Select Case j
                    Case 1
                        kolumny(i, j) = "Wartoœæ w walucie"
                    Case 2
                        kolumny(i, j) = "kurs"
                    Case 3
                        kolumny(i, j) = "w PLN"
                End Select
            ElseIf i > 7 And i < 10 Then
                Select Case j
                    Case 2
                        kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R6C[-2],kursy!R1C1:R400C1,1),0),1)"
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select
            ElseIf i > 18 And i < 23 Then
                Select Case j
                    Case 2
                        kolumny(i, j) = "=IFERROR(HLOOKUP(RC5,kursy!R1:R400,MATCH(R6C[-2],kursy!R1C1:R400C1,1),0),1)"
                    Case 3
                        kolumny(i, j) = "=RC[-2]*RC[-1]"
                End Select
            End If
        Next i
    Next j
    With Sheets("dostêpny kredyt i zaleg³e p³")
        .Range(.Cells(6, .Range("F7").End(xlToRight).Column + 1), .Cells(UBound(kolumny, 1), .Range("F7").End(xlToRight).Column + 3)).FormulaR1C1 = kolumny
    End With
End Sub

Private Sub podsumowaniePerBank()
    Call ileDni
    Dim i, j, k, dzien As Integer
    
    Dim kolumny As Variant
    ReDim kolumny(3 To 54, 1 To ileDnii)
    
    k = Sheets("detale").Range("F4").End(xlToRight).Column
    For j = 1 To ileDnii
        k = k + 3
        dzien = j
        Call dataSub(dzien)
        For i = 3 To UBound(kolumny, 1)
            If i = 3 Or i = 17 Or i = 23 Or i = 30 Or i = 36 Or i = 43 Or i = 51 Then
                kolumny(i, j) = data
            ElseIf i > 3 And i < 13 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie per bank'!R4C1:R12C1,detale!C3,'Podsumowanie per bank'!RC4)"
            ElseIf i = 13 Then
                kolumny(i, j) = "=SUM(R[-9]C:R[-1]C)"
            ElseIf i = 14 Then
                kolumny(i, j) = "=(SUMIFS(detale!R26C" & k & ":R80C" & k & ",detale!R26C4:R80C4,'Podsumowanie per bank'!R11C1)-'Podsumowanie per bank'!R[-1]C)"
            ElseIf i = 18 Then
                kolumny(i, j) = "=SUMIFS(detale!R81C" & k & ":R94C" & k & ",detale!R81C4:R94C4,'Podsumowanie per bank'!R18C3,detale!R81C3:R94C3,'Podsumowanie per bank'!R18C4)"
            ElseIf i = 19 Then
                kolumny(i, j) = "=SUM(R[-1]C)"
            ElseIf i = 20 Then
                kolumny(i, j) = "=SUM(detale!R81C" & k & ":R94C" & k & ")-'Podsumowanie per bank'!R[-1]C"
            ElseIf i = 24 Then
                kolumny(i, j) = "=SUMIFS(detale!R95C" & k & ",detale!R95C4,'Podsumowanie per bank'!R24C3,detale!R95C3,'Podsumowanie per bank'!R24C4)"
            ElseIf i = 25 Then
                kolumny(i, j) = "=SUM(R[-1]C)"
            ElseIf i = 26 Then
                kolumny(i, j) = "=detale!R95C" & k & "-'Podsumowanie per bank'!R[-1]C"
            ElseIf i = 31 Then
                kolumny(i, j) = "=SUMIFS(detale!R96C" & k & ":R113C" & k & ",detale!R96C4:R113C4,'Podsumowanie per bank'!R31C3,detale!R96C3:R113C3,'Podsumowanie per bank'!R31C4)"
            ElseIf i = 32 Then
                kolumny(i, j) = "=SUM(R[-1]C)"
            ElseIf i = 33 Then
                kolumny(i, j) = "=(SUMIFS(detale!R96C" & k & ":R113C" & k & ",detale!R96C4:R113C4,'Podsumowanie per bank'!R31C3,detale!R96C3:R113C3,'Podsumowanie per bank'!R31C4))-'Podsumowanie per bank'!R[-1]C"
            ElseIf i = 37 Then
                kolumny(i, j) = "=SUMIFS(detale!R114C" & k & ":R115C" & k & ",detale!R114C4:R115C4,'Podsumowanie per bank'!R37C3,detale!R114C3:R115C3,'Podsumowanie per bank'!R37C4)"
            ElseIf i = 38 Then
                kolumny(i, j) = "=SUM(R[-1]C)"
            ElseIf i = 39 Then
                kolumny(i, j) = "=(SUMIFS(detale!R114C" & k & ":R115C" & k & ",detale!R114C4:R115C4,'Podsumowanie per bank'!R37C3,detale!R114C3:R115C3,'Podsumowanie per bank'!R37C4))-'Podsumowanie per bank'!R[-1]C"
            ElseIf i > 43 And i < 46 Then
                kolumny(i, j) = "=SUMIFS(detale!R116C" & k & ":R121C" & k & ",detale!R116C4:R121C4,'Podsumowanie per bank'!RC1,detale!R116C3:R121C3,'Podsumowanie per bank'!RC4)"
            ElseIf i = 46 Then
                kolumny(i, j) = "=SUM(R[-2]C:R[-1]C)"
            ElseIf i = 47 Then
                kolumny(i, j) = "=(SUMIFS(detale!R116C" & k & ":R121C" & k & ",detale!R116C4:R121C4,'Podsumowanie per bank'!R45C1))-'Podsumowanie per bank'!R[-1]C"
            ElseIf i = 52 Then
                kolumny(i, j) = "=SUMIFS(detale!R122C" & k & ",detale!R122C4,'Podsumowanie per bank'!R52C3,detale!R122C3,'Podsumowanie per bank'!R52C4)"
            ElseIf i = 53 Then
                kolumny(i, j) = "=SUM(R[-1]C)"
            ElseIf i = 54 Then
                kolumny(i, j) = "=detale!R122C" & k & "-'Podsumowanie per bank'!R[-1]C"
            End If
        Next i
    Next j
    With Sheets("Podsumowanie per bank")
        .Range(.Cells(3, .Range("d3").End(xlToRight).Column + 1), .Cells(UBound(kolumny, 1), .Range("d3").End(xlToRight).Column + ileDnii)).FormulaR1C1 = kolumny
    End With
End Sub

Private Sub podsumowanieWWalutach()
    Call ileDni
    Dim i, j, k, dzien As Integer
    
    Dim kolumny As Variant

    ReDim kolumny(4 To 96, 1 To ileDnii)

    
    k = Sheets("detale").Range("F4").End(xlToRight).Column - 2
    For j = 1 To ileDnii
        k = k + 3
        dzien = j
        Call dataSub(dzien)
        For i = 4 To UBound(kolumny, 1)

            If i = 4 Or i = 32 Or i = 51 Or i = 66 Or i = 82 Then
                kolumny(i, j) = data
            ElseIf i > 4 And i < 24 Then
                kolumny(i, j) = "=SUMIFS(detale!R5C" & k & ":R123C" & k & ",detale!R5C5:R123C5,'Podsumowanie w walutach'!RC3)"
            ElseIf i = 24 Then
                kolumny(i, j) = "=SUM(R[-19]C:R[-1]C)"
            ElseIf i = 25 Then
                kolumny(i, j) = "=detale!R124C" & k & "-'Podsumowanie w walutach'!R[-1]C"
            ElseIf i > 32 And i < 47 Then
                kolumny(i, j) = "=SUMIFS(detale!R133C" & k & ":R146C" & k & ",detale!R133C5:R146C5,'Podsumowanie w walutach'!RC3)"
            ElseIf i = 47 Then
                kolumny(i, j) = "=SUM(R[-14]C:R[-1]C)"
            ElseIf i = 48 Then
                kolumny(i, j) = "=detale!R147C" & k & "-'Podsumowanie w walutach'!R[-1]C"
            ElseIf i > 51 And i < 58 Then
                kolumny(i, j) = "=SUMIFS(detale!R154C" & k & ":R159C" & k & ",detale!R154C5:R159C5,'Podsumowanie w walutach'!RC3)"
            ElseIf i = 58 Then
                kolumny(i, j) = "=SUM(R[-6]C:R[-1]C)"
            ElseIf i = 59 Then
                kolumny(i, j) = "=detale!R160C" & k & "-'Podsumowanie w walutach'!R[-1]C"
            ElseIf i > 66 And i < 74 Then
                kolumny(i, j) = "=SUMIFS(detale!R172C" & k & ":R178C" & k & ",detale!R172C5:R178C5,'Podsumowanie w walutach'!RC3)"
            ElseIf i = 74 Then
                kolumny(i, j) = "=SUM(R[-7]C:R[-1]C)"
            ElseIf i = 75 Then
                kolumny(i, j) = "=detale!R179C" & k & "-'Podsumowanie w walutach'!R[-1]C"
            ElseIf i > 82 And i < 91 Then
                kolumny(i, j) = "=SUMIFS(detale!R187C" & k & ":R194C" & k & ",detale!R187C5:R194C5,'Podsumowanie w walutach'!RC3)"
            ElseIf i = 91 Then
                kolumny(i, j) = "=SUM(R[-8]C:R[-1]C)"
            ElseIf i = 92 Then
                kolumny(i, j) = "=detale!R195C" & k & "-'Podsumowanie w walutach'!R[-1]C"
            ElseIf i = 96 Then

                kolumny(i, j) = "=R[-72]C+R[-49]C+R[-38]C+R[-22]C+R[-5]C"
            End If
        Next i
    Next j
    With Sheets("Podsumowanie w walutach")
        .Range(.Cells(4, .Range("c4").End(xlToRight).Column + 1), .Cells(UBound(kolumny, 1), .Range("c4").End(xlToRight).Column + ileDnii)).FormulaR1C1 = kolumny
    End With
End Sub

Private Sub podsumowanieWPLN()
    Call ileDni
    Dim i, j, k, m, dzien As Integer
    
    Dim kolumny As Variant

    ReDim kolumny(4 To 121, 1 To ileDnii)

    
    k = Sheets("detale").Range("F4").End(xlToRight).Column
    m = Sheets("Podsumowanie per bank").Range("D3").End(xlToRight).Column
    For j = 1 To ileDnii
        k = k + 3
        m = m + 1
        dzien = j
        Call dataSub(dzien)
        For i = 4 To UBound(kolumny, 1)

            If i = 4 Or i = 32 Or i = 51 Or i = 66 Or i = 82 Then
                kolumny(i, j) = data
            ElseIf i > 4 And i < 24 Then
                kolumny(i, j) = "=SUMIFS(detale!R5C" & k & ":R123C" & k & ",detale!R5C5:R123C5,'Podsumowanie w PLN'!RC3)"
            ElseIf i = 24 Then
                kolumny(i, j) = "=SUM(R[-19]C:R[-1]C)"
            ElseIf i = 25 Then
                kolumny(i, j) = "=detale!R124C" & k & "-'Podsumowanie w PLN'!R[-1]C"
            ElseIf i > 32 And i < 47 Then
                kolumny(i, j) = "=SUMIFS(detale!R133C" & k & ":R146C" & k & ",detale!R133C5:R146C5,'Podsumowanie w PLN'!RC3)"
            ElseIf i = 47 Then
                kolumny(i, j) = "=SUM(R[-14]C:R[-1]C)"
            ElseIf i = 48 Then
                kolumny(i, j) = "=detale!R147C" & k & "-'Podsumowanie w PLN'!R[-1]C"
            ElseIf i > 51 And i < 58 Then
                kolumny(i, j) = "=SUMIFS(detale!R154C" & k & ":R159C" & k & ",detale!R154C5:R159C5,'Podsumowanie w PLN'!RC3)"
            ElseIf i = 58 Then
                kolumny(i, j) = "=SUM(R[-6]C:R[-1]C)"
            ElseIf i = 59 Then
                kolumny(i, j) = "=detale!R160C" & k & "-'Podsumowanie w PLN'!R[-1]C"
            ElseIf i > 66 And i < 74 Then
                kolumny(i, j) = "=SUMIFS(detale!R172C" & k & ":R178C" & k & ",detale!R172C5:R178C5,'Podsumowanie w PLN'!RC3)"
            ElseIf i = 74 Then
                kolumny(i, j) = "=SUM(R[-7]C:R[-1]C)"
            ElseIf i = 75 Then
                kolumny(i, j) = "=detale!R179C" & k & "-'Podsumowanie w PLN'!R[-1]C"
            ElseIf i > 82 And i < 91 Then
                kolumny(i, j) = "=SUMIFS(detale!R187C" & k & ":R194C" & k & ",detale!R187C5:R194C5,'Podsumowanie w PLN'!RC3)"
            ElseIf i = 91 Then
                kolumny(i, j) = "=SUM(R[-8]C:R[-1]C)"
            ElseIf i = 92 Then
                kolumny(i, j) = "=detale!R195C" & k & "-'Podsumowanie w PLN'!R[-1]C"
            ElseIf i = 97 Then
                kolumny(i, j) = "=R[-73]C+R[-50]C+R[-39]C+R[-23]C+R[-6]C"
            ElseIf i = 98 Then
                kolumny(i, j) = "=R[-1]C-R[3]C"
            ElseIf i = 100 Then
                kolumny(i, j) = "=R[-26]C+R[-53]C+'Podsumowanie per bank'!R13C" & m & "-R[1]C+R[-9]C"
            ElseIf i = 101 Then
                kolumny(i, j) = "=detale!R63C" & k & "+detale!R64C" & k
            ElseIf i = 103 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R103C3)"
            ElseIf i = 105 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R105C3)"
            ElseIf i = 107 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R107C3)"
            ElseIf i = 109 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R109C3)"
            ElseIf i = 111 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R111C3)"
            ElseIf i = 113 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R113C3)"
            ElseIf i = 115 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R115C3)"
            ElseIf i = 117 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R117C3)"
            ElseIf i = 119 Then
                kolumny(i, j) = "=SUMIFS(detale!C" & k & ",detale!C4,'Podsumowanie w PLN'!R119C3)"
            ElseIf i = 121 Then

                kolumny(i, j) = "=SUM(R[-21]C,R[-18]C,R[-16]C,R[-14]C,R[-12]C,R[-10]C,R[-8]C,R[-6]C,R[-4]C,R[-2]C)-R[-24]C+R[-20]C"
            End If
        Next i
    Next j
        
    With Sheets("Podsumowanie w PLN")
        .Range(.Cells(4, .Range("c4").End(xlToRight).Column + 1), .Cells(UBound(kolumny, 1), .Range("c4").End(xlToRight).Column + ileDnii)).FormulaR1C1 = kolumny
    End With
End Sub

Private Sub wykresy()
    Call ileDni
    Dim i, j, k, dzien As Integer
    
    Dim kolumny As Variant
    ReDim kolumny(2 To 5, 1 To ileDnii)
    
    k = Sheets("dla banku").Range("D8").End(xlToRight).Column
    For j = 1 To ileDnii
        k = k + 1
        dzien = j
        Call dataSub(dzien)
        For i = 2 To UBound(kolumny, 1)
            Select Case i
                Case 2
                    kolumny(i, j) = data
                Case 3
                    kolumny(i, j) = "='dla banku'!R31C" & k
                Case 4
                    kolumny(i, j) = "='dla banku'!R32C" & k
                Case 5
                    kolumny(i, j) = "='dla banku'!R32C" & k & "+'dla banku'!R50C" & k & "+'dla banku'!R51C" & k & "-'dla banku'!R33C" & k & "-'dla banku'!R35C" & k & "-'dla banku'!R34C" & k
            End Select
        Next i
    Next j
    
    With Sheets("wykresy")
        .Range(.Cells(2, .Range("b2").End(xlToRight).Column + 1), .Cells(UBound(kolumny, 1), .Range("b2").End(xlToRight).Column + ileDnii)).FormulaR1C1 = kolumny
    End With
End Sub

Private Sub dlaBanku()
    Call ileDni
    Dim i, j, podWpln, dosKreZal, det, dzien As Integer
    
    Dim kolumny As Variant
    ReDim kolumny(8 To 188, 1 To ileDnii)
    
    podWpln = Sheets("Podsumowanie w PLN").Range("c4").End(xlToRight).Column
    dosKreZal = Sheets("dostêpny kredyt i zaleg³e p³").Range("f7").End(xlToRight).Column
    det = Sheets("detale").Range("F4").End(xlToRight).Column
    For j = 1 To ileDnii
        kolDlaBanku = kolDlaBanku + 1
        podWpln = podWpln + 1
        dosKreZal = dosKreZal + 3
        det = det + 3
        dzien = j
        Call dataSub(dzien)
        For i = 8 To UBound(kolumny, 1)
            If i = 8 Or i = 48 Or i = 57 Or i = 67 Or i = 98 Or i = 130 Or i = 162 Then
                kolumny(i, j) = data
            ElseIf i > 8 And i < 31 Then
                Select Case i
                    Case Is = 25
                        kolumny(i, j) = "=SUMIFS('Podsumowanie w PLN'!C" & podWpln & ",'Podsumowanie w PLN'!C3,'dla banku'!RC4)-IF(R49C" & kolDlaBanku & "<0,R49C" & kolDlaBanku & ",0)"
                    Case Else
                        kolumny(i, j) = "=SUMIFS('Podsumowanie w PLN'!C" & podWpln & ",'Podsumowanie w PLN'!C3,'dla banku'!RC4)"
                End Select
            ElseIf i = 31 Then
                kolumny(i, j) = "=SUM(R[-22]C:R[-1]C)"
            ElseIf i = 32 Then
                kolumny(i, j) = "=R[-1]C-R[18]C-R[19]C+IF(R[17]C>0,0,R[17]C)"
            ElseIf i = 35 Then
                kolumny(i, j) = "=SUM(R[1]C:R[4]C)"
            ElseIf i = 36 Then
                kolumny(i, j) = "='dostêpny kredyt i zaleg³e p³'!R19C" & dosKreZal
            ElseIf i = 37 Then
                kolumny(i, j) = "='dostêpny kredyt i zaleg³e p³'!R22C" & dosKreZal
            ElseIf i = 38 Then
                kolumny(i, j) = "='dostêpny kredyt i zaleg³e p³'!R20C" & dosKreZal
            ElseIf i = 39 Then
                kolumny(i, j) = "='dostêpny kredyt i zaleg³e p³'!R21C" & dosKreZal

            ElseIf i = 40 Then
                kolumny(i, j) = "=R[-9]C-R[-5]C-R[-7]C+R[1]C-R[-6]C"
            ElseIf i = 43 Then
                kolumny(i, j) = "=detale!R198C" & det & "-'dla banku'!R[-12]C-R[6]C"
            ElseIf i = 44 Then
                kolumny(i, j) = "='Podsumowanie w PLN'!R79C" & podWpln & "-R[-13]C"
            ElseIf i = 49 Then
                kolumny(i, j) = "=detale!R63C" & det & "+detale!R64C" & det

            ElseIf i = 50 Then
                kolumny(i, j) = "='dostêpny kredyt i zaleg³e p³'!R8C" & dosKreZal
            ElseIf i = 51 Then
                kolumny(i, j) = "='dostêpny kredyt i zaleg³e p³'!R9C" & dosKreZal
            ElseIf i = 53 Then
                kolumny(i, j) = "=R49C" & kolDlaBanku & "-R50C" & kolDlaBanku & "=R32C" & kolDlaBanku & "-R31C" & kolDlaBanku
            ElseIf i > 67 And i < 90 Then
                Select Case i
                    Case 84

                        kolumny(i, j) = "=SUMIFS('Podsumowanie w PLN'!R5C" & podWpln & ":R23C" & podWpln & ",'Podsumowanie w PLN'!R5C3:R23C3,'dla banku'!RC4)-R[-35]C"
                    Case Else
                        kolumny(i, j) = "=SUMIFS('Podsumowanie w PLN'!R5C" & podWpln & ":R23C" & podWpln & ",'Podsumowanie w PLN'!R5C3:R23C3,'dla banku'!RC4)"

                End Select
            ElseIf i = 90 Then
                kolumny(i, j) = "=SUM(R[-22]C:R[-1]C)"
            ElseIf i > 98 And i < 121 Then

                kolumny(i, j) = "=SUMIFS('Podsumowanie w PLN'!R67C" & podWpln & ":R73C" & podWpln & ",'Podsumowanie w PLN'!R67C3:R73C3,'dla banku'!RC4)"
            ElseIf i = 121 Then
                kolumny(i, j) = "=SUM(R[-22]C:R[-1]C)"
            ElseIf i > 130 And i < 153 Then
                kolumny(i, j) = "=SUMIFS('Podsumowanie w PLN'!R33C" & podWpln & ":R57C" & podWpln & ",'Podsumowanie w PLN'!R33C3:R57C3,'dla banku'!RC4)"
            ElseIf i = 153 Then
                kolumny(i, j) = "=SUM(R[-22]C:R[-1]C)"
            ElseIf i > 162 And i < 185 Then
                kolumny(i, j) = "=SUMIFS('Podsumowanie w PLN'!R83C" & podWpln & ":R90C" & podWpln & ",'Podsumowanie w PLN'!R83C3:R90C3,'dla banku'!RC4)"

            ElseIf i = 185 Then
                kolumny(i, j) = "=SUM(R[-22]C:R[-1]C)"
            ElseIf i = 188 Then
                kolumny(i, j) = "=R[-2]C+R[-34]C+R[-65]C-R[-124]C+R[-106]C"
            End If
        Next i
    Next j
    
    With Sheets("dla banku")
        .Range(.Cells(8, .Range("d8").End(xlToRight).Column + 1), .Cells(UBound(kolumny, 1), .Range("d8").End(xlToRight).Column + ileDnii)).FormulaR1C1 = kolumny
    End With
End Sub



Private Sub testRunTime()
    Dim startTime, endTime As Double
    startTime = Timer
    
    Call dodWszystko
    
    endTime = Timer - startTime
    Debug.Print "time: " & endTime
End Sub
