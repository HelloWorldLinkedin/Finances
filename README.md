# Finances
Personal Finance Management System
Sub CalcExp()
'Current Variables Used
Dim Lastrow         As Long
Dim i               As Integer
Dim ws              As Integer
'Copy Values
Dim Calendar        As Date
Dim Value           As Double
Dim Previous        As Double
Dim Category        As Integer
Dim Description     As String
    
    Sheets("FPosition").Range("E47:E65").ClearContents      'Clear Reserves
    Sheets("Calculations").Range("B2:M19").ClearContents    'Chart Calculations
    Sheets("Stock").Range("A2:M200").ClearContents          'Chart Stocks
    Sheets("SNum").Range("A2:M200").ClearContents           'Chart Stocks Number
    Sheets("Salary").Range("A2:C80").ClearContents          'Salary
    Sheets("Salary").Range("F2:H80").ClearContents          'Dividend
    Sheets("Salary").Range("K2:M80").ClearContents          'Interest
    
    For ws = 2 To 7 'No longer has CBA
        
        'Reset Calculation Values
        Call ClearContent(ws)
        Value = 0
        
        Lastrow = Worksheets(ws).Range("C" & Rows.Count).End(xlUp).Row
        For i = 2 To Lastrow
            Category = rowV(Worksheets(ws).Cells(i, 3)) 'determine the value of category
            Value = Worksheets(ws).Cells(i, 4).Value
            Previous = 0
            
            'Calculating the account transactions
            If Category < 18 Then
                Previous = Worksheets(ws).Cells(Category + 1, 9).Value
                Worksheets(ws).Cells(Category + 1, 9).Value = Previous + Value
                If Category < 17 Then
                    Call ChartCal(Category, Worksheets(ws).Cells(i, 1), Value)
                End If
            Else
                Calendar = Worksheets(ws).Cells(i, 1).Value
                Description = Worksheets(ws).Cells(i, 2).Value
                If Category = 90 Then           'Dividend
                    Call DividendCal(Calendar, Description, Value, ws)
                ElseIf Category = 91 Then       'Interest
                    Call InterestCal(Calendar, Description, Value, ws)
                ElseIf Category = 92 Then       'Fee
                    Previous = Worksheets(ws).Cells(5, 16)
                    Worksheets(ws).Cells(5, 16) = Previous + Value
                ElseIf Category = 93 Then       'Transfers
                    Call TransferCal(Calendar, Description, Value, ws)
                ElseIf Category = 94 Then       'Pending
                    Worksheets(ws).Cells(i, 4).Interior.Color = RGB(200, 0, 0)
                    Call PendingCal(Calendar, Description, Value, ws)
                ElseIf Category = 95 Then       'Salary
                    Call SalaryCal(Calendar, Description, Value, ws)
                ElseIf Category = 96 Then       'Buy
                    Call Portfolio(Calendar, Description, Value, ws, "Buy", Worksheets(ws).Cells(i, 5).Value)
                ElseIf Category = 97 Then       'Sell
                    Call Portfolio(Calendar, Description, Value, ws, "Sell", Worksheets(ws).Cells(i, 5).Value)
                End If
            End If
        Next i
    Next ws
End Sub

Function Portfolio(Calendar As Date, Description As String, Value As Double, ws As Integer, trade As String, stocks As Integer)
    Dim Lastrow As Integer
    Dim Found As Boolean
    Dim Price As Double
    Dim PreviousFee As Double
    Dim NewTotal As Integer
    Dim i As Integer
    
    Lastrow = Sheets("Stock").Range("A" & Rows.Count).End(xlUp).Row

    MonthCol = CalMonthCol(Calendar)

    'Portfolio Analysis
    For i = 2 To Lastrow
        If StrComp(Sheets("Stock").Cells(i, 1), Description) = 0 Then
            Found = True
            If trade = "Buy" Then
                NewTotal = Sheets("SNum").Cells(i, MonthCol) + (Value - 19.99) / stocks
                Price = (Sheets("Stock").Cells(i, MonthCol) * Sheets("SNum").Cells(i, MonthCol) + (Value - 19.99)) / NewTotal
                Sheets("Stock").Cells(i, MonthCol) = Price
                Sheets("SNum").Cells(i, MonthCol) = Sheets("SNum").Cells(i, MonthCol) + stocks
            ElseIf trade = "Sell" Then
                Sheets("SNum").Cells(i, MonthCol) = Sheets("SNum").Cells(i, MonthCol) - stocks
            End If
            Exit For
        End If
    Next i
    If Found = False Then
        Sheets("Stock").Cells(Lastrow + 1, 1) = Description
        Sheets("SNum").Cells(Lastrow + 1, 1) = Description
        Sheets("Stock").Cells(Lastrow + 1, MonthCol) = (Value - 19.99) / stocks
        Sheets("SNum").Cells(Lastrow + 1, MonthCol) = stocks
        'add a new row to homepage
        Lastrow = Sheets("FPosition").Range("E" & Rows.Count).End(xlUp).Row
        Sheets("FPosition").Cells(Lastrow + 1, 5) = Description
    End If
    
    'Calculating Accumulated Value for Investment
    If trade = "Buy" Then
        Previous = Worksheets(ws).Cells(6, 15).Value
        Worksheets(ws).Cells(6, 15) = Previous + Value
    Else
        Previous = Worksheets(ws).Cells(6, 16).Value
        Worksheets(ws).Cells(6, 16) = Previous + Value
    End If
    
    'Adding on Brokeage Fee
    PreviousFee = Worksheets(ws).Cells(7, 16).Value
    Worksheets(ws).Cells(7, 16) = PreviousFee + 19.99
End Function

Function rowV(Value As String)
Dim result As Integer

    If Value = "Telecom" Then
        rowV = 2
    ElseIf Value = "Entertainment" Then
        rowV = 3
    ElseIf Value = "Medication" Then
        rowV = 4
    ElseIf Value = "Groceries" Then
        rowV = 5
    ElseIf Value = "Restaurant" Then
        rowV = 6
    ElseIf Value = "Transportation" Then
        rowV = 7
    ElseIf Value = "Withdrawal" Then
        rowV = 8
    ElseIf Value = "Clothing" Then
        rowV = 9
    ElseIf Value = "Technology" Then
        rowV = 10
    ElseIf Value = "Education" Then
        rowV = 11
    ElseIf Value = "Miscellaneous" Then
        rowV = 16
    ElseIf Value = "Dividend" Then
        rowV = 90
    ElseIf Value = "Interest" Then
        rowV = 91
    ElseIf Value = "Fee" Then
        rowV = 92
    ElseIf Value = "Transfer" Then
        rowV = 93
    ElseIf Value = "Pending" Then
        rowV = 94
    ElseIf Value = "Salary" Then
        rowV = 95
    ElseIf Value = "Buy" Then
        rowV = 96
    ElseIf Value = "Sell" Then
        rowV = 97
    Else 'Bank Transfer or ATO
        rowV = 17 'error check
    End If
    
End Function
Sub ClearContent(ws As Integer)
Dim Lastrow As Integer
       
    Worksheets(ws).Range("F21:K50").ClearContents           'Transfers
    Worksheets(ws).Range("M17:R23").ClearContents           'Pending
    Worksheets(ws).Columns(4).Interior.Color = xlNone       'Pending Colour
    Worksheets(ws).Range("I3:I20").ClearContents            'Expenses
    Worksheets(ws).Cells(3, 15).Value = 0                   'Income
    Worksheets(ws).Cells(4, 15).Value = 0                   'Interest
    Worksheets(ws).Cells(5, 15).Value = 0                   'Dividend
    Worksheets(ws).Cells(6, 15).Value = 0                   'Sell
    Worksheets(ws).Cells(6, 16).Value = 0                   'Buy
    Worksheets(ws).Cells(5, 16).Value = 0                   'Fee
    Worksheets(ws).Cells(7, 16).Value = 0                   'Brokeage Fee
    
End Sub
Function CalMonthCol(Calendar As Date) As Integer
    Dim CalendarM As Integer
    
    CalendarM = Month(Calendar)
    If CalendarM < 7 Then           'JAN - JUN
        CalMonthCol = CalendarM + 7
    ElseIf CalendarM < 13 Then    'JUL - DEC
        CalMonthCol = CalendarM - 5
    End If

End Function
Sub ChartCal(MonthRow As Integer, Calendar As Date, Value As Double)

    MonthCol = CalMonthCol(Calendar)

    Sheets("Calculations").Cells(MonthRow, MonthCol) = Sheets("Calculations").Cells(MonthRow, MonthCol) + Value * -1
End Sub
Sub InterestCal(Calendar As Date, Description As String, Value As Double, ws As Integer)
Dim Lastrow     As Integer
Dim Previous    As Double

    Lastrow = Sheets("Salary").Range("K" & Rows.Count).End(xlUp).Row
    
    Sheets("Salary").Cells(Lastrow + 1, 11) = Calendar
    Sheets("Salary").Cells(Lastrow + 1, 12) = Description
    Sheets("Salary").Cells(Lastrow + 1, 13) = Value
    
    Previous = Worksheets(ws).Cells(4, 15).Value
    Worksheets(ws).Cells(4, 15) = Previous + Value
End Sub
Sub DividendCal(Calendar As Date, Description As String, Value As Double, ws As Integer)
Dim Lastrow         As Integer
Dim Previous        As Double

    Lastrow = Sheets("Salary").Range("F" & Rows.Count).End(xlUp).Row
    
    Sheets("Salary").Cells(Lastrow + 1, 6) = Calendar
    Sheets("Salary").Cells(Lastrow + 1, 7) = Description
    Sheets("Salary").Cells(Lastrow + 1, 8) = Value
    
    Previous = Worksheets(ws).Cells(5, 15)
    Worksheets(ws).Cells(5, 15) = Previous + Value
End Sub
Sub SalaryCal(Calendar As Date, Description As String, Value As Double, ws As Integer)
Dim Lastrow     As Integer
Dim Previous    As Double
    
    Lastrow = Sheets("Salary").Range("A" & Rows.Count).End(xlUp).Row
    
    Sheets("Salary").Cells(Lastrow + 1, 1) = Calendar
    Sheets("Salary").Cells(Lastrow + 1, 2) = Description
    Sheets("Salary").Cells(Lastrow + 1, 3) = Value
    
    Previous = Worksheets(ws).Cells(3, 15).Value
    Worksheets(ws).Cells(3, 15) = Previous + Value
End Sub
Sub TransferCal(Calendar As Date, Description As String, Value As Double, ws As Integer)
Dim Lastrow     As Integer

    If Value > 0 Then
        Lastrow = Worksheets(ws).Range("F" & Rows.Count).End(xlUp).Row
        Worksheets(ws).Cells(Lastrow + 1, 6) = Calendar
        Worksheets(ws).Cells(Lastrow + 1, 7) = Description
        Worksheets(ws).Cells(Lastrow + 1, 8) = Value
    Else
        Lastrow = Worksheets(ws).Range("J" & Rows.Count).End(xlUp).Row
        Worksheets(ws).Cells(Lastrow + 1, 11) = Calendar
        Worksheets(ws).Cells(Lastrow + 1, 10) = Description
        Worksheets(ws).Cells(Lastrow + 1, 9) = Value
    End If
End Sub
Sub PendingCal(Calendar As Date, Description As String, Value As Double, ws As Integer)
Dim Lastrow     As Integer

    If Value > 0 Then
        Lastrow = Worksheets(ws).Range("M" & Rows.Count).End(xlUp).Row
        Worksheets(ws).Cells(Lastrow + 1, 13) = Calendar
        Worksheets(ws).Cells(Lastrow + 1, 14) = Description
        Worksheets(ws).Cells(Lastrow + 1, 15) = Value
    Else
        Lastrow = Worksheets(ws).Range("Q" & Rows.Count).End(xlUp).Row
        Worksheets(ws).Cells(Lastrow + 1, 18) = Calendar
        Worksheets(ws).Cells(Lastrow + 1, 17) = Description
        Worksheets(ws).Cells(Lastrow + 1, 16) = Value
    End If
End Sub
Function TestDates(pDate1 As Date) As Long
    
    pDate2 = Format(Now(), "dd/mm/yyyy")
    'If pDate1 Is Nothing Then
    If pDate1 = 0 Then
        TestDates = 0
    Else
        TestDates = DateDiff("ww", pDate1, pDate2)
    End If

End Function
