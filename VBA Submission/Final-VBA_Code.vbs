Sub AnalyzeStockData()
Application.ScreenUpdating = False


Dim ws As Worksheet
Dim year_opening, year_closing  As Single
Dim select_index, first_row, select_row, last_row As Double
Dim volume As Double


For Each ws In Sheets
    Worksheets(ws.Name).Activate
    select_index = 2
    select_row = 2
    first_row = 2
    last_row = WorksheetFunction.CountA(ActiveSheet.Columns(1))
    volume = 0

'Assign headers etc to columns

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


'testing auto fit
Columns("I:Q").EntireColumn.AutoFit

'Identify unique tickers and add them to new column

For i = first_row To last_row

    tickers = Cells(i, 1).Value
    
    tickers2 = Cells(i - 1, 1).Value
    
    If tickers <> tickers2 Then
    
    Cells(select_row, 9).Value = tickers

    select_row = select_row + 1

    End If
    
       Next i


'volume assignment

For i = first_row To last_row + 1

    tickers = Cells(i, 1).Value
    
    tickers2 = Cells(i - 1, 1).Value

        If tickers = tickers2 And i > 2 Then
        
        volume = volume + Cells(i, 7).Value

        ElseIf i > 2 Then
    
        Cells(select_index, 12).Value = volume

        select_index = select_index + 1

        volume = 0

        Else: volume = volume + Cells(i, 7).Value

        End If

    Next i
    

'opening and closing year assignment

select_index = 2

For i = first_row To last_row
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

    year_closing = Cells(i, 6).Value

        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then

        year_opening = Cells(i, 3).Value

    End If

    If year_opening > 0 And year_closing > 0 Then

        increase = year_closing - year_opening

        percent_increase = increase / year_opening
 
        Cells(select_index, 10).Value = FormatPercent(increase)

        Cells(select_index, 11).Value = FormatPercent(percent_increase)
    
        year_closing = 0

        year_opening = 0

        select_index = select_index + 1

        End If

    Next i

'find min and max values, then assign each value a proper cell

    max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))

    min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))

    max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))

    Range("Q2").Value = FormatPercent(max_per)
    
    Range("Q3").Value = FormatPercent(min_per)

    Range("Q4").Value = (max_vol)
    
    Columns("I:Q").EntireColumn.AutoFit

    

'minamun and maximum filter and assign to new colomn

    For i = first_row To last_row

    If max_per = Cells(i, 11).Value Then
    
    Range("P2").Value = Cells(i, 9).Value

        ElseIf min_per = Cells(i, 11).Value Then
    
        Range("P3").Value = Cells(i, 9).Value

        ElseIf max_vol = Cells(i, 12).Value Then
    
        Range("P4").Value = Cells(i, 9).Value

        End If

    Next i

'loops through column 10 then applies earlier green or red interior


    For i = first_row To last_row

    If IsEmpty(Cells(i, 10).Value) Then
    
    End If

    If Cells(i, 10).Value > 0 Then

    Cells(i, 10).Interior.ColorIndex = 10

    Else

    Cells(i, 10).Interior.ColorIndex = 3

    End If

    Next i
    
 'color format the percentagechange colomn
 
 For i = first_row To last_row

    If IsEmpty(Cells(i, 11).Value) Then
    
    End If

    If Cells(i, 11).Value > 0 Then

    Cells(i, 11).Interior.ColorIndex = 4

    Else

    Cells(i, 11).Interior.ColorIndex = 9

    End If

    Next i
    

    Next ws
    
    'Columns("I:Q").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    Columns("I:Q").EntireColumn.AutoFit
    
    'here it only applies to the last sheet
    'Columns("I:Q").EntireColumn.AutoFit

    End Sub









