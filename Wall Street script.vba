' Stock data extraction VBA
' Written by Bryan Paynich
' 11/28/2018

' This VBA script processes stock data, aggregates the volume,
' yearly change, and percentage gain/loss for each stock symbol provided.
' The script also finds the highest volume stock, highest performance stock,
' and lowest performing stock in the list provided.

Sub RUN_ALL_SHEETS()
'
' RUNS MACROS ON EACH SHEET
Sheets("2016").Select
Call STOCK_DATA_EXTRACTION
Sheets("2015").Select
Call STOCK_DATA_EXTRACTION
Sheets("2014").Select
Call STOCK_DATA_EXTRACTION
End Sub

Sub STOCK_DATA_EXTRACTION()

'SET VARIABLES
Dim temp_ticker As String
Dim current_ticker As String
Dim temp_vol As Double
Dim vol_total As Double
Dim open_value As Double
Dim close_value As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim outcnt As Long
Dim i As Long
Dim first_item As Integer
Dim LastCol As Integer

'FIND LAST ROW
With ActiveSheet
    ColRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

'INITALIZE COUNTERS
first_item = 0
vol_total = 0
outcnt = 2
    
'PRINT HEADERS
 Range("I1") = "Ticker"
 Range("J1") = "Yearly Change"
 Range("K1") = "Percentage Change"
 Range("L1") = "Total Stock Volume"

'LOOPING THROUGH DATA
For i = 2 To ColRow
   
'FIRST ITEM INSERTION
If first_item = 0 Then
    current_ticker = Cells(2, 1)
    Range("I2") = Cells(2, 1)
    open_value = Cells(i, 3)
    first_item = 1
End If
'SET FIRST ITEM IN LIST TO TEMP VARIABLE FOR INTERATIVE COMPARISON
temp_ticker = Cells(i, 1)
temp_vol = Cells(i, 7)

'IF/ELSE LOGIC TO SUMMARIZE VOLUME TOTAL AND CALCULATE DIFFERENCE IN OPEN/CLOSE VALUES
If (temp_ticker = current_ticker) Then
    vol_total = vol_total + temp_vol
    close_value = Cells(i, 6)
Else
    Cells(outcnt, 9) = current_ticker
    Cells(outcnt, 12) = vol_total
    yearly_change = close_value - open_value
    Cells(outcnt, 10) = yearly_change
    
    If (close_value = 0 And open_value = 0) Then
        percentage_change = 0
    ElseIf (close_value = 0) Then
        percentage_change = 0
    Else
        percentage_change = (close_value - open_value) / close_value
    End If
    
    Cells(outcnt, 11) = percentage_change
    current_ticker = temp_ticker
    vol_total = 0
    open_value = Cells(i, 3)
    outcnt = outcnt + 1
End If
        
Next i

'FORMATTING AFTER INITIAL PROCESSING
Columns("J:J").Select
Selection.Style = "Currency"
Columns("K:K").Select
Selection.Style = "Percent"
Columns("L:L").Select
Selection.Style = "Comma"
Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
Range("J2").Select
Selection.End(xlDown).Select
Selection.End(xlUp).Select
Range("J2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
    Formula1:="=0"
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 5287936
    .TintAndShade = 0
End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWindow.SmallScroll Down:=-135
    
'AUTOFIT COLUMNS FOR CORRECT WIDTH
Columns("J:J").EntireColumn.AutoFit
Columns("K:K").EntireColumn.AutoFit
Columns("L:L").EntireColumn.AutoFit

temp_greatest_percentage_increase = Cells(i, 10)
temp_greatest_percentage_decrease = Cells(i, 11)
temp_greatest_volume = Cells(2, 12)
current_ticker = Cells(2, 9)

'HEADERS FOR FINAL VALUES OF GREATEST PERFORMANCE STOCKS
Range("O1") = "Ticker"
Range("P1") = "Value"
Range("N2") = "Greatest % Increase"
Range("N3") = "Greatest % Decrease"
Range("N4") = "Greatest Total Volume"

'FIND LAST ROW
With ActiveSheet
    SecRow = .Cells(.Rows.Count, "I").End(xlUp).Row
End With

'FIND GREATEST VOLUME
For i = 3 To SecRow

current_volume = Cells(i, 12)
current_ticker = Cells(i, 9)
If (current_volume > temp_greatest_volume) Then
temp_greatest_volume = current_volume
temp_ticker = current_ticker
End If

Next i

Range("O4") = temp_ticker
Range("P4") = temp_greatest_volume

'FIND GREATEST INCREASE

temp_increase = Cells(2, 11)

For i = 3 To ColRow

current_increase = Cells(i, 11)
current_ticker = Cells(i, 9)
If (current_increase > temp_increase) Then
temp_increase = current_increase
temp_ticker = current_ticker
End If

Next i

Range("O2") = temp_ticker
Range("P2") = temp_increase

'FIND GREATEST DECREASE

temp_decrease = Cells(2, 11)

For i = 3 To ColRow

current_decrease = Cells(i, 11)
current_ticker = Cells(i, 9)
If (current_decrease < temp_decrease) Then
temp_decrease = current_decrease
temp_ticker = current_ticker
End If

Next i

Range("O3") = temp_ticker
Range("P3") = temp_decrease

'FINAL FORMATTING
Columns("N:N").EntireColumn.AutoFit
Columns("O:O").EntireColumn.AutoFit
Columns("P:P").EntireColumn.AutoFit
Range("P2:P3").Select
Selection.Style = "Percent"
Range("P4").Select
Selection.Style = "Comma"
Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
Range("A1").Select
End Sub
