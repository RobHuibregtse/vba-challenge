Attribute VB_Name = "Module11"
Sub calc_all_sheets()

Dim currentsheet As Worksheet

Application.ScreenUpdating = False

For Each currentsheet In Worksheets

    currentsheet.Select

    Call calc_current_sheet

Next

Application.ScreenUpdating = True

End Sub

Sub calc_current_sheet()

'Define worksheet variable
Dim ws As Worksheet

'Print column headers to sheet and set column widths
Range("I1").Value = "Ticker"
Range("J1").Value = "Quarterly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Volume"
Range("I1:L1").ColumnWidth = 15

'Declare variables needed for calculations and looping
Dim totalcount As Integer
Dim lastrow As Long
Dim currticker As String
Dim nextticker As String
Dim currtickercount As Integer
Dim currtickernumber As Integer
Dim currtickeropen As Double
Dim currtickerclose As Double
Dim currtickervolume As Double
Dim quarterlychange As Double
Dim quarterlypercentchange As Double

'Set initial values for all counters
totalcount = 2
lastrow = Cells(Rows.count, 1).End(xlUp).Row
currtickercount = 2
currtickernumber = 2

'Set initial open value
currtickeropen = Cells(totalcount, 3).Value

'Loop through all rows in sheet
For totalcount = 2 To lastrow

    'Set ticker and next ticker values
    currticker = Cells(totalcount, 1)
    nextticker = Cells(totalcount + 1, 1)
    
    'If still counting the current ticker
    If currticker = nextticker Then
    
    'Add current row volume to the current volume
    currtickervolume = currtickervolume + Cells(totalcount, 7).Value
    
    'Once the end of the current ticker is reached
    ElseIf currticker <> nextticker Then
    
    'Add current row volume to the current volume
    currtickervolume = currtickervolume + Cells(totalcount, 7).Value
    
    'Calculate quarterly change
    currtickerclose = Cells(totalcount, 6).Value
    quarterlychange = currtickerclose - currtickeropen
    quarterlypercentchange = quarterlychange / currtickeropen
    
    'Print calculated values to sheet
    Cells(currtickernumber, 9).Value = currticker
    Cells(currtickernumber, 10).Value = quarterlychange
    Cells(currtickernumber, 11).Value = quarterlypercentchange
    Cells(currtickernumber, 11).NumberFormat = "0.00%"
    Cells(currtickernumber, 12).Value = currtickervolume
    
    'Color quarterly change and quarterly percent cells based on positive or negative value
    If quarterlychange > 0 Then
    Cells(currtickernumber, 10).Interior.ColorIndex = 4
    Cells(currtickernumber, 11).Interior.ColorIndex = 4
    ElseIf quarterlychange < 0 Then
    Cells(currtickernumber, 10).Interior.ColorIndex = 3
    Cells(currtickernumber, 11).Interior.ColorIndex = 3
    Else
    Cells(currtickernumber, 10).Interior.ColorIndex = 6
    Cells(currtickernumber, 11).Interior.ColorIndex = 6
    End If
    
    'Set counters in prepraration for the next ticker
    currtickervolume = 0
    currtickeropen = Cells((totalcount + 1), 3).Value
    currtickernumber = currtickernumber + 1
    
    Else
    End If

Next totalcount

'Print summary headers
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N1").ColumnWidth = 20
Range("O1:P1").ColumnWidth = 15

'Find summary values, format cells
Range("P2").Value = Application.WorksheetFunction.Max(Range("K:K"))
Range("P2").NumberFormat = "0.00%"
Range("O2").Value = Application.WorksheetFunction.XLookup(Range("P2"), Range("K:K"), Range("I:I"), "Error", 1)

Range("P3").Value = Application.WorksheetFunction.Min(Range("K:K"))
Range("P3").NumberFormat = "0.00%"
Range("O3").Value = Application.WorksheetFunction.XLookup(Range("P3"), Range("K:K"), Range("I:I"), "Error", 1)

Range("P4").Value = Application.WorksheetFunction.Max(Range("L:L"))
Range("P4").NumberFormat = "0"
Range("O4").Value = Application.WorksheetFunction.XLookup(Range("P4"), Range("L:L"), Range("I:I"), "Error", 1)
End Sub


