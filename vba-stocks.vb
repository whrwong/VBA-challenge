Sub vba_stocks()

    ' Declare variable 'ticker' as string
    Dim ticker As String
    ' Declare variable 'N' as row # for output
    Dim N As Integer
    ' Declare variable 'open_price' as double
    Dim open_price As Double
    ' Declare variable 'close_price' as double
    Dim close_price As Double
    ' Declare variable 'total_vol' as integer
    Dim total_vol As Long
    total_vol = 0
    ' Declare variable ' i ' as integer for input
    Dim i As Long

    
    ' Declare and set variables for workbook and worksheets
    Dim ThisWorkbook As Workbook: Set ThisWorkbook = ActiveWorkbook
    Dim sht As Worksheet: Set sht = ThisWorkbook.ActiveSheet
    
    ' Look through all sheets in this workbook
    For Each sht In ThisWorkbook.Worksheets

   ' Set column headers for output
     sht.Cells(1, 9).Value = "Ticker"
     sht.Cells(1, 10).Value = "Yearly Change"
     sht.Cells(1, 11).Value = "Percent Change"
     sht.Cells(1, 12).Value = "Total Volume"

        ' Start writing the values in the second row
        N = 2
        ' Write ticker values into the sheet 1 row 2, col 9
        sht.Cells(N, 9).Value = sht.Cells(2, 1).Value
        open_price = sht.Cells(2, 3).Value
        total_vol = sht.Cells(2, 7).Value
        ' Loops through rows 2 to the last row of data
        
        ' Count rows 2 to end of the data of input data
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

       
            ' Check if we are still within the same ticker, if we are not then
            If sht.Cells(i, 1).Value <> sht.Cells(i + 1, 1).Value Then
                close_price = sht.Cells(i, 6).Value
                sht.Cells(N, 10).Value = close_price - open_price
                sht.Cells(N, 11).Value = Round(((close_price - open_price) / open_price) * 100, 0) & "%"
                sht.Cells(N, 12).Value = total_vol
                           
                total_vol = 0
                
                ' We move to the next row for writing the next ticker
                N = N + 1
                ' Write the new input ticker in the output ticker
                sht.Cells(N, 9).Value = sht.Cells(i + 1, 1).Value
                
            Else
            ' Add the next volume to the 'total_vol'
            total_vol = total_vol + sht.Cells(i, 7).Value / 1000
    
            End If
         Next i
    Next sht
End Sub

