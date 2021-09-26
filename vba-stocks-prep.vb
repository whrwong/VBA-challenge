Sub vba_stocks()

  ' Set an initial variable for holding the ticker
  Dim ticker As String

  ' Set an initial variable for holding the yearly change
  Dim yearly_change as double

  ' Set an initial variable for holding the percent change
  Dim percent_change as double

  ' Set an initial variable for holding the total volume
  Dim total_volume as integer

  ' Loop through all tickers in the same year
  For tab = 1 To 7
	For row = 2 to Cells(Rows.Count,1).End(xlUp).Row

            ' Hold the first value in the "open" column in a variable called open_price
            ' Hold the first value in the "volumn" column in a variable called volumn

    		    ' Check if we are still within the same credit card brand, if we are not...
    		    If Cells(row, 1).Value == Cells(row+1, 1).Value Then
            
                    ' Add the second consecutive value to the first value in volumn
      		    Else 

                  ' Get the last value in the "close" column in a variable called close_price
                  ' Get the next ticker
                  ' Store the opening_price
                  ' Store the volumm 
                End If

    Next row 

  Next tab
