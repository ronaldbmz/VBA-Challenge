Sub ticker_calculation()

'Declare variables assign values to be used later

Dim sh As Integer
Dim r As Variant
Dim last_row As Variant

Dim sheet_count As Integer
Dim sheet_name As Variant


Dim ticker As Variant
Dim price_change As Variant
Dim pct_price_change As Variant
Dim vol As Variant

Dim start_price As Variant
Dim close_price As Variant
Dim new_ticker_start_row_number As Variant

Dim start_row As Variant 'this variable stores the starting position of our ouput

Dim opt As Integer

'Initializing variables

vol = 0
new_ticker_start_row_number = 2

'Step 1: To get the count of total sheets

sheet_count = ActiveWorkbook.Worksheets.Count

'Step 2: Iterate through each sheet

For sh = 1 To sheet_count

    'Getting the sheet name
    sheet_name = ActiveWorkbook.Worksheets(sh).Name
    
    'Activating the sheet
    Sheets(sheet_name).Activate
    
        'clear the existing values in the columns J to M
    Range("J:M").ClearContents
    Range("J:M").ClearFormats
    
    'creating required headers
    
    Range("J1").Value = "Ticker Stock Name"
    Range("K1").Value = "Yearly Price Change"
    Range("L1").Value = "Percentage Change"
    Range("M1").Value = "Volume"
    
    Columns("J:M").AutoFit
    
'Step 3: iterating through each row of the data of each sheet

    last_row = Range("A1").End(xlDown).Row 'getting the last row from the current activated sheet.
    

    
    For r = 2 To last_row
        'checking if the ticker value is the same in the next row or not
        If Range("A" & r).Value = Range("A" & r + 1).Value Then
            vol = Range("G" & r).Value + vol  'getting the cummulative sume of volume
            
        Else 'calculating values needed
        start_price = Range("C" & new_ticker_start_row_number).Value
        close_price = Range("F" & r).Value
        price_change = close_price - start_price
        
        
        If start_price = 0 Then
        pct_price_change = 0
        Else
        pct_price_change = (close_price - start_price) / start_price
        End If
        
        ticker = Range("A" & r).Value
        vol = Range("G" & r).Value + vol
        
        'changing new ticker start position
        
        new_ticker_start_row_number = r + 1
        
        'adding all the required calculated values to the sheet
        
        start_row = Range("J100000").End(xlUp).Row + 1
        
        Range("J" & start_row).Value = ticker
        Range("K" & start_row).Value = price_change
        Range("L" & start_row).Value = pct_price_change
        
        'converting decimal value to percentage
        Range("L" & start_row).Value = Format(Range("L" & start_row).Value, "#.##%")
        
        Range("M" & start_row).Value = vol
        
        're-initializing variable vol to 0 so that we can get the cumulative sum for new ticker
        vol = 0
    
        'conditional formatting based on price change value
        
        If Range("K" & start_row).Value < 0 Then
             Range("K" & start_row).Interior.Color = RGB(255, 0, 0)
        Else:
             Range("K" & start_row).Interior.Color = RGB(169, 208, 142)
             
        End If
        
      End If
      
    Next r
    
    'bonus part
        'clear old contents
        Range("O:Q").ClearContents
        Range("O:Q").ClearFormats
        
        'giving headers
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'using formulas to get the values of greatest % increase and decrease and also greatest volume
        max_value_pct = Application.WorksheetFunction.Max(Range("L2:L" & start_row + 1))
        min_value_pct = Application.WorksheetFunction.Min(Range("L2:L" & start_row))
        max_value_vol = Application.WorksheetFunction.Max(Range("M2:M" & start_row + 1))
        
        'getting the corresponding ticket based on greatest values
        max_pct_ticker = Application.WorksheetFunction.Index(Range("J2:J" & start_row + 1), Application.WorksheetFunction.Match(max_value_pct, Range("L2:L" & start_row + 1), 0))
        min_pct_ticker = Application.WorksheetFunction.Index(Range("J2:J" & start_row + 1), Application.WorksheetFunction.Match(min_value_pct, Range("L2:L" & start_row + 1), 0))
        max_vol_ticker = Application.WorksheetFunction.Index(Range("J2:J" & start_row + 1), Application.WorksheetFunction.Match(max_value_vol, Range("M2:M" & start_row + 1), 0))
        
        'copying the values in the respective cells
        Range("P2").Value = max_pct_ticker
        Range("P3").Value = min_pct_ticker
        Range("P4").Value = max_vol_ticker
        
        Range("Q2").Value = Format(max_value_pct, "#.####%")
        Range("Q3").Value = Format(min_value_pct, "#.####%")
        Range("Q4").Value = max_value_vol
        
        Columns("O:Q").AutoFit
        Columns("J:M").AutoFit
        

    
    
    
    opt = MsgBox("Tab: " & sheet_name & " is completed. Do you want to continue to next tab?", vbYesNo + vbQuestion, "Continue Yes/No")
    
    If opt = vbNo Then
        GoTo exit_macro
    End If

new_ticker_start_row_number = 2

Next sh


exit_macro:
MsgBox ("Marco run is completed!")


End Sub

