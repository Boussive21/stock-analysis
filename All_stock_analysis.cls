Attribute VB_Name = "Module1"


Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    'Title of Analysis
    Range("A1").Value = "DAQO (Ticker:DQ)"
    
    'Create a Header Row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
       
    
    '--------Activating 2018 Worksheet---------------
    
    Worksheets("2018").Activate
    
    'Set initial 2018 DQ Volume to zero
    totalVolume = 0
    
    'Creating a Variable for Starting & Ending Price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    
    'Find number of rows to loop over
    rowStart = 2
    'DELETE: rowEnd = 3013
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    

        
    'Loop over all rows between rowStart and rowEnd
    For i = rowStart To rowEnd
            
            'increase totalVolume
            If Cells(i, 1).Value = "DQ" Then
                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(i, 8).Value
            End If
            
            
            'Identify first DQ row
            If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
                'set starting price
                startingPrice = Cells(i, 6).Value
            End If
            
            'Identify last DQ row
            If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
                'set ending price
                endingPrice = Cells(i, 6).Value
            End If
    Next i
    
    
    
    
    
    
    
   '--------Change to DQ Analysis Worksheet---------------
    
    Worksheets("DQ Analysis").Activate
    'Row header
    Cells(4, 1).Value = 2018
    'Sum of 2018 DQ Volume
    Cells(4, 2).Value = totalVolume
    '2018 DQ Return Value
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    
    

End Sub
