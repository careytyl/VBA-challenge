Attribute VB_Name = "Module1"
'Tyler Carey
'11.23.2019
'VBA Exercise

Sub Main():

    Dim WS_Count As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    
    For i = 1 To WS_Count
        
        'Cycles through each worksheet, starting with the first
        Worksheets(i).Select


        'Sets up titles
        Layout
    
    
        'Does the calculations
        Calculations


        'Lists top % increase, decrease, and volume
        Top_Numbers
    
     
        'Adds conditional formatting
        Color_Code
    
    Next i

End Sub

'Quick sub to format the layout of the title cells (assuming the provided data stays in the same layout)
Sub Layout()

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greates Total Volume"

End Sub

'Does the calculations for all parts
Sub Calculations()

    'Setup for For loop to allow for an unknown/changing number of rows
    numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count + 1
    
    'Used to keep track of which cell to insert each unique ticker
    Dim CellLocation As Double
        
    'Used to compare calues in provided data
    Dim TickerValue As String
    
    'Used to determine Yearly Change and % change
    Dim YearOpen As Double
    Dim YearClose As Double

    'Used to add stock volume
    Dim Volume As Double

    'Necessary initial values
    CellLocation = 2
    YearOpen = Cells(2, 3).Value
    
    
        
    For i = 2 To numrows

        'Compares the cells of the first column to each one below it
        If Cells(i, 1).Value = Cells(i - 1, 1).Value Then
            
            'Adds the volume for each cell with matching ticker value
            Volume = Volume + Cells(i, 7).Value
        
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value And i > 2 Or IsEmpty(Cells(i, 1)) = True Then
            
            'Lists the value of the ticker in the cell before it changes
            TickerValue = Cells(i - 1, 1).Value
            Cells(CellLocation, 9).Value = TickerValue
            Cells(CellLocation, 12).Value = Volume
            
            'Calculates and populates yearly change & % change cells
            'Makes calculation before updating new YearOpen value
            YearClose = Cells(i - 1, 6).Value
            Cells(CellLocation, 10).Value = YearClose - YearOpen
            Cells(CellLocation, 11).NumberFormat = "0%"
                If YearOpen = 0 Then
                    Cells(CellLocation, 11).Value = 0
                Else
                    Cells(CellLocation, 11).Value = (YearClose - YearOpen) / YearOpen
                End If
            YearOpen = Cells(i, 3).Value

            CellLocation = CellLocation + 1
            Volume = 0
            
            If IsEmpty(Cells(i, 1)) = True Then
                Exit For
            End If
            
        End If
        
    Next i

End Sub

'Finds the min and max values from the calculations sub and puts them in their own area
Sub Top_Numbers()

    numrows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    'Values for max %
    Dim PercentIncrease As Double
    Dim TickerIncrease As String
    
    'Values for min %
    Dim PercentDecrease As Double
    Dim TickerDecrease As String
    
    'Values for max volume
    Dim TopVolume As Double
    Dim VolumeTicker As String
    
    
    temp = 0
    
    For i = 2 To numrows
        
        'Checks for empty cells to exit loop
        If IsEmpty(Cells(i + 1, 11)) = True Then
            Exit For
            
        Else
        
            'Determines the max value
            If Cells(i, 11).Value > PercentIncrease Then
                PercentIncrease = Cells(i, 11).Value
                TickerIncrease = Cells(i, 9).Value
            End If
            
            'Determines the min value
            If Cells(i, 11).Value < PercentDecrease Then
                PercentDecrease = Cells(i, 11).Value
                TickerDecrease = Cells(i, 9).Value
            End If
            
            'Determines the max volume
            If Cells(i, 12).Value > TopVolume Then
                TopVolume = Cells(i, 12).Value
                VolumeTicker = Cells(i, 9).Value
            End If
            
        End If
    
    Next i
    

    'Displays results of all min/max values
    Cells(2, 16).Value = TickerIncrease
    Cells(2, 17).NumberFormat = "0%"
    Cells(2, 17) = PercentIncrease
    
    Cells(3, 16).Value = TickerDecrease
    Cells(3, 17).NumberFormat = "0%"
    Cells(3, 17) = PercentDecrease
    
    Cells(4, 16).Value = VolumeTicker
    Cells(4, 17) = TopVolume

End Sub

'Colors cells based on positive or negative yearly change
Sub Color_Code()

    numrows = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    For i = 2 To numrows
      
        If Cells(i, 10).Value >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4        'green
        Else
            Cells(i, 10).Interior.ColorIndex = 3        'red
        End If
    
        If IsEmpty(Cells(i + 1, 10)) = True Then        'exits loop when reacing a blank cell
            Exit For
        End If
    
    Next i

End Sub


