Attribute VB_Name = "Module1"


Sub Percent_Change_Volume()

    Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate
    Dim Tickecr As String
    Dim Total_Stock_Volume As LongLong
    Dim Last_Row As Long
    Dim First_Open As Double
    Dim Last_Close As Double
    Dim Yearly_Change As Double
    Dim Summary_Table_Row As Long
    Dim Input_Row As Long
    Dim Yearly_Change_Frac As Double
    
    'Set initial Values for Variables
    
    First_Data_Row = 2
    
    Open_Col = 3
    
    Ticker_Col = 1
    
    Input_Vol_Col = 7
    
    Close_Col = 6
    
    
    First_Open = Cells(First_Data_Row, Open_Col).Value
    Total_Stock_Volume = 0
    Summary_Table_Row = First_Data_Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'For loop ticker/changes/volume
        
        
    For Input_Row = First_Data_Row To LastRow
        Ticker = Cells(Input_Row, Ticker_Col).Value
        Total_Stock_Volume = Total_Stock_Volume + Cells(Input_Row, Input_Vol_Col).Value
        If Cells(Input_Row + 1, Ticker_Col).Value <> Ticker Then
        
        'Inputs
        
        Last_Close = Cells(Input_Row, Close_Col).Value
        
        'Calculations
        
        Yearly_Change = Last_Close - First_Open
        Yearly_Change_Frac = Yearly_Change / First_Open
        
        'Outputs
        
        Range("i" & Summary_Table_Row).Value = Ticker
        Range("j" & Summary_Table_Row).Value = Yearly_Change
        Range("k" & Summary_Table_Row).Value = FormatPercent(Yearly_Change_Frac)
        Range("l" & Summary_Table_Row).Value = Total_Stock_Volume
        
        'Nested loops for conditional formatting
        
        Dim FormatNeg As Long
        For FormatNeg = 1 To 1
            If Range("j" & Summary_Table_Row).Value < 0 Then
            Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
            Range("k" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
            
        Next FormatNeg
        Dim FormatPos As Long
        For FormatPos = 1 To 1
            If Range("j" & Summary_Table_Row).Value > 0 Then
            Range("j" & Summary_Table_Row).Interior.ColorIndex = 4
            Range("k" & Summary_Table_Row).Interior.ColorIndex = 4
        
            End If
            
        Next FormatPos
        
         'set up for next row in stock table
        First_Open = Cells(Input_Row + 1, Open_Col).Value
        Summary_Table_Row = Summary_Table_Row + 1
        Total_Stock_Volume = 0
        
        

        End If
        
        Next Input_Row
        
        Next
        
        ' Bonus
        
        
        
   End Sub
   

        
        
       
        
         
        
        

            
            
            
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    
    
    


