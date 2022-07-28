# VBA-challenge
Excel VBA assignment 2
Sub Stock_Caculate()


' Set Ws as a worksheet object variable.
    Dim Ws As Worksheet
    Dim Ticker_Name As String
    Dim Total_Ticker_Volume As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Change_Price As Double
    Dim Change_Percentage As Double
    
    For Each Ws In Worksheets
    
        Total_Ticker_Volume = 0
        Change_Price = 0
        Close_Price = 0
        Change_Price = 0
        
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        
        Dim Lastrow As Integer
        Dim i As Integer
        
        Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

            Ws.Range("I1").Value = "Ticker"
            Ws.Range("J1").Value = "Yearly Change"
            Ws.Range("K1").Value = "Percent Change"
            Ws.Range("L1").Value = "Total Stock Volume"
            
            Ws.Range("O2").Value = "Greatest % Increase"
            Ws.Range("O3").Value = "Greatest % Decrease"
            Ws.Range("O4").Value = "Greatest Total Volume"
            Ws.Range("P1").Value = "Ticker"
            Ws.Range("Q1").Value = "Value"

        Open_Price = Ws.Cells(2, 3).Value
        'Loop
        
        For i = 2 To Lastrow
            '  Same ticker name
            
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            Ticker_Name = Ws.Cells(i, 1).Value
            Close_Price = Ws.Cells(i, 6).Value
            Change_Price = Close_Price - Open_Price
            Change_Percentage = Change_Price / Open_Price * 100
            Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
              
              'Column I
              Ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
              'Column J
              Ws.Range("J" & Summary_Table_Row).Value = Change_Price
              'Column K
              Ws.Range("K" & Summary_Table_Row).Value = (CStr(Change_Percentage) & "%")
              ' Column L
              Ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
              
              'GREEN
              If (Change_Price > 0) Then
              Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
              
              'RED
              ElseIf (Change_Price <= 0) Then
              Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
              
              End If
                
              Summary_Table_Row = Summary_Table_Row + 1
              Change_Price = 0
              Close_Price = 0
              Open_Price = Ws.Cells(i + 1, 3).Value
                  
            ' Bonus
                
            Dim Max_Ticker As String
            Dim Min_Ticker As String
            Dim Max_percentage As Double
            Dim Min_percentage As Double
            Dim Max_volume_Ticker As String
            Dim Max_volume As Double
                 
            If (Change_Percentage > Max_percentage) Then
                    Max_percentage = Change_Percentage
                    Max_Ticker = Ticker_Name
            ElseIf (Change_Percentage < Min_percentage) Then
                    Min_percentage = Change_Percentage
                    Min_Ticker = Ticker_Name
            End If
                       
            If (Total_Ticker_Volume > Max_volume) Then
                    Max_volume = Total_Ticker_Volume
                    Max_volume_Ticker = Ticker_Name
            End If

                Change_Percentage = 0
                
                Total_Ticker_Volume = 0
                
                Ws.Range("Q2").Value = (CStr(Max_percentage) & "%")
                Ws.Range("Q3").Value = (CStr(Min_percentage) & "%")
                Ws.Range("P2").Value = Max_Ticker
                Ws.Range("P3").Value = Min_Ticker
                Ws.Range("Q4").Value = Max_volume
                Ws.Range("P4").Value = Max_volume_Ticker
            Else
                
                Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
            End If
        Next i
     Next Ws

End Sub

