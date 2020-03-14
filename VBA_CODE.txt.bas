Attribute VB_Name = "Module1"
Sub Multi_year_stock_data()
 
 ' Set CurrentWs

    Dim CurrentWs As Worksheet
    Dim Need_Summary_Table_Header As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    Need_Summary_Table_Header = False
    COMMAND_SPREADSHEET = True
    
    
    
        ' Set Ticker name

        Dim TICKER_Name As String
        TTICKER_Name = " "
        
        ' Set an initial variable

        Dim Total_TICKER_Volume As Double
        Total_TICKER_Volume = 0
        
        ' Set new variable

        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Delta_Price As Double
        Delta_Price = 0
        Dim Delta_Percent As Double
        Delta_Percent = 0

        ' Set new variables

        Dim MAX_TICKER_NAME As String
        MAX_TICKER_NAME = " "
        Dim MIN_TICKER_NAME As String
        MIN_TICKER_NAME = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0
         
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        ' The Results
        If Need_Summary_Table_Header Then

            CurrentWs.Range("I1").Value = "TICKER"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"

            ' Set Additional Titles
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "TICKER"
            CurrentWs.Range("Q1").Value = "Value"
        Else
            
            Need_Summary_Table_Header = True
        End If
        
        ' Set initial value CurrentWs,
        ' The rest ticker
        Open_Price = CurrentWs.Cells(2, 3).Value
        
        For i = 2 To Lastrow
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                ' Insert this ticker name data
                TICKER_Name = CurrentWs.Cells(i, 1).Value
                
                Close_Price = CurrentWs.Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
                ' Check Division
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                Else
                    MsgBox ("For " & TICKER_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
                ' Add Ticker name total volume
                Total_TICKER_Volume = Total_TICKER_Volume + CurrentWs.Cells(i, 7).Value
              
                
                ' Print the Ticker Name , Column i
                CurrentWs.Range("I" & Summary_Table_Row).Value = TICKER_Name

                CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price

                ' Fill Green and Red colors, ÒYEARLY CHANGEÓ

                If (Delta_Price > 0) Then
                    'GREEN color - good
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Delta_Price <= 0) Then
                    'RED color - bad
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")

                ' Print the Ticker Name,Column J
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_TICKER_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Delta_Price = 0
                
                'The for loop Delta_Percent = 0
                Close_Price = 0
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
                
                If (Delta_Percent > MAX_PERCENT) Then
                    MAX_PERCENT = Delta_Percent
                    MAX_TICKER_NAME = TICKER_Name
                    
                ElseIf (Delta_Percent < MIN_PERCENT) Then
                    MIN_PERCENT = Delta_Percent
                    MIN_TICKER_NAME = TICKER_Name
                End If
                       
                If (Total_TICKER_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_TICKER_Volume
                    MAX_VOLUME_TICKER = TICKER_Name
                End If
                
             
                Delta_Percent = 0
                Total_TICKER_Volume = 0
                
         
            ' Add to Total Ticker Volume
            Else
                Total_TICKER_Volume = Total_TICKER_Volume + CurrentWs.Cells(i, 7).Value
            End If
           
        Next i

            If Not COMMAND_SPREADSHEET Then
            
                CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                CurrentWs.Range("P2").Value = MAX_TICKER_NAME
                CurrentWs.Range("P3").Value = MIN_TICKER_NAME
                CurrentWs.Range("Q4").Value = MAX_VOLUME
                CurrentWs.Range("P4").Value = MAX_VOLUME_TICKER
                
            Else
                COMMAND_SPREADSHEET = False
            End If
        
     Next CurrentWs
End Sub

