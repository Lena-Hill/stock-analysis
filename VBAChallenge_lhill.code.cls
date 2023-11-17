VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBAChallenge()

' Set variables
Dim CurrentWs As Worksheet
Dim Need_Summary_Table_Header As Boolean
Dim COMMAND_SPREADSHEET As Boolean 'also from github
'Set header flags. got this code from github
Need_Summary_Table_Header = False
COMMAND_SPREADSHEET = True
        
' Loop through all of the worksheets, I hope.
For Each CurrentWs In Worksheets
        
' Set variable to hold the ticker name
Dim Ticker_Name As String
Ticker_Name = " "
            
' Set variable to hold total per ticker symbol
Dim Total_Ticker_Volume As Double
Total_Ticker_Volume = 0
            
' Set new variables
Dim Open_Price As Double
Open_Price = 0
Dim Close_Price As Double
Close_Price = 0
Dim Delta_Price As Double
Delta_Price = 0
Dim Delta_Percent As Double
Delta_Percent = 0

Dim max_ticker_name As String
max_ticker_name = " "
Dim min_ticker_name As String
min_ticker_name = " "
Dim max_percent As Double
max_percent = 0
Dim min_percent As Double
min_percent = 0
Dim max_volume_ticker As String
max_volume_ticker = " "
Dim max_volume As Double
max_volume = 0
             
' Keep track of the location for each ticker name. code from github
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
            
' Set row count
Dim Lastrow As Long
Dim i As Long
            
'Find the last row
Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

' For all worksheet except the first one, the Results
If Need_Summary_Table_Header Then      'this bit is from github
' Set Titles for the Summary Table for current worksheet
CurrentWs.Range("I1").Value = "Ticker"
CurrentWs.Range("J1").Value = "Yearly Change"
CurrentWs.Range("K1").Value = "Percent Change"
CurrentWs.Range("L1").Value = "Total Stock Volume"
 
CurrentWs.Range("O2").Value = "Greatest % Increase"
CurrentWs.Range("O3").Value = "Greatest % Decrease"
CurrentWs.Range("O4").Value = "Greatest Total Volume"
CurrentWs.Range("P1").Value = "Ticker"
CurrentWs.Range("Q1").Value = "Value"
    Else
'This is the first, resulting worksheet, reset flag for the rest of worksheets
Need_Summary_Table_Header = True 'this bit is also from github
    End If
            
' Set value of Open Price
Open_Price = CurrentWs.Cells(2, 3).Value
            
' Loop from the beginning of worksheet until last row
For i = 2 To Lastrow
                       
'write results to summary table
If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                
' Set the ticker name and insert data. code from stack overflow
Ticker_Name = CurrentWs.Cells(i, 1).Value
                    
' Calculate price change and percent change
Close_Price = CurrentWs.Cells(i, 6).Value
Delta_Price = Close_Price - Open_Price
' Check math. got this idea from github
If Open_Price <> 0 Then
Delta_Percent = (Delta_Price / Open_Price) * 100
Else

End If
                    
' Add volume to ticker name
Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                  
                    
' Print the ticker symbol in column I
CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
' Print the price change in column J
CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price
'Color code "yearly change"
    If (Delta_Price > 0) Then
'Fill column green
CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    ElseIf (Delta_Price <= 0) Then
'Fill column red
CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
                    
' Print percent change in column K
CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
' Print total volume in column L
CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                    
' Add 1 to the summary table row count. Got this from github
Summary_Table_Row = Summary_Table_Row + 1
' Reset price and percent change values
Delta_Price = 0
Delta_Percent = 0
Close_Price = 0
' Capture next Ticker's Open_Price, github
Open_Price = CurrentWs.Cells(i + 1, 3).Value
                  
                    
'Populate second summary table
If (Delta_Percent > max_percent) Then
max_percent = Delta_Percent
max_ticker_name = Ticker_Name
    ElseIf (Delta_Percent < min_percent) Then
min_percent = Delta_Percent
min_ticker_name = Ticker_Name
    End If
                           
If (Total_Ticker_Volume > max_volume) Then
max_volume = Total_Ticker_Volume
max_volume_ticker = Ticker_Name
End If

'reset counters again
Delta_Percent = 0
Total_Ticker_Volume = 0
                    
'Increase the total ticker volume
Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
    End If
          
Next i

If Not COMMAND_SPREADSHEET Then
                
CurrentWs.Range("Q2").Value = (CStr(max_percent) & "%")
CurrentWs.Range("Q3").Value = (CStr(min_percent) & "%")
CurrentWs.Range("P2").Value = max_ticker_name
CurrentWs.Range("P3").Value = min_ticker_name
CurrentWs.Range("Q4").Value = max_volume
CurrentWs.Range("P4").Value = max_volume_ticker
                    
Else
COMMAND_SPREADSHEET = True
End If
            
Next CurrentWs
End Sub



