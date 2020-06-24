VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickertotaler_moderate()


Dim ws As Worksheet
Dim ticker As String
Dim vol As Integer
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Need_Summary_Table_Header As Boolean
Dim COMMAND_SPREADSHEET As Boolean
    
Need_Summary_Table_Header = False       'Set Header flag
COMMAND_SPREADSHEET = True              'Hard part flag

For Each ws In Worksheets



 Dim Summary_Table_Row As Long
        Summary_Table_Row = 2

    'loop
     Dim Lastrow As Long
        Dim i As Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                    
 

            'insert values into summary
     If Need_Summary_Table_Header Then
     
            ' Set Titles for the Summary Table for current worksheet
           ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ' Set Additional Titles for new Summary Table on the right for current worksheet
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
        Else
            'This is the first, resulting worksheet, reset flag for the rest of worksheets
            Need_Summary_Table_Header = True
        End If

    
ws.Columns("K").NumberFormat = "0.00%"


   ' The rest ticker's open price will be initialized within the for loop below
        Open_Price = ws.Cells(2, 3).Value
        
        ' Loop from the beginning of the current worksheet(Row2) till its last row
        For i = 2 To Lastrow
        
      
            ' Check if we are still within the same ticker name,
            ' if not - write results to summary table
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the ticker name, we are ready to insert this ticker name data
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Calculate Delta_Price and Delta_Percent
                Close_Price = ws.Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
                ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                Else
                    ' Unlikely, but it needs to be checked to avoid program crushing
                    MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
                ' Add to the Ticker name total volume
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
              
                
                ' Print the Ticker Name in the Summary Table, Column I
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ' Print the Ticker Name in the Summary Table, Column I
                ws.Range("J" & Summary_Table_Row).Value = Delta_Price
                ' Fill "Yearly Change", i.e. Delta_Price with Green and Red colors
                If (Delta_Price > 0) Then
                    'Fill column with GREEN color - good
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Delta_Price <= 0) Then
                    'Fill column with RED color - bad
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Ticker Name in the Summary Table, Column I
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
                ' Print the Ticker Name in the Summary Table, Column J
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                ' Reset Delta_rice and Delta_Percent holders, as we will be working with new Ticker
                Delta_Price = 0
                ' Hard part,do this in the beginning of the for loop Delta_Percent = 0
                Close_Price = 0
                ' Capture next Ticker's Open_Price
                Open_Price = ws.Cells(i + 1, 3).Value
              
                
                
                ' Keep track of all extra hard counters and do calculations within the current spreadsheet
                If (Delta_Percent > MAX_PERCENT) Then
                    MAX_PERCENT = Delta_Percent
                    MAX_TICKER_NAME = Ticker_Name
                ElseIf (Delta_Percent < MIN_PERCENT) Then
                    MIN_PERCENT = Delta_Percent
                    MIN_TICKER_NAME = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Ticker_Volume
                    MAX_VOLUME_TICKER = Ticker_Name
                End If
                
                ' Hard part adjustments to resetting counters
                Delta_Percent = 0
                Total_Ticker_Volume = 0
                
            
            Else
                ' Encrease the Total Ticker Volume
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            End If
            ' For debugging MsgBox (CurrentWs.Rows(i).Cells(2, 1))
      
        Next i


            ' Record all new counts to the new summary table on the right of the current spreadsheet
            If Not COMMAND_SPREADSHEET Then
            
                ws.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                ws.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                ws.Range("P2").Value = MAX_TICKER_NAME
                ws.Range("P3").Value = MIN_TICKER_NAME
                ws.Range("Q4").Value = MAX_VOLUME
                ws.Range("P4").Value = MAX_VOLUME_TICKER
                
            Else
                COMMAND_SPREADSHEET = False
            End If
        
     Next ws
     
End Sub
