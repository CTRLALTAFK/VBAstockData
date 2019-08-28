Attribute VB_Name = "Module1"
Sub WallStreetVBA()
Dim ws As Worksheet

   For Each ws In Worksheets
        
             'Headers

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percernt Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Smallest % Increase"
ws.Cells(4, 15).Value = "Greatest Total Volume"

Dim ticker As String
Dim counter As Double
Dim totalVolume As Double
Dim Final As Double  ' closing stock value
Dim First As Double ' opening stock value
Dim yearChange As Double ' closing - opening stock value
Dim maxTicker As String



First = ws.Cells(2, 3).Value
ticker = ws.Cells(2, 1).Value
counter = 2
totalVolume = 0

   
    ' Starting Loop
    For i = 2 To ws.Range("a" & Rows.Count).End(xlUp).Row + 1
       If ticker <> ws.Cells(i, 1).Value Then ' runs when stock ticker changes
            
      ' Calculating Yearly Change
      Final = ws.Cells(i - 1, 6).Value
      yearChange = Final - First
            
        ' Calculating and Printing Yearly Change
          If First <> 0 Then
                ws.Cells(counter, 11) = Format((yearChange / First), "percent")
              
                             
                ' Formating Color
                If ws.Cells(counter, 10) > 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = 4
                
                ElseIf ws.Cells(counter, 10) < 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = 3
                            
                
                    
             Else
                ws.Cells(counter, 10) = "NA"
                ws.Cells(counter, 10).Interior.ColorIndex = 45
                
            End If
          End If
          
      ' Printing Ticker and Volume
       ws.Cells(counter, 9).Value = ticker
       ws.Cells(counter, 10).Value = yearChange
       ws.Cells(counter, 12).Value = totalVolume
       
                           
      ' Updating values
        counter = counter + 1
        ticker = ws.Cells(i, 1).Value
        totalVolume = 0
        First = ws.Cells(i, 3).Value
        
      
         
         End If
     
     
  'Adding total volume
   totalVolume = ws.Cells(i, 7) + totalVolume
   
        
        Next i
        
   ' Finding Greatest Percent Change and Printing Ticker
        Dim rowMax As Integer
        Dim MaxValue As Double
        Dim MaxName As String
                
                MaxValue = WorksheetFunction.Max(ws.Columns("K"))
                rowMax = ws.Range(WorksheetFunction.Index(ws.Columns("K"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Columns("K")), ws.Columns("K"), 0)).Address).Row
                MaxName = ws.Cells(rowMax, "I")
                ws.Cells(2, 16).Value = MaxName
                ws.Cells(2, 17).Value = Format(MaxValue, "percent")
          
  
   ' Finding Smallest Percent Change and Printing Ticker
            Dim rowMin As Integer
            Dim MinValue As Double
            Dim MinName As String
                                
                
                rowMin = ws.Range(WorksheetFunction.Index(ws.Columns("K"), WorksheetFunction.Match(WorksheetFunction.Min(ws.Columns("K")), ws.Columns("K"), 0)).Address).Row
                MinValue = WorksheetFunction.Min(ws.Columns("K"))
                MinName = ws.Cells(rowMin, "I")
                ws.Cells(3, 16).Value = MinName
                ws.Cells(3, 17).Value = Format(MinValue, "percent")
  
  
  
  'Max Total Volume
        Dim totMax As Integer
        Dim totValue As Double
        Dim totName As String
      

        totMax = ws.Range(WorksheetFunction.Index(ws.Columns("L"), WorksheetFunction.Match(WorksheetFunction.Max(ws.Columns("L")), ws.Columns("L"), 0)).Address).Row
        totValue = WorksheetFunction.Max(ws.Columns("L"))
        totName = ws.Cells(totMax, "I")
        ws.Cells(4, 16).Value = totName
        ws.Cells(4, 17).Value = totValue
    
    Next
     
End Sub
