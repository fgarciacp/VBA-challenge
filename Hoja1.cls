VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub QChange():
Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
 
        'defining variables type
        Dim WorksheetName As String
        Dim LastRow1 As Long
        Dim LastRow2 As Long
        Dim i As Long
        Dim z As Long
        Dim TickCount As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Create columns
        ws.Cells(1, 9).Value = "Symbol"
        ws.Cells(1, 10).Value = "triQL Change"
        ws.Cells(1, 11).Value = "% Change"
        ws.Cells(1, 12).Value = "TSK Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "% Increase"
        ws.Cells(3, 15).Value = "% Decrease"
        ws.Cells(4, 15).Value = "GreatTotal Vol"
        
        'Counters
        TC = 2
         z = 2
        
        ' Find the last row in the worksheet
        LastRow1 = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       
        
        'Loop through all rows
          For i = 2 To LastRow1
            
           'Check if ticker name changed
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'Write ticker in column  9
                ws.Cells(TC, 9).Value = ws.Cells(i, 1).Value
                
                'write Quartely Change in column 10
                ws.Cells(TC, 10).Value = ws.Cells(i, 6).Value - ws.Cells(z, 3).Value
                
                    
                    'Calculate and print porcent column  11
                    If ws.Cells(z, 3).Value <> 0 Then
                    PChang = ((ws.Cells(i, 6).Value - ws.Cells(z, 3).Value) / ws.Cells(z, 3).Value)
                    ws.Cells(TC, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TC, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                    'Conditional rule to red color
                    If ws.Cells(TC, 10).Value < 0 Then
                    ws.Cells(TC, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'the other cause to color green
                    ws.Cells(TC, 10).Interior.ColorIndex = 4
                
                    End If
                    
                'print total volume in colum 12
                ws.Cells(TC, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(z, 7), ws.Cells(i, 7)))
                
                TC = TC + 1
                
                z = i + 1
                
                End If
            
            Next i
            
        'Find last value in column
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        Vol = ws.Cells(2, 12).Value
        Incr = ws.Cells(2, 11).Value
        Decr = ws.Cells(2, 11).Value
        
            'Loop summary
            For i = 2 To LastRow2
            
                'For  volume conmditions
                If Vol < ws.Cells(i, 12).Value Then
                Vol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Vol = Vol
                
                End If
                
                'For increase conditions
                If ws.Cells(i, 11).Value > Incr Then
                Incr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Incr = Incr
                
                End If
                
                'For conditions decrease
                If ws.Cells(i, 11).Value < Decr Then
                Decr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Decr = Decr
                
                End If
                
            'Print results
            ws.Cells(2, 17).Value = Format(Incr, "Percent")
            ws.Cells(3, 17).Value = Format(Decr, "Percent")
            ws.Cells(4, 17).Value = Format(Vol, "Scientific")
            
            Next i
            
        'Djust column width automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
            
    Next ws
        
End Sub
