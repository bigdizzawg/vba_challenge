VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()

 ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
  ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Set a variables for specifying the column of interest, volume, opening price, closing price, quaterly change, and percent quarterly change, index will be used to print in the appropriate rows
  
  Dim column As Integer
  Dim volume As Double
  Dim oprice As Double
  Dim cprice As Double
  Dim index As Integer
  Dim qchange As Double
  Dim pchange As Double
  
  
  'Initializes variables to values needed for
  column = 1
  volume = 0
  oprice = ws.Cells(2, 3).Value
  index = 2
  
  'creates columns headings
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  
  
  ' Loop through rows in the column
  For i = 2 To LastRow
    ' Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        'adds last value for total volume
        volume = volume + ws.Cells(i, 7)
        
        'sets closing price before calculating quarterly change and percent change
        cprice = ws.Cells(i, 6).Value
        
        'calculates quarteryly change and percent change
        qchange = cprice - oprice
        pchange = (cprice - oprice) / oprice
        
        'print everything to the appropriate cell
        ws.Cells(index, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(index, 10) = qchange
        ws.Cells(index, 11) = pchange
            'conditional formatting
            If pchange > 0 Then
                ws.Cells(index, 11).Interior.ColorIndex = 4
            ElseIf pchange < 0 Then
                ws.Cells(index, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(index, 11).Interior.ColorIndex = 0
            End If
        ws.Cells(index, 12) = volume
        
        'reset oprice and volume for next ticker
        oprice = ws.Cells(i + 1, 3).Value
        volume = 0
        
        'advance index
        index = index + 1
                
      ' Message Box the value of the current cell and value of the next cell
        'MsgBox (Cells(i, column).Value & " and then " & Cells(i + 1, column).Value & " " & volume)
    Else
        'add all volume for each loop iteration except transition
        volume = volume + ws.Cells(i, 7)
    
    End If

  Next i
' Determine the Last Row for analysis
        LastRow1 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
 
 Dim maxp As Double
 Dim Minp As Double
 Dim MaxVol As Double
 
 maxp = WorksheetFunction.Max(ws.Range("k2:k1501"))
 MsgBox (maxp)
 Minp = WorksheetFunction.Min(ws.Range("k2:k1501"))
 MsgBox (Minp)
 MaxVol = WorksheetFunction.Max(ws.Range("L2:L1501"))
 MsgBox (MaxVol)
 
            ws.Cells(2, 16).Value = maxp
            ws.Cells(3, 16).Value = Minp
            ws.Cells(4, 16).Value = MaxVol
  
 
        For j = 2 To LastRow1
                    
            If Minp = ws.Cells(j, 11).Value Then
                ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
            End If
            If maxp = ws.Cells(j, 11).Value Then
                ws.Cells(2, 15).Value = ws.Cells(j, 9).Value
            End If
           If MaxVol = ws.Cells(j, 12).Value Then
                ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
           End If
        Next j
        
Next ws

End Sub
