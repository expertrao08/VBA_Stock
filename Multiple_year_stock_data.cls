VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub easy_solution():
'Easy
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.

'Declaring Variable
    Dim i As Double
    Dim Total_ST_MKT
    Dim rowCTN As Double
    
    'Assigning header value
    Worksheets("2016").Cells(1, 9).Value = "Tracking"
    Worksheets("2016").Cells(1, 10).Value = "Total Stock Market"
    rowCTN = 2
    
    Worksheets("2016").Cells(rowCTN, 9).Value = Worksheets("2016").Cells(rowCTN, 1).Value
    
    LastRow = Worksheets("2016").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Looping through all row to find Total stock Markets
    For i = 2 To LastRow
        If Worksheets("2016").Cells(i, 1).Value = Worksheets("2016").Cells(rowCTN, 9) Then
            Total_ST_MKT = Total_ST_MKT + Worksheets("2016").Cells(i, 7).Value
        Else
            Worksheets("2016").Cells(rowCTN, 10).Value = Total_ST_MKT
            Total_ST_MKT = Worksheets("2016").Cells(i, 7).Value
            rowCTN = rowCTN + 1
            Worksheets("2016").Cells(rowCTN, 9).Value = Worksheets("2016").Cells(i, 1).Value
        
        End If
     Next i
      Worksheets("2016").Cells(rowCTN, 10).Value = Total_ST_MKT
       
End Sub
Sub moderate_solution():
'Moderate
'Create a script that will loop through all the stocks and take the following info.
'Yearly change from what the stock opened the year at to what the closing price was.
'The percent change from the what it opened the year at to what it closed.
'The total Volume of the stock
'Ticker Symbol
'You should also have conditional formatting that will highlight positive change in green and negative change in red.


'Declaring Variable
    Dim i As Long
    Dim Total_ST_MKT
    Dim rowCTN As Double
    Dim minopenStockDate As Variant
    Dim maxcloseStockDate As Variant
    Dim rng As Range
    Dim dblMax As Variant
    Dim maxValue As Double
    Dim Ticker As String
    
    'Assigning header value
    
    Worksheets("2015").Cells(1, 9).Value = "Tracking"
    Worksheets("2015").Cells(1, 10).Value = "Yearly Change"
    Worksheets("2015").Cells(1, 11).Value = "Percentage Change"
    Worksheets("2015").Cells(1, 12).Value = "Total Stock Market"
    rowCTN = 2
    
    Worksheets("2015").Cells(rowCTN, 9).Value = Worksheets("2015").Cells(rowCTN, 1).Value
    minopenStockDate = Worksheets("2015").Cells(rowCTN, 3).Value
    LastRow = Worksheets("2015").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Looping through all coumnns
    For i = 2 To LastRow
        If Worksheets("2015").Cells(i, 1).Value = Worksheets("2015").Cells(rowCTN, 9) Then
            Total_ST_MKT = Total_ST_MKT + Worksheets("2015").Cells(i, 7).Value
            maxcloseStockDate = Worksheets("2015").Cells(i, 6).Value
        Else
            Worksheets("2015").Cells(rowCTN, 10).Value = maxcloseStockDate - minopenStockDate
            If maxcloseStockDate <= 0 Then
                Worksheets("2015").Cells(rowCTN, 11).Value = 0
            Else
                Worksheets("2015").Cells(rowCTN, 11).Value = (maxcloseStockDate / minopenStockDate) - 1
            End If
            
            If Worksheets("2015").Cells(rowCTN, 11).Value >= 0 Then
                Worksheets("2015").Cells(rowCTN, 10).Interior.ColorIndex = 4
            Else
                Worksheets("2015").Cells(rowCTN, 10).Interior.ColorIndex = 3
            End If
            Worksheets("2015").Cells(rowCTN, 12).Value = Total_ST_MKT
            Total_ST_MKT = Worksheets("2015").Cells(i, 7).Value
            rowCTN = rowCTN + 1
            Worksheets("2015").Cells(rowCTN, 9).Value = Worksheets("2015").Cells(i, 1).Value
        
        End If
     Next i
     
            Worksheets("2015").Cells(rowCTN, 10).Value = maxcloseStockDate - minopenStockDate
            If maxcloseStockDate <= 0 Then
                Worksheets("2015").Cells(rowCTN, 11).Value = 0
            Else
                Worksheets("2015").Cells(rowCTN, 11).Value = (maxcloseStockDate / minopenStockDate) - 1
            End If
            
            If Worksheets("2015").Cells(rowCTN, 11).Value >= 0 Then
                Worksheets("2015").Cells(rowCTN, 10).Interior.ColorIndex = 4
            Else
                Worksheets("2015").Cells(rowCTN, 10).Interior.ColorIndex = 3
            End If
            Worksheets("2015").Cells(rowCTN, 12).Value = Total_ST_MKT
            
End Sub
Sub Hard_Solutions()

'Hard
'Your solution will include everything from the moderate challenge.
'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

'Declaring Variable
    Dim i As Long
    Dim Total_ST_MKT
    Dim rowCTN As Double
    Dim minopenStockDate As Variant
    Dim maxcloseStockDate As Variant
    Dim rng As Range
    Dim dblMax As Variant
    Dim maxValue As Double
    Dim Ticker As String
    
'Assigning Variable
    Worksheets("2014").Cells(1, 9).Value = "Tracking"
    Worksheets("2014").Cells(1, 10).Value = "Yearly Change"
    Worksheets("2014").Cells(1, 11).Value = "Percentage Change"
    Worksheets("2014").Cells(1, 12).Value = "Total Stock Market"
    Worksheets("2014").Cells(2, 15).Value = "Greatest % Increase"
    Worksheets("2014").Cells(3, 15).Value = "Greatest % Decrease"
    Worksheets("2014").Cells(4, 15).Value = "Greatest Total Volume"
    Worksheets("2014").Cells(1, 16).Value = "Ticker"
    Worksheets("2014").Cells(1, 17).Value = "Value"
   
    rowCTN = 2
    
    Worksheets("2014").Cells(rowCTN, 9).Value = Worksheets("2014").Cells(rowCTN, 1).Value
    minopenStockDate = Worksheets("2014").Cells(rowCTN, 3).Value
    LastRow = Worksheets("2014").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Looping throught All Rows
    For i = 2 To LastRow
        If Worksheets("2014").Cells(i, 1).Value = Worksheets("2014").Cells(rowCTN, 9) Then
            Total_ST_MKT = Total_ST_MKT + Worksheets("2014").Cells(i, 7).Value
            maxcloseStockDate = Worksheets("2014").Cells(i, 6).Value
        Else
            Worksheets("2014").Cells(rowCTN, 10).Value = maxcloseStockDate - minopenStockDate
            If maxcloseStockDate <= 0 Then
                Worksheets("2014").Cells(rowCTN, 11).Value = 0
            Else
                Worksheets("2014").Cells(rowCTN, 11).Value = (maxcloseStockDate / minopenStockDate) - 1
            End If
            
            If Worksheets("2014").Cells(rowCTN, 11).Value >= 0 Then
                Worksheets("2014").Cells(rowCTN, 10).Interior.ColorIndex = 4
            Else
                Worksheets("2014").Cells(rowCTN, 10).Interior.ColorIndex = 3
            End If
            Worksheets("2014").Cells(rowCTN, 12).Value = Total_ST_MKT
            Total_ST_MKT = Worksheets("2014").Cells(i, 7).Value
            rowCTN = rowCTN + 1
            Worksheets("2014").Cells(rowCTN, 9).Value = Worksheets("2014").Cells(i, 1).Value
        
        End If
     Next i
     
            Worksheets("2014").Cells(rowCTN, 10).Value = maxcloseStockDate - minopenStockDate
            If maxcloseStockDate <= 0 Then
                Worksheets("2014").Cells(rowCTN, 11).Value = 0
            Else
                Worksheets("2014").Cells(rowCTN, 11).Value = (maxcloseStockDate / minopenStockDate) - 1
            End If
            
            If Worksheets("2014").Cells(rowCTN, 11).Value >= 0 Then
                Worksheets("2014").Cells(rowCTN, 10).Interior.ColorIndex = 4
            Else
                Worksheets("2014").Cells(rowCTN, 10).Interior.ColorIndex = 3
            End If
            Worksheets("2014").Cells(rowCTN, 12).Value = Total_ST_MKT
            
            Volume_Greatest_Decrease = Worksheets("2014").Cells(2, 11).Value
            Ticker_Greatest_Decrease = Worksheets("2014").Cells(2, 9).Value
            
            Volume_Greatest_Increase = Worksheets("2014").Cells(2, 11).Value
            Ticker_Greatest_Increase = Worksheets("2014").Cells(2, 9).Value
            
            Volume_Greatest_Total_Volume = Worksheets("2014").Cells(2, 12).Value
            Ticker_Greatest_Total_Volume = Worksheets("2014").Cells(2, 9).Value
            
            
            'Based on RowCTN finding which one Gratest increase or Decrease in Percentage and what is Greatest highest total Volumn
            
            For j = 3 To rowCTN
                If (Worksheets("2014").Cells(j, 11).Value < Volume_Greatest_Decrease) Then
                    Volume_Greatest_Decrease = Worksheets("2014").Cells(j, 11).Value
                    Ticker_Greatest_Decrease = Worksheets("2014").Cells(j, 9).Value
                End If
                
                If (Worksheets("2014").Cells(j, 11).Value > Volume_Greatest_Increase) Then
                    Volume_Greatest_Increase = Worksheets("2014").Cells(j, 11).Value
                    Ticker_Greatest_Increase = Worksheets("2014").Cells(j, 9).Value
                End If
                If (Worksheets("2014").Cells(j, 12).Value > Volume_Greatest_Total_Volume) Then
                    Volume_Greatest_Total_Volume = Worksheets("2014").Cells(j, 12).Value
                    Ticker_Greatest_Total_Volume = Worksheets("2014").Cells(j, 9).Value
                End If
                
                
            Next
            
            'Assigning increase and Decrease value to cells
            Worksheets("2014").Cells(2, 16).Value = Ticker_Greatest_Increase
            Worksheets("2014").Cells(2, 17).Value = Volume_Greatest_Increase
            Worksheets("2014").Cells(3, 16).Value = Ticker_Greatest_Decrease
            Worksheets("2014").Cells(3, 17).Value = Volume_Greatest_Decrease
            Worksheets("2014").Cells(4, 16).Value = Ticker_Greatest_Total_Volume
            Worksheets("2014").Cells(4, 17).Value = Volume_Greatest_Total_Volume
End Sub
     
     


