Attribute VB_Name = "Module1"
Sub WorksheetLoop()
'Set MainWs as worksheet object variable
Dim headers() As Variant
Dim MainWs As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

'Set header Info
headers() = Array("Ticker ", "Date ", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly_Change", "Percent_Change", "Stock_Volume", " ", " ", " ", "Ticker", "Value")

For Each MainWs In wb.Sheets
    With MainWs
    .Rows(1).Value = ""
    For i = LBound(headers()) To UBound(headers())
    .Cells(1, 1 + i).Value = headers(i)
    
    Next i
    .Rows(1).Font.Bold = True
    .Rows(1).VerticalAlignment = xlCenter
    End With
Next MainWs

    ' Loop through all of the worksheets in the workbook
    For Each MainWs In Worksheets
    
    'Set initial variables for calculations
       Dim Ticker_Name As String
       Ticker_Name = " "
       Dim Total_Ticker_Volume As Double
       Total_Ticker_Volume = O
       Dim Beg_Price As Double
       Beg_Price = 0
       Dim End_Price As Double
       End_Price -0
       Dim Yearly_Price_Change As Double
       Yealy_Price_Change = 0
       Dim Yealy_Price_Change_Percent As Double
       Yealy_Price_Change_Percent = 0
       Dim Max_Ticker_Name As String
       Max_Ticker_Name = " "
       Dim Min_Ticker_Name As String
       Min_Ticker_Name = " "
       Dim Max_Percent As Double
       Max_Percent = 0
       Dim Min_Percent As Double
       Min Percent = 0
        






End Sub
