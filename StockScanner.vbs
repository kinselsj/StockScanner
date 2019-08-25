VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data():

   Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub

Option Explicit

Sub RunCode()
            'Set initial variable for holding stock name
        Dim Stock_Name As String
    
        'Set initial variable for holding total volume per stock
        Dim Total_Volume As Double
        Total_Volume = 0
    
        'Location for each stock
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        'Loop through all yearly stock data
        For i = 2 To 797711
    
            'Check if Stock is the same
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Set the Stock Name
                Stock_Name = Cells(i, 1).Value
            
                'Add to the Total Volume
                Total_Volume = Total_Volume + Cells(i, 7).Value
            
                'Print the Stock in the Summary Table
                Range("I" & Summary_Table_Row).Value = Stock_Name
            
                'Print the Total Volume in the Summary Table
                Range("J" & Summary_Table_Row).Value = Total_Volume
            
                'Add one to the Summary Table Row
                Summary_Table_Row = Summary_Table_Row + 1
            
                'Reset the Total Volume
                Total_Volume = 0
            
            'If stock is the same
            Else
        
                'Add to Total Volume
                Total_Volume = Total_Volume + Cells(i, 7).Value
            
            End If
        
        Next i
        
End Sub
