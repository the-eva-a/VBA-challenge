Attribute VB_Name = "Module2"
Sub ClearAllFormatting()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Apply changes to all cells in the worksheet
        With ws.Cells
            ' Reset number formatting to General
            .NumberFormat = "General"
            
            ' Remove any interior color fill
            .Interior.ColorIndex = xlColorIndexNone
            
            ' Remove any added Values
            Range("I:Q").Value = ""
        End With
    Next ws
End Sub

