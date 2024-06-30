Attribute VB_Name = "Module3"
Sub CombineDataClean()

    '================================================================================
    'Add a column to each worksheet and insert the worksheet name into the new column
    '================================================================================

    'Create variables
    Dim ws As Worksheet
    Dim lastrow As Long
    
    'Loop through all sheets
    For Each ws In ActiveWorkbook.Sheets
    
        'Insert a new column
        Sheets(ws.Name).Range("A:A").Insert
    
        'Insert column header
        Sheets(ws.Name).Range("A1").Value = "Sheet"
        
        'Count rows of column B
        lastrow = ws.Cells(Rows.Count, "B").End(xlUp).row
        
        'Fill the rows with the sheetname
        Sheets(ws.Name).Range("A2:A" & lastrow).Value = ws.Name
        
        Next

    '=====================================
    'Combine data from each sheet into one

    'Create variables
    Dim CombinedSheet As Worksheet
    Dim LastRowSheet As Long

    'Add a new sheet
    Sheets.Add.Name = "Combined_Data"
    
    'Make the new sheet the first one
    Sheets("Combined_Data").Move Before:=Sheets(1)
    
    'Specify the location
    Set CombinedSheet = Worksheets("Combined_Data")

    'Loop through all sheets
    For Each ws In Worksheets

        'Find the last row of the combined sheet after each paste and create first empty row of next
        lastrow = CombinedSheet.Cells(Rows.Count, "A").End(xlUp).row + 1

        'Find the last row of each worksheet and subtract the first row header
        LastRowSheet = ws.Cells(Rows.Count, "A").End(xlUp).row - 1

        'Copy the contents of each sheet into the combined sheet
        CombinedSheet.Range("A" & lastrow & ":G" & ((LastRowSheet - 1) + lastrow)).Value = ws.Range("A2:G" & (LastRowSheet + 1)).Value

    Next ws

    'Copy the headers from sheet 1
    CombinedSheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value

    'Autofit to display data
    CombinedSheet.Columns("A:G").AutoFit


End Sub

