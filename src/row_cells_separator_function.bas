Attribute VB_Name = "row_cells_separator_function"
Function cellsSeparator(rng As Range) As Variant
   Dim cells_array, cell As Range, i As Long, C As Range
   
   'ReDim the dynamic array to store the range cells number.
   ReDim cells_array(rng.Cells.Count - 1) ' -1 count cells backward
                                           
   ' Loop through each area- there is one object in each area
   For Each cell In rng.Areas
        'Loop through cells within each area
        For Each C In cell.Cells
            cells_array(i) = C.Value: i = i + 1 'put each cell value in the array
        Next
   Next
   cellsSeparator = cells_array
End Function
