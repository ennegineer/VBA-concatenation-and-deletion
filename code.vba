Sub Domains()
    ' Add new column header
    ' **(Will need to know the final file structure to replace A4)
    Cells(1, 4).Value = "Domains"
    
    ' Set a variable for the column of interest (ID)
    Dim ID_column As Integer
    ID_column = 1
    
    ' Count rows
    RCount = Selection.Rows.Count
    
    ' Set a variable to hold the domains temporarily
    Dim domain As String
    
    For i = 2 To RCount
        ' First, copy the domain column to a new column
        Cells(i, 4).Value = Cells(i, 3).Value
    Next i
    
    ' Loop through rows
    For i = 2 To RCount
        
        ' Check to see if the company is listed again
        If Cells(i + 1, ID_column).Value = Cells(i, ID_column).Value Then
        
            ' Set the domain
            domain = Cells(i, 4).Value
        
            ' Add domain into new column
            Cells(i + 1, 4).Value = Cells(i + 1, 4).Value & "; " & domain
        End If
        
    Next i
    
    ' Loop through all rows from the beginning again
    For i = RCount To 2 Step -1
        ' If row ID matches the next row ID, delete it
        If Cells(i, ID_column).Value = Cells(i - 1, ID_column).Value Then
        
            ' Delete the row!
            Selection.Rows(i - 1).EntireRow.Delete
            
        End If
    Next i

End Sub

