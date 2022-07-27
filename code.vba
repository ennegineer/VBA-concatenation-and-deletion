Sub Domains()

    ' ***** SET LOCATIONS FIRST! *****
    ' Identify location of new column! Enter the number (ex: column D is 4)
    Dim newCol As Integer
    newCol = 4

    ' Identify the location of the domain column!
    Dim domainCol As Integer
    domainCol = 3

    ' Set a variable for the column of interest (ID)
    Dim ID_column As Integer
    ID_column = 1

    '********************************
    '********************************


    ' Add new column header
    Cells(1, newCol).Value = "Domains"

    ' Count rows
    RCount = Cells(Rows.Count, 1).End(xlUp).Row

    
    ' Set a variable to hold the domains temporarily
    Dim domain As String
    
    For i = 2 To RCount
        ' First, copy the domain column to a new column
        Cells(i, newCol).Value = Cells(i, domainCol).Value
    Next i
    
    ' Loop through rows
    For i = 2 To RCount
        
        ' Check to see if the company is listed again
        If Cells(i + 1, ID_column).Value = Cells(i, ID_column).Value Then
        
            ' Set the domain
            domain = Cells(i, newCol).Value
        
            ' Add domain into new column
            Cells(i + 1, newCol).Value = Cells(i + 1, newCol).Value & "; " & domain
        End If
        
    Next i
    
    ' Loop through all rows from the beginning again
    For i = RCount To 2 Step -1
        ' If row ID matches the next row ID, delete it
        If Cells(i, ID_column).Value = Cells(i - 1, ID_column).Value Then
        
            ' Delete the row!
            Cells(i - 1, 1).EntireRow.Delete
            
        End If
    Next i

End Sub


