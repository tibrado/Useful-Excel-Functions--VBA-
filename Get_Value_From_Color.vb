''' Args* : Range & color as Long Int. example red=255, white=16777215
Function Get_Value_From_Color(my_range As Range, my_color As Long) As String
    Dim cell_value As String
    
    cell_value = ""                         	 ''' String to hold values

    For Each cell In my_range                    ''' Get the ranges
        If cell.Interior.Color = my_color Then   ''' check my color
            cell_value = cell_value & cell.Value & " "   '''get all values and place them in one cell. sperate by commas
        End If
    Next cell

    Get_Value_From_Color = cell_value       ''' Return values found
End Function


