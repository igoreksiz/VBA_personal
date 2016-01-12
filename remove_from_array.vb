Public Function remove_from_array(ByVal my_array As Variant, ByVal l_to_remove As Long) As Variant

    Dim l_counter           As Long
    Dim arr_result()        As Long
    Dim b_found             As Boolean

    ReDim arr_result(UBound(my_array) - 1)
    
    
    For l_counter = LBound(my_array) To UBound(my_array)
    
        
        If (my_array(l_counter) = l_to_remove And Not b_found) Then
            b_found = True
        Else
            If b_found Then
                arr_result(l_counter - 1) = my_array(l_counter)
            Else
                arr_result(l_counter) = my_array(l_counter)
            End If
        End If
        
    Next l_counter
    
    remove_from_array = arr_result
    Call print_array(arr_result)
End Function
