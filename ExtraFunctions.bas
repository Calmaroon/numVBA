Attribute VB_Name = "ExtraFunctions"
Function intTuple(ParamArray values() As Variant) As Integer()
'As VBA cannot handle tuples, we need to have this function to take the input and interpret it as a collection
'This function will be used to pass a tuple of integers, have to do a check to see if each input is an integer, otherwise throw an error
    On Error GoTo err
    
    Dim arrTupleInt() As Integer
    ReDim arrTupleInt(0 To UBound(values))
    
    Dim i As Integer
    For i = LBound(values) To UBound(values)
        If VarType(values(i)) <> 2 And VarType(values(i)) <> 5 Then
            MsgBox "All input values must be Integers!"
            Exit Function
        End If
        
        arrTupleInt(i) = CInt(values(i))
    Next i

     
     intTuple = arrTupleInt()
    Exit Function
err:
    Debug.Print err.Number & ":" & err.Description
End Function
Function ndim(arrInput() As Double) As Integer
'returns the number of dimensions of an array.
    On Error GoTo err
    Dim intDimensions As Integer
    
    Dim i As Long
    Dim intCheck As Integer
    
    For i = 1 To 5 'for now let's handle up to 5 dimensions
        intCheck = LBound(arrInput, i)
    Next i
    
    ndim = 5
    Exit Function
err:
    ndim = i - 1
End Function
