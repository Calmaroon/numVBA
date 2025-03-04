Attribute VB_Name = "ManipulateArray"
Function ndim(arrInput As Variant) As Integer
'returns the number of dimensions of an array.
    On Error GoTo err
    Dim intDimensions As Integer
    
    Dim i As Integer
    Dim intCheck As Integer
    
    For i = 1 To 5 'for now let's handle up to 5 dimensions
        intCheck = LBound(arrInput, i)
    Next i
    
    ndim = 5
    Exit Function
err:
    ndim = i - 1
End Function
Function shape(arrInput As Variant) As Integer()
'returns the shape of an array as an array of integers; currently handling up to 5 dimensions
'to replicate NumPy functionality can use the JoinArray function, e.g. JoinArray(shape(zeros(inttuple(3,3,3))))
    Dim intShape() As Integer
    
    Dim intDimensions As Integer
    intDimensions = ndim(arrInput)
    ReDim intShape(0 To intDimensions - 1)
    
    Dim i As Integer
    For i = 0 To UBound(intShape)
        intShape(i) = UBound(arrInput, i + 1) - LBound(arrInput, i + 1) + 1
    Next i
    
    shape = intShape()
End Function
Sub test()
    Dim arr() As Double
    
    arr = Zeros(intTuple(50, 50))
    arr = Linspace(0, 100)
    Debug.Print shape(arr)(0)
End Sub
