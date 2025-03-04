Attribute VB_Name = "ArrayFunctions"
Sub ShapeArray(ByRef arr As Variant, intShape() As Long)
'Expand these 3 up to 10 dimensions; just generate the code in excel
    Select Case UBound(intShape)
        Case 0: ReDim arr(intShape(0))
        Case 1: ReDim arr(intShape(0), intShape(1))
        Case 2: ReDim arr(intShape(0), intShape(1), intShape(2))
        Case 3: ReDim arr(intShape(0), intShape(1), intShape(2), intShape(3))
    End Select
End Sub
Function getArrayValue(ByRef arr As Variant, Pos() As Long) As Variant
    Select Case UBound(Pos)
        Case 0: getArrayValue = arr(Pos(0))
        Case 1: getArrayValue = arr(Pos(0), Pos(1))
        Case 2: getArrayValue = arr(Pos(0), Pos(1), Pos(2))
        Case 3: getArrayValue = arr(Pos(0), Pos(1), Pos(2), Pos(3))
    End Select
End Function
Function setArrayValue(ByRef arr As Variant, Pos() As Long, fill As Variant) As Variant
    Select Case UBound(Pos)
        Case 0: arr(Pos(0)) = fill
        Case 1: arr(Pos(0), Pos(1)) = fill
        Case 2: arr(Pos(0), Pos(1), Pos(2)) = fill
        Case 3: arr(Pos(0), Pos(1), Pos(2), Pos(3)) = fill
    End Select
End Function
Public Function ArraySize(ByRef arr As Variant) As Double
    Dim Dimensions As Integer
    Dimensions = nDim(arr)
    
    Dim DblArraySize As Double
    DblArraySize = 1
    
    Dim i As Integer
    For i = 1 To Dimensions
        If LBound(arr, i) = 0 Then
            DblArraySize = DblArraySize * (UBound(arr, i) - LBound(arr, i) + 1)
        Else
            DblArraySize = DblArraySize * (UBound(arr, i) - LBound(arr, 1))
        End If
    Next
    ArraySize = DblArraySize
End Function
Function getArrayStrides(ByRef mods() As Long, intShape() As Long) As Long()
    'This will calculate the dimensional multiplier
    ReDim mods(UBound(intShape))
             
    mods(UBound(intShape)) = 1 ' Last dimension multiplier is always 1
    For i = UBound(intShape) - 1 To 0 Step -1
        mods(i) = mods(i + 1) * (intShape(i + 1) + 1)
    Next i
End Function
Function getNormalizedShape(shape As Variant) As Long()
    Dim intShape() As Long
    Dim i As Long
    
    ReDim intShape(0 To UBound(shape) - LBound(shape)) ' Adjust array size
    
    For i = LBound(shape) To UBound(shape)
        If LBound(shape) = 0 Then
            intShape(i) = shape(i) - 1 ' Convert to zero-based
        Else
            intShape(i - 1) = shape(i) - 1
        End If
    Next
    
    getNormalizedShape = intShape
End Function
Function nDim(arr As Variant) As Integer
On Error GoTo err
    'Returns the number of dimensions in an array
    Dim intDim As Integer
    Dim lngRun As Long

    For intDim = 1 To 10
        lngRun = UBound(arr, intDim)
    Next
    nDim = 10 'Max out at 10 Dimensions for now
Exit Function
err:    nDim = intDim - 1
End Function

