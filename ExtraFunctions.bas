Attribute VB_Name = "ExtraFunctions"
Public Function NumberArray(ParamArray values() As Variant) As Long()
    On Error GoTo err
    
    Dim arrResult() As Long
    ReDim arrResult(UBound(values))
    
    Dim i As Integer
    For i = 0 To UBound(values)
        If VarType(values(i)) <> 2 And VarType(values(i)) <> 5 Then
            MsgBox "All input values must be Integers!"
            Exit Function
        End If
        
        arrResult(i) = CLng(values(i))
    Next i

    NumberArray = arrResult()
    Exit Function
err:
    Debug.Print err.Number & ":" & err.Description
End Function
Public Function IsVector(ByRef arr As Variant) As Boolean
    ' Check if array is 1-dimensional
    Dim Dimensions As Long
    On Error GoTo ErrorHandler
    

    IsVector = (nDim(arr) = 1)
    Exit Function
    
ErrorHandler:
    IsVector = False
End Function

Public Function IsMatrix(ByRef arr As Variant) As Boolean
    ' Check if array is 2-dimensional
    Dim Dimensions As Long
    On Error GoTo ErrorHandler
    
    IsMatrix = (nDim(arr) = 2)
    Exit Function
    
ErrorHandler:
    IsMatrix = False
End Function

Public Function JoinArray(ByRef arr As Variant, Optional Delimiter As String = ",") As String
    Dim i As Long
    Dim result As String
    If nDim(arr) > 0 Then
        For i = LBound(arr) To UBound(arr)
            If i = LBound(arr) Then
                result = arr(i)
            Else
                result = result & Delimiter & arr(i)
            End If
        Next i
    Else
        result = arr
    End If
    JoinArray = result
End Function
Public Function Min(ParamArray values() As Variant) As Variant
    Dim i As Long
    Dim currentMin As Variant
    
    If LBound(values) > UBound(values) Then
        Min = Empty
        Exit Function
    End If

    currentMin = values(0)
    
    For i = LBound(values) + 1 To UBound(values)
        If IsNumeric(values(i)) Then
            If IsEmpty(currentMin) Or values(i) < currentMin Then
                currentMin = values(i)
            End If
        End If
    Next i
    
    Min = currentMin
End Function
Public Function Max(ParamArray values() As Variant) As Variant
    Dim i As Long
    Dim currentMax As Variant
    
    If LBound(values) > UBound(values) Then
        Min = Empty
        Exit Function
    End If

    currentMax = values(0)
    
    For i = LBound(values) + 1 To UBound(values)
        If IsNumeric(values(i)) Then
            If IsEmpty(currentMin) Or values(i) > currentMax Then
                currentMax = values(i)
            End If
        End If
    Next i
    
    Max = currentMax
End Function
