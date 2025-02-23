Attribute VB_Name = "Windowing"
Function Bartlett(intPoints As Integer) As Double()
'intPoints = Number of points in the output arrWindow. If zero or less, an empty array is returned.
    If intPoints < 1 Then
        err.Raise 27471, , "arrWindow size N must be greather than 0."
        Bartlett = Array()
        Exit Function
    End If
        
    Dim arrWindow() As Double
    ' Initialize the array with the same size as N
    ReDim arrWindow(intPoints - 1)
    
    'Handle the special case where intPoints = 1
    If intPoints = 1 Then
        arrWindow(0) = 1
        Bartlett = arrWindow()
        Exit Function
    End If

    Dim i As Integer
    For i = 0 To intPoints - 1
        If i <= (intPoints - 1) / 2 Then
            arrWindow(i) = 2 * i / (intPoints - 1)
        Else
            arrWindow(i) = 2 * (intPoints - 1 - i) / (intPoints - 1)
        End If
    Next i

    Bartlett = arrWindow()
End Function
