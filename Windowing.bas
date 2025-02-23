Attribute VB_Name = "Windowing"
Function Bartlett(intPoints As Integer) As Double()
'intPoints = Number of points in the output window. If zero or less, an empty array is returned.
    If intPoints < 1 Then
        err.Raise 27471, , "Window size N must be greather than 0."
        BarlettWindow = Array()
        Exit Function
    End If
    
        
    Dim window() As Double
    ' Initialize the array with the same size as N
    ReDim window(intPoints - 1)
    
    'Handle the special case where intPoints = 1
    If intPoints = 1 Then
        window(0) = 1
        Bartlett = window()
        Exit Function
    End If

        
    Dim i As Integer
    For i = 0 To intPoints - 1
        If i <= (intPoints - 1) / 2 Then
            window(i) = 2 * i / (intPoints - 1)
        Else
            window(i) = 2 * (intPoints - 1 - i) / (intPoints - 1)
        End If
    Next i

    Bartlett = window()
End Function
