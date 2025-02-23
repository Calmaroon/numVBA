Attribute VB_Name = "Windowing"
Function Bartlett(intPoints As Integer) As Double()
'intPoints = Number of points in the output arrWindow. If zero or less, an empty array is returned.
    If intPoints < 1 Then
        err.Raise 24601, , "Window size N must be greather than 0."
        Bartlett = Array()
        Exit Function
    End If
        
    Dim arrWindow() As Double
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
Function Blackman(intPoints As Integer) As Double()
    If intPoints < 1 Then
        err.Raise 24601, , "Window size must be greater than 0."
        Blackman = Array(0)
        Exit Function
    End If
    
    Dim arrWindow() As Double
    ReDim arrWindow(intPoints - 1)
    
    ' Handle the special case where intPoints = 1
    If intPoints = 1 Then
        arrWindow(0) = 1
        Blackman = arrWindow()
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To intPoints - 1
        arrWindow(i) = 0.42 - 0.5 * Cos(2 * pi * i / (intPoints - 1)) + 0.08 * Cos(4 * pi * i / (intPoints - 1))
    Next i
    
    Blackman = arrWindow()
End Function

Function Hamming(intPoints As Integer) As Double()
'The Hamming window is a taper formed by using a weighted cosine.
    If intPoints < 1 Then
        err.Raise 24601, , "Window size must be greater than 0."
        Hamming = Array(0)
        Exit Function
    End If
    
    Dim arrWindow() As Double
    ReDim arrWindow(intPoints - 1)
    
    ' Handle the special case where intPoints = 1
    If intPoints = 1 Then
        arrWindow(0) = 1
        Hamming = arrWindow()
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To intPoints - 1
        arrWindow(i) = 0.54 - 0.46 * Cos(2 * pi * i / (intPoints - 1))
    Next i
    
    Hamming = arrWindow()
End Function
Function Hanning(intPoints As Integer) As Double()
'The Hanning window is a taper formed by using a weighted cosine.
    If intPoints < 1 Then
        err.Raise 24601, , "Window size must be greater than 0."
        Hanning = Array(0)
        Exit Function
    End If
    
    Dim arrWindow() As Double
    ReDim arrWindow(intPoints - 1)
    
    ' Handle the special case where intPoints = 1
    If intPoints = 1 Then
        arrWindow(0) = 1
        Hanning = arrWindow()
        Exit Function
    End If
    
    Dim i As Integer
    For i = 0 To intPoints - 1
        arrWindow(i) = 0.5 - 0.5 * Cos(2 * pi * i / (intPoints - 1))
    Next i
    
    Hanning = arrWindow()
End Function

