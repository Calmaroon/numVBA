Attribute VB_Name = "CreateArray"
Function arrLinspace(dblStart As Double, dblEnd As Double, Optional lngIntervals As Long = 50, Optional boolEndpoint = True) As Double()
    If lngIntervals < 2 Then
        MsgBox "Number of points must be at least 2", vbExclamation
        Exit Function
    End If
    
    Dim arrResult() As Double
    
    Dim dblStepSize As Double
    If boolEndpoint Then
        dblStepSize = (dblEnd - dblStart) / (lngIntervals - 1)
    Else
        dblStepSize = (dblEnd - dblStart) / (lngIntervals)
    End If
    
    If dblStart > dblEnd Then
        dblStepSize = -dblStepSize
    End If
    
    ReDim arrResult(0 To lngIntervals - 1)
    
    Dim i As Long
    For i = 0 To lngIntervals - 1
        arrResult(i) = dblStart + i * dblStepSize
    Next i
    arrLinspace = arrResult()
End Function
