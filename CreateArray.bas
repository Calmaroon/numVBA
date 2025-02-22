Attribute VB_Name = "CreateArray"
Option Explicit
Function Linspace(dblStart As Double, dblStop As Double, Optional lngNum As Long = 50, Optional boolEndpoint As Boolean = True) As Double()
'Return evenly spaced numbers over a specified interval.
'Returns lngNum evenly spaced samples, calculated over the interval [dblStart, dblStop].
'The endpoint of the interval can optionally be excluded.

    If lngNum < 2 Then
        MsgBox "Number of points must be more than 1", vbExclamation
        Exit Function
    End If
    
    Dim arrResult() As Double
    
    Dim dblStepSize As Double
    If boolEndpoint Then
        dblStepSize = (dblStop - dblStart) / (lngNum - 1)
    Else
        dblStepSize = (dblStop - dblStart) / lngNum
    End If
    
    If dblStart > dblStop Then
        dblStepSize = -dblStepSize
    End If
    
    ReDim arrResult(0 To lngNum - 1)
    
    Dim i As Long
    For i = 0 To lngNum - 1
        arrResult(i) = dblStart + i * dblStepSize
    Next i
    
    Linspace = arrResult()
End Function
Function Logspace(dblStart As Double, dblStop As Double, Optional lngNum As Long = 50, Optional boolEndpoint As Boolean = True, Optional dblBase As Double = 10) As Double()
'Return numbers spaced evenly on a log scale.
'In linear space, the sequence starts at base ** start (base to the power of start) and ends with base ** stop (see endpoint below).
'arrLogspace is equivalent to the code base ^ linspace(dblStart, dblStop, lngNum ,boolEndpoint)

    If lngNum < 2 Then
        MsgBox "Number of points must be more than 1", vbExclamation
        Exit Function
    End If

    Dim arrLinspace() As Double
    arrLinspace = Linspace(dblStart, dblStop, lngNum, boolEndpoint)
        
    Dim arrResult() As Double
    ReDim arrResult(0 To lngNum - 1)
    
    Dim i As Long
    For i = 0 To lngNum - 1
        arrResult(i) = dblBase ^ arrLinspace(i)
    Next i
    
    Logspace = arrResult()
End Function
