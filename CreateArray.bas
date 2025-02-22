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

    ReDim arrResult(0 To lngNum - 1)
    
    Dim i As Long
    For i = 0 To lngNum - 1
        arrResult(i) = dblStart + (i * dblStepSize)
    Next i
    
    Linspace = arrResult()
End Function
Function Logspace(dblStart As Double, dblStop As Double, Optional lngNum As Long = 50, Optional boolEndpoint As Boolean = True, Optional dblBase As Double = 10) As Double()
'Return numbers spaced evenly on a log scale.
'In linear space, the sequence starts at base ** start (base to the power of start) and ends with base ** stop
'Logspace is equivalent to the code base ^ linspace(dblStart, dblStop, lngNum ,boolEndpoint)

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
Function Geomspace(dblStart As Double, dblStop As Double, Optional lngNum As Long = 50, Optional boolEndpoint As Boolean = True) As Double()
'Return numbers spaced evenly on a log scale (a geometric progression).
'This is similar to logspace, but with endpoints specified directly. Each output sample is a constant multiple of the previous

    If lngNum < 2 Then
        MsgBox "Number of points must be more than 1", vbExclamation
        Exit Function
    End If
    
    If dblStart = 0 Or dblStop = 0 Then
        MsgBox "Geometric sequence cannot include zero", vbExclamation
        Exit Function
    End If
    
    If (dblStart < 0 And dblStop > 0) Or (dblStart > 0 And dblStop < 0) Then
        MsgBox "Geometric sequence start and stop must have the same sign", vbExclamation
        Exit Function
    End If
    
    Dim dblSign As Double
    dblSign = IIf(dblStart < 0, -1, 1)
    
    Dim dblLogStart As Double
    Dim dblLogStop As Double
    
    If dblStart < 0 Then dblStart = Abs(dblStart)
    If dblStop < 0 Then dblStop = Abs(dblStop)
    
    dblLogStart = Math.Log(dblStart) / Log(10)
    dblLogStop = Math.Log(dblStop) / Log(10)
    
    Dim arrLogspace() As Double
    arrLogspace = Logspace(dblLogStart, dblLogStop, lngNum, boolEndpoint) 'base = 10.0
        
    Dim arrResult() As Double
    ReDim arrResult(0 To lngNum - 1)
    
    Dim i As Long
    For i = 0 To lngNum - 1
        arrResult(i) = arrLogspace(i) * dblSign
    Next i
    
    Geomspace = arrResult()
End Function
