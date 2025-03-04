Attribute VB_Name = "CreateArray"
Option Explicit
'1D Returns
Public Function Arange(dblStart As Double, Optional dblStop As Double = Empty, Optional dblStep As Double = 1) As Double()
    Dim Length As Double
    Dim arrResult() As Double
    Dim i As Long

    If dblStep = 0 Then Exit Function 'Step cannot be 0
    
    'Check if Start and Stop the step direction matches the start/stop direction
    If (dblStop - dblStart) / dblStep < 0 Then
        Arange = arrResult()
        Exit Function
    End If
    
    Length = ((dblStop - dblStart) / dblStep)
    Length = Int(Length)
    
    ReDim arrResult(0 To Length)
    For i = 0 To Length
        arrResult(i) = dblStart + (i * dblStep)
    Next i
    
    Arange = arrResult()
End Function
Function Linspace(dblStart As Double, dblStop As Double, Optional lngNum As Long = 50, Optional boolEndpoint As Boolean = True) As Double()
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
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Eye(rows As Integer, Optional columns As Integer = -1, Optional diagonal As Integer = 0) As Double()
    If columns <= 0 Then columns = rows

    Dim arrResult() As Double
    ReDim arrResult(0 To rows - 1, 0 To columns - 1)
    
    Dim i As Integer, j As Integer
    For i = 0 To rows - 1
        For j = 0 To columns - 1
            arrResult(i, j) = 0
        Next j
    Next i

    For i = 0 To rows - 1
        j = i + diagonal
        If j >= 0 And j < columns Then
            arrResult(i, j) = 1
        End If
    Next i
    
    Eye = arrResult()
End Function
Function Zeros(shape As Variant) As Double()
    Dim arrResult() As Double
    Dim i As Long

    Select Case VarType(shape)
        Case 2, 3, 4, 5: 'Single Numbers
            If shape < 1 Then shape = 1
            ReDim arrResult(shape - 1)
            
            For i = 0 To UBound(arrResult) - 1
                arrResult(i) = 0
            Next
        Case Else:
            Dim intShape() As Long: intShape = getNormalizedShape(shape)
            Dim mods() As Long: Call ArrayFunctions.getArrayStrides(mods(), intShape())
            Call ShapeArray(arrResult(), intShape())

            Dim dblPrevious As Double: Dim t As Long: Dim Pos() As Long
            ReDim Pos(UBound(intShape))
            For i = 0 To ArraySize(arrResult) - 1
                dblPrevious = i
                For t = 0 To UBound(mods)
                    Pos(t) = Int(dblPrevious / mods(t))
                    dblPrevious = dblPrevious Mod mods(t)
                Next t
                Call setArrayValue(arrResult(), Pos, 0)
            Next
    End Select
    Zeros = arrResult()
End Function
Function Ones(shape As Variant) As Double()
    Dim arrResult() As Double
    Dim i As Long

    Select Case VarType(shape)
        Case 2, 3, 4, 5: 'Single Numbers
            If shape < 1 Then shape = 1
            ReDim arrResult(shape - 1)
            
            For i = 0 To UBound(arrResult) - 1
                arrResult(i) = 0
            Next
        Case Else:
            Dim intShape() As Long: intShape = getNormalizedShape(shape)
            Dim mods() As Long: Call ArrayFunctions.getArrayStrides(mods(), intShape())
            Call ShapeArray(arrResult(), intShape())

            Dim dblPrevious As Double: Dim t As Long: Dim Pos() As Long
            ReDim Pos(UBound(intShape))
            For i = 0 To ArraySize(arrResult) - 1
                dblPrevious = i
                For t = 0 To UBound(mods)
                    Pos(t) = Int(dblPrevious / mods(t))
                    dblPrevious = dblPrevious Mod mods(t)
                Next t
                Call setArrayValue(arrResult(), Pos, 1)
            Next
    End Select
    Ones = arrResult()
End Function
Function NormalRandomFill(shape As Variant, Mean As Double, stdDev As Double) As Double()
    Dim arrResult() As Double
    Dim i As Long

    Select Case VarType(shape)
        Case 2, 3, 4, 5: 'Single Numbers
            If shape < 1 Then shape = 1
            ReDim arrResult(shape - 1)
            
            For i = 0 To UBound(arrResult) - 1
                arrResult(i) = 0
            Next
        Case Else:
            Dim intShape() As Long: intShape = getNormalizedShape(shape)
            Dim mods() As Long: Call ArrayFunctions.getArrayStrides(mods(), intShape())
            Call ShapeArray(arrResult(), intShape())

            Dim dblPrevious As Double: Dim t As Long: Dim Pos() As Long
            ReDim Pos(UBound(intShape))
            For i = 0 To ArraySize(arrResult) - 1
                dblPrevious = i
                For t = 0 To UBound(mods)
                    Pos(t) = Int(dblPrevious / mods(t))
                    dblPrevious = dblPrevious Mod mods(t)
                Next t
                Call setArrayValue(arrResult(), Pos, Normal(Mean, stdDev))
            Next
    End Select
    NormalRandomFill = arrResult()
End Function
Function IntegerRandomFill(shape As Variant, RangeMin As Long, RangeMax As Long) As Double()
    Dim arrResult() As Double
    Dim i As Long

    Select Case VarType(shape)
        Case 2, 3, 4, 5: 'Single Numbers
            If shape < 1 Then shape = 1
            ReDim arrResult(shape - 1)
            
            For i = 0 To UBound(arrResult) - 1
                arrResult(i) = 0
            Next
        Case Else:
            Dim intShape() As Long: intShape = getNormalizedShape(shape)
            Dim mods() As Long: Call ArrayFunctions.getArrayStrides(mods(), intShape())
            Call ShapeArray(arrResult(), intShape())

            Dim dblPrevious As Double: Dim t As Long: Dim Pos() As Long
            ReDim Pos(UBound(intShape))
            For i = 0 To ArraySize(arrResult) - 1
                dblPrevious = i
                For t = 0 To UBound(mods)
                    Pos(t) = Int(dblPrevious / mods(t))
                    dblPrevious = dblPrevious Mod mods(t)
                Next t
                Call setArrayValue(arrResult(), Pos, RandomInteger(RangeMin, RangeMax))
            Next
    End Select
    IntegerRandomFill = arrResult()
End Function

