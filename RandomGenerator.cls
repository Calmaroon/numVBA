VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RandomGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Initialize()
    Randomize ' Initialize the random number generator
End Sub
Public Function Uniform() As Double
    Uniform = Rnd
End Function
Public Function UniformRange(min As Double, max As Double) As Double
    UniformRange = min + (max - min) * Rnd
End Function
Public Function Normal(mean As Double, stdDev As Double) As Double
    Dim u1 As Double, u2 As Double
    Dim z0 As Double
    
    ' Generate two independent uniform random variables
    u1 = Rnd
    u2 = Rnd
    
    ' Box-Muller transform to get a normally distributed value
    z0 = Sqr(-2 * Log(u1)) * Cos(2 * pi * u2)
    
    ' Return the normal distribution with the given mean and standard deviation
    Normal = mean + z0 * stdDev
End Function
Public Function RandomInteger(min As Integer, max As Integer, Optional intSize = 1) As Variant
    ' Generates a random integer between min and max (inclusive)
    If intSize < 0 Then intSize = 0

    If intSize = 0 Or intSize = 1 Then
        RandomInteger = Int((max - min + 1) * Rnd + min)
    Else
        Dim intRandom() As Variant
        ReDim intRandom(intSize)
        Dim i As Integer
        For i = 0 To intSize
            intRandom(i) = Int((max - min + 1) * Rnd + min)
        Next
        
        RandomInteger = intRandom()
    End If
End Function
Public Function Random(Optional intSize As Integer = 1) As Variant
    If intSize < 1 Then intSize = 1
    If intSize = 1 Then
        Random = Rnd()
    Else
        Dim intRandom() As Variant
        ReDim intRandom(intSize)
        
        Dim i As Integer
        For i = 0 To intSize
            intRandom(i) = Rnd()
        Next
        
        Random = intRandom()
    End If
End Function
