Attribute VB_Name = "RandomizationFunctions"
Public Function Normal(Mean As Double, stdDev As Double) As Double
    Dim u1 As Double, u2 As Double
    Dim z0 As Double
    
    u1 = Rnd
    u2 = Rnd
    
    z0 = Sqr(-2 * Log(u1)) * Cos(2 * pi * u2)
    
    Normal = Mean + z0 * stdDev
End Function
Public Function RandomInteger(Min As Long, Max As Long) As Double
    RandomInteger = IIf(Max > Min, Int((Max - Min + 1) * Rnd + Min), Int((Min - Max + 1) * Rnd + Max))
End Function
Public Function Random(Optional intSize As Long = 1) As Double
    Random = Rnd() * intSize
End Function
Function Gumbel(Mu As Double, Optional Beta As Double = 1) As Double
'Mu = Location; beta = Scale
    Randomize
    Dim u As Double
    u = Rnd()
    
    If u = 0 Then u = 0.0000001
    
    Gumbel = Mu - Beta * Log(-1 * Log(u))
End Function
Function Binomial(n As Long, p As Double) As Double
'n = Number of times test is run, p = chance of success; should be 0<p<1
    Dim i As Integer, successes As Integer
    count = 0
    
    For i = 1 To n
        If Rnd() < p Then
            successes = successes + 1
        End If
    Next i
    Binomial = successes
End Function
