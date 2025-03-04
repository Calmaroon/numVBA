Attribute VB_Name = "RandomizationFunctions"
Public Function Normal(Mean As Double, stdDev As Double) As Double
    Dim u1 As Double, u2 As Double
    Dim z0 As Double
    
    u1 = Rnd
    u2 = Rnd
    
    z0 = Sqr(-2 * Log(u1)) * Cos(2 * pi * u2)
    
    Normal = Mean + z0 * stdDev
End Function
Public Function RandomInteger(Min As Long, Max As Long) As Long
    RandomInteger = IIf(Max > Min, Int((Max - Min + 1) * Rnd + Min), Int((Min - Max + 1) * Rnd + Max))
End Function
Public Function Random(Optional intSize As Integer = 1) As Double
    Random = Rnd()
End Function

