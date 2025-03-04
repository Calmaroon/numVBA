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
Function Gumbel(mu As Double, Optional beta As Double = 1) As Double
'Mu = Location; beta = Scale
    Randomize
    Dim u As Double
    u = Rnd()
    
    If u = 0 Then u = 0.0000001
    
    Gumbel = mu - beta * Log(-1 * Log(u))
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
Function Exponential(beta As Double) As Double
    Dim u As Double
    u = Rnd()
    
    If u = 0 Then u = 0.0000001
    
    Exponential = -beta * Log(u)
End Function
Function Bytes(n As Long) As Variant
    Dim i As Integer
    Dim ByteArray() As Bytes
    
    ReDim ByteArray(1 To n)
    
    For i = 1 To n
        ByteArray(i) = Int(Rnd() * 256)
    Next i
    
    Bytes = ByteArray
End Function
Function Logistic(mu As Double, s As Double) As Double
    Dim u As Double
    u = Rnd()
    
    If u = 1 Then u = 0.999999
    
    Logistic = mu + s * Log((u / (1 - u)))
End Function
Function Hypergeometric(nGood As Long, nBad As Long, nSample As Long) As Long
    Dim RemainingGood As Long
    Dim RemainingTotal As Long
    Dim numSelected As Long
    
    Dim i As Long
    
    RemainingGood = nGood
    RemainingTotal = nGood + nBad
    
    For i = 1 To nSample
        If Rnd() < (RemainingGood / RemainingTotal) Then
            numSelected = numSelected + 1
            RemainingGood = RemainingGood - 1
        End If
        RemainingTotal = RemainingTotal - 1
    Next
    
    Hypergeometric = numSelected
End Function
