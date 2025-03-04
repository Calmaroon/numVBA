# numVBA
Replicate NumPy functionality within VBA

This projects consists of a number of modules to generate Arrays in VBA as well as a class (nArray) that incorporates as many NumPy array functions as possible.

So far it consists of various functions to build arrays in VBA as well as some Windowing functions and functions to generate random numbers based on different distributions.

For example, to declare a 4D dimensional and fill it with zeros

Sub declareZeroArray()
    Dim ZeroArray() As Double
    ZeroArray = Zeros(NumberArray(5, 5, 5, 5))
    Debug.print ZeroArray(3,3,3,3) 'Returns 0
End Sub
