# numVBA
Replicate NumPy functionality within VBA

This projects consists of a number of modules to generate Arrays in VBA as well as a class (nArray) that incorporates as many NumPy array functions as possible.

So far it consists of various functions to build arrays in VBA as well as some Windowing functions and functions to generate random numbers based on different distributions.

For example, to declare a 4D Array filedl with zeros [Use NumberArray(n1,n2...,n#) as an equivilent to tuples]
```
Sub declareZeroArray()
    Dim ZeroArray() As Double
    ZeroArray = Zeros(NumberArray(5, 5, 5, 5))

    debug.print ArraySize(ZeroArray) 'Returns the number of elements in the array -  625
    debug.print Ndim(ZeroArray)      'Returns the number of dimensions in the array - 4
    debug.print ZeroArray(3,3,3,3)   'Returns 0
End Sub
```
