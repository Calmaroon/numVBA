VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "nArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim nArray As Variant


Public Sub SetArray(newArray As Variant)
    Dim i, j, k As Long
    Select Case ndim(newArray)
        Case 1:
            ReDim nArray(0 To UBound(newArray))
            For i = 0 To UBound(newArray, 1)
                nArray(i) = newArray(i)
            Next i
    End Select
End Sub

Public Property Get Item(index As Integer) As Variant
	Attribute Item.VB_UserMemId = 0
    Item = nArray(index)
End Property

Public Sub Clear()
    Erase nArray
End Sub

Public Function numDim() As Integer
    numDim = ndim(nArray)
End Function
Public Function size() As Long
    Dim i As Long
    Dim lngSize As Long
    For i = 1 To ndim(nArray)
        If i = i Then
            lngSize = UBound(nArray, i) + 1
        Else
            lngSize = lngSize * (UBound(nArray, i) + 1)
        End If
    Next i
    
    size = lngSize
End Function
Public Sub Display(Optional lngWidth As Long = 5)
    Select Case ndim(nArray)
        Case 1:
            Call PrintArray(lngWidth)
        Case 2:
           ' Call PrintMatrix
    End Select
End Sub
Private Sub PrintArray(lngWidth As Long)
    Dim i As Long
    Dim intBreak As Integer
    Dim strResult As String
    strResult = "["
    For i = LBound(nArray) To UBound(nArray)
        intBreak = intBreak + 1
        strResult = strResult & nArray(i) & ", "
        If intBreak = lngWidth Then
            strResult = strResult & Chr(10)
            intBreak = 0
        End If
    Next i
    
    strResult = Trim(strResult) & "]"
    Debug.Print strResult
End Sub
Function dblCast() As Double()
    Dim i, j, k As Long
    Dim dblArr() As Double
    Select Case ndim(nArray)
        Case 1:
            ReDim dblArr(0 To UBound(nArray))
            For i = 0 To UBound(nArray)
                dblArr(i) = CDbl(nArray(i))
            Next
    End Select
    dblCast = dblArr()
End Function
Function strCast() As String()
    Dim i, j, k As Long
    Dim strArr() As String
    Select Case ndim(nArray)
        Case 1:
            ReDim strArr(0 To UBound(nArray))
            For i = 0 To UBound(nArray)
                strArr(i) = CStr(nArray(i))
            Next
    End Select
    strCast = strArr()
End Function
