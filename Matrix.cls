VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Type")
'@PredeclaredId
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Dependency Graphs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ## Implementation ##
'
' Class_Initialize
' Self
' LengthX
' LengthY
' Resize
' IsIndexValid -> IndexValidate
' IsNumberType -> NumberTypeValidate
' ItemNoCheck
' Item
'
' ## Generators ##
'
' Emptie ---> EmptyLike
'         |-> Full -------> Repeat
'         |             |-> FullLike
'         |             |-> Zeros -> ZerosLike
'         |             |-> Ones --> OnesLike
'         |-> Arange
'         |-> LinSpace ---> LogSpace
'                       |-> GeomSpace
' ## V -> V ##
'
' (EmptyLike) ---> Absolute
'              |-> Power
'              |-> Translate
'              |-> Scaling -> Negate
'
' ## V -> F ##
'
' Max
' Min
' Sum
' Norm
'
' ## V x V -> V ##
'
' VectorLengthValidate -|
' (EmptyLike) --------------> Plus
'                         |-> Minus
'                         |-> Multiply
' ## V x V -> F ##
'
' (VectorLengthValidate) -> InnerProduct
'
' ## Constructors ##
'
' FromList
' FromArray
' FromCollection
'
' ## Output ##
'
' Clone
' ToImmediate ---> ToImmediateHead
'              |-> ToImmediateHead

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Storage
    Lists As List
    LengthX As Long
    LengthY As Long
End Type

Private This As Storage

Private Sub Class_Initialize()
    Set This.Lists = New List
End Sub

Public Property Get Self() As Matrix
    Set Self = Me
End Property

Public Property Get LengthX() As Long
    Let LengthX = This.LengthX
End Property

Public Property Get LengthY() As Long
    Let LengthY = This.LengthY
End Property

'Public Property Get Shape() As Long()
'    Dim shapeNew(2) As Long
'    Let shapeNew(1) = meLengthX
'    Let shapeNew(2) = meLengthY
'    Let Shape = shapeNew
'End Property

Public Sub Resize(ByVal newLengthX As Long, ByVal newLengthY As Long)
    Call This.Lists.Resize(newLengthX, True)

    Dim i As Long
    
    If newLengthX > This.LengthX Then
        For i = This.LengthX + 1 To newLengthX
            Set This.Lists(i) = New List
        Next i
    End If
    
    Let This.LengthX = newLengthX
    
    For i = 1 To newLengthX
        Call This.Lists.ItemNoCheck(i).Resize(newLengthY)
    Next i
    
    Let This.LengthY = newLengthY
End Sub

Private Function IsIndexValid(ByVal indexX As Long, ByVal indexY As Long) As Boolean
    If (1 <= indexX) And (indexX <= This.LengthX) And (1 <= indexY) And (indexY <= This.LengthY) Then
        Let IsIndexValid = True
    End If
End Function

Private Sub IndexValidate(ByVal indexX As Long, ByVal indexY As Long)
    If Not IsIndexValid(indexX, indexY) Then
        Call Err.Raise(9)
    End If
End Sub

Private Function IsNumberType(ByRef what As Variant) As Boolean
    Dim typeOfWhat As Long
    Let typeOfWhat = VarType(what)
    
    If typeOfWhat = 3 Then
        Let IsNumberType = True
    ElseIf typeOfWhat = 4 Then
        Let IsNumberType = True
    ElseIf typeOfWhat = 5 Then
        Let IsNumberType = True
    ElseIf typeOfWhat = 20 Then
        Let IsNumberType = True
    ElseIf typeOfWhat = 2 Then
        Let IsNumberType = True
    ElseIf typeOfWhat = 17 Then
        Let IsNumberType = True
    ElseIf typeOfWhat = 14 Then
        Let IsNumberType = True
    End If
End Function

Private Sub NumberTypeValidate(ByRef what As Variant)
    If Not IsNumberType(what) Then
        Call Err.Raise(1, , "Expected a number")
    End If
End Sub

Public Property Get ItemNoCheck(ByVal indexX As Long, ByVal indexY As Long) As Variant
    Let ItemNoCheck = This.Lists.ItemNoCheck(indexX).ItemNoCheck(indexY)
End Property

Public Property Let ItemNoCheck(ByVal indexX As Long, ByVal indexY As Long, ByVal what As Variant)
    Let This.Lists.ItemNoCheck(indexX).ItemNoCheck(indexY) = what
End Property

'@DefaultMember
Public Property Get Item(ByVal indexX As Long, ByVal indexY As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Let Item = This.Lists(indexX)(indexY)
End Property

Public Property Let Item(ByVal indexX As Long, ByVal indexY As Long, ByVal what As Variant)
    Call NumberTypeValidate(what)
    Let This.Lists(indexX)(indexY) = what
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generators
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Emptie(ByVal newLengthX As Long, ByVal newLengthY As Long) As Matrix
With New Matrix
    Call .Resize(newLengthX, newLengthY)
    Set Emptie = .Self
End With
End Function

Public Function EmptyLike(ByVal mat As Matrix) As Matrix
    Set EmptyLike = Emptie(mat.LengthX, mat.LengthY)
End Function

Public Function Full(ByVal newLengthX As Long, ByVal newLengthY As Long, ByVal what As Variant) As Matrix
    Call NumberTypeValidate(what)
    
    With Emptie(newLengthX, newLengthY)
        Dim j As Long
        Dim i As Long
        For i = 1 To newLengthX
            For j = 1 To newLengthY
                Let .ItemNoCheck(i, j) = what
            Next j
        Next i
        
        Set Full = .Self
    End With
End Function

Public Function FullLike(ByVal mat As Matrix, ByVal what As Variant) As Matrix
    Set FullLike = Full(mat.LengthX, mat.LengthY, what)
End Function

Public Function Zeros(ByVal newLengthX As Long, ByVal newLengthY As Long) As Matrix
    Set Zeros = Full(newLengthX, newLengthY, 0)
End Function

Public Function ZerosLike(ByVal mat As Matrix) As Matrix
    Set ZerosLike = Zeros(mat.LengthX, mat.LengthY)
End Function

Public Function Ones(ByVal newLengthX As Long, ByVal newLengthY As Long) As Matrix
    Set Ones = Full(newLengthX, newLengthY, 1)
End Function

Public Function OnesLike(ByVal mat As Matrix) As Matrix
    Set OnesLike = Ones(mat.LengthX, mat.LengthY)
End Function

Public Function Eye(ByVal newLengthX As Long, Optional ByVal newLengthY As Long = -1, _
                    Optional ByVal diagonalOffset As Long) As Matrix
    If newLengthY = -1 Then
        Let newLengthY = newLengthX
    End If
    
    With Zeros(newLengthX, newLengthY)
        Dim i As Long
        If diagonalOffset >= 0 Then
            For i = 1 + diagonalOffset To IIf(newLengthX <= newLengthY, newLengthX, newLengthY)
                Let .ItemNoCheck(i - diagonalOffset, i) = 1
            Next i
        Else ' diagonalOffset < 0
            Let diagonalOffset = -diagonalOffset
            
            For i = 1 + diagonalOffset To IIf(newLengthX <= newLengthY, newLengthX, newLengthY)
                Let .ItemNoCheck(i, i - diagonalOffset) = 1
            Next i
        End If
        
        Set Eye = .Self
    End With
End Function

Public Function Identity(ByVal newLengthX As Long) As Matrix
    Set Identity = Eye(newLengthX)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' M -> M
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' M -> F
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' M x M -> M
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Plus(ByVal mat1 As Matrix, ByVal mat2 As Matrix) As Matrix

End Function

Public Function Minus(ByVal mat1 As Matrix, ByVal mat2 As Matrix) As Matrix

End Function

Public Function Multiply(ByVal mat1 As Matrix, ByVal mat2 As Matrix) As Matrix
With Emptie(mat1.LengthX, mat2.LengthY)
    If mat1.LengthY <> mat2.LengthX Then
        Call Err.Raise(1, , "Error")
    End If
    
    Dim temp As Variant
    
    Dim k As Long
    Dim j As Long
    Dim i As Long
    For i = 1 To mat1.LengthX
        For j = 1 To mat2.LengthY
            Let temp = 0
            
            For k = 1 To mat1.LengthY
                Let temp = temp + mat1.ItemNoCheck(i, k) * mat2.ItemNoCheck(k, j)
            Next k
            
            Let .ItemNoCheck(i, j) = temp
        Next j
    Next i
    
    Set Multiply = .Self
End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' M x M -> F
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FromArray(ByVal arr As Variant) As Matrix
    If List.ArrayDimension(arr) <> 2 Then
        Call Err.Raise(1, , "Required array of dimension 2")
    End If

    With New Matrix
        Call Resize(List.ArrayLen(arr, 1), List.ArrayLen(arr, 2))
        
        Dim offsetX As Long: Let offsetX = LBound(arr, 1) - 1
        Dim offsetY As Long: Let offsetY = LBound(arr, 2) - 1
        
        Dim i As Long
        Dim j As Long
        For j = 1 To .LengthY
            For i = 1 To .LengthX
                Let .Self(i, j) = arr(i + offsetX, j + offsetY)
            Next i
        Next j

        Set FromArray = .Self
    End With
End Function

Public Function FromArrayOfArrays(ByVal arr As Variant) As Matrix
    If List.ArrayDimension(arr) <> 1 Then
        Call Err.Raise(1, , "Required array of dimension 1")
    End If
    
    Dim subArr As Variant
    For Each subArr In arr
        If List.ArrayDimension(subArr) <> 1 Then
            Call Err.Raise(1, , "Required array of dimension 1")
        End If
    Next subArr
    
    Dim width As Long
    Let width = List.ArrayLen(arr(LBound(arr)))
    
    For Each subArr In arr
        If List.ArrayLen(subArr) <> width Then
            Call Err.Raise(1, , "Error")
        End If
    Next subArr
    
    With Emptie(List.ArrayLen(arr), width)
        Dim offsetX As Long: Let offsetX = LBound(arr, 1) - 1
        Dim offsetY As Long: Let offsetY = LBound(arr(LBound(arr))) - 1
        
        Dim i As Long
        Dim j As Long
        For j = 1 To .LengthY
            For i = 1 To .LengthX
                Let .Self(i, j) = arr(i + offsetX)(j + offsetY)
            Next i
        Next j
        
        Set FromArrayOfArrays = .Self
    End With
End Function

Public Function FromArgumentsOfArrays(ParamArray arr() As Variant) As Matrix
    Set FromArgumentsOfArrays = FromArrayOfArrays(arr)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As Matrix
With New Matrix
    Call .Resize(This.LengthX, This.LengthY)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To This.LengthX
        For j = 1 To This.LengthY
            Let .ItemNoCheck(i, j) = This.Lists.ItemNoCheck(i).ItemNoCheck(j)
        Next j
    Next i
    
    Set Clone = .Self
End With
End Function

Public Function ToArray() As Variant()

End Function

Public Function ToArrayOfArrays() As Variant()

End Function

Public Sub ToImmediate()
    Const n As Long = 8
    Const nLess2 As Long = n - 2
    
    Dim lengthOfn As String * n
    Dim lengthOfnLess2 As String * nLess2
    
    Dim j As Long
    Dim temp As String
    Dim i As Long
    
    If LengthY = 1 Then
        For i = 1 To LengthX
            If i = 1 Then
                Let temp = "["
            Else
                Let temp = Space(1)
            End If
            
            Let temp = temp & "[ "
            
            LSet lengthOfnLess2 = ItemNoCheck(i, 1)
            
            Let temp = temp & lengthOfnLess2 & " ]"
            
            If i = LengthX Then
                Let temp = temp & "]"
            End If
            
            Debug.Print temp
        Next i

    Else
        For i = 1 To LengthX
            If i = 1 Then
                Let temp = "["
            Else
                Let temp = Space(1)
            End If
            
            Let temp = temp & "[ "
            
            LSet lengthOfn = Left(ItemNoCheck(i, 1), nLess2) & ", "
            
            Let temp = temp & lengthOfn
            
            For j = 2 To LengthY - 1
                LSet lengthOfn = Left(ItemNoCheck(i, j), nLess2) & ", "
                Let temp = temp & lengthOfn
            Next j
            
            LSet lengthOfnLess2 = ItemNoCheck(i, LengthY)
            Let temp = temp & lengthOfnLess2 & " ]"
            
            If i = LengthX Then
                Let temp = temp & "]"
            End If
            
            Debug.Print temp
        Next i
    End If
End Sub
