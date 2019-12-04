VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
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
' Length
' Resize
' IsNumberType -> NumberTypeValidate
' Item
' ItemNoCheck
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
Private vecList As List

Private Sub Class_Initialize()
    Set vecList = New List
End Sub

Public Property Get Self() As Vector
    Set Self = Me
End Property

Public Property Get Length() As Long
    Let Length = vecList.Length
End Property

Public Sub Resize(ByVal newLength As Long)
    Call vecList.Resize(newLength, True)
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

'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Let Item = vecList(index)
End Property

Public Property Let Item(ByVal index As Long, ByVal what As Variant)
    Call NumberTypeValidate(what)
    
    Let vecList(index) = what
End Property

Public Property Get ItemNoCheck(ByVal index As Long) As Variant
    Let ItemNoCheck = vecList.ItemNoCheck(index)
End Property

Public Property Let ItemNoCheck(ByVal index As Long, ByVal what As Variant)
    Let vecList.ItemNoCheck(index) = what
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generators
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Emptie(ByVal inputLength As Long) As Vector
With New Vector
    Call .Resize(inputLength)
    
    Set Emptie = .Self
End With
End Function

Public Function EmptyLike(ByVal vec As Vector) As Vector
    Set EmptyLike = Emptie(vec.Length)
End Function

Public Function Full(ByVal inputLength As Long, ByVal what As Variant) As Vector
    Call NumberTypeValidate(what)
    
    With Emptie(inputLength)
        Dim i As Long
        For i = 1 To inputLength
            Let .ItemNoCheck(i) = what
        Next i
        
        Set Full = .Self
    End With
End Function

Public Function Repeat(ByVal what As Variant, ByVal times As Long) As Vector
    Set Repeat = Full(times, what)
End Function

Public Function FullLike(ByVal vec As Vector, ByVal what As Variant) As Vector
    Set FullLike = Full(vec.Length, what)
End Function

Public Function Zeros(ByVal inputLength As Long) As Vector
    Set Zeros = Full(inputLength, 0)
End Function

Public Function ZerosLike(ByVal vec As Vector) As Vector
    Set ZerosLike = Zeros(vec.Length)
End Function

Public Function Ones(ByVal inputLength As Long) As Vector
    Set Ones = Full(inputLength, 1)
End Function

Public Function OnesLike(ByVal vec As Vector) As Vector
    Set OnesLike = Ones(vec.Length)
End Function

Public Function Arange(ByVal start As Variant, ByVal ends As Variant, _
                       Optional ByVal step As Variant = 1) As Vector
    Dim extraLength As Long
    Let extraLength = Int((ends - start) / step)

    If extraLength < 0 Then Let extraLength = 0

    With Emptie(1 + extraLength)
        Let .ItemNoCheck(1) = start

        Dim i As Long
        For i = 2 To .Length
            Let .ItemNoCheck(i) = .ItemNoCheck(i - 1) + step
        Next i

        Set Arange = .Self
    End With
End Function

Public Function LinSpace(ByVal start As Variant, ByVal ends As Variant, _
                         Optional ByVal num As Long = 50) As Vector
With Emptie(num + 1)
    Dim width As Double
    Let width = (ends - start) / num

    Let .ItemNoCheck(1) = start

    Dim i As Long
    For i = 2 To num
        Let .ItemNoCheck(i) = .ItemNoCheck(1) + (i - 1) * width ' .ItemNoCheck(i - 1) + width
    Next i

    Let .ItemNoCheck(num + 1) = ends

    Set LinSpace = .Self
End With
End Function

Public Function LogSpace(ByVal start As Variant, ByVal ends As Variant, _
                         Optional ByVal num As Long = 50, _
                         Optional ByVal base As Double) As Vector
    If base = 0 Then Let base = Math.Exp(1)

    With LinSpace(start, ends, num)
        Dim i As Long
        For i = 1 To .Length
            Let .ItemNoCheck(i) = base ^ (.ItemNoCheck(i))
        Next i

        Set LogSpace = .Self
    End With
End Function

Public Function GeomSpace(ByVal start As Variant, ByVal ends As Variant, _
                          Optional ByVal num As Long = 50, _
                          Optional ByVal base As Double) As Vector
    If base = 0 Then Let base = Math.Exp(1)

    With LinSpace(Log(start) / Log(base), Log(ends) / Log(base), num)
        Let .ItemNoCheck(1) = start

        Dim i As Long
        For i = 2 To .Length - 1
            Let .ItemNoCheck(i) = base ^ (.ItemNoCheck(i))
        Next i

        Let .ItemNoCheck(.Length) = ends

        Set GeomSpace = .Self
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' V -> V
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Absolute(ByVal vec As Vector) As Vector
With EmptyLike(vec)
    Dim i As Long
    For i = 1 To vec.Length
        Let .ItemNoCheck(i) = Abs(vec.ItemNoCheck(i))
    Next i
    Set Absolute = .Self
End With
End Function

Public Function Power(ByVal vec As Vector, ByVal p As Variant) As Vector
With EmptyLike(vec)
    Dim i As Long
    For i = 1 To vec.Length
        Let .ItemNoCheck(i) = vec.ItemNoCheck(i) ^ p
    Next i
    Set Power = .Self
End With
End Function

Public Function Translate(ByVal vec As Vector, ByVal scalar As Variant) As Vector
With EmptyLike(vec)
    Dim i As Long
    For i = 1 To vec.Length
        Let .ItemNoCheck(i) = scalar + vec.ItemNoCheck(i)
    Next i

    Set Translate = .Self
End With
End Function

Public Function Scaling(ByVal vec As Vector, ByVal scalar As Variant) As Vector
With EmptyLike(vec)
    Dim i As Long
    For i = 1 To vec.Length
        Let .ItemNoCheck(i) = scalar * vec.ItemNoCheck(i)
    Next i

    Set Scaling = .Self
End With
End Function

Public Function Negate(ByVal vec As Vector) As Vector
    Set Negate = Scaling(vec, -1)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' V -> F
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Max(ByVal vec As Vector) As Variant
With vec
    Let Max = .Self(1)
    
    Dim i As Long
    For i = 2 To vec.Length
        If .ItemNoCheck(i) > Max Then
            Let Max = .ItemNoCheck(i)
        End If
    Next i
End With
End Function

Public Function Min(ByVal vec As Vector) As Variant
With vec
    Let Min = .Self(1)
    
    Dim i As Long
    For i = 2 To vec.Length
        If .ItemNoCheck(i) < Min Then
            Let Min = .ItemNoCheck(i)
        End If
    Next i
End With
End Function

Public Function Sum(ByVal vec As Vector) As Variant
With vec
    Dim i As Long
    For i = 1 To vec.Length
        Let Sum = Sum + .ItemNoCheck(i)
    Next i
End With
End Function

Public Function Norm(ByVal vec As List, Optional ByVal p As Variant = 2) As Variant
    If p = -1 Then
        Let Norm = Max(Absolute(vec))
    ElseIf p = 0 Then
    ElseIf p = 1 Then
        Let Norm = Sum(Absolute(vec))
    Else
        Let Norm = Sum(Power(Absolute(vec), p)) ^ (1 / p)
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' V x V -> V
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub VectorLengthValidate(ByVal vec1 As Vector, ByVal vec2 As Vector)
    If vec1.Length <> vec2.Length Then
        Call Err.Raise(1, , "Vector length mismatch")
    End If
End Sub

Public Function Plus(ByVal vec1 As Vector, ByVal vec2 As Vector) As Vector
    Call VectorLengthValidate(vec1, vec2)
    
    With EmptyLike(vec1)
        Dim i As Long
        For i = 1 To vec1.Length
            Let .ItemNoCheck(i) = vec1.ItemNoCheck(i) + vec2.ItemNoCheck(i)
        Next i

        Set Plus = .Self
    End With
End Function

Public Function Minus(ByVal vec1 As Vector, ByVal vec2 As Vector) As Vector
    Call VectorLengthValidate(vec1, vec2)
    
    With EmptyLike(vec1)
        Dim i As Long
        For i = 1 To vec1.Length
            Let .ItemNoCheck(i) = vec1.ItemNoCheck(i) - vec2.ItemNoCheck(i)
        Next i

        Set Minus = .Self
    End With
End Function

Public Function Multiply(ByVal vec1 As Vector, ByVal vec2 As Vector) As Vector
    Call VectorLengthValidate(vec1, vec2)

    With EmptyLike(vec1)
        Dim i As Long
        For i = 1 To vec1.Length
            Let .ItemNoCheck(i) = vec1.ItemNoCheck(i) * vec2.ItemNoCheck(i)
        Next i

        Set Multiply = .Self
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' V x V -> F
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InnerProduct(ByVal vec1 As Vector, ByVal vec2 As Vector) As Variant
    Call VectorLengthValidate(vec1, vec2)

    Let InnerProduct = Sum(Multiply(vec1, vec2))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FromList() As Vector

End Function

Public Function FromArray() As Vector

End Function

Public Function FromCollection() As Vector

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As Vector
With New Vector
    Call .Resize(Me.Length)
    
    Dim i As Long
    For i = 1 To Me.Length
        Let .ItemNoCheck(i) = Me.ItemNoCheck(i)
    Next i
    
    Set Clone = .Self
End With
End Function

Public Sub ToImmediate(Optional ByVal head As Long = 50, Optional ByVal tail As Long = 50)
    Call vecList.ToImmediate(head, tail)
End Sub

Public Sub ToImmediateHead(Optional ByVal head As Long = 50)
    Call vecList.ToImmediateHead(head)
End Sub

Public Sub ToImmediateTail(Optional ByVal tail As Long = 50)
    Call vecList.ToImmediateTail(tail)
End Sub
