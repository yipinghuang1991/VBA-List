VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule
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
' Item
' Clone
'

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
    Call vecList.ResizeJust(newLength)
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

Public Function FullLike(ByVal vec As Vector, ByVal what As Variant) As Vector
    Set FullLike = Full(vec.Length, what)
End Function

Public Function Repeat(ByVal what As Variant, ByVal times As Long) As Vector
    Set Repeat = Full(times, what)
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
' Unary Operation
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

Public Function Negate(ByVal vec As Vector) As Vector
With EmptyLike(vec)
    Dim i As Long
    For i = 1 To vec.Length
        Let .ItemNoCheck(i) = -vec.ItemNoCheck(i)
    Next i
    Set Negate = .Self
End With
End Function

'Public Function VecNorm(ByVal vec As List, Optional ByVal p As Variant = 2) As Variant
'    If p = -1 Then
'        Let VecNorm = Max(VecAbs(lst))
'    ElseIf p = 0 Then
'    ElseIf p = 1 Then
'        Let VecNorm = Sum(VecAbs(lst))
'    Else
'        Let VecNorm = Sum(VecPow(VecAbs(lst), p)) ^ (1 / p)
'    End If
'End Function
'
'Public Function VecSum(ByVal vec1 As List, ByVal vec2 As List) As List
'    If list1.Length <> list2.Length Then
'        Call Err.Raise(1, , "Error")
'    End If
'
'    With List.EmptyList(list1.Length)
'        Dim i As Long
'        For i = 1 To list1.Length
'            Let .Self(i) = list1(i) + list2(i)
'        Next i
'
'        Set VecSum = .Self
'    End With
'End Function
'
'Public Function VecProd(ByVal list1 As List, ByVal list2 As List) As List
'    If list1.Length <> list2.Length Then
'        Call Err.Raise(1, , "Error")
'    End If
'
'    With List.EmptyList(list1.Length)
'        Dim i As Long
'        For i = 1 To list1.Length
'            Let .Self(i) = list1(i) * list2(i)
'        Next i
'
'        Set VecProd = .Self
'    End With
'End Function
'
'Public Function ScalarProd(ByVal what As Variant, ByVal lst As List) As Variant
'With List.EmptyList(lst.Length)
'    Dim i As Long
'    For i = 1 To lst.Length
'        Let .Self(i) = what * lst(i)
'    Next i
'
'    Set ScalarProd = .Self
'End With
'End Function
'
'Public Function VecDot(ByVal list1 As List, ByVal list2 As List) As Variant
'    If list1.Length <> list2.Length Then
'        Call Err.Raise(1, , "Error")
'    End If
'
'    Let VecDot = List.Sum(VecProd(list1, list2))
'End Function

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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IsNumberType(ByRef what As Variant) As Boolean
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
