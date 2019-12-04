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
' Implementation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Storage
    Vectors As List
    LengthX As Long
    LengthY As Long
End Type

Private This As Storage

Private Sub Class_Initialize()
    Set This.Vectors = New List
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
    Call This.Vectors.Resize(newLengthX, True)

    Dim i As Long
    
    If newLengthX > This.LengthX Then
        For i = This.LengthX + 1 To newLengthX
            Set This.Vectors(i) = New Vector
        Next i
    End If
    
    Let This.LengthX = newLengthX
    
    For i = 1 To newLengthX
        Call This.Vectors.ItemNoCheck(i).Resize(newLengthY)
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

'@DefaultMember
Public Property Get Item(ByVal indexX As Long, ByVal indexY As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Let Item = This.Vectors(indexX)(indexY)
End Property

Public Property Let Item(ByVal indexX As Long, ByVal indexY As Long, ByVal what As Variant)
    Call NumberTypeValidate(what)
    
    Let This.Vectors(indexX)(indexY) = what
End Property

Public Property Get ItemNoCheck(ByVal indexX As Long, ByVal indexY As Long) As Variant
    Let ItemNoCheck = This.Vectors.ItemNoCheck(indexX).ItemNoCheck(indexY)
End Property

Public Property Let ItemNoCheck(ByVal indexX As Long, ByVal indexY As Long, ByVal what As Variant)
    Call NumberTypeValidate(what)
    
    Let This.Vectors.ItemNoCheck(indexX).ItemNoCheck(indexY) = what
End Property

Public Function Copy() As Matrix
With New Matrix
    Call .Resize(This.LengthX, This.LengthY)
    
    Dim i As Long
    Dim j As Long
    For i = 1 To This.LengthX
        For j = 1 To This.LengthY
            Let .ItemNoCheck(i, j) = This.Vectors.ItemNoCheck(i).ItemNoCheck(j)
        Next j
    Next i
    
    Set Copy = .Self
End With
End Function

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
Public Function ToArray() As Variant()

End Function

Public Function ToArrayOfArrays() As Variant()

End Function

