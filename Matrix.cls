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
Option Base 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Storage
    VectorList As List
    LengthX As Long
    LengthY As Long
End Type

Private This As Storage

Private Sub Class_Initialize()
    Set This.VectorList = New List
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
    Call This.VectorList.Resize(newLengthY)

    Dim tempList As List
    Dim i As Long

    If newLengthY > This.LengthY Then
        For i = This.LengthY + 1 To newLengthY
            Set tempList = New List
            Set This.VectorList(i) = tempList
        Next i
    End If

    Let This.LengthY = newLengthY

    For i = 1 To newLengthY
        Call This.VectorList(i).Resize(newLengthX)
    Next i

    Let This.LengthX = newLengthX
End Sub

'Public Function Copy() As List
'With New List
'    Call .Resize(This.theLength)
'
'    Dim i As Long
'    For i = 1 To This.theLength
'        Call .LetOrSet(i, Me(i))
''        If IsObject(Me(i)) Then
''            Set .Self(i) = Me(i)
''        Else
''            Let .Self(i) = Me(i)
''        End If
'    Next i
'
'    Set Copy = .Self
'End With
'End Function

Private Function IsIndexValid(ByVal indexX As Long, ByVal indexY As Long) As Boolean
    If (1 <= indexY) And (indexY <= This.LengthY) And (1 <= indexX) And (indexX <= This.LengthX) Then
            Let IsIndexValid = True
    End If
End Function

Private Sub IndexValidate(ByVal indexX As Long, ByVal indexY As Long)
    If Not IsIndexValid(indexX, indexY) Then
        Call Err.Raise(9)
    End If
End Sub

'@DefaultMember
Public Property Get Item(ByVal indexX As Long, ByVal indexY As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Call IndexValidate(indexX, indexY)
    
    If IsObject(This.VectorList(indexX)(indexY)) Then
        Set Item = This.VectorList(indexX)(indexY)
    Else
        Let Item = This.VectorList(indexX)(indexY)
    End If
End Property

Public Property Let Item(ByVal indexX As Long, ByVal indexY As Long, ByVal what As Variant)
    Call IndexValidate(indexX, indexY)
    
    Let This.VectorList(indexX)(indexY) = what
End Property

Public Property Set Item(ByVal indexX As Long, ByVal indexY As Long, ByVal what As Variant)
    Call IndexValidate(indexX, indexY)
    
    Set This.VectorList(indexX)(indexY) = what
End Property

Public Property Let Assign(ByVal indexX As Long, ByVal indexY As Long, ByVal what As Variant)
    Call IndexValidate(indexX, indexY)
    
    If IsObject(what) Then
        Set Me(indexX, indexY) = what
    Else
        Let Me(indexX, indexY) = what
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Information
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Sub Clear(Optional ByVal indexX As Long, Optional ByVal indexY As Long)
'    Dim i As Long
'
'    If indexX And indexY Then ' (indexX <> 0) And (indexY <> 0)
'        Call This.listX(indexX).Clear(indexY)
'    ElseIf indexX And (indexY = 0) Then ' (indexX <> 0) And (indexY = 0)
'        Call This.listX(indexX).Clear
'    ElseIf (indexX = 0) And indexY Then ' (indexX = 0) And (indexY <> 0)
'        For i = 1 To This.theLengthX
'            Call This.listX(i).Clear(indexY)
'        Next i
'    Else ' (indexX = 0) And (indexY = 0)
'        For i = 1 To This.theLengthX
'            Call This.listX(i).Clear
'        Next i
'    End If
'End Sub

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
                Let .Assign(i, j) = arr(i + offsetX, j + offsetY)
            Next i
        Next j

        Set FromArray = .Self
    End With
End Function

Public Function FromArrayOfArrays(ByVal arr As Variant) As Matrix
    If List.ArrayDimension(arr) <> 1 Then
        Call Err.Raise(1, , "Required array of dimension 1")
    End If
    
    Dim v As Variant
    For Each v In arr
        If List.ArrayDimension(v) <> 1 Then
            Call Err.Raise(1, , "Required array of dimension 1")
        End If
    Next v
    
    With New Matrix
        Call Resize(List.ArrayLen(arr), List.ArrayLen(arr(LBound(arr))))
        
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function ToArray() As Variant()
'
'End Function

'Public Function ToArrayOfArrays() As Variant()
'
'End Function
