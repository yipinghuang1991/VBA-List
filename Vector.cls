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

'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Let Item = vecList(index)
End Property

Public Property Let Item(ByVal index As Long, ByVal what As Variant)
    If IsNumberType(what) Then
        Let vecList(index) = what
    Else
        Call Err.Raise(1, , "Not a number")
    End If
End Property

Public Function Clone() As Vector
With New Vector

End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VecAbs(ByVal lst As List) As List
With List.EmptyList(lst.Length)
    Dim i As Long
    For i = 1 To lst.Length
        Let .Self(i) = Abs(lst(i))
    Next i
    
    Set VecAbs = .Self
End With
End Function

Public Function VecPow(ByVal lst As List, ByVal p As Variant) As List
With List.EmptyList(lst.Length)
    Dim i As Long
    For i = 1 To lst.Length
        Let .Self(i) = lst(i) ^ p
    Next i
    
    Set VecPow = .Self
End With
End Function

Public Function VecNorm(ByVal lst As List, Optional ByVal p As Variant = 2) As Variant
    If p = -1 Then
        Let VecNorm = Max(VecAbs(lst))
    ElseIf p = 0 Then
    ElseIf p = 1 Then
        Let VecNorm = Sum(VecAbs(lst))
    Else
        Let VecNorm = Sum(VecPow(VecAbs(lst), p)) ^ (1 / p)
    End If
End Function

Public Function VecSum(ByVal list1 As List, ByVal list2 As List) As List
    If list1.Length <> list2.Length Then
        Call Err.Raise(1, , "Error")
    End If
    
    With List.EmptyList(list1.Length)
        Dim i As Long
        For i = 1 To list1.Length
            Let .Self(i) = list1(i) + list2(i)
        Next i
        
        Set VecSum = .Self
    End With
End Function

Public Function VecProd(ByVal list1 As List, ByVal list2 As List) As List
    If list1.Length <> list2.Length Then
        Call Err.Raise(1, , "Error")
    End If
    
    With List.EmptyList(list1.Length)
        Dim i As Long
        For i = 1 To list1.Length
            Let .Self(i) = list1(i) * list2(i)
        Next i
        
        Set VecProd = .Self
    End With
End Function

Public Function ScalarProd(ByVal what As Variant, ByVal lst As List) As Variant
With List.EmptyList(lst.Length)
    Dim i As Long
    For i = 1 To lst.Length
        Let .Self(i) = what * lst(i)
    Next i
    
    Set ScalarProd = .Self
End With
End Function

Public Function VecDot(ByVal list1 As List, ByVal list2 As List) As Variant
    If list1.Length <> list2.Length Then
        Call Err.Raise(1, , "Error")
    End If
    
    Let VecDot = List.Sum(VecProd(list1, list2))
End Function



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

