VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
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
' Self
' Capacity -----> Class_Initialize
'             |-> Resize
'             --> TrimToSize
' Length = Count
' IsIndexValid -> IndexValidate
' ItemNoCheck ~~> Item
'
' ## Generators ##
'
' Emptie ---> EmptyLike
'         |-> Full ---> FullLike
'                   |-> Repeat
'
' ## Information ##
'
' (IndexValidate) ~~~> LastIndexOf
'                  |~> IndexOf ---> Contains
'                               |-> NumberOf
'                               --> IndexOfEach <-(Add)
'
' ## Operations that Preserves the List Length ##
'
' AsType
' SetRange
' Reverse
' Invert
' BitInvert
' Wipe
' Copy
' Move
'
' (IndexValidate) ~> Shift
'
' Permute ~> Swap
' SwapForSort ~> Sort
'        |-----> SortAndGetPermutation
'
' ## Operations that Alters the List Length ##
'
' Clear
'                     |-> InsertRange -> AddRange
' (IndexValidate) --  |-> Insert ------> Prepend
'                  |--|              |-> Append = Add
' (Shift) ----------  |-> RemoveRange
'                     |-> RemoveAt ----> Pop <-(IndexValidate)
'                                    |-> Remove <-(IndexOf)
'
' (Full) ----------  |-------> GetUniqueIndex --
' (GetIndexList   |--|                         |-> Unique
'  FromBoolList) --  |--+--|-> Keep ------------
' Map ------------------|  |-> KeepNot ----------> RemoveEach <-(IndexOfEach)
'
' ## Information between Lists ##
'
'                     |------------------------------------------
'                     |--------------------------|              |
'                     |                          V              V
' (Contains) -> IsSubsetOf -> IsSupersetOf -> IsEqualTo ---> IsProperSubsetOf
'                  |             |                       |-> IsProperSubsetOf
'                  |             |                              ^
'                  |             -------------------------------|
'                  |
'                  |  |------------------------------------------
'                  |  |--------------------------|              |
'                  V  |                          V              V
'               IsSubset ---> IsSuperset ---> IsEqual -----> IsProperSubset
'                                |                       |-> IsProperSubset
'                                |                              ^
'                                |------------------------------|
'
' ## Operations and Constructors among Lists ##
'
' (Full) -----------  |-> IntersectWith --> Intersect <-------
' (Contains) ------|  |                                      |
' (Keep) ----------|--|                        |--------- (Clone) -------|
' (GetIndexList    |  |                        V                         V
'  FromBoolList) ---  |-> DifferenceWith -> Difference -> UnionWith -> Union
'                                                            ^
' (InsertRange) ---------------------------------------------|
'
' ## Constructors ##
'
' (Clone) -------|
' (InsertRange) ---> Concatenate
'
' (Emptie) -------  |-> MapTo
'                |--|-> FromCollection
' (ItemNoCheck) --  |-> FromArray -> FromArguments
'                          ^
' (ArrayDimension) --------|
'
' (Add) -> FromEnumerable
'
' ## Output ##
'
' Clone
' GetRange
' CopyTo (Pending)
' ToArray
' ToCollection
' ToImmediate ---> ToImmediateHead
'              |-> ToImmediateTail
'
' ## Array Helper Functions ##
'
' ArrayDimension -> ArrayLen -> IsArrayEmpty
'
' ## Math Functions ##
'
' Floor
' Ceil
' Round (Pending)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Storage
    arr() As Variant
    Capacity As Long
    Length As Long
End Type

Private This As Storage

Public Property Get Self() As List
    Set Self = Me
End Property

Public Property Get Capacity() As Long
    Let Capacity = This.Capacity
End Property

Public Property Let Capacity(ByVal newCapacity As Long)
    If newCapacity > 0 Then
        Let This.Capacity = newCapacity
        ReDim Preserve This.arr(1 To This.Capacity)
        
        If This.Capacity < This.Length Then
            Let This.Length = This.Capacity
        End If
    ElseIf newCapacity = 0 Then
        Let This.Capacity = 1
        ReDim This.arr(1 To This.Capacity)
        Let This.Length = 0
    Else ' newCapacity < 0
        Call Err.Raise(1, , "Expected nonnegative integer")
    End If
End Property

Private Sub Class_Initialize()
    Let Capacity = 0
End Sub

Public Sub Resize(ByVal newLength As Long, Optional ByVal exact As Boolean)
    If exact Then
        If newLength > 0 Then
            Let This.Capacity = newLength
            ReDim Preserve This.arr(1 To This.Capacity)
            Let This.Length = newLength
        ElseIf newLength = 0 Then
            Let Capacity = 0
        Else ' newLength < 0
            Call Err.Raise(1, , "Expected nonnegative integer")
        End If
    Else
        If newLength >= 0 Then
            If newLength > This.Length Then
                If newLength > This.Capacity Then
                    Do
                        Let This.Capacity = 2 * This.Capacity
                    Loop Until newLength <= This.Capacity
                    
                    ReDim Preserve This.arr(1 To This.Capacity)
                End If
            Else ' newLength <= This.Length, equality is redundant
                Dim i As Long
                For i = newLength + 1 To This.Length
                    Let This.arr(i) = Empty
                Next i
            End If
            
            Let This.Length = newLength
        Else ' newLength < 0
            Call Err.Raise(1, , "Expected nonnegative integer")
        End If
    End If
End Sub

Public Sub TrimToSize()
    Let Capacity = This.Length
End Sub

Public Property Get Length() As Long
    Let Length = This.Length
End Property

Public Property Get Count() As Long
    Let Count = This.Length
End Property

Public Function IsIndexValid(ByVal index As Long) As Boolean ' Default to False
    If (1 <= index) And (index <= This.Length) Then
        Let IsIndexValid = True
    End If
End Function

Public Sub IndexValidate(ByVal index As Long)
    If Not IsIndexValid(index) Then
        Call Err.Raise(9)
    End If
End Sub

Public Property Get ItemNoCheck(ByVal index As Long)
    If IsObject(This.arr(index)) Then
        Set ItemNoCheck = This.arr(index)
    Else
        Let ItemNoCheck = This.arr(index)
    End If
End Property

Public Property Let ItemNoCheck(ByVal index As Long, ByVal what As Variant)
    If IsObject(what) Then
        Set This.arr(index) = what
    Else
        Let This.arr(index) = what
    End If
End Property

'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If index <= This.Length Then
        If IsObject(This.arr(index)) Then ' Let Item = ItemNoCheck(index)
            Set Item = This.arr(index)
        Else
            Let Item = This.arr(index)
        End If
    Else
        Call Err.Raise(9)
    End If
End Property

Public Property Let Item(ByVal index As Long, ByVal what As Variant)
    If index <= This.Length Then
        If IsObject(what) Then ' Let ItemNoCheck(index) = what
            Set This.arr(index) = what
        Else
            Let This.arr(index) = what
        End If
    Else
        Call Err.Raise(9)
    End If
End Property

' Redundant
Public Property Set Item(ByVal index As Long, ByVal what As Variant)
    If index <= This.Length Then
        Set This.arr(index) = what
    Else
        Call Err.Raise(9)
    End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generators
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Emptie(ByVal inputLength As Long) As List
With New List
    Call .Resize(inputLength, True)
    Set Emptie = .Self
End With
End Function

Public Function EmptyLike(ByVal lst As List) As List
    Set EmptyLike = Emptie(lst.Length)
End Function

Public Function Full(ByVal inputLength As Long, ByVal what As Variant) As List
With Emptie(inputLength)
    Dim i As Long
    For i = 1 To inputLength
        Let .ItemNoCheck(i) = what
    Next i
    
    Set Full = .Self
End With
End Function

Public Function FullLike(ByVal lst As List, ByVal what As Variant) As List
    Set FullLike = Full(lst.Length, what)
End Function

Public Function Repeat(ByVal what As Variant, ByVal times As Long) As List
    Set Repeat = Full(times, what)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Information
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastIndexOf(ByVal what As Variant, _
                            Optional ByVal indexStart As Long, _
                            Optional ByVal indexLength As Long = -1) As Long ' Default to 0
    If indexStart Then ' indexStart <> 0
        If (indexStart < 1) Or (indexStart > This.Length) Then ' Call IndexValidate(indexStart)
            Call Err.Raise(9)
        End If
    Else ' indexStart = 0
        Let indexStart = This.Length ' Default
    End If
    
    Dim indexEnd As Long
    
    If indexLength = -1 Then
        Let indexEnd = 1
    Else
        Let indexEnd = indexStart - indexLength + 1
        
        If indexEnd < 1 Then
            Let indexEnd = 1
        End If
    End If
    
    Dim i As Long
    
    If IsObject(what) Then
        For i = indexStart To indexEnd Step -1
            If IsObject(This.arr(i)) Then
                If This.arr(i) Is what Then
                    Let LastIndexOf = i
                    Exit For
                End If
            End If
        Next i
    Else
        For i = indexStart To indexEnd Step -1
            If Not IsObject(This.arr(i)) Then
                If This.arr(i) = what Then
                    Let LastIndexOf = i
                    Exit For
                End If
            End If
        Next i
    End If
End Function

Public Function IndexOf(ByVal what As Variant, _
                        Optional ByVal indexStart As Long, _
                        Optional ByVal indexLength As Long = -1) As Long ' Default to 0
    If indexStart Then ' indexStart <> 0
        If (indexStart < 1) Or (indexStart > This.Length) Then ' Call IndexValidate(indexStart)
            Call Err.Raise(9)
        End If
    Else ' indexStart = 0
        Let indexStart = 1 ' Default
    End If
    
    Dim indexEnd As Long
    
    If indexLength = -1 Then
        Let indexEnd = This.Length
    Else
        Let indexEnd = indexStart + indexLength - 1
        
        If indexEnd > This.Length Then
            Let indexEnd = This.Length
        End If
    End If
    
    Dim i As Long
    
    If IsObject(what) Then
        For i = indexStart To indexEnd
            If IsObject(This.arr(i)) Then
                If This.arr(i) Is what Then
                    Let IndexOf = i
                    Exit For
                End If
            End If
        Next i
    Else
        For i = indexStart To indexEnd
            If Not IsObject(This.arr(i)) Then
                If This.arr(i) = what Then
                    Let IndexOf = i
                    Exit For
                End If
            End If
        Next i
    End If
End Function

Public Function Contains(ByVal what As Variant) As Boolean
    Let Contains = IndexOf(what) ' <> 0
End Function

Public Function NumberOf(ByVal what As Variant) As Long ' Default to 0
    Dim thisLength As Long
    Let thisLength = This.Length
    
    Dim i As Long ' Default to 0
    Do
        Let i = IndexOf(what, i + 1)
        If i Then ' i > 0, i.e., i <> 0
            Let NumberOf = NumberOf + 1
        End If
    Loop While i And (i < thisLength)
End Function

Public Function IndexOfEach(ByVal what As Variant) As List
With New List
    Dim thisLength As Long
    Let thisLength = This.Length
    
    Dim i As Long ' Default to 0
    Do
        Let i = IndexOf(what, i + 1)
        If i Then ' i > 0, i.e., i <> 0
            Call .Add(i)
        End If
    Loop While i And (i < thisLength)
    
    Set IndexOfEach = .Self
End With
End Function

Public Function GetIndexListFromBoolList(ByVal boolList As List) As List
    Dim indexList As List
    Set indexList = List.Emptie(boolList.Length)
    
    Dim j As Long ' Default to 0
    Dim i As Long
    For i = 1 To boolList.Length
        If boolList.ItemNoCheck(i) Then
            Let j = j + 1
            Let indexList.ItemNoCheck(j) = i
        End If
    Next i
    
    Call indexList.Resize(j)
    
    Set GetIndexListFromBoolList = indexList
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Operations that Preserve the List Length
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AsType(ByVal whatType As Variant)
    Dim i As Long

    Select Case whatType
        Case VbVarType.vbBoolean
            For i = 1 To This.Length
                Let This.arr(i) = CBool(This.arr(i))
            Next i
        Case VbVarType.vbDate
            For i = 1 To This.Length
                Let This.arr(i) = CDate(This.arr(i))
            Next i
        Case VbVarType.vbDouble
            For i = 1 To This.Length
                Let This.arr(i) = CDbl(This.arr(i))
            Next i
        Case VbVarType.vbLong
            For i = 1 To This.Length
                Let This.arr(i) = CLng(This.arr(i))
            Next i
        Case VbVarType.vbLongLong
            For i = 1 To This.Length
                Let This.arr(i) = CLngLng(This.arr(i))
            Next i
        Case VbVarType.vbSingle
            For i = 1 To This.Length
                Let This.arr(i) = CSng(This.arr(i))
            Next i
        Case VbVarType.vbString
            For i = 1 To This.Length
                Let This.arr(i) = CStr(This.arr(i))
            Next i
        Case VbVarType.vbVariant
            For i = 1 To This.Length
                Let This.arr(i) = CVar(This.arr(i))
            Next i
        Case "Real", "Numeric", "Numerical", "Number"
            For i = 1 To This.Length
                Let This.arr(i) = Val(This.arr(i))
            Next i
        Case Else
            Call Err.Raise(1, , "Conversion not supported")
    End Select
End Sub

Public Sub SetRange(ByVal index As Long, ByVal lst As List)
    If index + lst.Length - 1 <= This.Length Then
        Dim offset As Long
        Let offset = index - 1
        
        Dim i As Long
        For i = index To lst.Length + offset
            Let ItemNoCheck(i) = lst(i - offset)
        Next i
    Else
        Call Err.Raise(1, , "Error")
    End If
End Sub

Public Sub Reverse()
    Dim thisLength As Long
    Let thisLength = This.Length
    
    If thisLength Then ' thisLength <> 0
        Dim arr() As Variant
        ReDim arr(1 To This.Capacity) As Variant
        
        Dim offset As Long
        Let offset = thisLength + 1
        
        Dim i As Long
        For i = 1 To thisLength
            If IsObject(This.arr(i)) Then
                Set arr(i) = This.arr(offset - i)
            Else
                Let arr(i) = This.arr(offset - i)
            End If
        Next i
        
        Let This.arr = arr
    End If
End Sub

Public Sub Invert()
    Dim i As Long
    For i = 1 To This.Length
        Let This.arr(i) = Not This.arr(i)
    Next i
End Sub

Public Sub BitInvert()
    Dim i As Long
    For i = 1 To This.Length
        If This.arr(i) Then ' This.Arr(i) <> 0
            Let This.arr(i) = 0
        Else
            Let This.arr(i) = 1
        End If
    Next i
End Sub

Public Sub Wipe(Optional ByVal index As Long)
    If index Then ' index <> 0
        If index <= This.Length Then
            Let This.arr(index) = Empty
        Else
            Call Err.Raise(9)
        End If
    Else ' index = 0
        Dim i As Long
        For i = 1 To This.Length
            Let This.arr(i) = Empty
        Next i
    End If
End Sub

Public Sub Copy(ByVal index As Long, ByVal indexFrom As Long)
    If (index <= This.Length) And (indexFrom <= This.Length) Then
        If index <> indexFrom Then
            Let This.arr(index) = This.arr(indexFrom)
        End If
    Else
        Call Err.Raise(9)
    End If
End Sub

Public Sub Move(ByVal index As Long, ByVal indexFrom As Long)
    If (index <= This.Length) And (indexFrom <= This.Length) Then
        If index <> indexFrom Then
            Let This.arr(index) = This.arr(indexFrom)
            Let This.arr(indexFrom) = Empty
        End If
    Else
        Call Err.Raise(9)
    End If
End Sub

Public Sub Shift(ByVal offset As Long, _
                 Optional ByVal indexStart As Long, _
                 Optional ByVal indexLength As Long = -1)
    Dim thisLength As Long
    Let thisLength = This.Length
    
    If indexStart Then ' indexStart <> 0
        If (indexStart < 1) Or (indexStart > thisLength) Then
            Call Err.Raise(9)
        End If
    Else ' indexStart = 0
        Let indexStart = 1 ' Default
    End If
    
    Dim indexEnd As Long
    
    If indexLength = -1 Then
        Let indexEnd = thisLength
    Else
        Let indexEnd = indexStart + indexLength - 1
        
        If indexEnd > thisLength Then
            Let indexEnd = thisLength
        End If
    End If
    
    Dim i As Long
    
    If offset > 0 Then
        For i = IIf(indexEnd + offset <= thisLength, indexEnd + offset, thisLength) To indexStart + offset Step -1
            If IsObject(This.arr(i - offset)) Then
                Set This.arr(i) = This.arr(i - offset)
            Else
                Let This.arr(i) = This.arr(i - offset)
            End If
        Next i
        
        For i = IIf(indexStart + offset - 1 <= thisLength, indexStart + offset - 1, thisLength) To indexStart Step -1
            Let This.arr(i) = Empty
        Next i
    ElseIf offset < 0 Then
        For i = IIf(indexStart + offset >= 1, indexStart + offset, 1) To indexEnd + offset
            If IsObject(This.arr(i - offset)) Then
                Set This.arr(i) = This.arr(i - offset)
            Else
                Let This.arr(i) = This.arr(i - offset)
            End If
        Next i
        
        For i = IIf(indexEnd + offset + 1 > 0, indexEnd + offset + 1, 1) To indexEnd
            Let This.arr(i) = Empty
        Next i
    End If
End Sub

Public Sub Permute(ByVal indexList As List)
    Dim thisLength As Long: Let thisLength = This.Length
    Dim listLength As Long: Let listLength = indexList.Length
    
    Dim i As Long
    For i = 1 To listLength
        If indexList(i) > thisLength Then
            Call Err.Raise(9)
        End If
    Next i
    
    If listLength Then ' listLength <> 0
        Dim temp As Variant
        
        If IsObject(This.arr(indexList(listLength))) Then
            Set temp = This.arr(indexList(listLength))
        Else
            Let temp = This.arr(indexList(listLength))
        End If
        
        For i = listLength To 2 Step -1
            If IsObject(This.arr(indexList(i - 1))) Then
                Set This.arr(indexList(i)) = This.arr(indexList(i - 1))
            Else
                Let This.arr(indexList(i)) = This.arr(indexList(i - 1))
            End If
        Next i
        
        If IsObject(temp) Then
            Set This.arr(indexList(1)) = temp
        Else
            Let This.arr(indexList(1)) = temp
        End If
    End If
End Sub

Public Sub Swap(ByVal index1 As Long, ByVal index2 As Long)
    If (index1 <= This.Length) And (index2 <= This.Length) Then
        Dim temp As Variant

        If IsObject(This.arr(index2)) Then
            Set temp = This.arr(index2)
        Else
            Let temp = This.arr(index2)
        End If

        If IsObject(This.arr(index1)) Then
            Set This.arr(index2) = This.arr(index1)
        Else
            Let This.arr(index2) = This.arr(index1)
        End If

        If IsObject(This.arr(index2)) Then
            Set This.arr(index1) = temp
        Else
            Let This.arr(index1) = temp
        End If
    Else
        Call Err.Raise(9)
    End If
End Sub

Public Sub SwapForSort(ByVal index1 As Long, ByVal index2 As Long)
    Dim temp As Variant
    Let temp = This.arr(index2)
    Let This.arr(index2) = This.arr(index1)
    Let This.arr(index1) = temp
End Sub

Public Sub Sort(Optional ByVal isDescending As Boolean) ' Bubble sort, ascending
    Dim swapped As Boolean
    Dim i As Long
    Dim temp As Variant
    
    If isDescending Then
        Do
            Let swapped = False
            For i = 1 To This.Length - 1
                If This.arr(i) < This.arr(i + 1) Then
                    Let temp = This.arr(i + 1)
                    Let This.arr(i + 1) = This.arr(i)
                    Let This.arr(i) = temp
                    Let swapped = True
                End If
            Next i
        Loop While swapped
    Else
        Do
            Let swapped = False
            For i = 1 To This.Length - 1
                If This.arr(i) > This.arr(i + 1) Then
                    Let temp = This.arr(i + 1)
                    Let This.arr(i + 1) = This.arr(i)
                    Let This.arr(i) = temp
                    Let swapped = True
                End If
            Next i
        Loop While swapped
    End If
End Sub

Public Function SortAndGetPermutation(Optional ByVal isDescending As Boolean) As List ' Bubble sort, ascending
With List.Emptie(This.Length)
    Dim i As Long
    For i = 1 To This.Length
        Let .ItemNoCheck(i) = i
    Next i
    
    Dim swapped As Boolean
    Dim temp As Variant
    
    If isDescending Then
        Do
            Let swapped = False
            For i = 1 To This.Length - 1
                If This.arr(i) < This.arr(i + 1) Then
                    Let temp = This.arr(i + 1)
                    Let This.arr(i + 1) = This.arr(i)
                    Let This.arr(i) = temp
                    'Call SwapForSort(i, i + 1)
                    Call .SwapForSort(i, i + 1)
                    Let swapped = True
                End If
            Next i
        Loop While swapped
    Else
        Do
            Let swapped = False
            For i = 1 To This.Length - 1
                If This.arr(i) > This.arr(i + 1) Then
                    Let temp = This.arr(i + 1)
                    Let This.arr(i + 1) = This.arr(i)
                    Let This.arr(i) = temp
                    'Call SwapForSort(i, i + 1)
                    Call .SwapForSort(i, i + 1)
                    Let swapped = True
                End If
            Next i
        Loop While swapped
    End If
    
    Set SortAndGetPermutation = .Self
End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Operations that Alter the List Length
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
    Call Resize(0)
End Sub

Public Sub InsertRange(ByVal lst As List, Optional ByVal index As Long = -1)
    If index = -1 Then
        Let index = This.Length ' Default value
    ElseIf index Then ' index <> 0. We include index = 0
        Call IndexValidate(index)
    End If
    
    Dim lstLength As Long
    Let lstLength = lst.Length
    
    Call Resize(This.Length + lstLength)
    
    If lstLength > 0 Then
        Call Shift(lstLength, index + 1)
    End If
    
    Dim i As Long
    For i = 1 To lstLength
        If IsObject(lst(i)) Then
            Set This.arr(index + i) = lst(i)
        Else
            Let This.arr(index + i) = lst(i)
        End If
    Next i
End Sub

Public Sub AddRange(ByVal lst As List)
    Call InsertRange(lst)
End Sub

Public Sub Insert(ByVal what As Variant, Optional ByVal index As Long = -1)
    If index = -1 Then
        Let index = This.Length ' Default value
    ElseIf index Then ' index <> 0. We include index = 0
        Call IndexValidate(index)
    End If
    
    Call Resize(This.Length + 1)
    Call Shift(1, index + 1)
    
    If IsObject(what) Then
        Set This.arr(index + 1) = what
    Else
        Let This.arr(index + 1) = what
    End If
End Sub

Public Sub Prepend(ByVal what As Variant)
    Call Insert(what, 0)
End Sub

Public Sub Append(ByVal what As Variant)
    Call Insert(what, This.Length)
End Sub

Public Sub Add(ByVal what As Variant)
    Call Insert(what, This.Length)
End Sub

Public Sub RemoveRange(ByVal index As Long, ByVal lengthToRemove As Long)
    Call IndexValidate(index)
    
    If index + lengthToRemove <= This.Length Then
        Call Shift(-lengthToRemove, index + lengthToRemove)
        Call Resize(This.Length - lengthToRemove)
    Else ' index + Length > This.Length
        Call Resize(index - 1)
    End If
End Sub

Public Sub RemoveAt(ByVal index As Long)
    Call IndexValidate(index)
    
    If index <> This.Length Then
        Call Shift(-1, index + 1)
        Call Resize(This.Length - 1)
    Else ' index = This.Length
        Call Resize(index - 1)
    End If
End Sub

Public Function Pop(Optional ByVal index As Long) As Variant
    If index Then ' index <> 0
        Call IndexValidate(index)
    Else ' index = 0
        Let index = This.Length ' Default value
    End If
    
    If IsObject(This.arr(index)) Then
        Set Pop = This.arr(index)
    Else
        Let Pop = This.arr(index)
    End If
    
    Call RemoveAt(index)
End Function

Public Function Remove(ByVal what As Variant) As Boolean
    Dim foundIndex As Long
    Let foundIndex = IndexOf(what)
    
    If foundIndex Then ' foundIndex > 0
        Call RemoveAt(foundIndex)
        Let Remove = True
    End If
End Function

Public Sub Map(ByVal indexList As List)
    Dim indexListLength As Long
    Let indexListLength = indexList.Length
    
    Dim newArray() As Variant
    ReDim newArray(1 To indexListLength) As Variant
    
    Dim i As Long
    For i = 1 To indexListLength
        If IsObject(This.arr(indexList(i))) Then
            Set newArray(i) = This.arr(indexList(i))
        Else
            Let newArray(i) = This.arr(indexList(i))
        End If
    Next i
    
    Let This.arr = newArray
    Let This.Capacity = indexListLength
    Let This.Length = indexListLength
End Sub

Public Sub Keep(ByVal indexList As List)
    Dim boolList As List
    Set boolList = List.Full(This.Length, False)
    
    Dim i As Long
    For i = 1 To indexList.Length
        Let boolList(indexList(i)) = True
    Next i
    
    Call Map(GetIndexListFromBoolList(boolList))
End Sub

Public Function GetUniqueIndex() As List
    Dim thisLength As Long
    Let thisLength = This.Length
    
    Dim boolList As List
    Set boolList = List.Full(thisLength, False)
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To thisLength - 1
        If boolList(i) Then
            For j = i + 1 To thisLength
                If boolList(j) Then
                    If IsObject(This.arr(i)) And IsObject(This.arr(j)) Then
                        If This.arr(i) Is This.arr(j) Then Let boolList(j) = False
                    ElseIf (Not IsObject(This.arr(i))) And (Not IsObject(This.arr(j))) Then
                        If This.arr(i) = This.arr(j) Then Let boolList(j) = False
                    End If
                End If
            Next j
        End If
    Next i
    
    Set GetUniqueIndex = GetIndexListFromBoolList(boolList)
End Function

Public Sub Unique()
    Call Keep(GetUniqueIndex())
End Sub

Public Sub KeepNot(ByVal indexList As List)
    Dim boolList As List
    Set boolList = List.Full(This.Length, True)
    
    Dim i As Long
    For i = 1 To indexList.Length
        Let boolList(indexList(i)) = False
    Next i
    
    Call Map(GetIndexListFromBoolList(boolList))
End Sub

Public Sub RemoveEach(ByVal what As Variant)
    Call KeepNot(IndexOfEach(what))
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Information Between Lists
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsSubsetOf(ByVal lst As List) As Boolean
With lst
    Let IsSubsetOf = True
    
    Dim i As Long
    For i = 1 To This.Length
        If Not .Contains(This.arr(i)) Then
            Let IsSubsetOf = False
            Exit For
        End If
    Next i
End With
End Function

Public Function IsSupersetOf(ByVal lst As List) As Boolean
    Let IsSupersetOf = lst.IsSubsetOf(Me)
End Function

Public Function IsEqualTo(ByVal lst As List) As Boolean
    Let IsEqualTo = IsSubsetOf(lst) And IsSupersetOf(lst)
End Function

Public Function IsProperSubsetOf(ByVal lst As List) As Boolean
    Let IsProperSubsetOf = IsSubsetOf(lst) And Not IsEqualTo(lst)
End Function

Public Function IsProperSupersetOf(ByVal lst As List) As Boolean
    Let IsProperSupersetOf = IsSupersetOf(lst) And Not IsEqualTo(lst)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IsSubset(ByVal listL As List, ByVal listR As List) As Boolean
    Let IsSubset = listL.IsSubsetOf(listR)
End Function

Public Function IsSuperset(ByVal listL As List, ByVal listR As List) As Boolean
    Let IsSuperset = IsSubset(listR, listL)
End Function

Public Function IsEqual(ByVal listL As List, ByVal listR As List) As Boolean
    Let IsEqual = IsSubset(listL, listR) And IsSuperset(listL, listR)
End Function

Public Function IsProperSubset(ByVal listL As List, ByVal listR As List) As Boolean
    Let IsProperSubset = IsSubset(listL, listR) And Not IsEqual(listL, listR)
End Function

Public Function IsProperSuperset(ByVal listL As List, ByVal listR As List) As Boolean
    Let IsProperSuperset = IsSuperset(listL, listR) And Not IsEqual(listL, listR)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Operations among Lists
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub IntersectWith(ByVal lst As List)
    Dim boolList As List
    Set boolList = List.Full(This.Length, False)
    
    Dim i As Long
    For i = 1 To This.Length
        If lst.Contains(This.arr(i)) Then
            Let boolList(i) = True
        End If
    Next i
    
    Call Keep(GetIndexListFromBoolList(boolList))
End Sub

Public Sub DifferenceWith(ByVal lst As List)
    Dim boolList As List
    Set boolList = List.Full(This.Length, True)
    
    Dim i As Long
    For i = 1 To This.Length
        If lst.Contains(This.arr(i)) Then
            Let boolList(i) = False
        End If
    Next i
    
    Call Keep(GetIndexListFromBoolList(boolList))
End Sub

Public Sub UnionWith(ByVal lst As List)
    Call InsertRange(List.Difference(lst, Me))
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constructors among Lists
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Intersect(ByVal listL As List, ByVal listR As List) As List
With listL.Clone
    Call .IntersectWith(listR)
    Set Intersect = .Self
End With
End Function

Public Function Difference(ByVal listL As List, ByVal listR As List) As List
With listL.Clone
    Call .DifferenceWith(listR)
    Set Difference = .Self
End With
End Function

Public Function Union(ByVal listL As List, ByVal listR As List) As List
With listL.Clone
    Call .UnionWith(listR)
    Set Union = .Self
End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Concatenate(ByVal listL As List, ByVal listR As List) As List
With listL.Clone
    Call .InsertRange(listR)
    Set Concatenate = .Self
End With
End Function

Public Function FromEnumerable(ByRef obj As Variant) As List
    On Error GoTo NotEnumerable
    
    With New List
        Dim element As Variant
        For Each element In obj
            Call .Add(element)
        Next element
        
        Set FromEnumerable = .Self
    End With
    
    GoTo Enumerable
NotEnumerable:
    Call Err.Raise(1, , "Input not enumerable")
Enumerable:
End Function

Public Function MapTo(ByVal indexList As List) As List
With List.Emptie(indexList.Length)
    Dim i As Long
    For i = 1 To indexList.Length
        Let .ItemNoCheck(i) = This.arr(indexList(i))
    Next i
    
    Set MapTo = .Self
End With
End Function

Public Function FromCollection(ByRef coll As Collection) As List
With List.Emptie(coll.Count)
    Dim i As Long
    Dim element As Variant
    For Each element In coll
        Let i = i + 1
        Let .ItemNoCheck(i) = element
    Next element
    
    Set FromCollection = .Self
End With
End Function

Public Function FromArray(ByVal arr As Variant) As List
    If ArrayDimension(arr) <> 1 Then
        Call Err.Raise(1, , "Required array of dimension 1")
    End If
    
    With List.Emptie(ArrayLen(arr))
        Dim i As Long
        Dim element As Variant
        For Each element In arr
            Let i = i + 1
            Let .ItemNoCheck(i) = element
        Next element
        
        Set FromArray = .Self
    End With
End Function

Public Function FromArguments(ParamArray argArray() As Variant) As List
    Set FromArguments = List.FromArray(argArray)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As List
With List.Emptie(This.Length)
    Dim i As Long
    For i = 1 To This.Length
        Let .ItemNoCheck(i) = This.arr(i)
    Next i
    
    Set Clone = .Self
End With
End Function

Public Function GetRange(ByVal indexStart As Long, ByVal indexLength As Long) As List
    If indexStart Then ' indexStart <> 0
        Call IndexValidate(indexStart)
    Else ' indexStart = 0
        Let indexStart = 1 ' Default
    End If
    
    Dim indexEnd As Long
    Let indexEnd = indexStart + indexLength - 1
        
    If indexEnd > This.Length Then
        Let indexEnd = This.Length
    End If
    
    With List.Emptie(indexEnd - indexStart + 1) ' Let newLength = indexEnd - indexStart + 1
        Dim i As Long
        For i = 1 To .Length
            Let .ItemNoCheck(i) = This.arr(i + indexStart - 1)
        Next i
    
        Set GetRange = .Self
    End With
End Function

'Public Function CopyTo() As List
'
'End Function

Public Function ToArray(Optional ByVal base As Long = 1) As Variant()
    Dim offset As Long
    Let offset = 1 - base
    
    Dim thisLength As Long
    Let thisLength = This.Length
    
    Dim outputArray() As Variant
    
    If thisLength Then ' thisLength > 0 <=> thisLength <> 0
        ReDim outputArray(base To thisLength - offset) As Variant
    Else
        Let outputArray = Array()
    End If
    
    Dim i As Long
    For i = 1 To thisLength
        If IsObject(This.arr(i)) Then
            Set outputArray(i - offset) = This.arr(i)
        Else
            Let outputArray(i - offset) = This.arr(i)
        End If
    Next i

    Let ToArray = outputArray
End Function

Public Function ToCollection() As Collection
    Set ToCollection = New Collection
    
    With ToCollection
        Dim i As Long
        For i = 1 To This.Length
            Call .Add(This.arr(i))
        Next i
    End With
End Function

Public Sub ToImmediate(Optional ByVal head As Long = 50, Optional ByVal tail As Long = 50)
    Dim sep0 As String * 2: Let sep0 = "| "
    Dim index As String * 6
    Dim sep1 As String * 3: Let sep1 = " | "
    Dim itemType As String * 12
    Dim sep4 As String * 2: Let sep4 = ": "
    Dim itemValue As String * 20
    Dim sep5 As String * 2: Let sep5 = " |"
    
    Dim totalLength As Long
    Let totalLength = Len(sep0) + Len(index) _
                    + Len(sep1) + Len(itemType) _
                    + Len(sep4) + Len(itemValue) _
                    + Len(sep5)
    
    Dim topAndButtom As String
    Let topAndButtom = String(totalLength, "=")
    
    Dim middle As String
    Let middle = String(totalLength, "-")
    
    Debug.Print "## List Printing Start ##"
    Debug.Print topAndButtom
    
    RSet index = "Idx"
    RSet itemType = "Item Type"
    RSet itemValue = "Item Value"
    
    Debug.Print sep0 & index _
          & sep1 & itemType _
          & sep4 & itemValue _
          & sep5
    
    Debug.Print middle
    
    Dim thisLength As Long
    Let thisLength = This.Length
    
    If thisLength = 0 Then
        RSet index = "N/A"
        RSet itemType = vbNullString
        RSet itemValue = vbNullString
    
        Debug.Print sep0 & index _
              & sep1 & itemType _
              & sep4 & itemValue _
              & sep5
    End If
    
    Dim i As Long
    
    For i = 1 To IIf(thisLength > head, head, thisLength)
        RSet index = i
        RSet itemType = TypeName(This.arr(i))
        
        If IsObject(This.arr(i)) Then
            RSet itemValue = "ObjectType"
        Else
            RSet itemValue = This.arr(i)
        End If
        
        Debug.Print sep0 & index _
                  & sep1 & itemType _
                  & sep4 & itemValue _
                  & sep5
    Next i
    
    If thisLength > head + tail Then
        RSet index = ":"
        RSet itemType = vbNullString
        RSet itemValue = vbNullString
    
        Debug.Print sep0 & index _
                  & sep1 & itemType _
                  & sep4 & itemValue _
                  & sep5
    End If
    
    For i = IIf(thisLength > head + tail, thisLength - tail + 1, head + 1) To thisLength
        RSet index = i
        RSet itemType = TypeName(This.arr(i))
        
        If IsObject(This.arr(i)) Then
            RSet itemValue = "ObjectType"
        Else
            RSet itemValue = This.arr(i)
        End If
        
        Debug.Print sep0 & index _
                  & sep1 & itemType _
                  & sep4 & itemValue _
                  & sep5
    Next i
    
    Debug.Print topAndButtom
    Debug.Print "## List Printing End ##"
End Sub

Public Sub ToImmediateHead(Optional ByVal head As Long = 50)
    Call ToImmediate(head, 0)
End Sub

Public Sub ToImmediateTail(Optional ByVal tail As Long = 50)
    Call ToImmediate(0, tail)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Array Helper Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ArrayDimension(ByRef arr As Variant) As Long
    ' Uninitialized arrays and arrays with no element both return 0.
    If Not IsArray(arr) Then
        Call Err.Raise(1, , "Input not an array.")
    End If
    
    ' Uninitialized arrays and erased arrays cause error
    On Error GoTo Done
    
    Dim tmp As Long
    Dim i As Long
    
    Do While True
        Let i = i + 1
        Let tmp = UBound(arr, i)
    Loop
    
Done:
    Let ArrayDimension = i - 1
End Function

Public Function ArrayLen(ByRef arr As Variant, _
                         Optional ByVal dimension As Long = 1) As Long
    If Not dimension > 0 Then
        Call Err.Raise(1, , "Dimension should be positive.")
    End If
    
    Dim arrayDim As Long
    Let arrayDim = ArrayDimension(arr)
    
    If arrayDim Then ' arrayDim > 0 <=> arrayDim <> 0
        Let ArrayLen = UBound(arr, dimension) - LBound(arr, dimension) + 1
    'ElseIf arrayDim = 0
    End If
End Function

Public Function IsArrayEmpty(ByRef arr As Variant) As Boolean
    If ArrayLen(arr) = 0 Then
        Let IsArrayEmpty = True
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Math Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Floor(ByVal what As Double) As Long
    Let Floor = Int(what)
End Function

Public Function Ceil(ByVal what As Double) As Long
    Dim temp As Long
    Let temp = Round(what)
    
    If temp >= what Then
        Ceil = temp
    Else
        Ceil = temp + 1
    End If
End Function

'Public Function Round(ByVal what As Double) As Long
'
'End Function



