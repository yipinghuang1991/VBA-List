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
' Class_Initialize
' Self
' Capacity
' Length -> Count
' TrimToSize
' Resize
' IsIndexValid -> IndexValidate -> Item (Default Member) -> LetOrSet
'
' ## Generators ##
'
' EmptyList ---> EmptyLike
'            |-> Full -------> Repeat
'            |             |-> FullLike
'            |             |-> Ones ---> OnesLike
'            |             |-> Zeros --> ZerosLike
'            |             |-> Trues --> TruesLike
'            |             --> Falses -> FalsesLike
'            |-> Arange
'            --> LinSpace ---> LogSpace
'                          |-> GeomSpace
' ## Information ##
'
' (IndexValidate) ---> LastIndexOf
'                  |-> IndexOf ---------> Contains
'                                     |-> NumberOf
'                                     --> IndexOfEach <-(Add)
' (IsNumberType) ----> IsNumberList
' (IsIntegerType) ---> IsIntegerList ---> IsBitList
'                         |v
' (IsIndexValid) ----> IsIndexList -----> IndexListValidate -----> GetBoolListFromIndexList <-(Falses)
' (IsBoolList) ------> IsBoolIndexList -> BoolIndexListValidate -> GetIndexListFromBoolList <-(Add)
'
' ## Operations that Preserves the List Length ##
'
' Clear
' AsType
' SetRange
' Reverse
' Invert ----------> BoolInvert
' (IsBitList) -----> BitInvert
'
' Wipe -------  |-> Move
'            |--|
' Copy -------  |-> Shift <-(IndexValidate)
'  |
'  |--------------> Permute -> Swap -> Sort <-(Arange)
'                      ^
' (IndexListValidate) -|
'
' ## Operations that Alters the List Length ##
'
' (IndexListValidate) ---> Map
'                      |-> MapTo
'
'                      |-> InsertRange ---> AddRange
' (IndexValidate) ---  |                |-> Insert -----> Prepend
'                   |--|                              |-> Append -> Add
' (Shift) -----------  |
'                      |-> RemoveRange ---> RemoveAt ---> Pop <-(IndexValidate)
'                                                     |-> Remove <-(IndexOf)
'
' (Trues) ------------------------------|
'                                       V
'                      ------------> GetUniqueIndex --|
'                      |                              V
' (GetIndexListFromBoolList) --  |-> Keep(ILV) -+-> Unique
'                             |--|              |v
' (GetBoolListFromIndexList) --  |--------------+-> KeepNot -> RemoveEach <-(IndexOfEach)
'                                                      ^
' (Invert) --------------------------------------------|
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
' (Contains) ------------------  |-> IntersectWith --> Intersect <-------
'                             |  |                                      |
' (GetIndexListFromBoolList) -|--|                        |--------- (Clone) -------|
'                             |  |                        V                         V
' (Keep) ----------------------  |-> DifferenceWith -> Difference -> UnionWith -> Union
'                                                                       ^
' (InsertRange) --------------------------------------------------------|
'
' ## Constructors ##
'
' (Clone) -------|
' (InsertRange) ---> Concatenate
'
' FromEnumerable
' FromCollection
'
' (ArrayDimension) -> FromArray -> FromArguments
'
' ## Output ##
'
' Clone
' GetRange
' CopyTo
' ToArray
' ToCollection
' ToImmediate ---> ToImmediateHead
'              |-> ToImmediateTail
'
' ## Array Helper Functions ##
'
' ArrayDimension -> ArrayLen -> IsArrayEmpty
'
' ## Type Helper Functions ##
'
' IsNumberType
' IsIntegerType
'
' ## Math Functions ##
'
' Floor
' Ceil
' Round
' Max
' Min
' Sum
'
' ## Vector Functions ##
'
' VecAbs
' VecPow
' VecNorm
' VecSum
' VecProd
' ScalarProd
' VecDot

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Storage
    theArray() As Variant
    Length As Long
    Capacity As Long
End Type

Private This As Storage

Private Sub Class_Initialize()
With This
    Let .Capacity = 1
    ReDim .theArray(1 To .Capacity)
End With
End Sub

Public Property Get Self() As List
    Set Self = Me
End Property

Public Property Get Capacity() As Long
    Let Capacity = This.Capacity
End Property

Public Property Get Length() As Long
    Let Length = This.Length
End Property

Public Property Get Count() As Long
    Let Count = Me.Length
End Property

Public Sub TrimToSize()
    Dim newArray() As Variant
    Let newArray = Me.ToArray()
    
    Let This.theArray = newArray
    Let This.Capacity = This.Length
End Sub

Public Sub Resize(ByVal newLength As Long)
With This
    If newLength > .Length Then
        If newLength > .Capacity Then
            Do
                Let .Capacity = 2 * .Capacity
            Loop Until newLength <= .Capacity
            
            ReDim Preserve .theArray(1 To .Capacity)
        End If
    Else ' newLength <= .Length, equality is redundant
        Dim i As Long
        For i = newLength + 1 To .Length
            Let .theArray(i) = Empty
        Next i
    End If
    
    Let .Length = newLength
End With
End Sub

Public Function IsIndexValid(ByVal index As Long) As Boolean ' Default to False
    If (1 <= index) And (index <= This.Length) Then
        Let IsIndexValid = True
    End If
End Function

Public Sub IndexValidate(ByVal index As Long)
    If Not IsIndexValid(index) Then
        Call Err.Raise(1, , "Index out of range.")
    End If
End Sub

'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
With This
    Call IndexValidate(index)
    
    If IsObject(.theArray(index)) Then
        Set Item = .theArray(index)
    Else
        Let Item = .theArray(index)
    End If
End With
End Property

Public Property Let Item(ByVal index As Long, ByVal what As Variant)
    Call IndexValidate(index)
    Let This.theArray(index) = what
End Property

Public Property Set Item(ByVal index As Long, ByVal what As Variant)
    Call IndexValidate(index)
    Set This.theArray(index) = what
End Property

Public Sub LetOrSet(ByVal index As Long, ByVal what As Variant)
    If IsObject(what) Then
        Set Me(index) = what
    Else
        Let Me(index) = what
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generators
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EmptyList(ByVal inputLength As Long) As List
With New List
    Call .Resize(inputLength)
    Set EmptyList = .Self
End With
End Function

Public Function EmptyLike(ByVal lst As List) As List
    Set EmptyLike = Me.EmptyList(lst.Length)
End Function

Public Function Full(ByVal inputLength As Long, ByVal what As Variant) As List
With Me.EmptyList(inputLength)
    Dim i As Long
    For i = 1 To inputLength
        Call .LetOrSet(i, what)
    Next i
    
    Set Full = .Self
End With
End Function

Public Function Repeat(ByVal what As Variant, ByVal times As Long) As List
    Set Repeat = Me.Full(times, what)
End Function

Public Function FullLike(ByVal lst As List, ByVal what As Variant) As List
    Set FullLike = Me.Full(lst.Length, what)
End Function

Public Function Ones(ByVal inputLength As Long) As List
    Set Ones = Me.Full(inputLength, 1)
End Function

Public Function OnesLike(ByVal lst As List) As List
    Set OnesLike = Me.Ones(lst.Length)
End Function

Public Function Zeros(ByVal inputLength As Long) As List
    Set Zeros = Me.Full(inputLength, 0)
End Function

Public Function ZerosLike(ByVal lst As List) As List
    Set ZerosLike = Me.Zeros(lst.Length)
End Function

Public Function Trues(ByVal inputLength As Long) As List
    Set Trues = Me.Full(inputLength, True)
End Function

Public Function TruesLike(ByVal lst As List) As List
    Set TruesLike = Me.Trues(lst.Length)
End Function

Public Function Falses(ByVal inputLength As Long) As List
    Set Falses = Me.Full(inputLength, False)
End Function

Public Function FalsesLike(ByVal lst As List) As List
    Set FalsesLike = Me.Falses(lst.Length)
End Function

Public Function Arange(ByVal start As Variant, ByVal ends As Variant, _
                       Optional ByVal step As Variant = 1) As List
    Dim extraLength As Long
    Let extraLength = Int((ends - start) / step)
    
    If extraLength < 0 Then Let extraLength = 0
    
    With Me.EmptyList(1 + extraLength)
        Let .Self(1) = start
        
        Dim i As Long
        For i = 2 To .Length
            Let .Self(i) = .Self(i - 1) + step
        Next i
        
        Set Arange = .Self
    End With
End Function

Public Function LinSpace(ByVal start As Variant, ByVal ends As Variant, _
                         Optional ByVal num As Long = 50) As List
With Me.EmptyList(num + 1)
    Dim width As Double
    Let width = (ends - start) / num
    
    Let .Self(1) = start
    
    Dim i As Long
    For i = 2 To num
        Let .Self(i) = .Self(1) + (i - 1) * width ' .Self(i - 1) + width
    Next i
    
    Let .Self(num + 1) = ends
    
    Set LinSpace = .Self
End With
End Function

Public Function LogSpace(ByVal start As Variant, ByVal ends As Variant, _
                         Optional ByVal num As Long = 50, _
                         Optional ByVal base As Double) As List
    If base = 0 Then Let base = Math.Exp(1)
    
    With Me.LinSpace(start, ends, num)
        Dim i As Long
        For i = 1 To .Length
            Let .Self(i) = base ^ (.Self(i))
        Next i
        
        Set LogSpace = .Self
    End With
End Function

Public Function GeomSpace(ByVal start As Variant, ByVal ends As Variant, _
                          Optional ByVal num As Long = 50, _
                          Optional ByVal base As Double) As List
    If base = 0 Then Let base = Math.Exp(1)
    
    With Me.LinSpace(Log(start) / Log(base), Log(ends) / Log(base), num)
        Let .Self(1) = start
        
        Dim i As Long
        For i = 2 To .Length - 1
            Let .Self(i) = base ^ (.Self(i))
        Next i
        
        Let .Self(.Length) = ends
        
        Set GeomSpace = .Self
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Information
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastIndexOf(ByVal what As Variant, _
                            Optional ByVal indexStart As Long, _
                            Optional ByVal indexLength As Long = -1) As Long ' Default to 0
    If indexStart Then ' indexStart <> 0
        Call IndexValidate(indexStart)
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
            If IsObject(Me(i)) Then
                If Me(i) Is what Then
                    Let LastIndexOf = i
                    Exit For
                End If
            End If
        Next i
    Else
        For i = indexStart To indexEnd Step -1
            If Not IsObject(Me(i)) Then
                If Me(i) = what Then
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
        Call IndexValidate(indexStart)
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
            If IsObject(Me(i)) Then
                If Me(i) Is what Then
                    Let IndexOf = i
                    Exit For
                End If
            End If
        Next i
    Else
        For i = indexStart To indexEnd
            If Not IsObject(Me(i)) Then
                If Me(i) = what Then
                    Let IndexOf = i
                    Exit For
                End If
            End If
        Next i
    End If
End Function

Public Function Contains(ByVal what As Variant) As Boolean
    Let Contains = Me.IndexOf(what) ' <> 0
End Function

Public Function NumberOf(ByVal what As Variant) As Long ' Default to 0
    Dim thisLength As Long
    Let thisLength = This.Length
    
    Dim i As Long ' Default to 0
    
    Do
        Let i = Me.IndexOf(what, i + 1)
        If i Then ' i > 0, i.e., i <> 0
            Let NumberOf = NumberOf + 1
        End If
    Loop While i And (i < thisLength)
End Function

Public Function IndexOfEach(ByVal what As Variant) As List
    Dim thisLength As Long
    Let thisLength = This.Length
    
    With New List
        Dim i As Long ' Default to 0
        Do
            Let i = Me.IndexOf(what, i + 1)
            If i Then ' i > 0, i.e., i <> 0
                Call .Add(i)
            End If
        Loop While i And (i < thisLength)
        
        Set IndexOfEach = .Self
    End With
End Function

Public Function IsNumberList(ByVal numberList As List) As Boolean
With numberList
    Let IsNumberList = True
    
    Dim i As Long
    For i = 1 To .Length
        If Not List.IsNumberType(.Self(i)) Then
            Let IsNumberList = False
            Exit For
        End If
    Next i
End With
End Function

Public Function IsIntegerList(ByVal integerList As List) As Boolean
With integerList
    Let IsIntegerList = True
    
    Dim i As Long
    For i = 1 To .Length
        If Not List.IsIntegerType(.Self(i)) Then
            Let IsIntegerList = False
            Exit For
        End If
    Next i
End With
End Function

Public Function IsBitList(ByVal bitList As List) As Boolean
With bitList
    If List.IsIntegerList(.Self) Then
        Let IsBitList = True
        
        Dim i As Long
        For i = 1 To .Length
            If (.Self(i) <> 0) And (.Self(i) <> 1) Then
                Let IsBitList = False
                Exit For
            End If
        Next i
    End If
End With
End Function

Public Function IsIndexList(ByVal indexList As List) As Boolean
With indexList
    If List.IsIntegerList(.Self) Then
        Let IsIndexList = True
        
        Dim i As Long
        For i = 1 To .Length
            If Not Me.IsIndexValid(.Self(i)) Then
                Let IsIndexList = False
                Exit For
            End If
        Next i
    End If
End With
End Function

Public Sub IndexListValidate(ByVal indexList As List)
    If Not Me.IsIndexList(indexList) Then
        Call Err.Raise(1, , "Invalid index list.")
    End If
End Sub

Public Function GetBoolListFromIndexList(ByVal indexList As List) As List
    Call Me.IndexListValidate(indexList)
    
    With List.Falses(This.Length)
        Dim i As Long
        For i = 1 To indexList.Length
            Let .Self(indexList(i)) = True
        Next i
        Set GetBoolListFromIndexList = .Self
    End With
End Function

Public Function IsBoolList(ByVal boolList As List) As Boolean
With boolList
    Let IsBoolList = True
    
    Dim i As Long
    For i = 1 To .Length
        If VarType(.Self(i)) <> vbBoolean Then
            Let IsBoolList = False
            Exit For
        End If
    Next i
End With
End Function

Public Function IsBoolIndexList(ByVal boolIndexList As List) As Boolean
With boolIndexList
    If .Length = This.Length Then
        Let IsBoolIndexList = IsBoolList(boolIndexList)
    End If
End With
End Function

Public Sub BoolIndexListValidate(ByVal boolIndexList As List)
    If Not Me.IsBoolIndexList(boolIndexList) Then
        Call Err.Raise(1, , "Invalid index list.")
    End If
End Sub

Public Function GetIndexListFromBoolList(ByVal boolList As List) As List
    Call Me.BoolIndexListValidate(boolList)
    
    With New List
        Dim i As Long
        For i = 1 To boolList.Length
            If boolList(i) Then
                Call .Add(i)
            End If
        Next i
        Set GetIndexListFromBoolList = .Self
    End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Operations that Preserve the List Length
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
    Call Me.Resize(0)
End Sub

Public Sub AsType(ByVal whatType As Variant)
    Dim i As Long

    Select Case whatType
        Case VbVarType.vbBoolean
            For i = 1 To This.Length
                Let Me(i) = CBool(Me(i))
            Next i
'        Case VbVarType.vbByte
'            For i = 1 To This.Length
'                Let Me(i) = CByte(Me(i))
'            Next i
'        Case VbVarType.vbCurrency
'            For i = 1 To This.Length
'                Let Me(i) = CCur(Me(i))
'            Next i
        Case VbVarType.vbDate
            For i = 1 To This.Length
                Let Me(i) = CDate(Me(i))
            Next i
        Case VbVarType.vbDouble
            For i = 1 To This.Length
                Let Me(i) = CDbl(Me(i))
            Next i
'        Case VbVarType.vbDecimal
'            For i = 1 To This.Length
'                Let Me(i) = CDec(Me(i))
'            Next i
'        Case VbVarType.vbInteger
'            For i = 1 To This.Length
'                Let Me(i) = CInt(Me(i))
'            Next i
        Case VbVarType.vbLong
            For i = 1 To This.Length
                Let Me(i) = CLng(Me(i))
            Next i
        Case VbVarType.vbLongLong
            For i = 1 To This.Length
                Let Me(i) = CLngLng(Me(i))
            Next i
        Case VbVarType.vbSingle
            For i = 1 To This.Length
                Let Me(i) = CSng(Me(i))
            Next i
        Case VbVarType.vbString
            For i = 1 To This.Length
                Let Me(i) = CStr(Me(i))
            Next i
        Case VbVarType.vbVariant
            For i = 1 To This.Length
                Let Me(i) = CVar(Me(i))
            Next i
        Case "Real", "Numeric", "Numerical", "Number"
            For i = 1 To This.Length
                Let Me(i) = Val(Me(i))
            Next i
        Case Else
            Call Err.Raise(1, , "Conversion not supported.")
    End Select
End Sub

Public Sub SetRange(ByVal index As Long, ByVal lst As List)
    If index + lst.Length - 1 > This.Length Then
        Call Err.Raise(1, , "Error")
    End If
    
    Dim i As Long
    For i = index To index + lst.Length - 1
        Call Me.LetOrSet(i, lst(i - index + 1))
    Next i
End Sub

Public Sub Reverse()
    Dim thisLength As Long
    Let thisLength = This.Length
    
    If thisLength Then ' thisLength <> 0
        Dim newArray() As Variant
        ReDim newArray(1 To This.Capacity) As Variant
    
        Dim i As Long
        For i = 1 To thisLength
            If IsObject(Me(i)) Then
                Set newArray(i) = Me(thisLength - i + 1)
            Else
                Let newArray(i) = Me(thisLength - i + 1)
            End If
        Next i
        
        Let This.theArray = newArray
    End If
End Sub

Public Sub Invert()
    Dim i As Long
    For i = 1 To This.Length
        Let Me(i) = Not Me(i)
    Next i
End Sub

Public Sub BoolInvert()
    If List.IsBoolList(Me) Then
        Call Me.Invert
    Else
        Call Err.Raise(1, , "Not a bool list.")
    End If
End Sub

Public Sub BitInvert()
    If List.IsBitList(Me) Then
        Dim i As Long
        For i = 1 To This.Length
            If Me(i) Then ' Me(i) <> 0, i.e., Me(i) = 1
                Let Me(i) = 0
            Else
                Let Me(i) = 1
            End If
        Next i
    Else
        Call Err.Raise(1, , "Not a bit list.")
    End If
End Sub

Public Sub Wipe(Optional ByVal index As Long)
    If index Then ' index <> 0
        Let Me(index) = Empty
    Else
        With Me
            Dim i As Long
            For i = 1 To .Length
                Call .Wipe(i) ' Recursively
            Next i
        End With
    End If
End Sub

Public Sub Copy(ByVal index As Long, ByVal indexFrom As Long)
    If index <> indexFrom Then
        Call LetOrSet(index, Me(indexFrom))
    End If
End Sub

Public Sub Move(ByVal index As Long, ByVal indexFrom As Long)
    If index <> indexFrom Then
        Call Me.Copy(index, indexFrom)
        Call Me.Wipe(indexFrom)
    End If
End Sub

Public Sub Shift(ByVal offset As Long, _
                 Optional ByVal indexStart As Long, _
                 Optional ByVal indexLength As Long = -1)
With Me
    Dim thisLength As Long
    Let thisLength = This.Length
    
    If indexStart Then ' indexStart <> 0
        Call IndexValidate(indexStart)
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
    
    If offset > 0 Then
        For i = IIf(indexEnd + offset <= thisLength, indexEnd + offset, thisLength) To indexStart + offset Step -1
            Call .Copy(i, i - offset)
        Next i
        
        For i = IIf(indexStart + offset - 1 <= thisLength, indexStart + offset - 1, thisLength) To indexStart Step -1
            Call .Wipe(i)
        Next i
    ElseIf offset < 0 Then
        For i = IIf(indexStart + offset >= 1, indexStart + offset, 1) To indexEnd + offset
            Call .Copy(i, i - offset)
        Next i
        
        For i = IIf(indexEnd + offset + 1 > 0, indexEnd + offset + 1, 1) To indexEnd
            Call .Wipe(i)
        Next i
    End If
End With
End Sub

Public Sub Permute(ByVal indexList As List)
With Me
    Call .IndexListValidate(indexList)
    
    Dim listLength As Long
    Let listLength = indexList.Length
    
    If listLength Then ' listLength <> 0
        Dim lastItem As Variant
        
        If IsObject(Me(indexList(listLength))) Then
            Set lastItem = Me(indexList(listLength))
        Else
            Let lastItem = Me(indexList(listLength))
        End If
        
        Dim i As Long
        For i = listLength To 2 Step -1
            Call .Copy(indexList(i), indexList(i - 1))
        Next i
        
        Call .LetOrSet(indexList(1), lastItem)
    End If
End With
End Sub

Public Sub Swap(ByVal index1 As Long, ByVal index2 As Long)
    Call Me.Permute(List.FromArguments(index1, index2))
End Sub

Public Function Sort(Optional ByVal isDescending As Boolean) As List ' Bubble sort, ascending
With Me.Arange(1, This.Length)
    Dim swapped As Boolean
    Dim i As Long
    
    Dim thisLengthLessOne As Long
    Let thisLengthLessOne = This.Length - 1
    
    If isDescending Then
        Do
            Let swapped = False
            For i = 1 To thisLengthLessOne
                If Me(i) < Me(i + 1) Then
                    Call Me.Swap(i, i + 1)
                    Call .Swap(i, i + 1)
                    Let swapped = True
                End If
            Next i
        Loop While swapped
    Else
        Do
            Let swapped = False
            For i = 1 To thisLengthLessOne
                If Me(i) > Me(i + 1) Then
                    Call Me.Swap(i, i + 1)
                    Call .Swap(i, i + 1)
                    Let swapped = True
                End If
            Next i
        Loop While swapped
    End If
    
    Set Sort = .Self
End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Operations that Alter the List Length
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub map(ByVal indexList As List)
    Call Me.IndexListValidate(indexList)
    
    Dim newArray() As Variant
    ReDim newArray(1 To indexList.Length) As Variant
    
    Dim i As Long
    For i = 1 To indexList.Length
        If IsObject(Me(indexList(i))) Then
            Set newArray(i) = Me(indexList(i))
        Else
            Let newArray(i) = Me(indexList(i))
        End If
    Next i
    
    Let This.theArray = newArray
    Let This.Length = indexList.Length
End Sub

Public Function MapTo(ByVal indexList As List) As List
With New List
    Call Me.IndexListValidate(indexList)
    
    Call .Resize(indexList.Length)
    
    Dim i As Long
    For i = 1 To indexList.Length
        If IsObject(Me(indexList(i))) Then
            Set .Self(i) = Me(indexList(i))
        Else
            Let .Self(i) = Me(indexList(i))
        End If
    Next i
    
    Set MapTo = .Self
End With
End Function

Public Sub InsertRange(ByVal lst As List, Optional ByVal index As Long = -1)
With Me
    Dim thisLength As Long
    Let thisLength = This.Length
    
    If index = -1 Then
        Let index = thisLength ' Default value
    ElseIf index Then ' index <> 0. We include index = 0
        Call IndexValidate(index)
    End If
    
    Dim lstLength As Long
    Let lstLength = lst.Length
    
    Call .Resize(thisLength + lstLength)
    
    If lstLength > 0 Then
        Call .Shift(lstLength, index + 1)
    End If
    
    Dim i As Long
    For i = 1 To lstLength
        Call .LetOrSet(index + i, lst(i))
    Next i
End With
End Sub

Public Sub AddRange(ByVal lst As List)
    Call Me.InsertRange(lst)
End Sub

Public Sub Insert(ByVal what As Variant, Optional ByVal index As Long = -1)
    Call Me.InsertRange(List.FromArguments(what), index)
End Sub

Public Sub Prepend(ByVal what As Variant)
    Call Me.Insert(what, 0)
End Sub

Public Sub Append(ByVal what As Variant)
    Call Me.Insert(what, This.Length)
End Sub

Public Sub Add(ByVal what As Variant)
    Call Me.Append(what)
End Sub

Public Function RemoveRange(ByVal index As Long, ByVal Length As Long) As Boolean
    Call IndexValidate(index)
    
    If index + Length > This.Length Then
        Call Me.Resize(index - 1)
    Else ' index + Length <= This.Length
        Call Me.Shift(-Length, index + Length)
        Call Me.Resize(This.Length - Length)
    End If
End Function

Public Sub RemoveAt(ByVal index As Long)
    Call RemoveRange(index, 1)
End Sub

Public Function Pop(Optional ByVal index As Long) As Variant
    If index Then ' index <> 0
        Call IndexValidate(index)
    Else ' index = 0
        Let index = This.Length ' Default value
    End If
    
    If IsObject(Me(index)) Then
        Set Pop = Me(index)
    Else
        Let Pop = Me(index)
    End If
    
    Call RemoveAt(index)
End Function

Public Function Remove(ByVal what As Variant) As Boolean
    Dim foundIndex As Long
    Let foundIndex = IndexOf(what)
    
    If foundIndex Then ' foundIndex > 0
        Call Me.RemoveAt(foundIndex)
        Let Remove = True
    End If
End Function

Public Function GetUniqueIndex() As List
With List.Trues(This.Length)
    Dim i As Long
    Dim j As Long
    
    For i = 1 To This.Length - 1
        If .Self(i) Then
            For j = i + 1 To This.Length
                If .Self(j) Then
                    If IsObject(Me(i)) And IsObject(Me(j)) Then
                        If Me(i) Is Me(j) Then Let .Self(j) = False
                    ElseIf (Not IsObject(Me(i))) And (Not IsObject(Me(j))) Then
                        If Me(i) = Me(j) Then Let .Self(j) = False
                    End If
                End If
            Next j
        End If
    Next i
    
    Set GetUniqueIndex = Me.GetIndexListFromBoolList(.Self)
End With
End Function

Public Sub Keep(ByVal indexList As List)
    Call Me.IndexListValidate(indexList)
    
    Dim boolList As List
    Set boolList = Me.GetBoolListFromIndexList(indexList)
        
    Set indexList = Me.GetIndexListFromBoolList(boolList)

    Dim i As Long
    For i = 1 To indexList.Length
        Call Me.LetOrSet(i, Me(indexList(i)))
    Next i

    Call Me.Resize(indexList.Length)
End Sub

Public Sub Unique()
    Call Me.Keep(Me.GetUniqueIndex())
End Sub

Public Sub KeepNot(ByVal indexList As List)
With Me.GetBoolListFromIndexList(indexList)
    Call .Invert
    Call Me.Keep(Me.GetIndexListFromBoolList(.Self))
End With
End Sub

Public Sub RemoveEach(ByVal what As Variant)
    Call Me.KeepNot(Me.IndexOfEach(what))
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Information Between Lists
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsSubsetOf(ByVal lst As List) As Boolean
With lst
    Let IsSubsetOf = True
    
    Dim i As Long
    For i = 1 To This.Length
        If Not .Contains(Me(i)) Then
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
With List.Falses(This.Length)
    Dim i As Long
    For i = 1 To This.Length
        If lst.Contains(Me(i)) Then
            Let .Self(i) = True
        End If
    Next i
    
    Call Me.Keep(Me.GetIndexListFromBoolList(.Self))
End With
End Sub

Public Sub DifferenceWith(ByVal lst As List)
With List.Trues(This.Length)
    Dim i As Long
    For i = 1 To This.Length
        If lst.Contains(Me(i)) Then
            Let .Self(i) = False
        End If
    Next i
    
    Call Me.Keep(Me.GetIndexListFromBoolList(.Self))
End With
End Sub

Public Sub UnionWith(ByVal lst As List)
    Call Me.InsertRange(List.Difference(lst, Me))
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
    
    If False Then
NotEnumerable:
        Call Err.Raise(1, , "Input not enumerable.")
    End If
End Function

Public Function FromCollection(ByRef coll As Collection) As List
With New List
    Call .Resize(coll.NumberOf)
    
    Dim i As Long
    For i = 1 To .Length
        Call .LetOrSet(i, coll(i))
    Next i
    
    Set FromCollection = .Self
End With
End Function

Public Function FromArray(ByVal sourceArray As Variant) As List
    If ArrayDimension(sourceArray) <> 1 Then
        Call Err.Raise(1, , "Error")
    End If
    
    With New List
        Call .Resize(ArrayLen(sourceArray))
    
        Dim indexOffset As Long
        Let indexOffset = LBound(sourceArray) - 1
    
        Dim i As Long
        For i = 1 To .Length
            Call .LetOrSet(i, sourceArray(i + indexOffset))
        Next i
        
        Set FromArray = .Self
    End With
End Function

Public Function FromArguments(ParamArray argumentsArray() As Variant) As List
    Set FromArguments = List.FromArray(argumentsArray)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Clone() As List
With New List
    Call .Resize(This.Length)
    
    Dim i As Long
    For i = 1 To This.Length
        Call .LetOrSet(i, Me(i))
    Next i
    
    Set Clone = .Self
End With
End Function

Public Function GetRange(ByVal indexStart As Long, ByVal indexLength As Long) As List
With New List
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
    
    Dim newLength As Long
    Let newLength = indexEnd - indexStart + 1
    
    If newLength > 0 Then
        Call .Resize(newLength)
        
        Dim i As Long
        For i = 1 To newLength
            Call .LetOrSet(i, Me(i + indexStart - 1))
        Next i
    End If
    
    Set GetRange = .Self
End With
End Function

Public Function CopyTo() As List
    
End Function

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
        If IsObject(Me(i)) Then
            Set outputArray(i - offset) = Me(i)
        Else
            Let outputArray(i - offset) = Me(i)
        End If
    Next i

    Let ToArray = outputArray
End Function

Public Function ToCollection() As Collection
    Set ToCollection = New Collection
    
    With ToCollection
        Dim i As Long
        For i = 1 To This.Length
            Call .Add(Me(i))
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
        RSet itemType = TypeName(Me(i))
        
        If IsObject(Me(i)) Then
            RSet itemValue = "ObjectType"
        Else
            RSet itemValue = Me(i)
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
        RSet itemType = TypeName(Me(i))
        
        If IsObject(Me(i)) Then
            RSet itemValue = "ObjectType"
        Else
            RSet itemValue = Me(i)
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
    Call Me.ToImmediate(head, 0)
End Sub

Public Sub ToImmediateTail(Optional ByVal tail As Long = 50)
    Call Me.ToImmediate(0, tail)
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
' Type Helper Functions
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

Public Function IsIntegerType(ByRef what As Variant) As Boolean
    Dim typeOfWhat As Long
    Let typeOfWhat = VarType(what)
    
    If typeOfWhat = 3 Then
        Let IsIntegerType = True
    ElseIf typeOfWhat = 20 Then
        Let IsIntegerType = True
    ElseIf typeOfWhat = 2 Then
        Let IsIntegerType = True
    ElseIf typeOfWhat = 17 Then
        Let IsIntegerType = True
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

Public Function Round(ByVal what As Double) As Long

End Function

Public Function Max(ByVal lst As List) As Variant
    Dim temp As Variant
    Let temp = lst(1)
    
    On Error GoTo NotComparable
    
    Dim i As Long
    For i = 2 To lst.Length
        If lst(i) > temp Then
            Let temp = lst(i)
        End If
    Next i
    
    Let Max = temp
    
    If False Then
NotComparable:
        Call Err.Raise(1, , "Not comparable")
    End If
End Function

Public Function Min(ByVal lst As List) As Variant
    Dim temp As Variant
    Let temp = lst(1)
    
    On Error GoTo NotComparable
    
    Dim i As Long
    For i = 2 To lst.Length
        If lst(i) < temp Then
            Let temp = lst(i)
        End If
    Next i
    
    Let Min = temp
    
    If False Then
NotComparable:
        Call Err.Raise(1, , "Not comparable")
    End If
End Function

Public Function Sum(ByVal lst As List) As Variant
    Dim i As Long
    For i = 1 To lst.Length
        Let Sum = Sum + lst(i)
    Next i
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Vector Functions
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

