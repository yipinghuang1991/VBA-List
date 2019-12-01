VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Type")
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Dependency Graphs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ## Implementation ##
'
' Class_Initialize
' Self
' Capacity
' Count
'
' IndexOfValue -> ContainsValue
'
' IndexOfKey ---> ContainsKey = Contains = Exists
'       |                |----> Add -----> Item -> LetOrSet
'       |                        ^
'       |------------------------|
'       |-------> Key
'       |-------> Remove
'                   ^
' RemoveAt ---------|
'
' Clear --------> RemoveAll
'
' GetKey -----|
' GetByIndex ---> Clone
'
' SetByIndex
' TrimToSize
' Reverse
' SortByValue
' SortByKey ----> Sort
'
' ## Constructors ##
'
' FromCollection (Pending)
' FromLists
'
' ## Output ##
'
' CopyTo (Pending)
' Keys
' Values = Items
' GetKeyList
' GetValueList
' ToCollection (Pending)
' ToImmediate ---> ToImmediateHead
'              |-> ToImmediateTail

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type Storage
    keyList As List
    valueList As List
End Type

Private This As Storage

Private Sub Class_Initialize()
    Set This.keyList = New List
    Set This.valueList = New List
End Sub

Public Property Get Self() As Mapping
    Set Self = Me
End Property

Public Property Get Capacity() As Long
    Let Capacity = This.keyList.Capacity
End Property

' Dictionary compatible
Public Property Get Count() As Long
    Let Count = This.keyList.Count
End Property

Public Function IndexOfValue(ByVal value As Variant) As Long
    Let IndexOfValue = This.valueList.IndexOf(value)
End Function

Public Function ContainsValue(ByVal value As Variant) As Boolean
    Let ContainsValue = IndexOfValue(value)
End Function

Public Function IndexOfKey(ByVal Key As Variant) As Long
    Let IndexOfKey = This.keyList.IndexOf(Key)
End Function

Public Function ContainsKey(ByVal Key As Variant) As Boolean
    Let ContainsKey = IndexOfKey(Key)
End Function

Public Function Contains(ByVal Key As Variant) As Boolean
    Let Contains = IndexOfKey(Key)
End Function

' Dictionary compatible
Public Function Exists(ByVal Key As Variant) As Boolean
    Let Exists = IndexOfKey(Key)
End Function

Public Sub Add(ByVal Key As Variant, ByVal value As Variant, Optional ByVal index As Long = -1)
    If ContainsKey(Key) Then
        Call Err.Raise(457)
    Else
        Call This.keyList.Insert(Key, index)
        Call This.valueList.Insert(value, index)
    End If
End Sub

' Dictionary compatible
'@DefaultMember
Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim index As Long
    Let index = IndexOfKey(Key)
    
    If index Then ' index <> 0, i.e., index > 0
        If IsObject(This.valueList(index)) Then
            Set Item = This.valueList(index)
        Else
            Let Item = This.valueList(index)
        End If
    Else
        Call Add(Key, Empty)
    End If
End Property

' Dictionary compatible
Public Property Let Item(ByVal Key As Variant, ByVal what As Variant)
    If IsObject(what) Then
        Call Err.Raise(450)
    End If
    
    Dim index As Long
    Let index = IndexOfKey(Key)
    
    If index Then ' index <> 0, i.e., index > 0
        Let This.valueList(index) = what
    Else
        Call Add(Key, what)
    End If
End Property

' Dictionary compatible
Public Property Set Item(ByVal Key As Variant, ByVal what As Object)
    Dim index As Long
    Let index = IndexOfKey(Key)
    
    If index Then ' index <> 0, i.e., index > 0
        Set This.valueList(index) = what
    Else
        Call Add(Key, what)
    End If
End Property

Public Sub LetOrSet(ByVal Key As Variant, ByVal what As Variant)
    If IsObject(what) Then
        Set Me(Key) = what
    Else
        Let Me(Key) = what
    End If
End Sub

' Dictionary compatible
Public Property Let Key(ByVal oldKey As Variant, ByVal newKey As Variant)
    Dim indexOld As Long
    Let indexOld = IndexOfKey(oldKey)
    
    If indexOld = 0 Then
        Call Err.Raise(32811, , "Method 'Key' of object 'Mapping' failed")
    End If
    
    Dim indexNew As Long
    Let indexNew = IndexOfKey(newKey)
    
    If indexOld <> indexNew Then
        If indexNew Then ' indexNew <> 0
            Call Err.Raise(457)
        Else ' indexNew = 0
            If IsObject(newKey) Then
                Set This.keyList(indexOld) = newKey
            Else
                Let This.keyList(indexOld) = newKey
            End If
        End If
    End If
End Property

' Redundant
Public Property Set Key(ByVal oldKey As Variant, ByVal newKey As Object)
    Dim indexOld As Long
    Let indexOld = IndexOfKey(oldKey)
        
    If indexOld = 0 Then
        Call Err.Raise(32811, , "Method 'Key' of object 'Mapping' failed")
    End If
    
    Dim indexNew As Long
    Let indexNew = IndexOfKey(newKey)
    
    If indexOld <> indexNew Then
        If indexNew Then ' indexNew <> 0
            Call Err.Raise(457)
        Else ' indexNew = 0
            Set This.keyList(indexOld) = newKey
        End If
    End If
End Property

Public Sub RemoveAt(ByVal index As Long)
    Call This.keyList.RemoveAt(index)
    Call This.valueList.RemoveAt(index)
End Sub

' Dictionary compatible
Public Function Remove(ByVal Key As Variant, Optional ByVal errOnAbsent As Boolean = True) As Boolean
    Dim index As Long
    Let index = IndexOfKey(Key)
    
    If index Then ' index <> 0, i.e., index > 0
        Call RemoveAt(index)
        Let Remove = True
    ElseIf errOnAbsent Then
        Call Err.Raise(32811, , "Method 'Remove' of object 'Mapping' failed")
    End If
End Function

Public Sub Clear()
    Call This.keyList.Clear
    Call This.valueList.Clear
End Sub

' Dictionary compatible
Public Sub RemoveAll()
    Call Clear
End Sub

Public Function GetKey(ByVal index As Long) As Variant
    If IsObject(This.keyList(index)) Then
        Set GetKey = This.keyList(index)
    Else
        Let GetKey = This.keyList(index)
    End If
End Function

Public Function GetByIndex(ByVal index As Long) As Variant
    If IsObject(This.valueList(index)) Then
        Set GetByIndex = This.valueList(index)
    Else
        Let GetByIndex = This.valueList(index)
    End If
End Function

Public Function Clone() As Mapping
With New Mapping
    Dim i As Long
    For i = 1 To Count
        Call .Add(GetKey(i), GetByIndex(i))
    Next i
    
    Set Clone = .Self
End With
End Function

Public Sub SetByIndex(ByVal index As Long, ByVal what As Variant)
    If IsObject(what) Then
        Set This.valueList(index) = what
    Else
        Let This.valueList(index) = what
    End If
End Sub

Public Sub TrimToSize()
    Call This.keyList.TrimToSize
    Call This.valueList.TrimToSize
End Sub

Public Sub Reverse()
    Call This.keyList.Reverse
    Call This.valueList.Reverse
End Sub

Public Sub SortByValue(Optional ByVal isDescending As Boolean)
    Call This.keyList.Map(This.valueList.SortAndReturn(isDescending))
End Sub

Public Sub SortByKey(Optional ByVal isDescending As Boolean)
    Call This.valueList.Map(This.keyList.SortAndReturn(isDescending))
End Sub

Public Sub Sort(Optional ByVal isDescending As Boolean)
    Call SortByKey(isDescending)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function FromCollection(ByVal coll As Collection) As Mapping
'
'End Function

Public Function FromLists(ByVal newKeyList As List, ByVal newValueList As List) As Mapping
With New Mapping
    If newKeyList.Count <> newValueList.Count Then
        Call Err.Raise(1, , "Length not match")
    End If
    
    Dim i As Long
    For i = 1 To newKeyList.Count
        Call .Add(newKeyList(i), newValueList(i))
    Next i
    
    Set FromLists = .Self
End With
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function CopyTo() As List
'
'End Function

' Dictionary compatible
Public Function Keys(Optional ByVal base As Long) As Variant()
    Let Keys = This.keyList.ToArray(base)
End Function

Public Function Values(Optional ByVal base As Long) As Variant()
    Let Values = This.valueList.ToArray(base)
End Function

' Dictionary compatible
Public Function Items(Optional ByVal base As Long) As Variant()
    Let Items = This.valueList.ToArray(base)
End Function

Public Function GetKeyList() As List
    Set GetKeyList = This.keyList.Clone
End Function

Public Function GetValueList() As List
    Set GetValueList = This.valueList.Clone
End Function

'Public Function ToCollection() As Collection
'
'End Function

Public Sub ToImmediate(Optional ByVal head As Long = 50, Optional ByVal tail As Long = 50)
    Dim sep0 As String * 2: Let sep0 = "| "
    Dim index As String * 4
    Dim sep1 As String * 3: Let sep1 = " | "
    Dim keyType As String * 8
    Dim sep2 As String * 2: Let sep2 = ": "
    Dim keyValue As String * 12
    Dim sep3 As String * 3: Let sep3 = " | "
    Dim itemType As String * 10
    Dim sep4 As String * 2: Let sep4 = ": "
    Dim itemValue As String * 12
    Dim sep5 As String * 2: Let sep5 = " |"
    
    Dim totalLength As Long
    Let totalLength = Len(sep0) + Len(index) _
                    + Len(sep1) + Len(keyType) _
                    + Len(sep2) + Len(keyValue) _
                    + Len(sep3) + Len(itemType) _
                    + Len(sep4) + Len(itemValue) _
                    + Len(sep5)
    
    Dim topAndButtom As String
    Let topAndButtom = String(totalLength, "=")
    
    Dim middle As String
    Let middle = String(totalLength, "-")
    
    Debug.Print "## Mapping Printing Start ##"
    
    Debug.Print topAndButtom
    
    RSet index = "Idx"
    RSet keyType = "Key Type"
    RSet keyValue = "Key"
    RSet itemType = "Item Type"
    RSet itemValue = "Item Value"
    
    Debug.Print sep0 & index _
          & sep1 & keyType _
          & sep2 & keyValue _
          & sep3 & itemType _
          & sep4 & itemValue _
          & sep5
    
    Debug.Print middle
    
    Dim meCount As Long
    Let meCount = Count
    
    If meCount = 0 Then
        RSet index = "N/A"
        RSet keyType = vbNullString
        RSet keyValue = vbNullString
        RSet itemType = vbNullString
        RSet itemValue = vbNullString
    
        Debug.Print sep0 & index _
              & sep1 & keyType _
              & sep2 & keyValue _
              & sep3 & itemType _
              & sep4 & itemValue _
              & sep5
    End If
    
    Dim i As Long
    
    For i = 1 To IIf(meCount > head, head, meCount)
        RSet index = i
        RSet keyType = TypeName(GetKey(i))
        
        If IsObject(GetKey(i)) Then
            RSet keyValue = "ObjectType"
        Else
            RSet keyValue = GetKey(i)
        End If
        
        RSet itemType = TypeName(GetByIndex(i))
        
        If IsObject(GetByIndex(i)) Then
            RSet itemValue = "ObjectType"
        Else
            RSet itemValue = GetByIndex(i)
        End If
        
        Debug.Print sep0 & index _
                  & sep1 & keyType _
                  & sep2 & keyValue _
                  & sep3 & itemType _
                  & sep4 & itemValue _
                  & sep5
    Next i
    
    If meCount > head + tail Then
        RSet index = ":"
        RSet keyType = vbNullString
        RSet keyValue = vbNullString
        RSet itemType = vbNullString
        RSet itemValue = vbNullString
    
        Debug.Print sep0 & index _
                  & sep1 & keyType _
                  & sep2 & keyValue _
                  & sep3 & itemType _
                  & sep4 & itemValue _
                  & sep5
    End If
    
    For i = IIf(meCount > head + tail, meCount - tail + 1, head + 1) To meCount
        RSet index = i
        RSet keyType = TypeName(GetKey(i))
        
        If IsObject(GetKey(i)) Then
            RSet keyValue = "ObjectType"
        Else
            RSet keyValue = GetKey(i)
        End If
        
        RSet itemType = TypeName(GetByIndex(i))
        
        If IsObject(GetByIndex(i)) Then
            RSet itemValue = "ObjectType"
        Else
            RSet itemValue = GetByIndex(i)
        End If
        
        Debug.Print sep0 & index _
                  & sep1 & keyType _
                  & sep2 & keyValue _
                  & sep3 & itemType _
                  & sep4 & itemValue _
                  & sep5
    Next i
    
    Debug.Print topAndButtom
    Debug.Print "## Mapping Printing End ##"
End Sub

Public Sub ToImmediateHead(Optional ByVal head As Long = 50)
    Call ToImmediate(head, 0)
End Sub

Public Sub ToImmediateTail(Optional ByVal tail As Long = 50)
    Call ToImmediate(0, tail)
End Sub
