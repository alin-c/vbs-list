' Represents a resizeable list, with a handy interface.
Class List
    Private aCollection
    Private iLastIndex
    Private iInitialCapacity

    ' Initializes the list.
    Private Sub Class_Initialize()
        iInitialCapacity = 100
        iLastIndex = - 1
        ReDim aCollection(iInitialCapacity - 1)
    End Sub

    ' Gets the highest index of the list.
    Public Property Get LastIndex()
        LastIndex = iLastIndex
    End Property

    ' Sets the highest index of the list.
    Private Property Let LastIndex(index)
        iLastIndex = index
    End Property

    ' Gets the number of elements in the list.
    Public Property Get Count()
        Count = LastIndex + 1
    End Property

    ' Accesses an element by index.
    Public Property Get Item(index)
        If Not IsInteger(index) Then
            Err.Raise 13, "List", "Index " & GetNotAnIntegerErrorText(index)
        End If
        If index < 0 Or index > LastIndex Then
            Err.Raise 9, "List", GetIndexOutOfBoundsErrorText(index)
        Else
            Item = aCollection(index)
        End If
    End Property

    ' Sets the value of an element at a given index.
    Public Property Let Item(index, vElement)
        If Not IsInteger(index) Then
            Err.Raise 13, "List", "Index " & GetNotAnIntegerErrorText(index)
        End If
        If index < 0 And index > LastIndex Then
            Err.Raise 9, "List", GetIndexOutOfBoundsErrorText(index)
        Else
            aCollection(index) = vElement
        End If
    End Property

    ' Gets the underlying array.
    Public Property Get Collection()
        Call Optimize()
        Collection = aCollection
    End Property

    ' Trims the unused elements in the underlying array.
    Public Sub Optimize()
        If Count > 0 Then
            ReDim Preserve aCollection(LastIndex)
        ElseIf Count = 0 Then
            Call Clear()
        End If
    End Sub

    ' Deletes all the elements.
    Public Sub Clear()
        ReDim aCollection(0)
        LastIndex = - 1
    End Sub

    ' Appends an item to the list.
    Public Sub Append(vElement)
        If Count = 0 Then
            ReDim aCollection(iInitialCapacity - 1)
        End If
        If LastIndex + 1 > UBound(aCollection) Then
            ReDim Preserve aCollection(UBound(aCollection) * 2 + 1)
        End If
        LastIndex = LastIndex + 1
        aCollection(LastIndex) = vElement
    End Sub

    ' Appends an array to the list. When the list is empty, this effectively builds a list with the contents of the array.
    Public Sub AppendArray(aArray)
        If Not IsArray(aArray) Then
            Err.Raise 13, "List", GetNotAnArrayErrorText(aArray)
        End If
        Dim vElement
        For Each vElement In aArray
            Call Append(vElement)
        Next
    End Sub

    ' Inserts an item at a given index into the list.
    Public Sub Insert(vElement, index)
        If Not IsInteger(index) Then
            Err.Raise 13, "List", "Index " & GetNotAnIntegerErrorText(index)
        End If
        If index < 0 Or index > LastIndex Then
            Err.Raise 9, "List", GetIndexOutOfBoundsErrorText(index)
        End If
        If LastIndex + 1 > UBound(aCollection) Then
            ReDim Preserve aCollection(UBound(aCollection) * 2 + 1)
        End If
        Dim i
        For i = LastIndex + 1 To index Step - 1
            aCollection(i + 1) = aCollection(i)
        Next
        aCollection(index) = vElement
        LastIndex = LastIndex + 1
    End Sub

    ' Deletes the element at a specific index.
    Public Sub Delete(index)
        If Not IsInteger(index) Then
            Err.Raise 13, "List", "Index " & GetNotAnIntegerErrorText(index)
        End If
        If index < 0 Or index > LastIndex Then
            Err.Raise 9, "List", GetIndexOutOfBoundsErrorText(index)
        Else
            Call Optimize()
            Dim i
            For i = index + 1 To LastIndex
                aCollection(i - 1) = aCollection(i)
            Next
            aCollection(LastIndex) = Null
            LastIndex = LastIndex - 1
        End If
    End Sub

    ' Gets an array containing a slice of the list. The last index is included in the output.
    Public Function Slice(iStart, iEnd)
        If Not IsInteger(iStart) Then
            Err.Raise 13, "List", "Slice bound " & GetNotAnIntegerErrorText(iStart)
        ElseIf Not IsInteger(iEnd) Then
            Err.Raise 13, "List", "Slice bound " & GetNotAnIntegerErrorText(iEnd)
        End If
        If iStart < 0 Then
            iStart = 0
        End If
        If iEnd > LastIndex Then
            iEnd = LastIndex
        End If
        If iStart > iEnd Then
            Slice = Array()
            Exit Function
        End If
        Dim aSliced(), i, j
        ReDim aSliced(iEnd - iStart)
        j = 0
        For i = iStart To iEnd
            aSliced(j) = aCollection(i)
            j = j + 1
        Next
        Slice = aSliced
    End Function

    ' Checks if a given element is in the list, case insensitive.
    Public Function Contains(vElement)
        Contains = False
        Dim i
        For i = 0 To LastIndex
            If LCase(aCollection(i)) = LCase(vElement) Then
                Contains = True
                Exit Function
            End If
        Next
    End Function

    ' Checks if a given element is in the list.
    Public Function ContainsCaseSensitive(vElement)
        ContainsCaseSensitive = (UBound(Filter(aCollection, vElement)) >= 0)
    End Function

    ' Checks if a given element is in the list and returns its index; if it is not found returns -1.
    Public Function IndexOf(vElement)
        IndexOf = - 1
        Dim i
        For i = 0 To LastIndex
            If LCase(aCollection(i)) = LCase(vElement) Then
                IndexOf = i
                Exit Function
            End If
        Next
    End Function

    ' Sorts the internal array using QuickSort.
    Public Sub Sort(sOrder)
        sOrder = LCase(sOrder)
        Select Case sOrder
            Case "asc", "desc"
            Case Else
                Err.Raise 5, "List", GetInvalidSortArgumentErrorText(sOrder)
        End Select
        Call Optimize()
        Call QuickSort(aCollection, LBound(aCollection), UBound(aCollection), sOrder)
    End Sub

    ' QuickSort utility.
    Private Sub QuickSort(aArray, iLow, iHigh, sOrder)
        If iLow < iHigh Then
            Dim iPartition
            iPartition = Partition(aArray, iLow, iHigh, sOrder)
            Call QuickSort(aArray, iLow, iPartition - 1, sOrder)
            Call QuickSort(aArray, iPartition + 1, iHigh, sOrder)
        End If
    End Sub

    ' QuickSort utility.
    Private Function Partition(aArray, iLow, iHigh, sOrder)
        Dim iPivot, i, j, iTemp
        iPivot = aArray(iHigh)
        i = iLow - 1
        For j = iLow To iHigh - 1
            If sOrder = "asc" Then
                If aArray(j) <= iPivot Then
                    i = i + 1
                    iTemp = aArray(i)
                    aArray(i) = aArray(j)
                    aArray(j) = iTemp
                End If
            ElseIf sOrder = "desc" Then
                If aArray(j) >= iPivot Then
                    i = i + 1
                    iTemp = aArray(i)
                    aArray(i) = aArray(j)
                    aArray(j) = iTemp
                End If
            End If
        Next
        iTemp = aArray(i + 1)
        aArray(i + 1) = aArray(iHigh)
        aArray(iHigh) = iTemp
        Partition = i + 1
    End Function

    ' Cleans up resources when the object is destroyed.
    Private Sub Class_Terminate()
        Erase aCollection
    End Sub

    ' Checks if a given value is an integer (strictly, i.e. also compares the type).
    Private Function IsInteger(vElement)
        IsInteger = False
        If IsNumeric(vElement) And Int(vElement) = vElement Then
            IsInteger = True
        End If
    End Function

    ' Gets the error text for IndexOutOfBounds.
    Private Function GetIndexOutOfBoundsErrorText(index)
        GetIndexOutOfBoundsErrorText = "Index " & index & " is out of bounds (0, " & LastIndex & ")."
    End Function

    ' Gets the error text for NotAnInteger.
    Private Function GetNotAnIntegerErrorText(vElement)
        GetNotAnIntegerErrorText = vElement & " (" & TypeName(vElement) & ") is not an integer."
    End Function

    ' Gets the error text for NotAnArray.
    Private Function GetNotAnArrayErrorText(vElement)
        GetNotAnArrayErrorText = vElement & " (" & TypeName(vElement) & ") is not an array."
    End Function

    ' Gets the error text for InvalidSortArgument.
    Private Function GetInvalidSortArgumentErrorText(vElement)
        GetInvalidSortArgumentErrorText = "Argument " & vElement & " is not valid (it must be either ""asc"" or ""desc"")."
    End Function
End Class