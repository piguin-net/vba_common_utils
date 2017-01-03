Sub p_set(ByRef dst, ByRef src)
    On Error GoTo setError
    Set dst = src
    Exit Sub
setError:
    dst = src
End Sub

Function CallByNameN(Object As Object, ProcName As String, CallType As VbCallType, ByRef args() As Variant) As Variant
    Select Case UBound(args)
        Case -1: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType))
        Case 0: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0)))
        Case 1: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1)))
        Case 2: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2)))
        Case 3: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2), args(3)))
        Case 4: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2), args(3), args(4)))
        Case 5: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2), args(3), args(4), args(5)))
        Case 6: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6)))
        Case 7: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)))
        Case 8: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8)))
        Case 9: Call p_set(CallByNameN, CallByName(Object, ProcName, CallType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9)))
        Case Else: Call Err.Raise(Number:=999, Source:="CallByNameN", Description:="指定可能な引数を超えています。")
    End Select
End Function

Function p_zip(ByRef list1() As Variant, ByRef list2() As Variant) As Variant()
    Dim result() As Variant
    Dim i As Long
    result = Array()
    For i = IIf(LBound(list1) < LBound(list2), LBound(list1), LBound(list2)) To IIf(UBound(list1) > UBound(list2), UBound(list1), UBound(list2))
        Call p_array_push(result, True, Array(list1(i), list2(i)))
    Next i
    Call p_set(p_zip, result)
End Function

Sub p_swap(ByRef left As Variant, ByRef right As Variant)
    Dim temp As Variant
    Call p_set(temp, left)
    Call p_set(left, right)
    Call p_set(right, temp)
End Sub

Sub p_unpack(ByRef list() As Variant, ParamArray args())
    Dim i As Long
    For i = LBound(list) To UBound(list)
        Call p_set(args(i), list(i))
    Next i
End Sub

Function p_array_concat(ParamArray args()) As Variant()
    Dim result(), arg, x
    result = Array()
    For Each arg In args
        If IsArray(arg) Then
            For Each x In arg
                ReDim Preserve result(UBound(result) + 1)
                Call p_set(result(UBound(result)), x)
            Next
        Else
            ReDim Preserve result(UBound(result) + 1)
            Call p_set(result(UBound(result)), arg)
        End If
    Next
    Call p_set(p_array_concat, result)
End Function

Function p_array_copy(ByRef list() As Variant) As Variant()
    Dim result() As Variant
    Dim item As Variant
    result = Array()
    For Each item In list
        ReDim Preserve result(UBound(result) + 1)
        Call p_set(result(UBound(result)), item)
    Next
    Call p_set(p_array_copy, result)
End Function

Function p_array_sort(ByRef list() As Variant, ByRef impl As Object, ByVal method As String, ByVal destructive As Boolean, ParamArray args() As Variant) As Variant()
    Dim result() As Variant
    Dim i, j As Long
    
    If destructive Then
        result = list
    Else
        result = p_array_copy(list)
    End If
    
    For i = LBound(result) To UBound(result): For j = UBound(result) To i Step -1
        If CallByNameN(impl, method, VbMethod, p_array_concat(result(i), result(j), args)) Then
            Call p_swap(result(i), result(j))
        End If
    Next j: Next i
    
    Call p_set(p_array_sort, result)
End Function

Function p_array_reverse(ByRef list() As Variant, ByVal destructive As Boolean) As Variant()
    Dim result() As Variant
    Dim i As Long
    
    If destructive Then
        result = list
    Else
        result = p_array_copy(list)
    End If
    
    For i = LBound(result) To UBound(result) / 2
        Call p_swap(result(i), result(UBound(result) - i + 1))
    Next i
    
    Call p_set(p_array_reverse, result)
End Function

Function p_array_unshift(ByRef list() As Variant, ByVal destructive As Boolean, ParamArray args() As Variant)
    Dim result() As Variant
    Dim n, i, j As Long
    
    If destructive Then
        result = list
    Else
        result = p_array_copy(list)
    End If
    
    n = UBound(result)
    ReDim Preserve result(n + UBound(args))
    
    For i = LBound(result) To n
        Call p_set(result(x + i), result(i))
    Next i
    
    For i = LBound(args) To UBound(args)
        Call p_set(result(i), args(i))
    Next i
    
    Call p_set(p_array_unshift, result)
End Function

Function p_array_push(ByRef list() As Variant, ByVal destructive As Boolean, ParamArray args() As Variant)
    Dim result() As Variant
    Dim i As Long
    
    If destructive Then
        result = list
    Else
        result = p_array_copy(list)
    End If
    
    ReDim Preserve result(UBound(result) + UBound(args))
    
    For i = LBound(args) To UBound(args)
        Call p_set(UBound(result), args(i))
    Next i
    
    Call p_set(p_array_push, result)
End Function

Function p_array_map(ByRef list() As Variant, ByRef impl As Object, ByVal method As String, ByVal destructive As Boolean, ParamArray args() As Variant) As Variant()
    Dim result() As Variant
    Dim i As Long
    
    If destructive Then
        result = list
    Else
        result = p_array_copy(list)
    End If
    
    For i = LBound(result) To UBound(result)
        Call p_set(result(i), CallByNameN(impl, method, VbMethod, p_array_concat(result(i), args)))
    Next
    
    Call p_set(p_array_map, result)
End Function

Function p_array_filter(ByRef list() As Variant, ByRef impl As Object, ByVal method As String, ByVal destructive As Boolean, ByVal getIndex As Boolean, ParamArray args() As Variant) As Variant()
    Dim result() As Variant
    Dim i, j As Long
    
    If destructive Then
        result = list
    Else
        result = p_array_copy(list)
    End If
    
    j = LBound(result)
    For i = LBound(result) To UBound(result)
        If CallByNameN(impl, method, VbMethod, p_array_concat(item, args)) Then
            result(j) = result(i)
            j = j + 1
        End If
    Next i
    ReDim Preserve result(j)
    
    Call p_set(p_array_filter, result)
End Function
