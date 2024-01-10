Attribute VB_Name = "Tool"

'==============================================================================
' 获取注册表中指定键的值.
' Input     | string   | 注册表路径.
' Output    | string   | 参数值.
'==============================================================================
Public Function GetConfigPath(path As String) As String

    Dim wscr As Object
    Set wscr = CreateObject("WScript.Shell")
    
    GetConfigPath = wscr.RegRead(path)
    
End Function


'==============================================================================
' Tells whether exists sheet.
' Input     | string   | Sheet name.
' Output    | bool     | Whether exists sheet.
'==============================================================================
Public Function ExistsSheet(inText As String) As Boolean

    Dim i As Integer
    For i = 1 To Worksheets.count
        If Worksheets(i).Name = inText Then
            ExistsSheet = True
            Exit Function
        End If
    Next i
    
    ExistsSheet = False

End Function


Public Function RemoveString(inText As String, inStrToRemove As String) As String

    Dim res As String
    res = inText
    
    Dim i As Integer
    i = InStr(res, inStrToRemove)
    
    Dim n As Integer
    n = Len(inStrToRemove)
    
    While (i > 0)
        res = Mid(res, 1, i - 1) & Mid(res, i + n)
        i = InStr(res, inStrToRemove)
    Wend
    
    RemoveString = res
    
End Function


Public Function GetParentId(inId As String) As String
    
    If inId = "" Then
        GetParentId = ""
        Exit Function
    End If
    
    Dim i As Integer
    i = InStrRev(inId, Separator)
    If i < 1 Then
        GetParentId = ""
        Exit Function
    End If
        
    GetParentId = Mid(inId, 1, i - 1)
   
End Function


Public Function GetNextId(inId As String) As String
    
On Error GoTo err
    
    If Len(inId) = 0 Then
        GetNextId = ""
        Exit Function
    End If
    
    Dim ix%, iId%
    ix = InStrRev(inId, Separator)
    If ix = 0 Then
        iId = CInt(inId)
        GetNextId = CStr(iId + 1)
    Else
        iId = CInt(Mid(inId, ix + 1))
        GetNextId = Mid(inId, 1, ix) & CStr(iId + 1)
    End If
    Exit Function

err:
    
    GetNextId = ""
    Exit Function

End Function


Public Function AppendArray(arr, val)
    
    Dim n As Integer
    n = SafeUBound(arr)
    ReDim Preserve arr(n + 1)
    arr(n + 1) = val
    
End Function


Public Function ExistInArray(arr, val) As Boolean
    
    ExistInArray = (SearchInArray(arr, val) <> -1)
    
End Function


Public Function RemoveArray(arr, ix)
    
    Dim i%, n%
    n = SafeUBound(arr)
    For i = ix To n - 1
        arr(i) = arr(i + 1)
    Next i
    
    ReDim Preserve arr(n - 1)
    
End Function


Public Function SearchInArray(arr, val) As Integer
    
    Dim i As Integer
    For i = SafeUBound(arr) To 1 Step -1
        If arr(i) = val Then
            SearchInArray = i
            Exit Function
        End If
    Next i
    
    SearchInArray = -1

End Function


Public Function SafeUBound(arr) As Integer

    On Error Resume Next
    
    Dim i&
    i = UBound(arr)
    If err = 0 Then
        SafeUBound = i
    Else
        SafeUBound = 0
    End If
    
    If SafeUBound < 0 Then
        SafeUBound = 0
    End If
    
End Function


Function InsertRow(rowIx As Integer, c1 As Integer, c2 As Integer, c3 As Integer, c4 As Integer)

    Rows(rowIx + 1).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Dim ad1$, ad2$
    If (c1 > 0 And c2 > c1) Then
        ad1 = GetAddress(rowIx + 1, c1)
        ad2 = GetAddress(rowIx + 1, c2)
        Range(ad1 & ":" & ad2).Select
        Selection.Merge
    End If
    
    If c3 > 0 And c4 > c3 Then
        ad1 = GetAddress(rowIx + 1, c3)
        ad2 = GetAddress(rowIx + 1, c4)
        Range(ad1 & ":" & ad2).Select
        Selection.Merge
    End If
    
End Function


Function CopyRow(srcRow As Integer, dstRow As Integer)

    Rows(srcRow).Copy
    Rows(dstRow).Insert Shift:=xlDown
    Application.CutCopyMode = False
    
End Function


Function SetBorder(startRow As Integer, endRow As Integer, startColumn As Integer, endColumn As Integer)

    Range(Cells(startRow, startColumn), Cells(endRow, endColumn)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Function


Public Function GetIdLevel(inId As String) As Integer
    
    GetIdLevel = CountOf(inId, Separator) + 1

End Function


Public Function TrimIdByLevel(inId As String, inLevel As Integer) As String

    Dim level%
    level = GetIdLevel(inId)
    If level = inLevel Then
    
        TrimIdByLevel = inId
        
    ElseIf level < inLevel Then
    
        TrimIdByLevel = ""
        
    Else
    
        Dim i%, ix%
        ix = 0
        For i = 1 To inLevel
            ix = InStr(ix + 1, inId, Separator)
            If ix < 1 Then GoTo err
        Next i
        
        TrimIdByLevel = Mid(inId, 1, ix - 1)
        
    End If
    
    Exit Function
    
err:

    TrimIdByLevel = ""

End Function


Public Function IsChildId(inId As String, inChildId As String, outLevel As Integer) As Boolean
    
    Dim rest As String
    outLevel = 0
    
    Dim le As Integer
    le = Len(inId)
        
    If le = 0 Then
        IsChildId = True
        outLevel = CountOf(rest, Separator)
        If CountOf(rest, Separator) = 0 Then
            outLevel = 1
            Exit Function
        End If
    End If
    
    If Len(inChildId) <= le Then
        IsChildId = False
        Exit Function
    End If

    If Mid(inChildId, 1, le) <> inId Then
        IsChildId = False
        Exit Function
    End If
        
    
    rest = Mid(inChildId, le + 1)
    If Mid(rest, 1, Len(Separator)) <> Separator Then
        IsChildId = False
        Exit Function
    End If
    
    outLevel = CountOf(rest, Separator)
    IsChildId = True
    
End Function


Public Function AppendArrayToArray(arr, arrToAppend)
    
    Dim n As Integer, n1 As Integer
    n = SafeUBound(arr)
    n1 = SafeUBound(arrToAppend)
    ReDim Preserve arr(n + n1)
    
    Dim i As Integer
    For i = 1 To n1
        arr(n + i) = arrToAppend(i)
    Next i
    
End Function


Function GetAddress(rowIx As Integer, columnIx As Integer) As String

On Error GoTo err

    Dim ad$
    ad = ActiveSheet.Cells(rowIx, columnIx).Address
    ad = Replace(ad, "$", "")
    GetAddress = ad
    Exit Function

err:
    GetAddress = ""
    Exit Function
    
End Function


Function MergeCells(rowIx As Integer, column1 As Integer, column2 As Integer)

    Dim ad1 As String, ad2 As String
    ad1 = Cells(rowIx, c1).Address
    ad1 = Replace(ad1, "$", "")
    ad2 = Cells(rowIx, c2).Address
    ad2 = Replace(ad2, "$", "")
    Range(ad1 & ":" & ad2).Merge
    
End Function


Public Function DeleteSheet(sheetName As String)

    Dim displayedAlerts As Boolean
    displayedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = displayedAlerts

End Function


Public Function DeleteRow(row As Integer)

    Rows(row).Delete Shift:=xlUp
    
End Function


Public Function CopySheetAfter(srcSheetName As String, newSheetName As String, Optional afterSheetName As String) As Worksheet
    
    Dim displayedAlerts As Boolean
    displayedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    If afterSheetName = "" Then
        Worksheets(srcSheetName).Copy After:=Worksheets(Worksheets.count)
    Else
        Worksheets(srcSheetName).Copy After:=Worksheets(afterSheetName)
    End If
    
    ActiveSheet.Name = AdjustSheetName(newSheetName)
    
    Set CopySheetAfter = ActiveSheet

    Application.DisplayAlerts = displayedAlerts
    
End Function


Public Function CopySheetBefore(srcSheetName As String, newSheetName As String, Optional beforeSheetName As String) As Worksheet
    
    Dim displayedAlerts As Boolean
    displayedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    If beforeSheetName = "" Then
        Worksheets(srcSheetName).Copy Before:=Worksheets(1)
    Else
        Worksheets(srcSheetName).Copy Before:=Worksheets(beforeSheetName)
    End If
    
    ActiveSheet.Name = AdjustSheetName(newSheetName)
    
    Set CopySheetBefore = ActiveSheet

    Application.DisplayAlerts = displayedAlerts

End Function


Public Function AdjustSheetName(newSheetName As String) As String

    Dim shName As String
    shName = newSheetName
    While (ExistsSheet(shName))
        shName = shName & "_1"
    Wend
    
    AdjustSheetName = shName
    
End Function


Public Function CountOf(text1 As String, text2 As String) As Integer

    CountOf = (Len(text1) - Len(Replace(text1, text2, ""))) / Len(text2)
    
End Function


Function IsEmptyProduct(inRow As Integer) As Boolean
    
    Dim id As String
    id = GetSourceId(inRow)
    IsEmptyProduct = (id = "")
    
End Function


Function IsProduct(inRow As Integer) As Boolean
    
    Dim id As String
    id = GetSourceId(inRow)
    If id = "" Then
        IsProduct = True
    End If
    
    IsProduct = CountOf(id, Separator) = 0
    
End Function


Function hasChild(sheet As Worksheet, row As Integer) As Boolean

    Dim id1$, id2$, id3$
    id1 = Trim(sheet.Cells(row, SourceIndex2Column).Text)
    id2 = Trim(sheet.Cells(row + 1, SourceIndex2Column).Text)

End Function


Function WriteText(filepath As String, content As String)

On Error GoTo err

    Open filepath For Output As #1
        Print #1, content
    Close #1
    
    Exit Function
    
err:
    
    MsgBox "Can not write to " & filepath

End Function
