Sub BoxChartFunctions()
    Dim ary As Variant
    ary = Range("B1:B19").Value
    Dim ans1, ans2, ans3, ans4, ans5 As Variant
    Dim index, values As Variant
    'Array
    ans1 = mdAry_sdAry(ary)
    ans2 = (boxChart(ans1))
    
    'Srting Range
    ans3 = strRng_sdAry("B1:B19", "Sheet1")
    ans4 = (boxChart(ans3))
    
    'BoxChart Multi Dim Array
    index = Range("A1:A19").Value
    values = Range("B1:B19").Value
    ans5 = BoxChart_Agg(index, values)
End Sub

Function BoxChart_Agg(index As Variant, values As Variant) As Variant
' Requires my Unique Function
Dim tempAry1, BoxChart_Agg_Temp, tempAry2 As Variant
Dim C As Long
Dim U As Boolean
Dim UnqAryVals As Variant
ReDim UnqAryVals(1 To UBound(index))

UnqAryVals(1) = index(1, 1)
C = 1
For i = 2 To UBound(index)
    U = True
    For j = 1 To UBound(UnqAryVals)
        If index(i, 1) = UnqAryVals(j) Then U = False
    Next j
    If U = True Then
        UnqAryVals(C + 1) = index(i, 1)
        C = C + 1
    End If
Next i

ReDim Preserve UnqAryVals(1 To C)
ReDim BoxChart_Agg_Temp(1 To C, 1 To 6)
For i = 1 To UBound(UnqAryVals)
    C = 1
    ReDim tempAry1(0 To UBound(values) - 1)
    ReDim tempAry2(0 To 4)
    For j = 1 To UBound(values)
        If index(j, 1) = UnqAryVals(i) Then
            tempAry1(C - 1) = values(j, 1)
            C = C + 1
        End If
    Next j
    ReDim Preserve tempAry1(0 To C - 2)
    tempAry2 = boxChart(tempAry1)
    BoxChart_Agg_Temp(i, 1) = UnqAryVals(i)
    BoxChart_Agg_Temp(i, 2) = tempAry2(0)
    BoxChart_Agg_Temp(i, 3) = tempAry2(1)
    BoxChart_Agg_Temp(i, 4) = tempAry2(2)
    BoxChart_Agg_Temp(i, 5) = tempAry2(3)
    BoxChart_Agg_Temp(i, 6) = tempAry2(4)
Next i

BoxChart_Agg = BoxChart_Agg_Temp
End Function


Public Function BubbleSrt(ArrayIn, Ascending As Boolean)
Dim SrtTemp As Variant
Dim i As Long
Dim j As Long

If Ascending = True Then
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) > ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
Else
    For i = LBound(ArrayIn) To UBound(ArrayIn)
         For j = i + 1 To UBound(ArrayIn)
             If ArrayIn(i) < ArrayIn(j) Then
                 SrtTemp = ArrayIn(j)
                 ArrayIn(j) = ArrayIn(i)
                 ArrayIn(i) = SrtTemp
             End If
         Next j
     Next i
End If

BubbleSrt = ArrayIn
End Function

Function mdAry_sdAry(ary As Variant) As Variant
    Dim mdAry_sdAry_temp As Variant
    ReDim mdAry_sdAry_temp(0 To UBound(ary) - 1)
    For i = 1 To UBound(ary)
        mdAry_sdAry_temp(i - 1) = ary(i, 1)
    Next i
    mdAry_sdAry = mdAry_sdAry_temp
End Function

Function strRng_sdAry(rng As String, ws As String) As Variant
    Dim ary As Variant
    Dim strRng_sdAry_temp
    ary = Worksheets(ws).Range(rng).Value
    ReDim strRng_sdAry_temp(0 To UBound(ary) - 1)
    For i = 1 To UBound(ary)
        strRng_sdAry_temp(i - 1) = ary(i, 1)
    Next i
    strRng_sdAry = strRng_sdAry_temp
End Function

Function boxChart(srtAry As Variant) As Variant
    Dim boxChart_temp As Variant
    ReDim boxChart_temp(0 To 4)
    srtAry = BubbleSrt(srtAry, True)
    boxChart_temp(4) = srtAry(UBound(srtAry))
    boxChart_temp(0) = srtAry(0)
    boxChart_temp(1) = Round(boxChart_temp(4) / 4, 0)
    boxChart_temp(3) = boxChart_temp(1) * 3
    boxChart_temp(2) = boxChart_temp(1) * 2
    boxChart = boxChart_temp
End Function


Function AryFromCell(WS As String, index As String, Optional values As String) As Variant

Dim tmpAry, rngAry As Variant
Dim H As Long
Dim rng As Range
Dim cAdd1, cAdd2 As String

If values = "" Then

AryFromCell = Worksheets(WS).Range(index, Worksheets(WS).Range(index).End(xlDown)).Value

Else

Set rng = Range(values)

cAdd1 = rng.Row & "," & rng.Column

rngAry = Split(cAdd1, ",")

tmpAry = Worksheets(WS).Range(index, Worksheets(WS).Range(index).End(xlDown)).Value

H = rngAry(0) + UBound(tmpAry)

cAdd2 = H & ", " & rngAry(1)

AryFromCell = Worksheets(WS).Range(Cells.Rows(Int(rngAry(0))).Columns(Int(rngAry(1))), _
                     Cells.Rows(Int(rngAry(0)) + UBound(tmpAry) - 1).Columns(Int(rngAry(1)))).Value

End If
End Function

Function TwoArraysToOneDataSet(ary1 As Variant, ary2 As Variant) As Variant

Dim ary As Variant
ReDim ary(1 To UBound(ary1), 1 To 2)

For i = 1 To UBound(ary1)

ary(i, 1) = ary1(i, 1)
ary(i, 2) = ary2(i, 1)

Next i

TwoArraysToOneDataSet = ary

End Function


Function mergeTwoArrays(agg As String, index As Variant, values As Variant) As Variant

Dim ary1, ary2, unique As Variant
Dim c As Long
Dim avg As Variant
Dim isSame As Boolean
ReDim unique(1 To UBound(index))


'Unique value set
unique(1) = index(1, 1)
c = 1

For i = 2 To UBound(index)
    isSame = False
    
    For j = 1 To UBound(unique)
        If unique(j) = index(i, 1) Then isSame = True
    Next j
    
    If isSame = False Then
    c = c + 1
    unique(c) = index(i, 1)
    End If
Next i

ReDim ary2(1 To c, 1 To 2)
For i = 1 To UBound(ary2)
ary2(i, 1) = unique(i)
Next i

ReDim unique(1 To c, 1 To 2)
unique = ary2
ary2 = Empty

'Aggregation
For i = 1 To UBound(unique)
    ReDim avg(1 To 2)
    avg(1) = 0
    avg(2) = 0
    For j = 1 To UBound(values)
        If index(j, 1) = unique(i, 1) Then
            avg(1) = avg(1) + 1
            If agg <> "count" Then avg(2) = avg(2) + values(j, 1)
        End If
    Next j
    Select Case True
        Case agg = "sum"
            unique(i, 2) = avg(2)
        Case agg = "average"
            unique(i, 2) = Round(avg(2) / avg(1), 0)
        Case agg = "count"
            unique(i, 2) = avg(1)
    End Select
Next i



'return
mergeTwoArrays = unique
End Function


Function BubbleDataPoints(ary As Variant) As String

BubbleDataPoints = ""

For i = 1 To UBound(ary)

    BubbleDataPoints = BubbleDataPoints & "{  name: '" & ary(i, 1) & "' x: " & ary(i, 2) & " y: " & ary(i, 3) & " z: " & ary(i, 4) & " }," & vbNewLine

Next i

End Function


Function TwoDataPointsNumber(ary As Variant) As String

TwoDataPointsNumber = ""

For i = 1 To UBound(ary)

    TwoDataPointsNumber = TwoDataPointsNumber & "{ label: " & ary(i, 1) & " y: " & ary(i, 2) & " }," & vbNewLine

Next i

End Function


Function TwoDataPointsText(ary As Variant) As String

TwoDataPointsText = ""

For i = 1 To UBound(ary)

    TwoDataPointsText = TwoDataPointsText & "{ label: '" & ary(i, 1) & "' y: " & ary(i, 2) & " }," & vbNewLine

Next i

End Function


Function BoxDataPoints(ary As Variant) As String

BoxDataPoints = ""

For i = 1 To UBound(ary)

    BoxDataPoints = BoxDataPoints & "{ label: '" & ary(i, 1) & "' y: " & ary(i, 2) & ", " & ary(i, 3) & ", " & ary(i, 4) & ", " & ary(i, 5) & ", " & ary(i, 6) & " }," & vbNewLine

Next i

End Function


Function RangePoints(ary As Variant) As String

RangePoints = ""

For i = 1 To UBound(ary)
    RangePoints = RangePoints & "{ x: " & i * 10 & ", y:[" & ary(i, 2) & ", " & ary(i, 3) & "], label: '" & ary(i, 1) & "'}," & vbNewLine
Next i

End Function

