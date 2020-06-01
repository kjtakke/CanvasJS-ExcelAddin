Function UNIQUEVALUES_Chart(ArryRange, Optional Sorted As Boolean)

Dim TableRange() As Variant, Unique As Variant
Dim counter As Long, InternalCounter As Long
Dim IsUnique As Boolean, UniqueItems As Long, ExternalCounter As Integer
Dim ArryRange2 As Range

TableRange = ArryRange.Value
ReDim Unique(1 To UBound(TableRange), 1 To 1)

UniqueItems = 1

    For counter = 1 To UBound(TableRange)
        IsUnique = True
        If TableRange(counter, 1) = "" Then GoTo En:
        
                For InternalCounter = 1 To UBound(Unique)
                    If TableRange(counter, 1) = Unique(InternalCounter, 1) Then
                        IsUnique = False
                        Exit For
                    End If
                Next InternalCounter

        If IsUnique = True Then
        Unique(UniqueItems, 1) = TableRange(counter, 1)
        UniqueItems = UniqueItems + 1
        End If
En:
    Next counter
    
    For counter = 1 To UBound(Unique)
    If Unique(counter, 1) = Err.Number <> 0 Then Unique(counter, 1) = ""
    
    Next counter
    
UniqueItems = UniqueItems - 2

If Sorted = True Then

    ReDim Sort(1 To UBound(Unique)) As Variant
    
    For i = 1 To UBound(Unique)
        Sort(i) = Unique(i, 1)
    Next i
    
    Quicksort_Chart vArray:=Sort, arrLbound:=LBound(Sort), arrUbound:=UBound(Sort)
    
    Dim Sp As Integer
    Sp = 1
    For i = 1 To UBound(Unique)
        If Sort(i) = "" Then
            GoTo Nx:
        Else
            Unique(Sp, 1) = Sort(i)
            Sp = Sp + 1
        End If
Nx:
    Next i

    
End If


UNIQUEVALUES_Chart = Unique

End Function

'*************************************************************************************************************



'*************************************************************************************************************
Sub Quicksort_Chart(vArray As Variant, arrLbound As Long, arrUbound As Long)
'Sorts a one-dimensional VBA array from smallest to largest
'using a very fast quicksort algorithm variant.
Dim pivotVal As Variant
Dim vSwap    As Variant
Dim tmpLow   As Long
Dim tmpHi    As Long
 
tmpLow = arrLbound
tmpHi = arrUbound
pivotVal = vArray((arrLbound + arrUbound) \ 2)
 
While (tmpLow <= tmpHi) 'divide
   While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
      tmpLow = tmpLow + 1
   Wend
  
   While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
      tmpHi = tmpHi - 1
   Wend
 
   If (tmpLow <= tmpHi) Then
      vSwap = vArray(tmpLow)
      vArray(tmpLow) = vArray(tmpHi)
      vArray(tmpHi) = vSwap
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If
Wend
 
  If (arrLbound < tmpHi) Then Quicksort_Chart vArray, arrLbound, tmpHi 'conquer
  If (tmpLow < arrUbound) Then Quicksort_Chart vArray, tmpLow, arrUbound 'conquer
End Sub
'*************************************************************************************************************

