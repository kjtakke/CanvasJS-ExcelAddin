
Sub LegendTest()

Debug.Print (Legend_Proprerty(legendEnabled:="True", StartColumn:=3))

End Sub
Function Legend_Proprerty(legendEnabled As String, StartColumn As Integer)

If legendEnabled = "False" Then GoTo enLegeond:
Legend = Legend & "legend: {" & vbNewLine
StartCell = 81

'LEGEND VERABLES
'##################################################################################################################################

'Legend Options
Dim Legend_Cursor As String
Dim Legend_maxWidth As String
Dim Legend_maxHeight As String
Dim Legend_itemWrap As String
Dim Legend_itemMaxWidth As String
Dim Legend_itemWidth As String
Dim Legend_dockInsidePlotArea As String
Dim Legend_fontSize As String
Dim Legend_fontFamily As String
Dim Legend_fontColor As String
Dim Legend_fontWeight As String
Dim Legend_fontStyle As String
Dim Legend_verticalAlign As String
Dim Legend_horizontalAlign As String
Dim Legend_itemclick As String
Dim Legend_reversed As String

'Option Locations
Dim wb As Workbook: Set wb = Workbooks("JS-Engine")
Dim ws As Worksheet: Set ws = wb.Worksheets("JS_Column")
Dim rng As Range: Set rng = ws.Range("A50:ZZ1000")


'Assign Options to Variables
Legend_Cursor = rng.Cells(StartCell + 0, StartColumn).Value
Legend_reversed = rng.Cells(StartCell + 1, StartColumn).Value
Legend_maxWidth = rng.Cells(StartCell + 2, StartColumn).Value
Legend_maxHeight = rng.Cells(StartCell + 3, StartColumn).Value
Legend_itemWrap = rng.Cells(StartCell + 4, StartColumn).Value
Legend_itemMaxWidth = rng.Cells(StartCell + 5, StartColumn).Value
Legend_itemWidth = rng.Cells(StartCell + 6, StartColumn).Value
Legend_dockInsidePlotArea = rng.Cells(StartCell + 7, StartColumn).Value
Legend_fontSize = rng.Cells(StartCell + 8, StartColumn).Value
Legend_fontFamily = rng.Cells(StartCell + 9, StartColumn).Value
Legend_fontColor = rng.Cells(StartCell + 10, StartColumn).Value
Legend_fontWeight = rng.Cells(StartCell + 11, StartColumn).Value
Legend_fontStyle = rng.Cells(StartCell + 12, StartColumn).Value
Legend_verticalAlign = rng.Cells(StartCell + 13, StartColumn).Value
Legend_horizontalAlign = rng.Cells(StartCell + 14, StartColumn).Value
Legend_itemclick = rng.Cells(StartCell + 15, StartColumn).Value
Legend_reversed = rng.Cells(StartCell + 16, StartColumn).Value


'CREATE JS
'##################################################################################################################################


'Cursor
If Legend_Cursor = "Pointer" Then
    Legend = Legend & "cursor: 'pointer'," & vbNewLine
    
ElseIf Legend_Cursor = "Crosshair" Then
    Legend = Legend & "cursor: 'crosshair'," & vbNewLine
Else
    Legend = Legend & "cursor: 'crosshair'," & vbNewLine
End If

'Reversed
If Legend_reversed = "True" Then
    Legend = Legend & "reversed: true," & vbNewLine
    
ElseIf Legend_reversed = "False" Then
    Legend = Legend & "reversed: false," & vbNewLine
Else
    Legend = Legend & "reversed: false," & vbNewLine
End If

'maxWidth
If Legend_maxWidth = "" Then
    GoTo en_LegendmaxWidth:
Else
    Legend = Legend & "maxWidth: " & Legend_maxWidth & "," & vbNewLine
End If
en_LegendmaxWidth:


'maxHeight
If Legend_maxHeight = "" Then
    GoTo en_LegendmaxHeight:
Else
    Legend = Legend & "maxHeight: " & Legend_maxHeight & "," & vbNewLine
End If
en_LegendmaxHeight:


'ItemWrap
If Legend_itemWrap = "True" Then
    Legend = Legend & "itemWrap: true," & vbNewLine
ElseIf Legend_itemWrap = "False" Then
    Legend = Legend & "itemWrap: false," & vbNewLine
Else
    Legend = Legend & "itemWrap: false," & vbNewLine
End If

'itemMaxWidth
If Legend_itemMaxWidth = "" Then
    GoTo en_LegenditemMaxWidth:
Else
    Legend = Legend & "itemMaxWidth: " & Legend_itemMaxWidth & "," & vbNewLine
End If
en_LegenditemMaxWidth:


'itemWidth
If Legend_itemWidth = "" Then
    GoTo en_LegenditemWidth:
Else
    Legend = Legend & "itemWidth: " & Legend_itemWidth & "," & vbNewLine
End If
en_LegenditemWidth:


'dockInsidePlotArea
If Legend_dockInsidePlotArea = "True" Then
    Legend = Legend & "dockInsidePlotArea: true," & vbNewLine
ElseIf Legend_dockInsidePlotArea = "False" Then
    Legend = Legend & "dockInsidePlotArea: false," & vbNewLine
Else
    Legend = Legend & "dockInsidePlotArea: false," & vbNewLine
End If

'fontSize
If Legend_fontSize = "" Then
    GoTo en_LegendfontSize:
Else
    Legend = Legend & "fontSize: " & Legend_fontSize & "," & vbNewLine
End If
en_LegendfontSize:

'fontFamily
If Legend_fontFamily = "" Then
    GoTo en_LegendFontFamily:
Else
    Legend = Legend & "fontFamily: '" & Legend_fontFamily & "'," & vbNewLine
End If
en_LegendFontFamily:

'fontColor
If Legend_fontColor = "" Then
    GoTo en_LegendFontColor:
Else
    Legend = Legend & "fontColor: '" & Legend_fontColor & "'," & vbNewLine
End If
en_LegendFontColor:

'fontWeight
If Legend_fontWeight = "" Then
    GoTo en_LegendFontWeight:
Else
    Legend = Legend & "fontWeight: '" & Legend_fontWeight & "'," & vbNewLine
End If
en_LegendFontWeight:

'fontStyle
If Legend_fontStyle = "" Then
    GoTo en_LegendFontStyle:
Else
    Legend = Legend & "fontStyle: '" & Legend_fontStyle & "'," & vbNewLine
End If
en_LegendFontStyle:


'verticalAlign
If Legend_verticalAlign = "" Then
    GoTo en_LegendVerticalAlign:
Else
    Legend = Legend & "verticalAlign: '" & Legend_verticalAlign & "'," & vbNewLine
End If
en_LegendVerticalAlign:


'horizontalAlign
If Legend_horizontalAlign = "" Then
    GoTo en_LegendHorizontalAlign:
Else
    Legend = Legend & "horizontalAlign: '" & Legend_horizontalAlign & "'," & vbNewLine
End If
en_LegendHorizontalAlign:


'itemclick
If Legend_itemclick = "" Then
    GoTo en_LegendItemclick:
Else
    Legend = Legend & "itemclick: toggleDataSeries," & vbNewLine
End If
en_LegendItemclick:

Legend = Legend & "}," & vbNewLine

Legend_Proprerty = Legend

enLegeond:

End Function

'ADD TO JS
'##################################################################################################################################
'Onclick
'function toggleDataSeries(e) {
'    if (typeof (e.dataSeries.visible) === "undefined" || e.dataSeries.visible) {
'        e.dataSeries.visible = false;
'    } else {
'        e.dataSeries.visible = true;
'    }
'    e.chart.render();
'}

'LEGEND PROPERTIES
'##################################################################################################################################
    'legend Properties  variable prefix Legend & LegendOption       https://canvasjs.com/docs/charts/chart-options/legend/
    '        cursor:                pointer, crosshair
    '        reversed:              true, false
    '        maxWidth:              #
    '        maxHeight:             #
    '        itemWrap:              true, false
    '        itemMaxWidth:          #
    '        itemWidth:             #
    '        itemTextFormatter:     //Uses a Function
    '        markerMargin:          #
    '        dockInsidePlotArea:    true, false
    '        fontSize:              #
    '        fontFamily:            calibri, arial, tahoma, verdana
    '        fontColor:             //Color red, blue... or #fff, #000...
    '        fontWeight:            normal, lighter, normal, bold, bolder
    '        fontStyle:             normal, italic, oblique
    '        verticalAlign:         top, center, bottom
    '        horizontalAlign:       left, right, center
    '        itemmouseover:         //Uses a Function
    '        itemmousemove:         //Uses a Function
    '        itemmouseout:          //Uses a Function
    '        itemclick:             //Uses a Function
