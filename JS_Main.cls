Sub testing()

Debug.Print ( _
    ChartOptions(StartColumn:=3))
End Sub



'Chart Options  variable prefix Chart & ChartOption
Function ChartOptions(StartColumn As Integer) As String

Dim StartCell As Integer
StartCell = 33

Dim container As String
Dim ChartinteractivityEnabled As String
Dim ChartanimationDuration As String
Dim ChartAnimationEnabled As String
Dim ChartexportFileName As String
Dim ChartexportEnabled As String
Dim ChartzoomEnabled As String
Dim ChartzoomType As String
Dim Charttheme As String
Dim ChartbackgroundColor As String
Dim ChartcolorSet As String
Dim Chartwidth As String
Dim ChartHeight As String
Dim ChartdataPointMaxWidth As String
Dim ChartdataPointMinWidth As String
Dim ChartdataPointWidth As String
Dim aggFunct As String
Dim label_01 As String
Dim label_02 As String
Dim label_03 As String
Dim label_04 As String
Dim label_05 As String
Dim yRng_01 As String
Dim yRng_02 As String
Dim yRng_03 As String
Dim yRng_04 As String
Dim yRng_05 As String
Dim xRng_01 As String
Dim xRng_02 As String
Dim xRng_03 As String
Dim xRng_04 As String
Dim xRng_05 As String
Dim name_01 As String
Dim name_02 As String
Dim name_03 As String
Dim name_04 As String
Dim name_05 As String
Dim datasetName_01 As String
Dim datasetName_02 As String
Dim datasetName_03 As String
Dim datasetName_04 As String
Dim datasetName_05 As String
Dim chartType_01 As String
Dim chartType_02 As String
Dim chartType_03 As String
Dim chartType_04 As String
Dim chartType_05 As String
Dim legendEnabled As String

Dim wb As Workbook: Set wb = Workbooks("JS-Engine")
Dim ws As Worksheet: Set ws = wb.Worksheets("JS_Column")
Dim rng As Range: Set rng = ws.Range("A1:ZZ1000")

container = rng.Cells(StartCell + 0, StartColumn).Value
ChartinteractivityEnabled = rng.Cells(StartCell + 1, StartColumn).Value
ChartanimationDuration = rng.Cells(StartCell + 2, StartColumn).Value
ChartAnimationEnabled = rng.Cells(StartCell + 3, StartColumn).Value
ChartexportFileName = rng.Cells(StartCell + 4, StartColumn).Value
ChartexportEnabled = rng.Cells(StartCell + 5, StartColumn).Value
ChartzoomEnabled = rng.Cells(StartCell + 6, StartColumn).Value
ChartzoomType = rng.Cells(StartCell + 7, StartColumn).Value
Charttheme = rng.Cells(StartCell + 8, StartColumn).Value
ChartbackgroundColor = rng.Cells(StartCell + 9, StartColumn).Value
ChartcolorSet = rng.Cells(StartCell + 10, StartColumn).Value
Chartwidth = rng.Cells(StartCell + 11, StartColumn).Value
ChartHeight = rng.Cells(StartCell + 12, StartColumn).Value
ChartdataPointMaxWidth = rng.Cells(StartCell + 13, StartColumn).Value
ChartdataPointMinWidth = rng.Cells(StartCell + 14, StartColumn).Value
ChartdataPointWidth = rng.Cells(StartCell + 15, StartColumn).Value
aggFunct = rng.Cells(StartCell + 16, StartColumn).Value
label_01 = rng.Cells(StartCell + 17, StartColumn).Value
label_02 = rng.Cells(StartCell + 18, StartColumn).Value
label_03 = rng.Cells(StartCell + 19, StartColumn).Value
label_04 = rng.Cells(StartCell + 20, StartColumn).Value
label_05 = rng.Cells(StartCell + 21, StartColumn).Value
yRng_01 = rng.Cells(StartCell + 22, StartColumn).Value
yRng_02 = rng.Cells(StartCell + 23, StartColumn).Value
yRng_03 = rng.Cells(StartCell + 24, StartColumn).Value
yRng_04 = rng.Cells(StartCell + 25, StartColumn).Value
yRng_05 = rng.Cells(StartCell + 26, StartColumn).Value
xRng_01 = rng.Cells(StartCell + 27, StartColumn).Value
xRng_02 = rng.Cells(StartCell + 28, StartColumn).Value
xRng_03 = rng.Cells(StartCell + 29, StartColumn).Value
xRng_04 = rng.Cells(StartCell + 30, StartColumn).Value
xRng_05 = rng.Cells(StartCell + 31, StartColumn).Value
name_01 = rng.Cells(StartCell + 32, StartColumn).Value
name_02 = rng.Cells(StartCell + 33, StartColumn).Value
name_03 = rng.Cells(StartCell + 34, StartColumn).Value
name_04 = rng.Cells(StartCell + 35, StartColumn).Value
name_05 = rng.Cells(StartCell + 36, StartColumn).Value
datasetName_01 = rng.Cells(StartCell + 37, StartColumn).Value
datasetName_02 = rng.Cells(StartCell + 38, StartColumn).Value
datasetName_03 = rng.Cells(StartCell + 39, StartColumn).Value
datasetName_04 = rng.Cells(StartCell + 40, StartColumn).Value
datasetName_05 = rng.Cells(StartCell + 41, StartColumn).Value
chartType_01 = rng.Cells(StartCell + 42, StartColumn).Value
chartType_02 = rng.Cells(StartCell + 43, StartColumn).Value
chartType_03 = rng.Cells(StartCell + 44, StartColumn).Value
chartType_04 = rng.Cells(StartCell + 45, StartColumn).Value
chartType_05 = rng.Cells(StartCell + 46, StartColumn).Value
legendEnabled = rng.Cells(StartCell + 47, StartColumn).Value






'DATASET SERIES ARRAYS
'##################################################################################################################################

'Dataset Arrays
On Error Resume Next
Dim datasetCount As Integer
datasetCount = 1
Dim labelSets(1 To 5) As Variant
labelSets(1) = label_01
labelSets(2) = label_02
labelSets(3) = label_03
labelSets(4) = label_04
labelSets(5) = label_05

    'Count number of dataset
    For i = 1 To 5
        If labelSets(i) = "" Then Exit For
        datasetCount = datasetCount + 1
    Next
    datasetCount = datasetCount - 1
    
'Y Series
Dim yRngSets(1 To 5) As Variant
yRngSets(1) = yRng_01
yRngSets(2) = yRng_02
yRngSets(3) = yRng_03
yRngSets(4) = yRng_04
yRngSets(5) = yRng_05
'X Series
Dim xRngSets(1 To 5) As Variant
xRngSets(1) = xRng_01
xRngSets(2) = xRng_02
xRngSets(3) = xRng_03
xRngSets(4) = xRng_04
xRngSets(5) = xRng_05

'Name Series
Dim nameSets(1 To 5) As Variant
nameSets(1) = name_01
nameSets(2) = name_02
nameSets(3) = name_03
nameSets(4) = name_04
nameSets(5) = name_05

'Chart Type Series
Dim chartTypeSets(1 To 5) As Variant
chartTypeSets(1) = chartType_01
chartTypeSets(2) = chartType_02
chartTypeSets(3) = chartType_03
chartTypeSets(4) = chartType_04
chartTypeSets(5) = chartType_05

'Dataset Name Series
Dim datasetNameSets(1 To 5) As Variant
datasetNameSets(1) = datasetName_01
datasetNameSets(2) = datasetName_02
datasetNameSets(3) = datasetName_03
datasetNameSets(4) = datasetName_04
datasetNameSets(5) = datasetName_05

'Reset Errors
On Error GoTo 0



'INITILISE INPUT VARIABLES
'##################################################################################################################################

'Aggrigration function sum or count
If aggFunct = "" Then aggFunct = "sum"

ChartOptions = "var chart = new CanvasJS.Chart('" & container & "'," & vbNewLine & "{" & vbNewLine

'        interactivityEnabled   false, true
    If Chart_interactivityEnabled = "" Then
        ChartOptions = ChartOptions & "interactivityEnabled: " & "true" & "," & vbNewLine
    Else
        ChartOptions = ChartOptions & "interactivityEnabled: " & Chart_interactivityEnabled & "," & vbNewLine
    End If

'        animationDuration      # 1200
    If Chart_animationDuration = "" Then
        ChartOptions = ChartOptions & "animationDuration: " & 1200 & "," & vbNewLine
    Else
        ChartOptions = ChartOptions & "animationDuration: " & Chart_animationDuration & "," & vbNewLine
    End If

'        AnimationEnabled       false, true
    If Chart_AnimationEnabled = "" Then
        ChartOptions = ChartOptions & "animationEnabled: " & "true" & "," & vbNewLine
    Else
        ChartOptions = ChartOptions & "animationEnabled: " & Chart_AnimationEnabled & "," & vbNewLine
    End If

'        exportFileName         text
    If Chart_exportFileName = "" Then
        ChartOptions = ChartOptions & "exportFileName: '" & "My Chart" & "'," & vbNewLine
    Else
        ChartOptions = ChartOptions & "exportFileName: '" & Chart_exportFileName & "'," & vbNewLine
    End If

'        exportEnabled          true, false
    If Chart_exportEnabled = "" Then
        ChartOptions = ChartOptions & "exportEnabled: " & "true" & "," & vbNewLine
    Else
        ChartOptions = ChartOptions & "exportEnabled: " & Chart_exportEnabled & "," & vbNewLine
    End If

'        zoomEnabled            true, false
    If Chart_zoomEnabled = "" Then
        ChartOptions = ChartOptions & "zoomEnabled: " & "true" & "," & vbNewLine
    Else
        ChartOptions = ChartOptions & "zoomEnabled: " & Chart_zoomEnabled & "," & vbNewLine
    End If

'        zoomType               x, y, xy
    If Chart_zoomType = "" Then
        ChartOptions = ChartOptions & "zoomType: '" & "x" & "'," & vbNewLine
    Else
        ChartOptions = ChartOptions & "zoomType: '" & Chart_zoomType & "'," & vbNewLine
    End If

'        theme                  light1, light2, dark1, dark2
    If Chart_theme = "" Then
        ChartOptions = ChartOptions & "theme: '" & "light1" & "'," & vbNewLine
    Else
        ChartOptions = ChartOptions & "theme: '" & Chart_theme & "'," & vbNewLine
    End If

'        backgroundColor        //Color red, blue... or #fff, #000...
    If Chart_backgroundColor = "" Then
        ChartOptions = ChartOptions & "backgroundColor: '" & "white" & "'," & vbNewLine
    Else
        ChartOptions = ChartOptions & "backgroundColor: '" & Chart_backgroundColor & "'," & vbNewLine
    End If

'        colorSet               colorSet1, colorSet2, colorSet3 or [//colorSet Array “#4661EE”, “#EC5657”, ...]);
    If Chart_colorSet = "" Then
        ChartOptions = ChartOptions & "colorSet: '" & "colorSet1" & "'," & vbNewLine
    Else
        ChartOptions = ChartOptions & "colorSet: '" & Chart_colorSet & "'," & vbNewLine
    End If

'        width                  #
    If Chart_width = "" Then
        ChartOptions = ChartOptions
    Else
        ChartOptions = ChartOptions & "width: " & Chart_width & "," & vbNewLine
    End If

'        Height                 #
    If Chart_Height = "" Then
        ChartOptions = ChartOptions
    Else
        ChartOptions = ChartOptions & "height: " & Chart_Height & "," & vbNewLine
    End If

'        dataPointMaxWidth      #
    If Chart_dataPointMaxWidth = "" Then
        ChartOptions = ChartOptions
    Else
        ChartOptions = ChartOptions & "dataPointMaxWidth: " & Chart_dataPointMaxWidth & "," & vbNewLine
    End If

'        dataPointMinWidth      #
    If Chart_dataPointMinWidth = "" Then
        ChartOptions = ChartOptions
    Else
        ChartOptions = ChartOptions & "dataPointMinWidth: " & Chart_dataPointMinWidth & "," & vbNewLine
    End If

'        dataPointWidth         #
    If Chart_dataPointWidth = "" Then
        ChartOptions = ChartOptions
    Else
        ChartOptions = ChartOptions & "dataPointWidth: " & Chart_dataPointWidth & "," & vbNewLine
    End If


'FUNCTIONS TO CREATE
'##################################################################################################################################

'LEGEND
'##################################################################################################################################

'XXXXXX___CHANGE RANGE VARIABLES___XXXXXX
ChartOptions = ChartOptions & Legend_Proprerty(legendEnabled:="True", StartColumn:=3)

'TITLE
'##################################################################################################################################
'title Properties

'SUBTITLE
'##################################################################################################################################
'subtitles

'X AXIS
'##################################################################################################################################
'axisX
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
        
'Y AXIS
'##################################################################################################################################
'axisY
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
        
'X AXIS 2
'##################################################################################################################################
'axisX2:
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
        
'Y AXIS 2
'##################################################################################################################################
'axisY2:
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
        
'TOOTTIP
'##################################################################################################################################
'tooltip

'DATA
'##################################################################################################################################
'data
Dim data  As String
Dim chartType As String
Dim tmpStr As String
Dim tmpStr2 As String
Dim lblWB, yRngWB, xRngWB, nameWB As Workbook
Dim lblWS, yRngWS, xRngWS, nameWS As Workbook
Dim lblRng, yRngRng, xRngRng, nameRng As Range







ChartOptions = ChartOptions & "data: [" & vbNewLine
'i = 1
'Debug.Print (labelSets(i))
'DATAPOINTS
'##################################################################################################################################
For i = 1 To datasetCount
    'yRngAdd = yRngSets(i)
    'lblAdd = labelSets(i)
    chartType = chartTypeSets(i)
    
    On Error Resume Next
    tmpStr = labelSets(i)
    Set lblWB = Workbooks(WB_Add(tmpStr))
    Set lblWS = Sheets(Sheet_Add(tmpStr))
    Set lblRng = lblWS.Range(Rng_Add(tmpStr))

    tmpStr2 = yRngSets(i)
    Set yRngWB = Workbooks(WB_Add(tmpStr2))
    Set yRngWS = Sheets(Sheet_Add(tmpStr2))
    Set yRngRng = yRngWS.Range(Rng_Add(tmpStr2))
    On Error GoTo 0
    
    'legendEnabled
    If legendEnabled = "True" Then data = data & "showInLegend: 'true'" & vbNewLine

    'Name
    If datasetNameSets(i) <> "" Then data = data & "name: '" & datasetNameSets(i) & "'" & vbNewLine

    'Chart Type
    data = data & ChartTypesFunction(chartType)

    'Add DataPoints
    data = data & "dataPoints: [" & vbNewLine
        data = data & CreateDatapoints( _
                                        labelRng:=lblWS.Range(Rng_Add(tmpStr)), _
                                        labelWB:=Workbooks(WB_Add(tmpStr)), _
                                        yRng:=yRngWS.Range(Rng_Add(tmpStr2)), _
                                        yWB:=Workbooks(WB_Add(tmpStr2)), _
                                        chartType:=chartType, _
                                        aggFunct:=aggFunct)

    data = data & vbNewLine & "]}," & vbNewLine
    
    'Debug.Print (data)
Next i


'CLOSE CHART OPTIONS
'##################################################################################################################################
data = data & vbNewLine & "]" & vbNewLine & "});" & vbNewLine & "chart.render();" & vbNewLine
ChartOptions = ChartOptions & data

End Function
'yRngSets(i)
'Z = Workbook(WB_Add(yRngSets(i))).Worksheet(Sheet_Add(yRngSets(i))).Range(Rng_Add(yRngSets(i)))
'Workbook(WB_Add(labelSets(i))).Worksheet(Sheet_Add(labelSets(i))).Range(Rng_Add(labelSets(i)))

Function WB_Add(add As String) As String
    Dim fnd2, add_len As Integer
    fnd2 = Application.WorksheetFunction.Find("]", add, 1)
    WB_Add = Mid(add, 2, fnd2 - 2)
End Function


Function Sheet_Add(add As String) As String
    Dim fnd, fnd2, add_len As Integer
    fnd = Application.WorksheetFunction.Find("!", add, 1)
    fnd2 = Application.WorksheetFunction.Find("]", add, 1)
    Sheet_Add = Mid(add, fnd2 + 1, fnd - fnd2 - 1)
End Function

Function Rng_Add(add As String) As String
    Dim fnd, fnd2, add_len As Integer
    fnd = Application.WorksheetFunction.Find("!", add, 1)
    fnd2 = Application.WorksheetFunction.Find("]", add, 1)
    Rng_Add = Right(add, fnd - fnd2 - 1)
End Function











'######Own Sub Set########
'        subtitles              text - subtitles:[{text: "This is a Subtitle", fontColor: "red", fontSize: 30}],
'        axisX2:                axisX2: {title: "Secondary X Axis",},
'        axisY2:                axisY2:{title: "axisY2 Title", interlacedColor: "#F8F1E4"},






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
    '        FontSize:              #
    '        fontFamily:            calibri, arial, tahoma, verdana
    '        fontColor:             //Color red, blue... or #fff, #000...
    '        fontWeight:            normal, lighter, normal, bold, bolder
    '        FontStyle:             normal, italic, oblique
    '        verticalAlign:         top, center, bottom
    '        horizontalAlign:       left, right, center
    '        itemmouseover:         //Uses a Function
    '        itemmousemove:         //Uses a Function
    '        itemmouseout:          //Uses a Function
    '        itemclick:             //Uses a Function
    '        reversed:              true, false
    
    
    'title Properties variable prefix Title & TitleOption           https://canvasjs.com/docs/charts/chart-options/title/
    '        text                   text
    '        wrap                   true, false
    '        maxWidth               #
    '        dockInsidePlotArea     true, false
    '        verticalAlign          top, center, bottom
    '        horizontalAlign        left, right, center
    '        FontSize               #
    '        fontFamily             calibri, arial, tahoma, verdana
    '        fontWeight             normal, lighter, normal, bold, bolder
    '        fontColor              //Color red, blue... or #fff, #000...
    '        FontStyle              normal, italic, oblique
    '        borderThickness        #
    '        cornerRadius           #
    '        BorderColor            //Color red, blue... or #fff, #000...
    '        backgroundColor        //Color red, blue... or #fff, #000...
    '        margin                 #
    '        padding                #
    
    'axisX   Properties variable prefix xisX & xisXOption
    '    title                      text
    '    titleWrap                  true, false
    '    titleMaxWidth              #
    '    titleFontColor             //Color red, blue... or #fff, #000...
    '    titleFontSize              #
    '    titleFontFamily            calibri, arial, tahoma, verdana
    '    titleFontWeight            #
    '    titleFontStyle             normal, italic, oblique
    '    margin                     #
    '    labelBackgroundColor
    '    labelMaxWidth
    '    labelWrap
    '    labelAutoFit
    '    labelAngle
    '    labelFontFamily
    '    labelFontColor
    '    labelFontSize
    '    labelFontWeight
    '    labelFormatter
    '    labelFontStyle
    '    prefix
    '    suffix
    '    valueFormatString
    '    minimum
    '    Maximum
    '    viewportMaximum
    '    viewportMinimum
    '    interval
    '    intervalType
    '    reversed
    '    logarithmic
    '    logarithmBase
    '    tickLength
    '    tickColor
    '    tickThickness
    '    lineColor
    '    lineThickness
    '    lineDashType
    '    gridDashType
    '    interlacedColor
    '    gridThickness
    '    gridColor

            'stripLines
    '            Value
    '            showOnTop
    '            lineDashType
    '            opacity
    '            labelFormatter
    '            StartValue
    '            endValue
    '            Thickness
    '            color
    '            label
    '            labelPlacement
    '            labelAlign
    '            labelBackgroundColor
    '            labelWrap
    '            labelMaxWidth
    '            labelFontFamily
    '            labelFontColor
    '            labelFontSize
    '            labelFontWeight
    '            labelFontStyle

            'crosshair
    '            Enabled
    '            snapToDataPoint
    '            color
    '            opacity
    '            Thickness
    '            lineDashType
    '            valueFormatString
    '            label
    '            labelWrap
    '            labelMaxWidth
    '            labelBackgroundColor
    '            labelFontFamily
    '            labelFontColor
    '            labelFontSize
    '            labelFontWeight
    '            labelFontStyle
    '            labelFormatter

            'scaleBreaks
    '            autoCalculate
    '            maxNumberOfAutoBreaks
    '            collapsibleThreshold
    '            Spacing
    '            type
    '            fillOpacity
    '            color
    '            lineColor
    '            lineThickness
    '            lineDashType

                'customBreaks
    '                StartValue
    '                endValue
    '                Spacing
    '                type
    '                fillOpacity
    '                color
    '                lineColor
    '                lineThickness
    '                lineDashType

    'axisY   Properties variable prefix xisX & xisXOption
    '    title
    '    titleMaxWidth
    '    titleWrap
    '    titleFontColor
    '    titleFontSize
    '    titleFontFamily
    '    titleFontWeight
    '    titleFontStyle
    '    margin
    '    labelBackgroundColor
    '    labelFormatter
    '    labelMaxWidth
    '    labelWrap
    '    labelAutoFit
    '    labelAngle
    '    labelFontFamily
    '    labelFontColor
    '    labelFontSize
    '    labelFontWeight
    '    labelFontStyle
    '    prefix
    '    suffix
    '    valueFormatString
    '    minimum
    '    Maximum
    '    interval
    '    reversed
    '    logarithmic
    '    logarithmBase
    '    viewportMaximum
    '    viewportMinimum
    '    tickLength
    '    tickColor
    '    tickThickness
    '    lineColor
    '    lineThickness
    '    lineDashType
    '    gridDashType
    '    interlacedColor
    '    gridThickness
    '    gridColor
    '    includeZero

        'stripLines
    '        Value
    '        StartValue
    '        endValue
    '        showOnTop
    '        opacity
    '        labelFormatter
    '        lineDashType
    '        Thickness
    '        color
    '        label
    '        labelPlacement
    '        labelAlign
    '        labelBackgroundColor
    '        labelWrap
    '        labelMaxWidth
    '        labelFontFamily
    '        labelFontColor
    '        labelFontSize
    '        labelFontWeight
    '        labelFontStyle

        'crosshair
    '        Enabled
    '        snapToDataPoint
    '        color
    '        opacity
    '        Thickness
    '        lineDashType
    '        valueFormatString
    '        label
    '        labelWrap
    '        labelMaxWidth
    '        labelBackgroundColor
    '        labelFontFamily
    '        labelFontColor
    '        labelFontSize
    '        labelFontWeight
    '        labelFontStyle
    '        labelFormatter

        'scaleBreaks
    '        autoCalculate
    '        collapsibleThreshold
    '        maxNumberOfAutoBreaks
    '        Spacing
    '        type
    '        fillOpacity
    '        color
    '        lineColor
    '        lineThickness
    '        lineDashType

            'customBreaks
        '        StartValue
        '        endValue
        '        Spacing
        '        type
        '        fillOpacity
        '        color
        '        lineColor
        '        lineThickness
        '        lineDashType

    'tooltip
    '    Enabled
    '    fontColor
    '    FontStyle
    '    FontSize
    '    fontFamily
    '    fontWeight
    '    borderThickness
    '    cornerRadius
    '    reversed
    '    contentFormatter
    '    backgroundColor
    '    shared
    '    Content
    '    AnimationEnabled
    '    BorderColor

    'data
'        name
'        Cursor
'        Visible
'        type
'        axisXType
'        axisYType
'        axisXIndex
'        axisYIndex
'        xValueType
'        yValueFormatString
'        zValueFormatString
'        percentFormatString
'        xValueFormatString
'        highlightEnabled
'        connectNullData
'        lineDashType
'        nullDataLineDashType
'        indexLabelLineDashType
'        indexLabelFormatter
'        color
'        lineColor
'        bevelEnabled
'        fillOpacity
'        startAngle
'        indexLabel
'        innerRadius
'        Radius
'        indexLabelPlacement
'        indexLabelMaxWidth
'        indexLabelWrap
'        indexLabelOrientation
'        neckWidth
'        indexLabelBackgroundColor
'        neckHeight
'        indexLabelFontStyle
'        reversed
'        valueRepresents
'        upperBoxColor
'        lowerBoxColor
'        indexLabelFontColor
'        whiskerLength
'        whiskerColor
'        whiskerThickness
'        indexLabelFontSize
'        whiskerDashType
'        indexLabelFontFamily
'        stemColor
'        indexLabelFontWeight
'        stemThickness
'        indexLabelLineColor
'        stemDashType
'        indexLabelLineThickness
'        linkedDataSeriesIndex
'        toolTipContent
'        lineThickness
'        markerType
'        markerColor
'        MarkerSize
'        markerBorderColor
'        markerBorderThickness
'        showInLegend
'        legendText
'        legendMarkerBorderColor
'        legendMarkerBorderThickness
'        legendMarkerType
'        explodeOnClick
'        Click
'        legendMarkerColor
'        mouseover
'        MouseMove
'        mouseout
'        risingColor
'        fallingColor

        'dataPoints
    '        x
    '        y
    '        Z
    '        isIntermediateSum
    '        name
    '        isCumulativeSum
    '        Cursor
    '        label
    '        highlightEnabled
    '        indexLabelLineDashType
    '        indexLabelFormatter
    '        indexLabel
    '        indexLabelPlacement
    '        indexLabelOrientation
    '        indexLabelBackgroundColor
    '        indexLabelFontColor
    '        indexLabelWrap
    '        indexLabelMaxWidth
    '        indexLabelFontSize
    '        indexLabelFontStyle
    '        indexLabelFontFamily
    '        indexLabelFontWeight
    '        indexLabelLineColor
    '        indexLabelLineThickness
    '        toolTipContent
    '        exploded
    '        color
    '        lineColor
    '        lineDashType
    '        showInLegend
    '        legendText
    '        legendMarkerType
    '        legendMarkerColor
    '        legendMarkerBorderColor
    '        legendMarkerBorderThickness
    '        markerType
    '        markerColor
    '        MarkerSize
    '        markerBorderColor
    '        markerBorderThickness
    '        Click
    '        mouseover
    '        MouseMove
    '        mouseout


