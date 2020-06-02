Sub testing()

Debug.Print ( _
    ChartOptions( _
        "chartContainer2", _
        chartType_01:="columnChart", _
        label_01:=Range("A2:A57"), _
        yRng_01:=Range("C2:C57"), _
        chartType_02:="lineChart", _
        label_02:=Range("A2:A57"), _
        yRng_02:=Range("D2:D57"), _
        chartType_03:="stackedAreaChart", _
        label_03:=Range("A2:A57"), _
        yRng_03:=Range("g2:g57"), _
        chartType_04:="stackedAreaChart", _
        label_04:=Range("A2:A57"), _
        yRng_04:=Range("f2:f57"), _
        aggFunct:="sum" _
    ) _
)

End Sub



'Chart Options  variable prefix Chart & ChartOption
Function ChartOptions( _
    container As String, Optional Chart_interactivityEnabled As String, Optional Chart_animationDuration As String, _
    Optional Chart_AnimationEnabled As String, Optional Chart_exportFileName As String, Optional Chart_exportEnabled As String, _
    Optional Chart_zoomEnabled As String, Optional Chart_zoomType As String, Optional Chart_theme As String, Optional Chart_backgroundColor As String, _
    Optional Chart_colorSet As String, Optional Chart_width As String, Optional Chart_Height As String, Optional Chart_dataPointMaxWidth As String, _
    Optional Chart_dataPointMinWidth As String, Optional Chart_dataPointWidth As String, Optional aggFunct As String, _
    Optional label_01 As Range, Optional label_02 As Range, Optional label_03 As Range, Optional label_04 As Range, Optional label_05 As Range, _
    Optional yRng_01 As Range, Optional yRng_02 As Range, Optional yRng_03 As Range, Optional yRng_04 As Range, Optional yRng_05 As Range, _
    Optional xRng_01 As Range, Optional xRng_02 As Range, Optional xRng_03 As Range, Optional xRng_04 As Range, Optional xRng_05 As Range, _
    Optional name_01 As Range, Optional name_02 As Range, Optional name_03 As Range, Optional name_04 As Range, Optional name_05 As Range, _
    Optional datasetName_01 As Range, Optional datasetName_02 As Range, Optional datasetName_03 As Range, Optional datasetName_04 As Range, Optional datasetName_05 As Range, _
    Optional chartType_01 As String, Optional chartType_02 As String, Optional chartType_03 As String, Optional chartType_04 As String, Optional chartType_05 As String, _
    Optional legendEnabled As String _
    ) As String



'Dim container As String
'Dim ChartinteractivityEnabled As String
'Dim ChartanimationDuration As String
'Dim ChartAnimationEnabled As String
'Dim ChartexportFileName As String
'Dim ChartexportEnabled As String
'Dim ChartzoomEnabled As String
'Dim ChartzoomType As String
'Dim Charttheme As String
'Dim ChartbackgroundColor As String
'Dim ChartcolorSet As String
'Dim Chartwidth As String
'Dim ChartHeight As String
'Dim ChartdataPointMaxWidth As String
'Dim ChartdataPointMinWidth As String
'Dim ChartdataPointWidth As String
'Dim aggFunct As String
'Dim label01 As String
'Dim label02 As String
'Dim label03 As String
'Dim label04 As String
'Dim label05 As String
'Dim yRng01 As String
'Dim yRng02 As String
'Dim yRng03 As String
'Dim yRng04 As String
'Dim yRng05 As String
'Dim xRng01 As String
'Dim xRng02 As String
'Dim xRng03 As String
'Dim xRng04 As String
'Dim xRng05 As String
'Dim name01 As String
'Dim name02 As String
'Dim name03 As String
'Dim name04 As String
'Dim name05 As String
'Dim datasetName01 As String
'Dim datasetName02 As String
'Dim datasetName03 As String
'Dim datasetName04 As String
'Dim datasetName05 As String
'Dim chartType01 As String
'Dim chartType02 As String
'Dim chartType03 As String
'Dim chartType04 As String
'Dim chartType05 As String
'Dim legendEnabled As String









'DATASET SERIES ARRAYS
'##################################################################################################################################

'Dataset Arrays
On Error Resume Next
Dim datasetCount As Integer
datasetCount = 1
Dim labelSets(1 To 5) As Variant
labelSets(1) = label_01.Address
labelSets(2) = label_02.Address
labelSets(3) = label_03.Address
labelSets(4) = label_04.Address
labelSets(5) = label_05.Address

    'Count number of dataset
    For i = 1 To 5
        If IsEmpty(labelSets(i)) Then Exit For
        datasetCount = datasetCount + 1
    Next
    datasetCount = datasetCount - 1
    
'Y Series
Dim yRngSets(1 To 5) As Variant
yRngSets(1) = yRng_01.Address
yRngSets(2) = yRng_02.Address
yRngSets(3) = yRng_03.Address
yRngSets(4) = yRng_04.Address
yRngSets(5) = yRng_05.Address

'X Series
Dim xRngSets(1 To 5) As Variant
xRngSets(1) = xRng_01.Address
xRngSets(2) = xRng_02.Address
xRngSets(3) = xRng_03.Address
xRngSets(4) = xRng_04.Address
xRngSets(5) = xRng_05.Address

'Name Series
Dim nameSets(1 To 5) As Variant
nameSets(1) = name_01.Address
nameSets(2) = name_02.Address
nameSets(3) = name_03.Address
nameSets(4) = name_04.Address
nameSets(5) = name_05.Address

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


'ChartOptions = ChartOptions & Legend

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
Dim data, chartType As String








ChartOptions = ChartOptions & "data: [" & vbNewLine

'DATAPOINTS
'##################################################################################################################################
For i = 1 To datasetCount
    chartType = chartTypeSets(i)
    
    'legendEnabled
    If legendEnabled = "True" Then data = data & "showInLegend: 'true'" & vbNewLine
    
    'Name
    If datasetNameSets(i) <> "" Then data = data & "name: '" & datasetNameSets(i) & "'" & vbNewLine
    
    'Chart Type
    data = data & ChartTypesFunction(chartType)
    
    'Add DataPoints
    data = data & "dataPoints: [" & vbNewLine
    data = data & CreateDatapoints(labelRng:=Range(labelSets(i)), yRng:=Range(yRngSets(i)), aggFunct:=aggFunct, chartType:=chartType)
    data = data & vbNewLine & "]}," & vbNewLine
Next i


'CLOSE CHART OPTIONS
'##################################################################################################################################
data = data & vbNewLine & "]" & vbNewLine & "});" & vbNewLine & "chart.render();" & vbNewLine
ChartOptions = ChartOptions & data

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


