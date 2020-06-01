Sub testing()
Debug.Print (ChartOptions("chartContainer2", chartType_01:="barChart", label_01:=Range("A2:A20"), yRng_01:=Range("B2:B20")))
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
    Optional chartType_01 As String, Optional chartType_02 As String, Optional chartType_03 As String, Optional chartType_04 As String, Optional chartType_05 As String _
    ) As String


'Count number of datasets
On Error Resume Next
Dim datasetCount As Integer
datasetCount = 1
Dim labelSets(1 To 5) As Variant
labelSets(1) = label_01
labelSets(2) = label_02
labelSets(3) = label_03
labelSets(4) = label_04
labelSets(5) = label_05
For Each iteme In datasets
    If IsEmpty(datasets(datasetCount)) Then Exit For
    datasetCount = datasetCount + 1
Next
datasetCount = datasetCount - 1
On Error GoTo 0
'datasets = Empty

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



'###Function to Create###
'legend Properties

'title Properties

'subtitles

'axisX
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
'axisY
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
        
'axisX2:
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
        
'axisY2:
    'stripLines
    'crosshair
    'scaleBreaks
        'customBreaks
        
'tooltip

'data
Dim data As String


ChartOptions = ChartOptions & "data: [" & vbNewLine '& "type: 'bar'," & vbNewLine
For i = 1 To datasetCount
    
    data = data & ChartTypesFunction(chartType_01)
        'Add DataPoints
        data = data & "dataPoints: [" & vbNewLine
            data = data & CreateDatapoints(labelRng:=label_01, yRng:=yRng_01, aggFunct:=aggFunct, chartType:=chartType_01)


data = data & vbNewLine & "]}," & vbNewLine

Next i
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
