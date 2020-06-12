Sub TestJSchatrs()

Dim WebDev As WebDev
Set WebDev = New WebDev
Dim WebPage As String

Dim chart_1_Legend  As String
chart_1_Legend = JSchartLegend(cursor:="pointer", _
                            reversed:=WebDev.JSfalse, _
                            maxWidth:="50", _
                            maxHeight:="50", _
                            itemWrap:=WebDev.JStrue, _
                            itemMaxWidth:="50", _
                            itemWidth:="50", _
                            markerMargin:="4", _
                            dockInsidePlotArea:=WebDev.JSfalse, _
                            fontSize:="12", _
                            fontFamily:="calibri", _
                            fontColor:="black", _
                            fontWeight:="bold", _
                            FontStyle:="italic", _
                            verticalAlign:="top", _
                            horizontalAlign:="center", _
                            itemclick:=True)
                            
Dim Chart_1_Title As String
Chart_1_Title = JSchartTitle(text:="My chart title", _
                        wrap:=WebDev.JSfalse, _
                        maxWidth:="100", _
                        dockInsidePlotArea:=WebDev.JSfalse, _
                        verticalAlign:="top", _
                        horizontalAlign:="center", _
                        fontSize:="72", _
                        fontFamily:="calibri", _
                        fontWeight:="bold", _
                        fontColor:="black", _
                        FontStyle:="italic", _
                        borderThickness:="1", _
                        cornerRadius:="15", _
                        borderColor:="black", _
                        backgroundColor:="white", _
                        margin:="5", _
                        padding:="5")



Dim Chart_1_Tooltip As String
Chart_1_Tooltip = JSchartTooltop(enabled:="", _
                                    fontColor:="black", _
                                    FontStyle:="italic", _
                                    fontSize:="10", _
                                    fontFamily:="calibri", _
                                    fontWeight:="bold", _
                                    borderThickness:="1", _
                                    cornerRadius:="15", _
                                    reversed:=WebDev.JSfalse, _
                                    backgroundColor:="white", _
                                    sharedValues:=WebDev.JStrue, _
                                    content:="{label}: {y}", _
                                    animationEnabled:=WebDev.JStrue, _
                                    borderColor:="black")

Debug.Print (JSchartMain( _
    container:="container1", interactivityEnabled:=WebDev.JStrue, animationDuration:="1200", animationEnabled:="true", exportFileName:="Mychart", _
    exportEnabled:=WebDev.JStrue, zoomEnabled:=WebDev.JStrue, zoomType:="xy", theme:="light1", backgroundColor:="white", ChartWidth:="50", _
    ChartHeight:="50", dataPointMaxWidth:="50", dataPointMinWidth:="50", dataPointWidth:="50", colorSet:="colorSet1", _
        ChartTitle:=Chart_1_Title, _
        ChartSubtitles:="***", _
        axisX:="***", _
        axisX2:="***", _
        axisY:="***", _
        axisY2:="***", _
        legend:=chart_1_Legend, _
        tooltip:=Chart_1_Tooltip, _
        data:="***"))
        
        

        
        
End Sub

Function JSchartMain(container As String, _
                        Optional interactivityEnabled As String, _
                        Optional animationDuration As String, _
                        Optional animationEnabled As String, _
                        Optional exportFileName As String, _
                        Optional exportEnabled As String, _
                        Optional zoomEnabled As String, _
                        Optional zoomType As String, _
                        Optional theme As String, _
                        Optional backgroundColor As String, _
                        Optional ChartWidth As String, _
                        Optional ChartHeight As String, _
                        Optional dataPointMaxWidth As String, _
                        Optional dataPointMinWidth As String, _
                        Optional dataPointWidth As String, _
                        Optional colorSet As String, _
                        Optional ChartTitle As String, _
                        Optional ChartSubtitles As String, _
                        Optional axisX As String, _
                        Optional axisX2 As String, _
                        Optional axisY As String, _
                        Optional axisY2 As String, _
                        Optional legend As String, _
                        Optional tooltip As String, _
                        Optional data As String) As String
                        
    'Main Chart Options
    JSchartMain = "var  chart =  new  CanvasJS.Chart('" & container & "'," & vbNewLine & "{" & vbNewLine
    If interactivityEnabled <> "" Then JSchartMain = JSchartMain & "interactivityEnabled: " & interactivityEnabled & "," & vbNewLine
    If animationDuration <> "" Then JSchartMain = JSchartMain & "animationDuration: " & animationDuration & "," & vbNewLine
    If animationEnabled <> "" Then JSchartMain = JSchartMain & "animationEnabled: " & animationEnabled & "," & vbNewLine
    If exportFileName <> "" Then JSchartMain = JSchartMain & "exportFileName: " & exportFileName & "," & vbNewLine
    If exportEnabled <> "" Then JSchartMain = JSchartMain & "exportEnabled: " & exportEnabled & "," & vbNewLine
    If zoomType <> "" Then JSchartMain = JSchartMain & "zoomType: '" & zoomType & "'," & vbNewLine
    If theme <> "" Then JSchartMain = JSchartMain & "theme: '" & theme & "'," & vbNewLine
    If backgroundColor <> "" Then JSchartMain = JSchartMain & "backgroundColor: '" & backgroundColor & "'," & vbNewLine
    If ChartWidth <> "" Then JSchartMain = JSchartMain & "width: " & ChartWidth & "," & vbNewLine
    If ChartWidth <> "" Then JSchartMain = JSchartMain & "height: " & ChartWidth & "," & vbNewLine
    If dataPointMaxWidth <> "" Then JSchartMain = JSchartMain & "dataPointMaxWidth: " & dataPointMaxWidth & "," & vbNewLine
    If dataPointMinWidth <> "" Then JSchartMain = JSchartMain & "dataPointMinWidth: " & dataPointMinWidth & "," & vbNewLine
    If colorSet <> "" Then JSchartMain = JSchartMain & "colorSet: '" & colorSet & "'," & vbNewLine
    
    'Has Sub functions
    If ChartTitle <> "" Then JSchartMain = JSchartMain & "title:{ " & ChartTitle & "}," & vbNewLine
    If ChartSubtitles <> "" Then JSchartMain = JSchartMain & "subtitles:[ " & ChartSubtitles & "]," & vbNewLine
    If axisX <> "" Then JSchartMain = JSchartMain & "axisX:[ " & axisX & "]," & vbNewLine
    If axisX2 <> "" Then JSchartMain = JSchartMain & "axisX2:[ " & axisX2 & "]," & vbNewLine
    If axisY <> "" Then JSchartMain = JSchartMain & "axisY:[ " & axisY & "]," & vbNewLine
    If axisY2 <> "" Then JSchartMain = JSchartMain & "axisY2:[ " & axisY2 & "]," & vbNewLine
    If legend <> "" Then JSchartMain = JSchartMain & "legend:{ " & legend & "}," & vbNewLine
    If tooltip <> "" Then JSchartMain = JSchartMain & "tooltip:{ " & tooltip & "}," & vbNewLine
    If data <> "" Then JSchartMain = JSchartMain & "data:[ " & data & "]," & vbNewLine
    
    'Render Chart
    JSchartMain = JSchartMain & "});" & vbNewLine & "chart.render();" & vbNewLine & vbNewLine
End Function


Function JSchartLegend(Optional cursor As String, _
                        Optional reversed As String, _
                        Optional maxWidth As String, _
                        Optional maxHeight As String, _
                        Optional itemWrap As String, _
                        Optional itemMaxWidth As String, _
                        Optional itemWidth As String, _
                        Optional markerMargin As String, _
                        Optional dockInsidePlotArea As String, _
                        Optional fontSize As String, _
                        Optional fontFamily As String, _
                        Optional fontColor As String, _
                        Optional fontWeight As String, _
                        Optional FontStyle As String, _
                        Optional verticalAlign As String, _
                        Optional horizontalAlign As String, _
                        Optional itemclick As Boolean _
                        ) As String
    JSchartLegend = ""
    If cursor <> "" Then JSchartLegend = JSchartLegend & "cursor: '" & cursor & "'," & vbNewLine
    If reversed <> "" Then JSchartLegend = JSchartLegend & "reversed: " & reversed & "," & vbNewLine
    If maxWidth <> "" Then JSchartLegend = JSchartLegend & "maxWidth: " & maxWidth & "," & vbNewLine
    If maxHeight <> "" Then JSchartLegend = JSchartLegend & "maxHeight: " & maxHeight & "," & vbNewLine
    If itemWrap <> "" Then JSchartLegend = JSchartLegend & "itemWrap: " & itemWrap & "," & vbNewLine
    If itemMaxWidth <> "" Then JSchartLegend = JSchartLegend & "itemMaxWidth: " & itemMaxWidth & "," & vbNewLine
    If itemWidth <> "" Then JSchartLegend = JSchartLegend & "itemWidth: " & itemWidth & "," & vbNewLine
    If markerMargin <> "" Then JSchartLegend = JSchartLegend & "markerMargin: " & markerMargin & "," & vbNewLine
    If dockInsidePlotArea <> "" Then JSchartLegend = JSchartLegend & "dockInsidePlotArea: " & dockInsidePlotArea & "," & vbNewLine
    If fontSize <> "" Then JSchartLegend = JSchartLegend & "FontSize: " & fontSize & "," & vbNewLine
    If fontFamily <> "" Then JSchartLegend = JSchartLegend & "fontFamily: '" & fontFamily & "'," & vbNewLine
    If fontColor <> "" Then JSchartLegend = JSchartLegend & "fontColor: '" & fontColor & "'," & vbNewLine
    If fontWeight <> "" Then JSchartLegend = JSchartLegend & "fontWeight: '" & fontWeight & "'," & vbNewLine
    If FontStyle <> "" Then JSchartLegend = JSchartLegend & "FontStyle: '" & FontStyle & "'," & vbNewLine
    If verticalAlign <> "" Then JSchartLegend = JSchartLegend & "verticalAlign: '" & verticalAlign & "'," & vbNewLine
    If horizontalAlign <> "" Then JSchartLegend = JSchartLegend & "horizontalAlign: '" & horizontalAlign & "'," & vbNewLine
    If itemclick = True Then JSchartLegend = JSchartLegend & "itemclick: toggleDataSeries," & vbNewLine
    If reversed <> "" Then JSchartLegend = JSchartLegend & "reversed: " & reversed & "," & vbNewLine

End Function

Function JSchartTitle(Optional text As String, _
                        Optional wrap As String, _
                        Optional maxWidth As String, _
                        Optional dockInsidePlotArea As String, _
                        Optional verticalAlign As String, _
                        Optional horizontalAlign As String, _
                        Optional fontSize As String, _
                        Optional fontFamily As String, _
                        Optional fontWeight As String, _
                        Optional fontColor As String, _
                        Optional FontStyle As String, _
                        Optional borderThickness As String, _
                        Optional cornerRadius As String, _
                        Optional borderColor As String, _
                        Optional backgroundColor As String, _
                        Optional margin As String, _
                        Optional padding As String) As String
                        
    JSchartTitle = ""
    If text <> "" Then JSchartTitle = JSchartTitle & "text: '" & text & "'," & vbNewLine
    If maxWidth <> "" Then JSchartTitle = JSchartTitle & "maxWidth: " & maxWidth & "," & vbNewLine
    If dockInsidePlotArea <> "" Then JSchartTitle = JSchartTitle & "dockInsidePlotArea: " & dockInsidePlotArea & "," & vbNewLine
    If verticalAlign <> "" Then JSchartTitle = JSchartTitle & "verticalAlign: '" & verticalAlign & "'," & vbNewLine
    If horizontalAlign <> "" Then JSchartTitle = JSchartTitle & "horizontalAlign: '" & horizontalAlign & "'," & vbNewLine
    If fontSize <> "" Then JSchartTitle = JSchartTitle & "fontSize: " & fontSize & "," & vbNewLine
    If fontFamily <> "" Then JSchartTitle = JSchartTitle & "fontFamily: '" & fontFamily & "'," & vbNewLine
    If fontWeight <> "" Then JSchartTitle = JSchartTitle & "fontWeight: '" & fontWeight & "'," & vbNewLine
    If fontColor <> "" Then JSchartTitle = JSchartTitle & "fontColor: '" & fontColor & "'," & vbNewLine
    If FontStyle <> "" Then JSchartTitle = JSchartTitle & "FontStyle: '" & FontStyle & "'," & vbNewLine
    If borderThickness <> "" Then JSchartTitle = JSchartTitle & "borderThickness: " & borderThickness & "," & vbNewLine
    If cornerRadius <> "" Then JSchartTitle = JSchartTitle & "cornerRadius: " & cornerRadius & "," & vbNewLine
    If borderColor <> "" Then JSchartTitle = JSchartTitle & "borderColor: '" & borderColor & "'," & vbNewLine
    If backgroundColor <> "" Then JSchartTitle = JSchartTitle & "backgroundColor: '" & backgroundColor & "'," & vbNewLine
    If margin <> "" Then JSchartTitle = JSchartTitle & "margin: " & margin & "," & vbNewLine
    If padding <> "" Then JSchartTitle = JSchartTitle & "padding: " & padding & "," & vbNewLine
End Function
    
Function JSchartTooltop(Optional enabled As String, _
                        Optional fontColor As String, _
                        Optional FontStyle As String, _
                        Optional fontSize As String, _
                        Optional fontFamily As String, _
                        Optional fontWeight As String, _
                        Optional borderThickness As String, _
                        Optional cornerRadius As String, _
                        Optional reversed As String, _
                        Optional backgroundColor As String, _
                        Optional sharedValues As String, _
                        Optional content As String, _
                        Optional animationEnabled As String, _
                        Optional borderColor As String) As String
                        
    JSchartTooltop = ""
    If text <> "" Then JSchartTooltop = JSchartTooltop & "enabled: " & enabled & "," & vbNewLine
    If fontColor <> "" Then JSchartTooltop = JSchartTooltop & "fontColor: '" & fontColor & "'," & vbNewLine
    If FontStyle <> "" Then JSchartTooltop = JSchartTooltop & "FontStyle: '" & FontStyle & "'," & vbNewLine
    If fontFamily <> "" Then JSchartTooltop = JSchartTooltop & "fontFamily: '" & fontFamily & "'," & vbNewLine
    If fontWeight <> "" Then JSchartTooltop = JSchartTooltop & "fontWeight: " & fontWeight & "," & vbNewLine
    If borderThickness <> "" Then JSchartTooltop = JSchartTooltop & "borderThickness: " & borderThickness & "," & vbNewLine
    If cornerRadius <> "" Then JSchartTooltop = JSchartTooltop & "cornerRadius: " & cornerRadius & "," & vbNewLine
    If reversed <> "" Then JSchartTooltop = JSchartTooltop & "reversed: " & reversed & "," & vbNewLine
    If backgroundColor <> "" Then JSchartTooltop = JSchartTooltop & "backgroundColor: '" & backgroundColor & "'," & vbNewLine
    If sharedValues <> "" Then JSchartTooltop = JSchartTooltop & "sharedValues: " & sharedValues & "," & vbNewLine
    If content <> "" Then JSchartTooltop = JSchartTooltop & "content: '" & content & "'," & vbNewLine
    If animationEnabled <> "" Then JSchartTooltop = JSchartTooltop & "animationEnabled: " & animationEnabled & "," & vbNewLine
    If borderColor <> "" Then JSchartTooltop = JSchartTooltop & "borderColor: '" & borderColor & "'," & vbNewLine
End Function

'######################   Yet to do Demo   ######################
Function JSchartAxisX(Optional axisTitle As String, Optional titleWrap As String, _
                        Optional titleMaxWidth As String, Optional titleFontColor As String, _
                        Optional titleFontSize As String, Optional titleFontFamily As String, _
                        Optional titleFontWeight As String, Optional titleFontStyle As String, _
                        Optional margin As String, Optional labelBackgroundColor As String, _
                        Optional labelMaxWidth As String, Optional labelWrap As String, _
                        Optional labelAutoFit As String, Optional labelAngle As String, _
                        Optional labelFontFamily As String, Optional labelFontColor As String, _
                        Optional labelFontSize As String, Optional labelFontWeight As String, _
                        Optional labelFormatter As String, Optional labelFontStyle As String, _
                        Optional Prefix As String, Optional suffix As String, _
                        Optional valueFormatString As String, Optional minimum As String, _
                        Optional maximum As String, Optional viewportMaximum As String, _
                        Optional viewportMinimum, Optional interval As String, _
                        Optional intervalType As String, Optional reversed As String, _
                        Optional logarithmic As String, Optional logarithmBase As String, _
                        Optional tickLength As String, Optional tickColor As String, _
                        Optional tickThickness As String, Optional lineColor As String, _
                        Optional lineThickness As String, Optional lineDashType As String, _
                        Optional gridDashType As String, Optional interlacedColor As String, _
                        Optional gridThickness As String, Optional gridColor As String, _
                        Optional stripLines As String, Optional crosshair As String, _
                        Optional scaleBreaks As String) As String
                        
    'Main Axis X Options
    JSchartAxisX = ""
    If axisTitle <> "" Then JSchartAxisX = JSchartAxisX & "title: '" & axisTitle & "'," & vbNewLine
    If titleWrap <> "" Then JSchartAxisX = JSchartAxisX & "titleWrap: " & titleWrap & "," & vbNewLine
    If titleFontSize <> "" Then JSchartAxisX = JSchartAxisX & "titleFontSize: " & titleFontSize & "," & vbNewLine
    If titleFontFamily <> "" Then JSchartAxisX = JSchartAxisX & "titleFontFamily: '" & titleFontFamily & "'," & vbNewLine
    If titleFontWeight <> "" Then JSchartAxisX = JSchartAxisX & "titleFontWeight: '" & titleFontWeight & "'," & vbNewLine
    If titleFontStyle <> "" Then JSchartAxisX = JSchartAxisX & "titleFontStyle: '" & titleFontStyle & "'," & vbNewLine
    If margin <> "" Then JSchartAxisX = JSchartAxisX & "margin: " & margin & "," & vbNewLine
    If labelBackgroundColor <> "" Then JSchartAxisX = JSchartAxisX & "labelBackgroundColor: '" & labelBackgroundColor & "'," & vbNewLine
    If labelMaxWidth <> "" Then JSchartAxisX = JSchartAxisX & "labelMaxWidth: " & labelMaxWidth & "," & vbNewLine
    If labelWrap <> "" Then JSchartAxisX = JSchartAxisX & "labelWrap: " & labelWrap & "," & vbNewLine
    If labelAutoFit <> "" Then JSchartAxisX = JSchartAxisX & "labelAutoFit: " & labelAutoFit & "," & vbNewLine
    If labelAngle <> "" Then JSchartAxisX = JSchartAxisX & "labelAngle: " & labelAngle & "," & vbNewLine
    If labelFontFamily <> "" Then JSchartAxisX = JSchartAxisX & "labelFontFamily: '" & labelFontFamily & "'," & vbNewLine
    If labelFontColor <> "" Then JSchartAxisX = JSchartAxisX & "labelFontColor: '" & labelFontColor & "'," & vbNewLine
    If labelFontSize <> "" Then JSchartAxisX = JSchartAxisX & "labelFontSize: " & labelFontSize & "," & vbNewLine
    If labelFontWeight <> "" Then JSchartAxisX = JSchartAxisX & "labelFontWeight: '" & labelFontWeight & "'," & vbNewLine
    If labelFormatter <> "" Then JSchartAxisX = JSchartAxisX & "labelFormatter: " & labelFormatter & "," & vbNewLine
    If labelFontStyle <> "" Then JSchartAxisX = JSchartAxisX & "labelFontStyle: '" & labelFontStyle & "'," & vbNewLine
    If Prefix <> "" Then JSchartAxisX = JSchartAxisX & "Prefix: '" & Prefix & "'," & vbNewLine
    If suffix <> "" Then JSchartAxisX = JSchartAxisX & "suffix: '" & suffix & "'," & vbNewLine
    If valueFormatString <> "" Then JSchartAxisX = JSchartAxisX & "valueFormatString: '" & valueFormatString & "'," & vbNewLine
    If minimum <> "" Then JSchartAxisX = JSchartAxisX & "minimum: " & minimum & "," & vbNewLine
    If maximum <> "" Then JSchartAxisX = JSchartAxisX & "maximum: " & maximum & "," & vbNewLine
    If viewportMinimum <> "" Then JSchartAxisX = JSchartAxisX & "viewportMinimum: " & viewportMinimum & "," & vbNewLine
    If viewportMaximum <> "" Then JSchartAxisX = JSchartAxisX & "viewportMaximum: " & viewportMaximum & "," & vbNewLine
    If interval <> "" Then JSchartAxisX = JSchartAxisX & "interval: " & interval & "," & vbNewLine
    If intervalType <> "" Then JSchartAxisX = JSchartAxisX & "intervalType: '" & intervalType & "'," & vbNewLine
    If reversed <> "" Then JSchartAxisX = JSchartAxisX & "reversed: " & reversed & "," & vbNewLine
    If logarithmic <> "" Then JSchartAxisX = JSchartAxisX & "logarithmic: " & logarithmic & "," & vbNewLine
    If logarithmBase <> "" Then JSchartAxisX = JSchartAxisX & "logarithmBase: " & logarithmBase & "," & vbNewLine
    If tickLength <> "" Then JSchartAxisX = JSchartAxisX & "tickLength: " & tickLength & "," & vbNewLine
    If tickColor <> "" Then JSchartAxisX = JSchartAxisX & "tickColor: '" & tickColor & "'," & vbNewLine
    If tickThickness <> "" Then JSchartAxisX = JSchartAxisX & "tickThickness: " & tickThickness & "," & vbNewLine
    If lineColor <> "" Then JSchartAxisX = JSchartAxisX & "lineColor: '" & lineColor & "'," & vbNewLine
    If lineThickness <> "" Then JSchartAxisX = JSchartAxisX & "lineThickness: " & lineThickness & "," & vbNewLine
    If lineDashType <> "" Then JSchartAxisX = JSchartAxisX & "lineDashType: '" & lineDashType & "'," & vbNewLine
    If gridDashType <> "" Then JSchartAxisX = JSchartAxisX & "gridDashType: '" & gridDashType & "'," & vbNewLine
    If interlacedColor <> "" Then JSchartAxisX = JSchartAxisX & "interlacedColor: '" & interlacedColor & "'," & vbNewLine
    If gridThickness <> "" Then JSchartAxisX = JSchartAxisX & "gridThickness: " & gridThickness & "," & vbNewLine
    If gridColor <> "" Then JSchartAxisX = JSchartAxisX & "gridColor: '" & gridColor & "'," & vbNewLine
    
    'Has Sub functions
    If stripLines <> "" Then JSchartAxisX = JSchartAxisX & "stripLines: [" & stripLines & "]," & vbNewLine
    If crosshair <> "" Then JSchartAxisX = JSchartAxisX & "crosshair: {" & crosshair & "}," & vbNewLine
    If scaleBreaks <> "" Then JSchartAxisX = JSchartAxisX & "scaleBreaks: {" & scaleBreaks & "}," & vbNewLine
    

End Function
    'axisX   Properties variable prefix xisX & xisXOption    
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


