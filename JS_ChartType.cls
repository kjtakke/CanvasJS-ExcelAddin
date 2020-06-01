

Function ChartTypesFunction(enumInput) As String
    Select Case True
        Case enumInput = "candlestickChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "candlestick" & "', " & vbNewLine
        
        Case enumInput = "stackedBar100Chart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "stackedBar100" & "', " & vbNewLine
        
        Case enumInput = "stackedBarChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "stackedBar" & "', " & vbNewLine
        
        Case enumInput = "pieChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "type: " & "pie" & "', " & vbNewLine
        
        Case enumInput = "lineChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "line" & "', " & vbNewLine
        
        Case enumInput = "areaChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "area" & "', " & vbNewLine
        
        Case enumInput = "stackedAreaChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "stackedArea" & "', " & vbNewLine
        
        Case enumInput = "rangeAreaChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "rangeArea" & "', " & vbNewLine
        
        Case enumInput = "barChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "bar" & "', " & vbNewLine
        
        Case enumInput = "ColumnChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "column" & "', " & vbNewLine
        
        Case enumInput = "ScatterPlotChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "scatter" & "', " & vbNewLine
        
        Case enumInput = "BoxChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "boxAndWhisker" & "', " & vbNewLine
        
        Case enumInput = "funnelChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "funnel" & "', " & vbNewLine
        
        Case enumInput = "DoughnutChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "Doughnut Chart" & "', " & vbNewLine
        
        Case enumInput = "pyramidChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "pyramid" & "', " & vbNewLine
        
        Case enumInput = "StackedColumn100Chart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "stackedColumn100" & "', " & vbNewLine
        
        Case enumInput = "rangeBarChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "rangeBar" & "', " & vbNewLine
        
        Case enumInput = "rangeColumnChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "rangeColumn" & "', " & vbNewLine
        
        Case enumInput = "stackedColumnChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "stackedColumn" & "', " & vbNewLine
        
        Case enumInput = "waterfallChart"
        ChartTypesFunction = "{" & vbNewLine & "type: '" & "waterfall" & "', " & vbNewLine
    
    End Select
End Function
