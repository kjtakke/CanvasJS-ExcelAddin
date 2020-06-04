Function CreateDatapoints( _
                                         chartType As String, _
                                Optional labelWB As Workbook, _
                                Optional labelRng As Range, _
                                Optional yWB As Workbook, _
                                Optional yRng As Range, _
                                Optional xWB As Workbook, _
                                Optional xRng As Range, _
                                Optional aggFunct As String, _
                                Optional nameWB As Workbook, _
                                Optional nameAdd As Range _
                          ) As String

'INITILISE
'##################################################################################################################################
Dim label, yValues, xValues, unique_label, nameValues As Variant
Dim chartData As String
On Error Resume Next
xValues = xRng.Value
label = labelRng.Value
yValues = yRng.Value
nameValues = nameAdd.Value
unique_label = UNIQUEVALUES_Chart(labelRng, True)
aggFunct = aggFunct
On Error GoTo 0


'DATASETS
'##################################################################################################################################
Dim labelCount As Long
For i = 1 To UBound(unique_label)
    If unique_label(i, 1) = "" Then Exit For
Next i
labelCount = i - 1

    'label
    Dim dataset As Variant
    ReDim dataset(1 To labelCount, 1 To 4)
    For i = 1 To labelCount
        dataset(i, 1) = unique_label(i, 1)
        dataset(i, 2) = 0
        dataset(i, 3) = 0
        dataset(i, 4) = 0
    Next i
    unique_label = Empty
    
    

 '2 POINT CHARTS
'##################################################################################################################################
        If chartType = "barChart" Or _
        chartType = "columnChart" Or _
        chartType = "lineChart" Or _
        chartType = "areaChart" Or _
        chartType = "pieChart" Or _
        chartType = "waterfallChart" Or _
        chartType = "pyramidChart" Or _
        chartType = "doughnutChart" Or _
        chartType = "funnelChart" Or _
        chartType = "stackedAreaChart" Or _
        chartType = "stackedBarChart" Or _
        chartType = "stackedBar100Chart" Or _
        chartType = "stackedColumn100Chart" Or _
        chartType = "stackedColumnChart" Then
            If aggFunct = "sum" Then
                'Values Sum
                For i = 1 To labelCount
                    For j = i To UBound(label)
                        If label(j, 1) = dataset(i, 1) Then dataset(i, 2) = dataset(i, 2) + yValues(j, 1)
                    Next j
                    
                Next i
            
            ElseIf aggFunct = "count" Then
                'Values Count
                For i = 1 To labelCount
                    For j = i To UBound(label)
                        If label(j, 1) = dataset(i, 1) Then dataset(i, 2) = dataset(i, 2) + 1
                    Next j
                Next i
            End If
    
            'Write Data Lines
            
            For i = 1 To labelCount
                chartData = chartData & ChartLineData(chartType:=chartType, y:=dataset(i, 2), label:=dataset(i, 1))
            Next i
            
'3 POINT CHARTS
'##################################################################################################################################
        ElseIf chartType = "Pie" Then




'4 POINT CHARTS
'##################################################################################################################################
        ElseIf chartType = "Bubble" Then




'4 POINT CHARTS
'##################################################################################################################################
        ElseIf chartType = "Box" Then



    End If

'RESULT
'##################################################################################################################################
CreateDatapoints = chartData
End Function



Function ChartLineData(chartType As String, y As Variant, Optional x As Variant, Optional label As Variant, Optional name As String) As Variant
    Select Case True
    '"barChart", "ColumnChart", "LineChart", "AreaChart", "pieChart", "waterfallChart", "pyramidChart", "DoughnutChart", "funnelChart"
    '"stackedAreaChart", "stackedBarChart", "stackedBar100Chart", "StackedColumn100Chart", "stackedColumnChart",
        'SIMPLE CHART
        'Bar Chart { y: 3, label: "Sweden" },
        
'2 POINT CHARTS
'##################################################################################################################################
        Case chartType = "barChart"
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Column Chart { x: '10', y: 71 },
        Case chartType = "columnChart"
        'ChartLineData = "{ x:'" & x & "', y:" & y & " }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Line Chart { y: 450 },
        Case chartType = "lineChart"
        'ChartLineData = "{ y:" & y & " }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Area Chart { x: 1, y: 74.4, label: "text" },
        Case chartType = "areaChart"
        'ChartLineData = "{ x: '" & x & "', y:" & y & ", label: '" & lable & "' }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        Case chartType = "pieChart"
        'Pie Chart { y: 51.08, label: "Chrome" },
        'ChartLineData = "{ y: " & y & ", label: '" & label & "'}," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Waterfall Chart { label: "Sales",  y: 1273 },
        Case chartType = "waterfallChart"
        'ChartLineData = "{ label:'" & label & "', y: " & y & " }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Funnel Chart { y: 1400, label: "Leads" },
        Case chartType = "funnelChart"
        'ChartLineData = "{ y:" & y & ", label: '" & lable & "' }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Doughnut Chart { y: 67, label: "Inbox" },
        Case chartType = "doughnutChart"
        'ChartLineData = "{ y:" & y & ", label: '" & lable & "' }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Doughnut Chart { y: 67, label: "Inbox" },
        Case chartType = "pyramidChart"
        'ChartLineData = "{ y:" & y & ", label: '" & lable & "' }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine

        'Stacked Area Chart { x: 1, y: 3000 },
        Case chartType = "stackedAreaChart"
        'ChartLineData = "{ x: '" & x & "', y:" & y & " }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Stacked Column 100% Chart { x: 'Date', y: 40 },
        Case chartType = "stackedColumn100Chart"
        'ChartLineData = "{ x:'" & x & "', y:" & y & " }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        'Stacked Column { y: 6.75, x: 'Date' },
        Case chartType = "stackedColumnChart"
        'ChartLineData = "{ y:" & y & ", x: '" & x & "' }," & vbNewLine
        ChartLineData = "{ y:" & y & ", label: '" & label & "' }," & vbNewLine
        
        
    
'3 POINT CHARTS
'##################################################################################################################################
        
        'Range Area Chart { x: 'txt'), y:[15, 21] },
        Case chartType = "rangeAreaChart"
        ChartLineData = "{ x: '" & x & "', y:[" & y & "] }," & vbNewLine
        
        'Scatter Plot Chart { x: 800, y: 350 },
        Case chartType = "ScatterPlotChart"
        ChartLineData = "{ x:" & x & ", y:" & y & " }," & vbNewLine
        
        'Range Bar Chart { x: 10, y:[80, 115], label: "Data Scientist" },
        Case chartType = "rangeBarChart"
        ChartLineData = "{ x: " & x & ", y:[" & y & "], label: '" & label & "' }," & vbNewLine
        
        'Range Column Chart { x: 'Date', y: [27.10, 38.99] },
        Case chartType = "rangeColumnChart"
        ChartLineData = "{ x: '" & x & "', y:[" & y & "] }," & vbNewLine
        
        
'4 POINT CHARTS
'##################################################################################################################################

        'CandlestickChart Chart { x: 'Date', y: [34.080002, 36.060001, 33.410000, 36.060001] },
        Case chartType = "candlestickChart"
        ChartLineData = "{ x: '" & x & "', y:[" & y & "] }," & vbNewLine
      
'5 POINT CHARTS
'##################################################################################################################################
                'Box Chart { x: 'Date',  y: [4, 6, 8, 9, 7] },
        Case chartType = "BoxChart"
        ChartLineData = "{ x: '" & x & "', y:[" & y & "] }," & vbNewLine
        
        
    End Select
End Function


