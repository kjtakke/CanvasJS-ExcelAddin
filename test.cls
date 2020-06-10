Sub Dashboard()
Dim HTML As String
Dim web As Worksheet
Dim MainDivStart As String
Dim MainDivEnd As String
Dim SubDivStart As String
Dim SubDivEnd As String
Dim JSToggleData As String
Dim JSChartRender As String
Dim Inputs As Variant

Set web = Sheets("Web")

Inputs = Worksheets("Array").Range("A2:K2", Worksheets("Array").Range("A1:K1").End(xlDown)).Value

MainDivStart = web.Range("E8").Value
MainDivEnd = web.Range("E9").Value
SubDivStart = web.Range("E11").Value
SubDivEnd = web.Range("E12").Value

JSToggleData = web.Range("E15").Value


HTML = web.Range("E2").Value & vbNewLine

    'Line Chart
    HTML = HTML & CreateDatapoints(container:="1", name:="Tasks Over Time", Title:="Tasks Over Time", AxisTitle:="Tasks", chart:="line", yTxt:="A4", xTxt:="B4")
    HTML = HTML & CreateDatapoints(container:="2", name:="Priotity", Title:="Tasks by Priority", AxisTitle:="Tasks", chart:="pie", yTxt:="F4", xTxt:="G4")
    HTML = HTML & CreateDatapoints(container:="3", name:="Lead", Title:="Tasks by Lead", AxisTitle:="Tasks", chart:="bar", yTxt:="K4", xTxt:="L4")

HTML = HTML & web.Range("E3").Value & vbNewLine

    'Line Chart
    HTML = HTML & "<div style='grid-template-columns: repeat(auto-fit, minmax(100%, 100%)); ' class='wrapper'>" & vbNewLine
        HTML = HTML & "<div id='chartContainer" & "1" & "' class='box a' style='height: 370px;'></div>" & vbNewLine
    HTML = HTML & "</div>" & vbNewLine
        
    HTML = HTML & "<div style='grid-template-columns: repeat(auto-fit, minmax(50%, 50%)); ' class='wrapper'>" & vbNewLine
        HTML = HTML & "<div id='chartContainer" & "2" & "' class='box a' style='height: 370px;'></div>" & vbNewLine
        HTML = HTML & "<div id='chartContainer" & "3" & "' class='box a' style='height: 370px;'></div>" & vbNewLine
    HTML = HTML & "</div>" & vbNewLine


HTML = HTML & web.Range("E4").Value
'Debug.Print (HTML)
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("C:\Users\kristopher.takken\Desktop\WebPage.html", True)
a.WriteLine HTML
a.Close




End Sub

Sub testDataPoints()

Debug.Print (CreateDatapoints(container:="2", name:="Tasks Over Time", Title:="Tasks Over Time", AxisTitle:="Tasks", chart:="line", yTxt:="A4", xTxt:="B4"))

End Sub

Function CreateDatapoints( _
container As String, _
chart As String, _
name As String, _
yTxt As String, _
xTxt As String, _
Title As String, _
AxisTitle As String _
) As String

Dim web As Worksheet
Set web = Sheets("Web")

Dim JSChartVar As String
Dim JSChartOptions1 As String
Dim JSChartOptionsName As String
Dim JSChartOptionsAxisName As String
Dim JSChartRender As String

Dim y As Variant
Dim x1 As Variant
Dim x2 As Variant
Dim x3 As Variant
Dim x4 As Variant
Dim x5 As Variant
Dim ws As Worksheet
Set ws = Sheets("ChartData")

JSChartVar = web.Range("E17").Value
JSChartOptions1 = web.Range("E18").Value
JSChartOptionsName = web.Range("E19").Value
JSChartOptionsAxisName = web.Range("E20").Value
JSChartRender = web.Range("E16").Value

CreateDatapoints = JSChartVar & container & "'"
CreateDatapoints = CreateDatapoints & JSChartOptions1 & Title & "'"
CreateDatapoints = CreateDatapoints & JSChartOptionsName & AxisTitle & "'"
CreateDatapoints = CreateDatapoints & JSChartOptionsAxisName

y = Worksheets("ChartData").Range(yTxt, Worksheets("ChartData").Range(yTxt).End(xlDown)).Value
x = Worksheets("ChartData").Range(xTxt, Worksheets("ChartData").Range(xTxt).End(xlDown)).Value


CreateDatapoints = CreateDatapoints & " type:'" & chart & "', name:'" & name & "', showInLegend: false, dataPoints: ["

      
For i = 1 To UBound(y)
    CreateDatapoints = CreateDatapoints & "{ y:" & " " & x(i, 1) & ", label: '" & y(i, 1) & "' }," & vbNewLine
Next i



CreateDatapoints = CreateDatapoints & "]},]}); chart.render();" & vbNewLine
End Function



