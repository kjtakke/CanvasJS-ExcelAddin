Sub Testing()
Dim WebDev As WebDev
Set WebDev = New WebDev
Dim WebPage As String

Dim Styles As String
Dim Scripts As String

Styles = "<style>body {background-color: black;}h1 {color: white;text-align: center;}p {font-family: verdana;font-size: 20px;}</style>"
Scripts = "<script>function toCelsius(fahrenheit) {return (5 / 9) * (fahrenheit - 32);}</script>"


WebPage = HTML_INIT( _
                        TITLE:="My Site", _
                        LINK:=StyleSheetLink(WebDev.Bootstrap4CSS), _
                        SCRIPT:=ScriptLinks(WebDev.Bootstrap4JS, WebDev.jQueryJS, WebDev.popperJS), _
                        MyStyles:=Styles, _
                        MyScripts:=Scripts _
                    )

WebPage = WebPage & HTML_Text( _
                                TextType:="H1", _
                                text:="This is my Heading", _
                                CSS_Class:="alert alert-success", _
                                JS_id:="Txt1", _
                                style:="color:red;font-size:12px;" _
                            )

WebPage = WebPage & HTML_Image( _
                                src:="https://objective:id:BQ13776676", _
                                style:="border: black solid 1px;", _
                                JS_id:="MyImg1", _
                                CSS_Class:="rounded" _
                            )

WebPage = WebPage & HTML_Iframe( _
                                src:="https://www.w3schools.com/html/html_iframe.asp", _
                                style:="border: black solid 1px;", _
                                JS_id:="MyFrame1", _
                                CSS_Class:="rounded" _
                            )

WebPage = WebPage & HTML_DIV_Open( _
                                style:="border: black solid 1px;", _
                                JS_id:="MyDiv1", _
                                CSS_Class:="container" _
                            )

WebPage = WebPage & HTML_Button( _
                                style:="border:black solid 1px;", _
                                JS_id:="MyBtn1", _
                                CSS_Class:=WebDev.BSfontLighter & WebDev.BStextInfo, _
                                text:="My Button", _
                                onclick:="window.location.href='https://w3docs.com';" _
                            )

WebPage = WebPage & WebDev.HTML_DIV_Close

WebPage = WebPage & HTML_Table( _
                                        hasHeader:=True, _
                                        sheetName:="Sheet1", _
                                        cellRange:="A1:C5", _
                                        style:="border:black solid 1px;", _
                                        JS_id:="Mytbl1", _
                                        CSS_Class:=WebDev.BStableDarkHover _
                                    )
                                    
WebPage = WebPage & WebDev.HTML_Close

Debug.Print (WebPage)
End Sub

Sub Testing2()
Debug.Print (HTML_Table(hasHeader:=True, sheetName:="Sheet1", cellRange:="A1:C5", style:="border:black solid 1px;", JS_id:="Mytbl1", CSS_Class:="table table-striped"))
End Sub



Function HTML_Table(hasHeader As Boolean, sheetName As String, cellRange As String, Optional style As String, Optional JS_id As String, Optional CSS_Class As String) As String
    'Get Data
    Dim tblRange As Variant
    tblRange = Worksheets(sheetName).Range(cellRange).Value
    Dim TempVar As String
    'Column Count
    On Error GoTo tblCnt:
    For j = 1 To 500
        TempVar = tblRange(1, j)
    Next j
tblCnt:
    On Error GoTo 0
    
    'Set Table Tags
    HTML_Table = "<table "
    If style <> "" Then HTML_Table = HTML_Table & "style='" & style & "' "
    If JS_id <> "" Then HTML_Table = HTML_Table & "id='" & JS_id & "' "
    If CSS_Class <> "" Then HTML_Table = HTML_Table & "class='" & CSS_Class & "' "
    HTML_Table = HTML_Table & ">" & vbNewLine
    
    'Table Headers
    If hasHeader = True Then
        HTML_Table = HTML_Table & "<thead>" & vbNewLine & "<tr>" & vbNewLine
        For i = 1 To j - 1
            HTML_Table = HTML_Table & "<th>" & tblRange(1, i) & "</th>" & vbNewLine
        Next i
        HTML_Table = HTML_Table & "</tr>" & vbNewLine & "</thead>" & vbNewLine
        
        
        HTML_Table = HTML_Table & "<tbody>" & vbNewLine
        For k = 2 To UBound(tblRange)
            HTML_Table = HTML_Table & "<tr>" & vbNewLine
            For i = 1 To j - 1
                HTML_Table = HTML_Table & "<td>" & tblRange(k, i) & "</td>" & vbNewLine
            Next i
            HTML_Table = HTML_Table & "</tr>" & vbNewLine
        Next k
        HTML_Table = HTML_Table & "</tbody>" & vbNewLine & "</table>" & vbNewLine
    Else
        HTML_Table = HTML_Table & "<tbody>" & vbNewLine
           For k = 1 To UBound(tblRange)
               HTML_Table = HTML_Table & "<tr>" & vbNewLine
               For i = 1 To j - 1
                   HTML_Table = HTML_Table & "<td>" & tblRange(k, i) & "</td>" & vbNewLine
               Next i
               HTML_Table = HTML_Table & "</tr>" & vbNewLine
           Next k
           HTML_Table = HTML_Table & "</tbody>" & vbNewLine & "</table>" & vbNewLine
    End If

End Function

Function HTML_Button(Optional text As String, Optional onclick As String, Optional style As String, Optional JS_id As String, Optional CSS_Class As String) As String
'<button type="button">Click Me!</button>
    HTML_Button = "<button "
    If onclick <> "" Then HTML_Button = HTML_Button & "onclick=""" & onclick & """ ' "
    If style <> "" Then HTML_Button = HTML_Button & "style='" & style & "' "
    If JS_id <> "" Then HTML_Button = HTML_Button & "id='" & JS_id & "' "
    If CSS_Class <> "" Then HTML_Button = HTML_Button & "class='" & CSS_Class & "' "
    If text <> "" Then
        HTML_Button = HTML_Button & ">" & text
    Else
        HTML_Button = HTML_Button & ">"
    End If
    HTML_Button = HTML_Button & "</button>" & vbNewLine
End Function


Function HTML_DIV_Open(Optional style As String, Optional JS_id As String, Optional CSS_Class As String) As String
    HTML_DIV_Open = "<div "
    If style <> "" Then HTML_DIV_Open = HTML_DIV_Open & "style='" & style & "' "
    If JS_id <> "" Then HTML_DIV_Open = HTML_DIV_Open & "id='" & JS_id & "' "
    If CSS_Class <> "" Then HTML_DIV_Open = HTML_DIV_Open & "class='" & CSS_Class & "' "
    HTML_DIV_Open = HTML_DIV_Open & ">" & vbNewLine
End Function


Function HTML_Iframe(src As String, Optional style As String, Optional JS_id As String, Optional CSS_Class As String) As String
    HTML_Iframe = "<iframe src='" & src & "' "
    If style <> "" Then HTML_Iframe = HTML_Iframe & "style='" & style & "' "
    If JS_id <> "" Then HTML_Iframe = HTML_Iframe & "id='" & JS_id & "' "
    If CSS_Class <> "" Then HTML_Iframe = HTML_Iframe & "class='" & CSS_Class & "' "
    HTML_Iframe = HTML_Iframe & "></iframe>" & vbNewLine
End Function

Function HTML_Image(src As String, Optional style As String, Optional JS_id As String, Optional CSS_Class As String) As String
    HTML_Image = "<img src='" & src & "' "
    If style <> "" Then HTML_Image = HTML_Image & "style='" & style & "' "
    If JS_id <> "" Then HTML_Image = HTML_Image & "id='" & JS_id & "' "
    If CSS_Class <> "" Then HTML_Image = HTML_Image & "class='" & CSS_Class & "' "
    HTML_Image = HTML_Image & ">" & vbNewLine
End Function

Function HTML_Text( _
                    TextType As String, _
                    text As String, _
                    Optional CSS_Class As String, _
                    Optional JS_id As String, _
                    Optional style As String) As String
    HTML_Text = "<" & TextType
    If CSS_Class <> "" Then HTML_Text = HTML_Text & " class='" & CSS_Class & "'"
    If JS_id <> "" Then HTML_Text = HTML_Text & " ID='" & JS_id & "'"
    If style <> "" Then HTML_Text = HTML_Text & " Style='" & style & "'"
    HTML_Text = HTML_Text & ">" & text & "</" & TextType & ">" & vbNewLine
End Function

Function ScriptLinks( _
                         HREF1 As String, _
                Optional HREF2 As String, _
                Optional HREF3 As String, _
                Optional HREF4 As String, _
                Optional HREF5 As String, _
                Optional HREF6 As String, _
                Optional HREF7 As String, _
                Optional HREF8 As String, _
                Optional HREF9 As String, _
                Optional HREF10 As String, _
                Optional HREF11 As String, _
                Optional HREF12 As String, _
                Optional HREF13 As String, _
                Optional HREF14 As String, _
                Optional HREF15 As String _
                ) As String
                
    ScriptLinks = "<script src='" & HREF1 & "'></script>" & vbNewLine
    If HREF2 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF2 & "'></script>" & vbNewLine
    If HREF3 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF3 & "'></script>" & vbNewLine
    If HREF4 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF4 & "'></script>" & vbNewLine
    If HREF5 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF5 & "'></script>" & vbNewLine
    If HREF6 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF6 & "'></script>" & vbNewLine
    If HREF7 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF7 & "'></script>" & vbNewLine
    If HREF8 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF8 & "'></script>" & vbNewLine
    If HREF9 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF9 & "'></script>" & vbNewLine
    If HREF10 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF10 & "'></script>" & vbNewLine
    If HREF11 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF11 & "'></script>" & vbNewLine
    If HREF12 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF12 & "'></script>" & vbNewLine
    If HREF13 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF13 & "'></script>" & vbNewLine
    If HREF14 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF14 & "'></script>" & vbNewLine
    If HREF15 <> "" Then ScriptLinks = ScriptLinks & "<script src='" & HREF15 & "'></script>" & vbNewLine
End Function
           
Function StyleSheetLink( _
                         HREF1 As String, _
                Optional HREF2 As String, _
                Optional HREF3 As String, _
                Optional HREF4 As String, _
                Optional HREF5 As String, _
                Optional HREF6 As String, _
                Optional HREF7 As String, _
                Optional HREF8 As String, _
                Optional HREF9 As String, _
                Optional HREF10 As String, _
                Optional HREF11 As String, _
                Optional HREF12 As String, _
                Optional HREF13 As String, _
                Optional HREF14 As String, _
                Optional HREF15 As String _
                ) As String
                
    StyleSheetLink = "<link rel='stylesheet' href='" & HREF1 & "' > " & vbNewLine
    If HREF2 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF2 & " > " & vbNewLine
    If HREF3 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF3 & " > " & vbNewLine
    If HREF4 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF4 & " > " & vbNewLine
    If HREF5 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF5 & " > " & vbNewLine
    If HREF6 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF6 & " > " & vbNewLine
    If HREF7 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF7 & " > " & vbNewLine
    If HREF8 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF8 & " > " & vbNewLine
    If HREF9 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF9 & " > " & vbNewLine
    If HREF10 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF10 & " > " & vbNewLine
    If HREF11 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF11 & " > " & vbNewLine
    If HREF12 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF12 & " > " & vbNewLine
    If HREF13 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF13 & " > " & vbNewLine
    If HREF14 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF14 & " > " & vbNewLine
    If HREF15 <> "" Then StyleSheetLink = StyleSheetLink & "<link rel='stylesheet' href='" & HREF15 & " > " & vbNewLine
End Function

Function HTML_INIT(Optional TITLE As String, Optional LINK As String, Optional MyStyles As String, Optional SCRIPT As String, Optional MyScripts As String) As String
    HTML_INIT = "<HTML>" & vbNewLine & "<HEAD>" & vbNewLine
    If TITLE <> "" Then HTML_INIT = HTML_INIT & "<TILTE>" & TITLE & "</TITLE>" & vbNewLine
    If LINK <> "" Then HTML_INIT = HTML_INIT & LINK & vbNewLine
    If MyStyles <> "" Then HTML_INIT = HTML_INIT & MyStyles & vbNewLine
    If SCRIPT <> "" Then HTML_INIT = HTML_INIT & SCRIPT & vbNewLine
    If MyScripts <> "" Then HTML_INIT = HTML_INIT & MyScripts & vbNewLine
    HTML_INIT = HTML_INIT & "</HEAD>" & vbNewLine & "<BODY>" & vbNewLine
End Function

