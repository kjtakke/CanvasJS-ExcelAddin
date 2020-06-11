Sub Testing()
Dim WebDev As WebDev
Set WebDev = New WebDev
Dim WebPage As String

Dim Styles As String
Dim Scripts As String

Styles = "<style>body {background-color: lightblue;}h1 {color: white;text-align: center;}p {font-family: verdana;font-size: 20px;}</style>"
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
                                Text:="This is my Heading", _
                                CSS_Class:="alert alert-success", _
                                JS_ID:="Txt1", _
                                Style:="color:red;font-size:12px;" _
                            )


Debug.Print (WebPage)
End Sub

Sub Testing2()
Debug.Print (HTML_Text(TextType:="H1", Text:="This is my Heading", CSS_Class:="alert alert-success", JS_ID:="Txt1", Style:="color:red; font-size:12px;"))
End Sub

Function HTML_Text( _
                    TextType As String, _
                    Text As String, _
                    Optional CSS_Class As String, _
                    Optional JS_ID As String, _
                    Optional Style As String) As String
    HTML_Text = "<" & TextType
    If CSS_Class <> "" Then HTML_Text = HTML_Text & " class='" & CSS_Class & "'"
    If JS_ID <> "" Then HTML_Text = HTML_Text & " ID='" & JS_ID & "'"
    If Style <> "" Then HTML_Text = HTML_Text & " Style='" & Style & "'"
    HTML_Text = HTML_Text & ">" & Text & "</" & TextType & ">" & vbNewLine
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
