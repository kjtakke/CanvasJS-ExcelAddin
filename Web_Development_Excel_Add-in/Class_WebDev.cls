Public Bootstrap4CSS As String
Public Bootstrap4JS As String
Public popperJS As String
Public jQueryJS As String
Public BR1 As String
Public BR2 As String
Public BR3 As String
Public BR4 As String
Public BR5 As String
Public HTML_DIV_Close
Public HTML_Close As String

Public BStablePlane As String
Public BStableStripped As String
Public BStableBorered As String
Public BStableHover As String
Public BStableDark As String
Public BStableDarkStripped As String
Public BStableDarkHover As String
Public BStableborderless As String
Public BStabletheadDark As String
Public BStabletheadLight As String
Public BStableSmall As String
Public BStableColorPrimary As String
Public BStableColorSuccess As String
Public BStableColorDanger As String
Public BStableColorInfo As String
Public BStableColorWarning As String
Public BStableColorActive As String
Public BStableColorSecondary As String
Public BStableColorLight As String
Public BStableColorDark As String

Public BScontainer As String
Public BScontainerSm As String
Public BScontainerMd As String
Public BScontainerLg As String
Public BScontainerXl As String
Public BScontainerFluid As String
Public BScontainerFluidSm As String
Public BScontainerFluidMd As String
Public BScontainerFluidLg As String
Public BScontainerFluidXl As String
Public BScontainerPrimary As String
Public BScontainerDark As String
Public BScontainerDanger As String
Public BScontainerInfo As String
Public BScontainerWarning As String
Public BScontainerSecondary As String

Public BStextWhite As String
Public BStextPrimary As String
Public BStextDark As String
Public BStextDanger As String
Public BStextInfo As String
Public BStextWarning As String
Public BStextSecondary As String

Public BSfontBold As String
Public BSfontBolder As String
Public BSfontItalic As String
Public BSfontLight As String
Public BSfontLighter As String
Public BSfontNormal As String
Public BSfontLead As String
Public BSfontSmall As String
Public BSfontLeft As String
Public BSfontLeftSM As String
Public BSfontLeftMD As String
Public BSfontLeftLG As String
Public BSfontLeftXL As String
Public BSfontBreak As String
Public BSfontCenter As String
Public BSfontCenterSM As String
Public BSfontCenterMD As String
Public BSfontCenterLG As String
Public BSfontCenterXL As String
Public BSfontPlane As String
Public BSfontRight As String
Public BSfontRightSM As String
Public BSfontRightMD As String
Public BSfontRightLG As String
Public BSfontRightXL As String
Public BSfontJustify As String
Public BSfontMonospace As String
Public BSfontNowrap As String
Public BSfontLowercase As String
Public BSfontReset As String
Public BSfontUppercase As String
Public BSfontCapitalize As String
Public BSfontInitialism As String
Public BSfontListUnstyled As String
Public BSfontListInline As String
Public BSfontPreScrollable As String

Public JStrue As String
Public JSfalse As String

Private Sub Class_Initialize()
'CSS
Bootstrap4CSS = "https://maxcdn.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css"
    'Text
    BSfontBold = " weight-bold"
    BSfontBolder = " font-weight-bolder"
    BSfontItalic = " font-italic"
    BSfontLight = " font-weight-light"
    BSfontLighter = " font-weight-lighter"
    BSfontNormal = " font-weight-normal"
    BSfontLead = " lead"
    BSfontSmall = " small"
    BSfontLeft = " text-left"
    BSfontLeftSM = " text-sm-left"
    BSfontLeftMD = " text-md-left"
    BSfontLeftLG = " text-lg-left"
    BSfontLeftXL = " text-xl-left"
    BSfontBreak = " text-break"
    BSfontCenter = " text-center"
    BSfontCenterSM = " text-sm-center"
    BSfontCenterMD = " text-md-center"
    BSfontCenterLG = " text-lg-center"
    BSfontCenterXL = " text-xl-center"
    BSfontPlane = " text-decoration-none"
    BSfontRight = " text-right"
    BSfontRightSM = " text-sm-right"
    BSfontRightMD = " text-md-right"
    BSfontRightLG = " text-lg-right"
    BSfontRightXL = " text-xl-right"
    BSfontJustify = " text-justify"
    BSfontMonospace = " text-monospace"
    BSfontNowrap = " text-nowrap"
    BSfontLowercase = " text-lowercase"
    BSfontReset = " text-reset"
    BSfontUppercase = " text-uppercase"
    BSfontCapitalize = " text-capitalize"
    BSfontInitialism = " initialism"
    BSfontListUnstyled = " list-unstyled"
    BSfontListInline = " list-inline"
    BSfontPreScrollable = " pre-scrollable"
    
    'Container
    BScontainer = " container"
    BScontainerSm = " container-sm"
    BScontainerMd = " container-md"
    BScontainerLg = " container-lg"
    BScontainerXl = " container-xl"
    
    BScontainerFluid = " container-fluid"
    BScontainerFluidSm = " container-fluid-sm"
    BScontainerFluidMd = " container-fluid-md"
    BScontainerFluidLg = " container-fluid-lg"
    BScontainerFluidXl = " container-fluid-xl"
    
    BScontainerPrimary = " bg-primary"
    BScontainerDark = " bg-dark"
    BScontainerDanger = " bg-danger"
    BScontainerInfo = " bg-info"
    BScontainerWarning = " bg-warning"
    BScontainerSecondary = " bg-secondary"
    
    'Text
    BStextWhite = " text-white"
    BStextPrimary = " text-primary"
    BStextDark = " text-dark"
    BStextDanger = " text-danger"
    BStextInfo = " text-info"
    BStextWarning = " text-warning"
    BStextSecondary = " text-secondary"


    'Bootstrap Table
    BStablePlane = " table"
    BStableStripped = " table table-striped"
    BStableBorered = " table table-bord"
    BStableHover = " table table-hover"
    BStableDark = " table table-dark"
    BStableDarkStripped = " table table-dark table-striped"
    BStableDarkHover = " table table-dark table-hover"
    BStableborderless = " table table-borderless"
    BStabletheadDark = " table thead-dark"
    BStabletheadLight = " table thead-light"
    BStableSmall = " table table-sm"
    BStableColorPrimary = " table table-primary"
    BStableColorSuccess = " table table-success"
    BStableColorDanger = " table table-danger"
    BStableColorInfo = " table table-info"
    BStableColorWarning = " table table-warning"
    BStableColorActive = " table table-active"
    BStableColorSecondary = " table table-secondary"
    BStableColorLight = " table table-light"
    BStableColorDark = " table table-dark"
    
    
'JS
Bootstrap4JS = "https://maxcdn.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js"
popperJS = "https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"
jQueryJS = "https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"

    'Variables
    JStrue = "true"
    JSfalse = "false"
    
'Formatting
BR1 = "<BR>"
BR2 = "<BR><BR>"
BR3 = "<BR><BR><BR>"
BR4 = "<BR><BR><BR><BR>"
BR5 = "<BR><BR><BR><BR><BR>"

HTML_DIV_Close = "</div>" & vbNewLine
HTML_Close = "</body>" & vbNewLine & "</html>" & vbNewLine
End Sub
