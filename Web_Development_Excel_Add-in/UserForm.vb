'Decorations##########################################################################################################################################################

'Saved File Location
Public myfile As String

'Webpage Text
Public HTML As String
Public HTMLPageData As String
Public HTMLRowData As String
Public HTMLColumnData As String

'Public Arrays
Public WDdata As Variant            'Website Data Aarray
Public ColumnText As Variant        'Column Text Field
Public JavaScriptText As Variant    'JavaScript Field

'Deliminators
Public MD As String
Public PP As String
Public PD As String
Public RR As String
Public RD As String
Public CC As String
Public CD As String
Public TD As String
Public JD As String

'Array Dimentions
Public AD1 As Integer
Public AD2 As Integer
Public AD3 As Integer
Public AD4 As Integer

'Data Fields
Public DM As Integer
Public DP As Integer
Public DR As Integer
Public DC As Integer
Public XLWD As String


'Validation Variables
    'IsNumeric
    Public WDROWWidthIsNumeric
    Public WDRowHeightIsNumeric
    Public WDRowPaddingTopIsNumeric
    Public WDRowPaddingBottomIsNumeric
    Public WDRowMarginTopIsNumeric
    Public WDRowMarginBottomIsNumeric
    Public WDRowBorderRadiousIsNumeric
    Public WDRowBorderThicknessIsNumeric
    Public WDMartricsTextSize1IsNumeric
    Public WDMartricsTextSize2IsNumeric
    Public WDMartricsTextSize3IsNumeric
    Public WDMartricsTextSize4IsNumeric
    Public WDMartricsTextSize5IsNumeric
    Public WDMartricsTextSize6IsNumeric
    Public WDMartricsTextSize7IsNumeric
    Public WDMartricsTextSize8IsNumeric
    Public WDMartricsTextSize9IsNumeric
    Public WDMartricsTextSize10IsNumeric
    Public WDMartricsWidth1IsNumeric
    Public WDMartricsWidth2IsNumeric
    Public WDMartricsWidth3IsNumeric
    Public WDMartricsWidth4IsNumeric
    Public WDMartricsWidth5IsNumeric
    Public WDMartricsWidth6IsNumeric
    Public WDMartricsWidth7IsNumeric
    Public WDMartricsWidth8IsNumeric
    Public WDMartricsWidth9IsNumeric
    Public WDMartricsWidth10IsNumeric
    Public WDMartricsColumnPaddingLeftIsNumeric
    Public WDMartricsColumnPaddingRightIsNumeric
    Public WDMartricsColumnMarginLeftIsNumeric
    Public WDMartricsColumnMarginRightIsNumeric
    Public WDMartricsColumnHeightIsNumeric
    Public WDMartricsColumnWidthIsNumeric
    Public WDColumnTableTextSizeIsNumeric
    Public WDColumnTablePaddingLeftIsNumeric
    Public WDColumnTablePaddingRightIsNumeric
    Public WDColumnTableMarginLeftIsNumeric
    Public WDColumnTableMarginRightIsNumeric
    Public WDColumnTableWidthIsNumeric
    Public WDColumnTextHeightIsNumeric
    Public WDColumnTextWidthIsNumeric
    Public WDChartTitleTextSizeIsNumeric
    Public WDChartSubtitleTextSizeIsNumeric
    Public WDChartHeightIsNumeric
    Public WDChartWidthIsNumeric
    Public WDChartAnimationLenghtIsNumeric
    Public WDChartTooltipTextSizeIsNumeric
    Public WDChartLegendTextSizeIsNumeric
    Public WDChartTooltipBorderRadiousIsNumeric




'Spare Delimiters https://passwordsgenerator.net/
    '7,}fu?[K3b;Cnc:(=&-?')9dzr!?$C35}B2&dZSQF-QW9/q\~+':yBVA86R3D7E*U!2bX^(CxdB#kM>K7m>v6gj6p&([F=jjgt3bJ6D&7zF)>&L!xMhaQq:2CF$qM$>
    'n&:>}spv2k;jP)XjyAMF5V.W[Ds'K~awWTZRHP3gX7Y5pj3~5K,=)*d6_ENaz-Q:x5A^zm4~DkR`$;!DKP<^v=39}Q96.GJ,H5&)u'g)<<[\br3DAJ!W/-duuk?h;sfU
    'AT)e7E\3tq4.CxY27/#pBkjR8hsQ@+b2f^2v9u4]@V~ew/9GKTwTVA(=A*K:Tf7f\eMy\Z~f-*=}SQzuBq(->k#HzQ#:qj3'C`Tge)Vxa@N4ErQV!M,*zHJ,jUWav@eH
    'S>FT\'hPU#;E*g7[}Jb?D$WnK7XtG{tH[sQ<(Dsh~xK&w@R3!=^pN\2vzU?kXC\6%^R#R~n}n`Qzc~7E-uJ-77pJw9&4v:*,SHB,PL[B!(cx\b8^RS;/S<zA-EeqQcM.
    '9AxZ*)kYTARH+Ttu#HF%SQ!,9{Z4sa-@[F~J[c:P;tXu6L&G5G^LQ`Hj>/ESP[Bc;M26z2V4N<6Zz5/Bbk6pRz'.zCh'rAA:)j@j~j75bK9z%n:kzK237=N~wPCYkVQc
    'U%ATa?kA^)\x^G=?(HdD!k??mBSJ?A%qn}3_&^:{$ck\Y>ck%Zcj<^9k96PSx>wCZD7y'.d=K[A5CwxEpmuRt3Gqj/NDUwgv9}r;+g6G&c\=!Q'r?4q+Y3KEvz^S?eQ<
    'xJa<z6+[/p%X8^Qm\N,+y2`5d+{}N\u/[+Y!zvgb]6-,aa~!^$`pa%J/)4nj["vxM}:"_4H<[DCGWwCEHnJnZzG7='5rU`:?<xY]w/Py!Fyjy%C6RSEb~?xhT6ZQXRe_
    'T2@A/}YfBNz[(_x}\Feku+F~Qf%gFcWH%**u=,NE$*z?\uHSuUZYe)-k?znm\9A]Rv5GJMFQJj?f$DaHtAM9#DvLaLfx',"Jz~Wt^9skUqM&x#Uy)6{RP$3"x-y5`K7d
    'acp&]'E2WcWEdxpg8U]gxbx5GTa#<~<6W+k'Z{=Sk?h#/=-](x8^\,eH*`&AEtMHN7bbN/P\u_9"cL#um.C3xTkCx]KHBaddA{Yh&Fu{>3T8Rh9x`/!=%s9uUfz@P\yt
    '~CZQ[Y3/fqHx,(({#aPC]+S.w42MQedqQc4]QJFbB&A6'TJ"Dmm#N=aX7A73"bENt$!2!Z+Bq{*3%eA6%Sa~suEc>6x+9U^NVn>CwhJ~^UJ):S{P@563j^UUv,2Te+bN

'Open/Activate/Initiate###############################################################################################################################################
Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Initialize()
    'HideBar Me
    Me.WDPage = 1
    Me.WDRow = 1
    Me.WDColumn = 1
    Call BuildWDDataArray
       
    Dim GetFile As Integer
      
    GetFile = MsgBox("Load Existing XLWD File" & vbNewLine & "(Excel Web Development File)?", vbYesNo, "New or Load")
    If GetFile = 6 Then
    
        'Get File
        Dim textline As String
        On Error GoTo en:
        myfile = Application.GetOpenFilename(FileFilter:="XLWD File (*.xlwd), *.xlwd")
        Open myfile For Input As #1
        Application.Wait (Now + TimeValue("0:00:05"))
        Do Until EOF(1)
            Line Input #1, textline
            XLWD = XLWD & textline
        Loop
        
        Call SplitWDDataArray
        Call LoadWDDataArrayToForm
    End If
    
    'Formatting Adjustments
en:
    Call ColumnTypeOptions
    Call ListBoxesLoad
End Sub

Private Sub WDCreate_Click()
Call LoadWDDataToArray
Call WriteHTMLDocument
End Sub

'Buttons##############################################################################################################################################################
Private Sub WDSave_Click()
    Call LoadWDDataToArray
    Call ConcatinateWDDataArray
    
    If myfile = "" Then
        Dim sfolder As String
        myfile = InputBox("File Name")
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = -1 Then ' if OK is pressed
                sfolder = .SelectedItems(1)
            End If
        End With
        myfile = sfolder & "\" & myfile & ".xlwd"
        Call SaveWDData
    Else
        Call SaveWDData
    End If
    
End Sub

Private Sub WDSaveAS_Click()
    Dim sfolder, myTempFileName As String
    
    myTempFileName = InputBox("File Name")
    If myTempFileName = "" Then Exit Sub
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sfolder = .SelectedItems(1)
        End If
    End With
    myTempFileName = sfolder & "\" & myTempFileName & ".xlwd"
    myfile = myTempFileName
    Call SaveWDData
End Sub

Private Sub WDOpen_Click()
'Get File
Dim textline, myTempFile As String
On Error GoTo en:
myTempFile = Application.GetOpenFilename(FileFilter:="XLWD File (*.xlwd), *.xlwd")
Open myTempFile For Input As #1

Do Until EOF(1)
    Line Input #1, textline
    XLWD = XLWD & textline
Loop

Call SplitWDDataArray
Call LoadWDDataArrayToForm
'Formatting Adjustments
myfile = myTempFile
en:
Call ColumnTypeOptions
End Sub

Private Sub WDClose_Click()
'UserForm_QueryClose 0, 0
Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub

Private Sub WDQuit_Click()
ShowBar Me
Me.Hide
Me.Show

End Sub

'Change_Page/Row/Column###############################################################################################################################################
'Page
Private Sub WDSelectPage_SpinDown()
    If Me.WDPage = 1 Then Exit Sub
    Call LoadWDDataToArray
    Me.WDPage = Me.WDPage - 1
    Me.WDRow = 1
    Me.WDColumn = 1
    Call LoadWDDataArrayToForm
    Call ChoiceFieldsBlank
End Sub

Private Sub WDSelectPage_SpinUp()
    If Me.WDPage = AD1 Then Exit Sub
    Call LoadWDDataToArray
    Me.WDPage = Me.WDPage + 1
    Me.WDRow = 1
    Me.WDColumn = 1
    Call LoadWDDataArrayToForm
    Call ChoiceFieldsBlank
End Sub

'Row
Private Sub WDSelectRow_SpinDown()
    If Me.WDRow = 1 Then Exit Sub
    Call LoadWDDataToArray
    Me.WDRow = Me.WDRow - 1
    Me.WDColumn = 1
    Call LoadWDDataArrayToForm
    Call ChoiceFieldsBlank
End Sub

Private Sub WDSelectRow_SpinUp()
    If Me.WDRow = AD2 Then Exit Sub
    Call LoadWDDataToArray
    Me.WDRow = Me.WDRow + 1
    Me.WDColumn = 1
    Call LoadWDDataArrayToForm
    Call ChoiceFieldsBlank
End Sub

'Column
Private Sub WDSelectColumn_SpinDown()
    If Me.WDColumn = 1 Then Exit Sub
    Call LoadWDDataToArray
    Me.WDColumn = Me.WDColumn - 1
    Call LoadWDDataArrayToForm
    Call ChoiceFieldsBlank
End Sub

Private Sub WDSelectColumn_SpinUp()
    If Me.WDColumn = AD3 Then Exit Sub
    Call LoadWDDataToArray
    Me.WDColumn = Me.WDColumn + 1
    Call LoadWDDataArrayToForm
    Call ChoiceFieldsBlank
End Sub

Sub ChoiceFieldsBlank()
    If WDdata(Me.WDPage, 0, 0, 2) = "" Then WDRibbonBar = False
    If WDdata(Me.WDPage, Me.WDRow, 0, 14) = "" Then WDIncludeRow = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 191) = "" Then WDOptionText = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 192) = "" Then WDOptionMetrics = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 193) = "" Then WDOptionCharts = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 194) = "" Then WDOptionTable = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 195) = "" Then WDIncludeColumn = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 136) = "" Then WDChartLayerInclude1 = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 137) = "" Then WDChartLayerInclude2 = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 138) = "" Then WDChartLayerInclude3 = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 139) = "" Then WDChartLayerInclude4 = False
    If WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 140) = "" Then WDChartLayerInclude5 = False
End Sub

'INIT_FUNCTIONS#######################################################################################################################################################

Sub SiteDevelopment()
    WebpageDevelopment.Show
End Sub

Sub SaveWDData()
    Dim myTxt, fileName, fileExt, add As String
    On Error GoTo en:
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(myfile, True)
    a.WriteLine XLWD
    a.Close
en:
End Sub

Sub BuildWDDataArray()

    MD = "C:2s3ZpnC8A,S{T/H)SWZ'24\mmuv3Egb%M/QDA86AUer`zn=Z'u@8;tTry{gqYa5VK`.(y9LvR~&PTs\=RQW2<}A@s:#Lr>V(W2;-s4~$Wbq9~NT'},Q.bm*Rj'7Nve" 'Metadata
    PP = "^cv8CR(U<3wbvh2>*ee.bK'6b:ZQqwj@s#?EQLhU:U>4Q:^[pALeg,/a+/]R$ZuG48_rTuC9)kQyKUZUe:#jv_.DK$3fm}g%*]~/,`A&$V;5;[yAz$BPw}TV`yXqB~G%" 'Page
    PD = "C3`:j~52,`/Bt:b:y~y[^PRtznp8^XE-vSA:93=#LjLR>M~8%%$jB<x<G;5)*cB4sPFV9#}/Rd5E8^)<@NazNjEX8S~ND&Qk/Mt_n&3?Y5Dbxx[GNG#En,GZ&k-3RhD:" 'Page Data
    RR = "2@Nf<S>GH3VQEvZY+GSw:*-@(?%DV_h{#6AZp'6{DL`~w.cM<U$;8e'BqhyCpSZ2WQ'%}]N+6]xf`pT@,_b@a-g2[]*Hh!8}U4ngnYFVWgyV$y?::]D&bBw[fWD}Y~GF" 'Row
    RD = "ED!>g9b)u&;4>?~>f'#3_x=S=Pjmstm:CZ/r'tY4A[fMk5~P%C>*77*)<u^9'sUXGWhKpZ9RtJ{%{zrABU4~Mrmh4MuS,,pGsSDEv4)[~F$M6PbCUEdA9gGgP'tbQzPn" 'Row Data
    CC = "`YE%%8y8zGX_d7<y*FDSG3!h.\KF2qQf%A#z8[v\@ML~bU#ehM<U+aV3t,7YdwgU>ydR_E^V>4xzGXfP;c3j.a45QFJwRxv/pD:=5QEK~4@Q7m5]KkaD5;!#q_T6t$&'" 'Column
    CD = "7,)zbZ,>-_*/*6yeLSX&~@@y,k6$HksXyX~ex}#g&(AL\yDY(kcn9!`xQg9$[GyVq;(a/vC$4T=^+jB?yL3M8m'u]76F)v/XaEV'#>K?f5g=5]7>^h!y4:%c{_SW*fky" 'Column Data
    TD = "auY5xUf454Ajc6~~}.CS.<DZU7bB?Ee.+;YZ5$J?N9!68.~fgrquYj]{A,5Rfe$(;=caBe*\g!$%b4REtwkn6w]]cT>N[T([VE_J?%}$DNak`w)@:58zse[<4M#d.Zp6" 'Text Column Data
    JD = "Hn:w_N<mSpmm,w~_DrC,~}:6$yneD:9+KhA>,nr3X+w-:jVQYCpND=]4?-,g[pA)wcN''zffZW(U=?&uXGj&~V%8N^5ryBN`@+MsY!<;x`;r6dd#y*'5@:x2{u_w}LH]" 'JavaScript Data

    AD1 = 100  'WDdata Dimention 1
    AD2 = 100  'WDdata Dimention 2
    AD3 = 30   'WDdata Dimention 3
    AD4 = 195  'WDdata Dimention 4
    
    DM = 4     'Metadata Fields
    DP = 18    'Page Fields
    DR = 14    'Row Fields
    DC = 195   'Column Fields
    
    ReDim WDdata(0 To AD1, 0 To AD2, 0 To AD3, 0 To AD4)
    
End Sub

Sub ConcatinateWDDataArray()
    'P = Page    R = Row      C = Column

    'Metadata    0.0.0.0-##
    'Page data   P.0.0.0-##   0 = 0(exclude) or 1(include)
    'Row data    P.R.0.0-##   0 = 0(exclude) or 1(include)
    'Column data P.R.C.0-###  0 = 0(exclude) or 1(include)

    Dim MTD As Variant
    Dim AllData As String
    'WDdata(0 to 100, 0 to 100, 0 to 30, 0 to 300)

    AllData = ""

    'Metadata
    For M = 0 To DM
        'If WDdata(0, 0, 0, M) = "" Then Exit For
        AllData = AllData & WDdata(0, 0, 0, M) & MM
    Next M
    AllData = AllData & PP
    'Pages
    For P = 1 To 100
        If WDdata(P, 0, 0, P) = "" Then Exit For
        'Page Data
        For PA = 0 To DP
            AllData = AllData & WDdata(P, 0, 0, PA) & PD
        Next PA
        
        'Rows
        For R = 1 To 100
            If WDdata(P, R, 0, R - 1) = False Then Exit For
            'Row Data
            For RA = 0 To DR
                AllData = AllData & WDdata(P, R, 0, RA) & RD
            Next RA

            'Columns
            For c = 1 To 30
                If WDdata(P, R, c, c - 1) = False Then Exit For
                'Column Data
                For CA = 0 To DC
                    AllData = AllData & WDdata(P, R, c, CA) & CD
                Next CA

                'Concatinate Columns
                AllData = AllData & CC
            Next c

            'Concatinate Rows
            AllData = AllData & RR
        Next R

        'Concatinate Pages
        AllData = AllData & PP
    Next P

    'Write to XLWD File
    XLWD = AllData
    
End Sub

Sub SplitWDDataArray()

    'P = Page    R = Row      C = Column

    'Metadata    0.0.0.0-##
    'Page data   P.0.0.0-##   0 = 0(exclude) or 1(include)
    'Row data    P.R.0.0-##   0 = 0(exclude) or 1(include)
    'Column data P.R.C.0-###  0 = 0(exclude) or 1(include)


    Dim MTD As Variant
    Dim PGS, RWS, CLS As Variant
    Dim PGSD, RWSD, CLSD As Variant

    'Metadata
    MTD = Split(XLWD, MM)
    For M = 1 To UBound(MTD) - 1
        WDdata(0, 0, 0, M - 1) = MTD(M, 1)
    Next M

    'Pages
    PGS = Split(XLWD, Me.PP)
    For P = 2 To UBound(PGS)

        'Page Data
        PGSD = Split(PGS(P - 1), Me.PD)
        For PA = 1 To UBound(PGSD) - 1
            WDdata(P - 1, 0, 0, PA) = PGSD(PA)
        Next PA
            
        'Rows
        RWS = Split(PGS(P - 1), Me.RR)
        For R = 0 To UBound(RWS)

            'Row Data
            RWSD = Split(RWS(R), Me.RD)
            For RA = 1 To UBound(RWSD) - 1
                WDdata(P - 1, R + 1, 0, RA) = RWSD(RA)
            Next RA

            'Columns
            CLS = Split(RWS(R), Me.CC)
            For c = 0 To UBound(CLS)

                'Column Data
                CLSD = Split(CLS(c), Me.CD)
                For CA = 1 To UBound(CLSD) - 1
                    WDdata(P - 1, R + 1, c + 1, CA) = CLSD(CA)
                Next CA
            Next c
        Next R
    Next P
    
    
End Sub

Sub LoadWDDataToArray()
    WDdata(Me.WDPage, 0, 0, 1) = Me.WDPageTabTitle
    WDdata(Me.WDPage, 0, 0, 2) = Me.WDRibbonBar
    WDdata(Me.WDPage, 0, 0, 3) = Me.WDPageBackgroundColor
    WDdata(Me.WDPage, 0, 0, 4) = Me.WDPageBackgroundImageURL
    WDdata(Me.WDPage, 0, 0, 5) = Me.WDPageJSLink1
    WDdata(Me.WDPage, 0, 0, 6) = Me.WDPageJSLink2
    WDdata(Me.WDPage, 0, 0, 7) = Me.WDPageCSSLink
    WDdata(Me.WDPage, 0, 0, 8) = Me.WDRibbinLinkText1
    WDdata(Me.WDPage, 0, 0, 9) = Me.WDRibbinLinkText2
    WDdata(Me.WDPage, 0, 0, 10) = Me.WDRibbinLinkText3
    WDdata(Me.WDPage, 0, 0, 11) = Me.WDRibbinLinkText4
    WDdata(Me.WDPage, 0, 0, 12) = Me.WDRibbinLinkText5
    WDdata(Me.WDPage, 0, 0, 13) = Me.WDRibbinLinkURL1
    WDdata(Me.WDPage, 0, 0, 14) = Me.WDRibbinLinkURL2
    WDdata(Me.WDPage, 0, 0, 15) = Me.WDRibbinLinkURL3
    WDdata(Me.WDPage, 0, 0, 16) = Me.WDRibbinLinkURL4
    WDdata(Me.WDPage, 0, 0, 17) = Me.WDRibbinLinkURL5
    WDdata(Me.WDPage, 0, 0, 18) = Me.WDJavaScript
    
    
    WDdata(Me.WDPage, Me.WDRow, 0, 0) = Me.WDIncludeRow
    
    WDdata(Me.WDPage, Me.WDRow, 0, 1) = Me.WDRowStyle
    WDdata(Me.WDPage, Me.WDRow, 0, 2) = Me.WDRowTextAlign
    WDdata(Me.WDPage, Me.WDRow, 0, 3) = Me.WDRowBackgroundColor
    WDdata(Me.WDPage, Me.WDRow, 0, 4) = Me.WDRowID
    WDdata(Me.WDPage, Me.WDRow, 0, 5) = Me.WDRowWidth
    WDdata(Me.WDPage, Me.WDRow, 0, 6) = Me.WDRowHeight
    WDdata(Me.WDPage, Me.WDRow, 0, 7) = Me.WDRowPaddingTop
    WDdata(Me.WDPage, Me.WDRow, 0, 8) = Me.WDRowPaddingBottom
    WDdata(Me.WDPage, Me.WDRow, 0, 9) = Me.WDRowMarginTop
    WDdata(Me.WDPage, Me.WDRow, 0, 10) = Me.WDRowMarginBottom
    WDdata(Me.WDPage, Me.WDRow, 0, 11) = Me.WDRowBorderColor
    WDdata(Me.WDPage, Me.WDRow, 0, 12) = Me.WDRowBorderRadious
    WDdata(Me.WDPage, Me.WDRow, 0, 13) = Me.WDRowBorderThickness
    
    WDdata(Me.WDPage, Me.WDRow, 0, 14) = Me.WDIncludeRow

    
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 0) = Me.WDIncludeColumn
    
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 1) = Me.WDColumnText
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 2) = Me.WDColumnTextHeight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 3) = Me.WDColumnTextWidth
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 4) = Me.WDColumnTableSheetName
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 5) = Me.WDColumnTableTopLeftCell
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 6) = Me.WDColumnTableTopRightCell
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 7) = Me.WDColumnTableAggregation
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 8) = Me.WDColumnTableStyle
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 9) = Me.WDColumnTableFont
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 10) = Me.WDColumnTableTextSize
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 11) = Me.WDColumnTableHeader
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 12) = Me.WDColumnTablePaddingLeft
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 13) = Me.WDColumnTablePaddingRight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 14) = Me.WDColumnTableMarginLeft
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 15) = Me.WDColumnTableMarginRight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 16) = Me.WDColumnTableWidth
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 17) = Me.WDMartricsText1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 18) = Me.WDMartricsText2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 19) = Me.WDMartricsText3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 20) = Me.WDMartricsText4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 21) = Me.WDMartricsText5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 22) = Me.WDMartricsText6
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 23) = Me.WDMartricsText7
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 24) = Me.WDMartricsText8
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 25) = Me.WDMartricsText9
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 26) = Me.WDMartricsText10
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 27) = Me.WDMartricsSheet1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 28) = Me.WDMartricsSheet2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 29) = Me.WDMartricsSheet3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 30) = Me.WDMartricsSheet4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 31) = Me.WDMartricsSheet5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 32) = Me.WDMartricsSheet6
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 33) = Me.WDMartricsSheet7
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 34) = Me.WDMartricsSheet8
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 35) = Me.WDMartricsSheet9
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 36) = Me.WDMartricsSheet10
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 37) = Me.WDMartricsCell1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 38) = Me.WDMartricsCell2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 39) = Me.WDMartricsCell3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 40) = Me.WDMartricsCell4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 41) = Me.WDMartricsCell5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 42) = Me.WDMartricsCell6
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 43) = Me.WDMartricsCell7
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 44) = Me.WDMartricsCell8
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 45) = Me.WDMartricsCell9
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 46) = Me.WDMartricsCell10
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 47) = Me.WDMartricsTextSize1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 48) = Me.WDMartricsTextSize2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 49) = Me.WDMartricsTextSize3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 50) = Me.WDMartricsTextSize4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 51) = Me.WDMartricsTextSize5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 52) = Me.WDMartricsTextSize6
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 53) = Me.WDMartricsTextSize7
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 54) = Me.WDMartricsTextSize8
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 55) = Me.WDMartricsTextSize9
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 56) = Me.WDMartricsTextSize10
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 57) = Me.WDMartricsWidth1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 58) = Me.WDMartricsWidth2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 59) = Me.WDMartricsWidth3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 60) = Me.WDMartricsWidth4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 61) = Me.WDMartricsWidth5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 62) = Me.WDMartricsWidth6
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 63) = Me.WDMartricsWidth7
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 64) = Me.WDMartricsWidth8
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 65) = Me.WDMartricsWidth9
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 66) = Me.WDMartricsWidth10
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 67) = Me.WDMartricsStyle1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 68) = Me.WDMartricsStyle2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 69) = Me.WDMartricsStyle3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 70) = Me.WDMartricsStyle4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 71) = Me.WDMartricsStyle5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 72) = Me.WDMartricsStyle6
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 73) = Me.WDMartricsStyle7
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 74) = Me.WDMartricsStyle8
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 75) = Me.WDMartricsStyle9
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 76) = Me.WDMartricsStyle10
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 77) = Me.WDMartricsColumnPaddingLeft
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 78) = Me.WDMartricsColumnPaddingRight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 79) = Me.WDMartricsColumnMarginLeft
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 80) = Me.WDMartricsColumnMarginRight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 81) = Me.WDMartricsColumnHeight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 82) = Me.WDMartricsColumnWidth
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 83) = Me.WDChartTitleText
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 84) = Me.WDChartTitleFontStyle
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 85) = Me.WDChartTitleTxtxColor
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 86) = Me.WDChartTitleTextSize
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 87) = Me.WDChartTitleFont
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 88) = Me.WDChartTitleFontWeight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 89) = Me.WDChartTitleVerticleAlign
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 90) = Me.WDChartTitleHorizontalAlign
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 91) = Me.WDChartSubtitleText
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 92) = Me.WDChartSubtitleFontStyle
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 93) = Me.WDChartSubtitleTextColor
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 94) = Me.WDChartSubtitleTextSize
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 95) = Me.WDChartSubtitleFont
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 96) = Me.WDChartSubtitleFontWeight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 97) = Me.WDChartSubtitleVerticleAlign
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 98) = Me.WDChartSubtitleHorizontalAlign
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 99) = Me.WDChartTooltipChoice
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 100) = Me.WDChartTooltipBackgroindColor
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 101) = Me.WDChartTooltipBorderRadious
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 102) = Me.WDChartTooltipTextColor
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 103) = Me.WDChartTooltipTextSize
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 104) = Me.WDChartTooltipFont
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 105) = Me.WDChartLegendChoice
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 106) = Me.WDChartLegendVerticleAlign
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 107) = Me.WDChartLegendHorizontalAlign
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 108) = Me.WDChartLegendTextColor
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 109) = Me.WDChartLegendTextSize
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 110) = Me.WDChartLegendFont
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 111) = Me.WDChartKPIChoice
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 112) = Me.WDChartKPIMeasures
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 113) = Me.WDChartHeight
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 114) = Me.WDChartWidth
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 115) = Me.WDChartBackgroundColor
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 116) = Me.WDChartDataPrefix
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 117) = Me.WDChartDataBoundries
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 118) = Me.WDChartAnimationLenght
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 119) = Me.WDChartDataSufix
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 120) = Me.WDChartExportFileTitle
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 121) = Me.WDChartTheme
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 122) = Me.WDChartColorsChoicePredefigned
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 123) = Me.WDChartColorsChoiceCustom
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 124) = Me.WDChartColorsPredefigned
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 125) = Me.WDChartColors1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 126) = Me.WDChartColors2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 127) = Me.WDChartColors3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 128) = Me.WDChartColors4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 129) = Me.WDChartColors5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 130) = Me.WDChartColors6
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 131) = Me.WDChartColors7
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 132) = Me.WDChartColors8
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 133) = Me.WDChartColors9
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 134) = Me.WDChartColors10
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 135) = Me.WDChartColors11
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 136) = Me.WDChartLayerInclude1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 137) = Me.WDChartLayerInclude2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 138) = Me.WDChartLayerInclude3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 139) = Me.WDChartLayerInclude4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 140) = Me.WDChartLayerInclude5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 141) = Me.WDChartLayerChartType1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 142) = Me.WDChartLayerChartType2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 143) = Me.WDChartLayerChartType3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 144) = Me.WDChartLayerChartType4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 145) = Me.WDChartLayerChartType5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 146) = Me.WDChartLayerAggregation1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 147) = Me.WDChartLayerAggregation2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 148) = Me.WDChartLayerAggregation3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 149) = Me.WDChartLayerAggregation4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 150) = Me.WDChartLayerAggregation5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 151) = Me.WDChartLayerSheetName1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 152) = Me.WDChartLayerSheetName2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 153) = Me.WDChartLayerSheetName3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 154) = Me.WDChartLayerSheetName4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 155) = Me.WDChartLayerSheetName5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 156) = Me.WDChartLayerIndexTopLeftCell1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 157) = Me.WDChartLayerIndexTopLeftCell2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 158) = Me.WDChartLayerIndexTopLeftCell3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 159) = Me.WDChartLayerIndexTopLeftCell4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 160) = Me.WDChartLayerIndexTopLeftCell5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 161) = Me.WDChartLayerValuesTopLeftCell1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 162) = Me.WDChartLayerValuesTopLeftCell2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 163) = Me.WDChartLayerValuesTopLeftCell3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 164) = Me.WDChartLayerValuesTopLeftCell4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 165) = Me.WDChartLayerValuesTopLeftCell5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 166) = Me.WDChartLayerColor1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 167) = Me.WDChartLayerColor2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 168) = Me.WDChartLayerColor3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 169) = Me.WDChartLayerColor4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 170) = Me.WDChartLayerColor5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 171) = Me.WDChartLayerName1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 172) = Me.WDChartLayerName2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 173) = Me.WDChartLayerName3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 174) = Me.WDChartLayerName4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 175) = Me.WDChartLayerName5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 176) = Me.WDChartLayerShowInLegend1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 177) = Me.WDChartLayerShowInLegend2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 178) = Me.WDChartLayerShowInLegend3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 179) = Me.WDChartLayerShowInLegend4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 180) = Me.WDChartLayerShowInLegend5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 181) = Me.WDChartLayerYValueStringFormat1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 182) = Me.WDChartLayerYValueStringFormat2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 183) = Me.WDChartLayerYValueStringFormat3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 184) = Me.WDChartLayerYValueStringFormat4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 185) = Me.WDChartLayerYValueStringFormat5
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 186) = Me.WDChartLayerYAxisType1
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 187) = Me.WDChartLayerYAxisType2
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 188) = Me.WDChartLayerYAxisType3
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 189) = Me.WDChartLayerYAxisType4
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 190) = Me.WDChartLayerYAxisType5
    
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 191) = Me.WDOptionText
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 192) = Me.WDOptionMetrics
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 193) = Me.WDOptionCharts
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 194) = Me.WDOptionTable
    
    WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 195) = Me.WDIncludeColumn


End Sub

Sub LoadWDDataArrayToForm()

    Me.WDPageTabTitle = WDdata(Me.WDPage, 0, 0, 1)
    'Debug.Print (WDdata(Me.WDPage, 0, 0, 1))
    Me.WDRibbonBar = WDdata(Me.WDPage, 0, 0, 2)
    Me.WDPageBackgroundColor = WDdata(Me.WDPage, 0, 0, 3)
    Me.WDPageBackgroundImageURL = WDdata(Me.WDPage, 0, 0, 4)
    Me.WDPageJSLink1 = WDdata(Me.WDPage, 0, 0, 5)
    Me.WDPageJSLink2 = WDdata(Me.WDPage, 0, 0, 6)
    Me.WDPageCSSLink = WDdata(Me.WDPage, 0, 0, 7)
    Me.WDRibbinLinkText1 = WDdata(Me.WDPage, 0, 0, 8)
    Me.WDRibbinLinkText2 = WDdata(Me.WDPage, 0, 0, 9)
    Me.WDRibbinLinkText3 = WDdata(Me.WDPage, 0, 0, 10)
    Me.WDRibbinLinkText4 = WDdata(Me.WDPage, 0, 0, 11)
    Me.WDRibbinLinkText5 = WDdata(Me.WDPage, 0, 0, 12)
    Me.WDRibbinLinkURL1 = WDdata(Me.WDPage, 0, 0, 13)
    Me.WDRibbinLinkURL2 = WDdata(Me.WDPage, 0, 0, 14)
    Me.WDRibbinLinkURL3 = WDdata(Me.WDPage, 0, 0, 15)
    Me.WDRibbinLinkURL4 = WDdata(Me.WDPage, 0, 0, 16)
    Me.WDRibbinLinkURL5 = WDdata(Me.WDPage, 0, 0, 17)
    Me.WDJavaScript = WDdata(Me.WDPage, 0, 0, 18)
    
    
    
    Me.WDIncludeRow = WDdata(Me.WDPage, Me.WDRow, 0, 0)

    
    Me.WDRowStyle = WDdata(Me.WDPage, Me.WDRow, 0, 1)
    Me.WDRowTextAlign = WDdata(Me.WDPage, Me.WDRow, 0, 2)
    Me.WDRowBackgroundColor = WDdata(Me.WDPage, Me.WDRow, 0, 3)
    Me.WDRowID = WDdata(Me.WDPage, Me.WDRow, 0, 4)
    Me.WDRowWidth = WDdata(Me.WDPage, Me.WDRow, 0, 5)
    Me.WDRowHeight = WDdata(Me.WDPage, Me.WDRow, 0, 6)
    Me.WDRowPaddingTop = WDdata(Me.WDPage, Me.WDRow, 0, 7)
    Me.WDRowPaddingBottom = WDdata(Me.WDPage, Me.WDRow, 0, 8)
    Me.WDRowMarginTop = WDdata(Me.WDPage, Me.WDRow, 0, 9)
    Me.WDRowMarginBottom = WDdata(Me.WDPage, Me.WDRow, 0, 10)
    Me.WDRowBorderColor = WDdata(Me.WDPage, Me.WDRow, 0, 11)
    Me.WDRowBorderRadious = WDdata(Me.WDPage, Me.WDRow, 0, 12)
    Me.WDRowBorderThickness = WDdata(Me.WDPage, Me.WDRow, 0, 13)
    
    Me.WDIncludeRow = WDdata(Me.WDPage, Me.WDRow, 0, 14)

    
    Me.WDIncludeColumn = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 0)
    
    Me.WDColumnText = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 1)
    Me.WDColumnTextHeight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 2)
    Me.WDColumnTextWidth = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 3)
    Me.WDColumnTableSheetName = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 4)
    Me.WDColumnTableTopLeftCell = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 5)
    Me.WDColumnTableTopRightCell = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 6)
    Me.WDColumnTableAggregation = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 7)
    Me.WDColumnTableStyle = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 8)
    Me.WDColumnTableFont = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 9)
    Me.WDColumnTableTextSize = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 10)
    Me.WDColumnTableHeader = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 11)
    Me.WDColumnTablePaddingLeft = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 12)
    Me.WDColumnTablePaddingRight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 13)
    Me.WDColumnTableMarginLeft = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 14)
    Me.WDColumnTableMarginRight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 15)
    Me.WDColumnTableWidth = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 16)
    Me.WDMartricsText1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 17)
    Me.WDMartricsText2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 18)
    Me.WDMartricsText3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 19)
    Me.WDMartricsText4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 20)
    Me.WDMartricsText5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 21)
    Me.WDMartricsText6 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 22)
    Me.WDMartricsText7 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 23)
    Me.WDMartricsText8 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 24)
    Me.WDMartricsText9 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 25)
    Me.WDMartricsText10 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 26)
    Me.WDMartricsSheet1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 27)
    Me.WDMartricsSheet2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 28)
    Me.WDMartricsSheet3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 29)
    Me.WDMartricsSheet4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 30)
    Me.WDMartricsSheet5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 31)
    Me.WDMartricsSheet6 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 32)
    Me.WDMartricsSheet7 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 33)
    Me.WDMartricsSheet8 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 34)
    Me.WDMartricsSheet9 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 35)
    Me.WDMartricsSheet10 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 36)
    Me.WDMartricsCell1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 37)
    Me.WDMartricsCell2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 38)
    Me.WDMartricsCell3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 39)
    Me.WDMartricsCell4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 40)
    Me.WDMartricsCell5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 41)
    Me.WDMartricsCell6 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 42)
    Me.WDMartricsCell7 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 43)
    Me.WDMartricsCell8 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 44)
    Me.WDMartricsCell9 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 45)
    Me.WDMartricsCell10 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 46)
    Me.WDMartricsTextSize1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 47)
    Me.WDMartricsTextSize2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 48)
    Me.WDMartricsTextSize3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 49)
    Me.WDMartricsTextSize4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 50)
    Me.WDMartricsTextSize5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 51)
    Me.WDMartricsTextSize6 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 52)
    Me.WDMartricsTextSize7 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 53)
    Me.WDMartricsTextSize8 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 54)
    Me.WDMartricsTextSize9 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 55)
    Me.WDMartricsTextSize10 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 56)
    Me.WDMartricsWidth1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 57)
    Me.WDMartricsWidth2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 58)
    Me.WDMartricsWidth3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 59)
    Me.WDMartricsWidth4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 60)
    Me.WDMartricsWidth5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 61)
    Me.WDMartricsWidth6 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 62)
    Me.WDMartricsWidth7 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 63)
    Me.WDMartricsWidth8 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 64)
    Me.WDMartricsWidth9 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 65)
    Me.WDMartricsWidth10 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 66)
    Me.WDMartricsStyle1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 67)
    Me.WDMartricsStyle2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 68)
    Me.WDMartricsStyle3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 69)
    Me.WDMartricsStyle4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 70)
    Me.WDMartricsStyle5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 71)
    Me.WDMartricsStyle6 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 72)
    Me.WDMartricsStyle7 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 73)
    Me.WDMartricsStyle8 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 74)
    Me.WDMartricsStyle9 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 75)
    Me.WDMartricsStyle10 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 76)
    Me.WDMartricsColumnPaddingLeft = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 77)
    Me.WDMartricsColumnPaddingRight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 78)
    Me.WDMartricsColumnMarginLeft = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 79)
    Me.WDMartricsColumnMarginRight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 80)
    Me.WDMartricsColumnHeight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 81)
    Me.WDMartricsColumnWidth = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 82)
    Me.WDChartTitleText = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 83)
    Me.WDChartTitleFontStyle = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 84)
    Me.WDChartTitleTxtxColor = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 85)
    Me.WDChartTitleTextSize = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 86)
    Me.WDChartTitleFont = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 87)
    Me.WDChartTitleFontWeight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 88)
    Me.WDChartTitleVerticleAlign = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 89)
    Me.WDChartTitleHorizontalAlign = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 90)
    Me.WDChartSubtitleText = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 91)
    Me.WDChartSubtitleFontStyle = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 92)
    Me.WDChartSubtitleTextColor = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 93)
    Me.WDChartSubtitleTextSize = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 94)
    Me.WDChartSubtitleFont = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 95)
    Me.WDChartSubtitleFontWeight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 96)
    Me.WDChartSubtitleVerticleAlign = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 97)
    Me.WDChartSubtitleHorizontalAlign = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 98)
    Me.WDChartTooltipChoice = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 99)
    Me.WDChartTooltipBackgroindColor = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 100)
    Me.WDChartTooltipBorderRadious = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 101)
    Me.WDChartTooltipTextColor = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 102)
    Me.WDChartTooltipTextSize = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 103)
    Me.WDChartTooltipFont = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 104)
    Me.WDChartLegendChoice = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 105)
    Me.WDChartLegendVerticleAlign = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 106)
    Me.WDChartLegendHorizontalAlign = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 107)
    Me.WDChartLegendTextColor = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 108)
    Me.WDChartLegendTextSize = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 109)
    Me.WDChartLegendFont = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 110)
    Me.WDChartKPIChoice = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 111)
    Me.WDChartKPIMeasures = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 112)
    Me.WDChartHeight = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 113)
    Me.WDChartWidth = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 114)
    Me.WDChartBackgroundColor = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 115)
    Me.WDChartDataPrefix = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 116)
    Me.WDChartDataBoundries = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 117)
    Me.WDChartAnimationLenght = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 118)
    Me.WDChartDataSufix = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 119)
    Me.WDChartExportFileTitle = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 120)
    Me.WDChartTheme = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 121)
    Me.WDChartColorsChoicePredefigned = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 122)
    Me.WDChartColorsChoiceCustom = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 123)
    Me.WDChartColorsPredefigned = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 124)
    Me.WDChartColors1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 125)
    Me.WDChartColors2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 126)
    Me.WDChartColors3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 127)
    Me.WDChartColors4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 128)
    Me.WDChartColors5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 129)
    Me.WDChartColors6 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 130)
    Me.WDChartColors7 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 131)
    Me.WDChartColors8 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 132)
    Me.WDChartColors9 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 133)
    Me.WDChartColors10 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 134)
    Me.WDChartColors11 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 135)
    Me.WDChartLayerInclude1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 136)
    Me.WDChartLayerInclude2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 137)
    Me.WDChartLayerInclude3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 138)
    Me.WDChartLayerInclude4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 139)
    Me.WDChartLayerInclude5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 140)
    Me.WDChartLayerChartType1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 141)
    Me.WDChartLayerChartType2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 142)
    Me.WDChartLayerChartType3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 143)
    Me.WDChartLayerChartType4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 144)
    Me.WDChartLayerChartType5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 145)
    Me.WDChartLayerAggregation1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 146)
    Me.WDChartLayerAggregation2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 147)
    Me.WDChartLayerAggregation3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 148)
    Me.WDChartLayerAggregation4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 149)
    Me.WDChartLayerAggregation5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 150)
    Me.WDChartLayerSheetName1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 151)
    Me.WDChartLayerSheetName2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 152)
    Me.WDChartLayerSheetName3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 153)
    Me.WDChartLayerSheetName4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 154)
    Me.WDChartLayerSheetName5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 155)
    Me.WDChartLayerIndexTopLeftCell1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 156)
    Me.WDChartLayerIndexTopLeftCell2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 157)
    Me.WDChartLayerIndexTopLeftCell3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 158)
    Me.WDChartLayerIndexTopLeftCell4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 159)
    Me.WDChartLayerIndexTopLeftCell5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 160)
    Me.WDChartLayerValuesTopLeftCell1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 161)
    Me.WDChartLayerValuesTopLeftCell2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 162)
    Me.WDChartLayerValuesTopLeftCell3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 163)
    Me.WDChartLayerValuesTopLeftCell4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 164)
    Me.WDChartLayerValuesTopLeftCell5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 165)
    Me.WDChartLayerColor1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 166)
    Me.WDChartLayerColor2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 167)
    Me.WDChartLayerColor3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 168)
    Me.WDChartLayerColor4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 169)
    Me.WDChartLayerColor5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 170)
    Me.WDChartLayerName1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 171)
    Me.WDChartLayerName2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 172)
    Me.WDChartLayerName3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 173)
    Me.WDChartLayerName4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 174)
    Me.WDChartLayerName5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 175)
    Me.WDChartLayerShowInLegend1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 176)
    Me.WDChartLayerShowInLegend2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 177)
    Me.WDChartLayerShowInLegend3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 178)
    Me.WDChartLayerShowInLegend4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 179)
    Me.WDChartLayerShowInLegend5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 180)
    Me.WDChartLayerYValueStringFormat1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 181)
    Me.WDChartLayerYValueStringFormat2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 182)
    Me.WDChartLayerYValueStringFormat3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 183)
    Me.WDChartLayerYValueStringFormat4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 184)
    Me.WDChartLayerYValueStringFormat5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 185)
    Me.WDChartLayerYAxisType1 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 186)
    Me.WDChartLayerYAxisType2 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 187)
    Me.WDChartLayerYAxisType3 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 188)
    Me.WDChartLayerYAxisType4 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 189)
    Me.WDChartLayerYAxisType5 = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 190)
    
    Me.WDOptionText = WDdata(Me.WDPage, Me.WDRow, WebpageDevelopment.WDColumn, 191)
    Me.WDOptionMetrics = WDdata(Me.WDPage, Me.WDRow, WebpageDevelopment.WDColumn, 192)
    Me.WDOptionCharts = WDdata(Me.WDPage, Me.WDRow, WebpageDevelopment.WDColumn, 193)
    Me.WDOptionTable = WDdata(Me.WDPage, Me.WDRow, WebpageDevelopment.WDColumn, 194)
    
    Me.WDIncludeColumn = WDdata(Me.WDPage, Me.WDRow, Me.WDColumn, 195)

End Sub

'FORMATTING###########################################################################################################################################################

Private Sub WDOptionCharts_Click()
Call ColumnTypeOptions
End Sub

Private Sub WDOptionMetrics_Click()
Call ColumnTypeOptions
End Sub

Private Sub WDOptionTable_Click()
Call ColumnTypeOptions
End Sub

Private Sub WDOptionText_Click()
Call ColumnTypeOptions
End Sub

Sub ColumnTypeOptions()
    ' This Sub changes the border colour of selected and unselected column frames

    'Variables
    Dim T, F As Variant
    T = &H405AF8
    F = &HD1CCCA

    'True
    If Me.WDOptionCharts = True Then Me.WDColumnChartFrame.BackColor = T
    If Me.WDOptionMetrics = True Then Me.WDColumnMetricsFrame.BackColor = T
    If Me.WDOptionMetrics = True Then Me.WDMetricsFrame.BackColor = T
    If Me.WDOptionText = True Then Me.WDColumnTextFrame.BackColor = T
    If Me.WDOptionTable = True Then Me.WDColumnTableFrame.BackColor = T
    
    'False
    If Me.WDOptionCharts = False Then Me.WDColumnChartFrame.BackColor = F
    If Me.WDOptionMetrics = False Then Me.WDColumnMetricsFrame.BackColor = F
    If Me.WDOptionMetrics = False Then Me.WDMetricsFrame.BackColor = F
    If Me.WDOptionText = False Then Me.WDColumnTextFrame.BackColor = F
    If Me.WDOptionTable = False Then Me.WDColumnTableFrame.BackColor = F

End Sub




'Validation###########################################################################################################################################################

'Numeric Warning
Sub MustBeNumeric()
    MsgBox ("Must Be Numeric")
End Sub




'Row Width
Private Sub WDRowWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDROWWidthIsNumeric = Me.WDRowWidth
End Sub
Private Sub WDRowWidth_Change()
    If IsNumeric(Me.WDRowWidth) = False And Me.WDRowWidth <> "" Then
        Me.WDRowWidth = WDROWWidthIsNumeric
        Call MustBeNumeric
    End If
End Sub

'Row Height
Private Sub WDRowHeight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDRowHeightIsNumeric = Me.WDRowHeight
End Sub
Private Sub WDRowHeight_Change()
    If IsNumeric(Me.WDRowHeight) = False And Me.WDRowHeight <> "" Then
        Me.WDRowHeight = WDRowHeightIsNumeric
        Call MustBeNumeric
    End If
End Sub


'Row Padding Top
Private Sub WDRowPaddingTop_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDRowPaddingTopIsNumeric = Me.WDRowPaddingTop
End Sub
Private Sub WDRowPaddingTop_Change()
    If IsNumeric(Me.WDRowPaddingTop) = False And Me.WDRowPaddingTop <> "" Then
        Me.WDRowPaddingTop = WDRowPaddingTopIsNumeric
        Call MustBeNumeric
    End If
End Sub

'Row Padding Bottom
Private Sub WDRowPaddingBottom_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDRowPaddingBottomTopIsNumeric = Me.WDRowPaddingBottom
End Sub
Private Sub WDRowPaddingBottom_Change()
    If IsNumeric(Me.WDRowPaddingBottom) = False And Me.WDRowPaddingBottom <> "" Then
        Me.WDRowPaddingBottom = WDRowPaddingBottomIsNumeric
        Call MustBeNumeric
    End If
End Sub



'Row Margin Top
Private Sub WDRowMarginTop_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDRowMarginTopIsNumeric = Me.WDRowMarginTop
End Sub
Private Sub WDRowMarginTop_Change()
    If IsNumeric(Me.WDRowMarginTop) = False And Me.WDRowMarginTop <> "" Then
        Me.WDRowMarginTop = WDRowMarginTopIsNumeric
        Call MustBeNumeric
    End If
End Sub

'Row Margin Bottom
Private Sub WDRowMarginBottom_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDRowMarginBottomTopIsNumeric = Me.WDRowMarginBottom
End Sub
Private Sub WDRowMarginBottom_Change()
    If IsNumeric(Me.WDRowMarginBottom) = False And Me.WDRowMarginBottom <> "" Then
        Me.WDRowMarginBottom = WDRowMarginBottomIsNumeric
        Call MustBeNumeric
    End If
End Sub


'Row Border Radious
Private Sub WDRowBorderRadious_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDRowBorderRadiousIsNumeric = Me.WDRowBorderRadious
End Sub
Private Sub WDRowBorderRadious_Change()
    If IsNumeric(Me.WDRowBorderRadious) = False And Me.WDRowBorderRadious <> "" Then
        Me.WDRowBorderRadious = WDRowBorderRadiousIsNumeric
        Call MustBeNumeric
    End If
End Sub


'Row Border Thickness
Private Sub WDRowBorderThickness_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    WDRowBorderThicknessIsNumeric = Me.WDRowBorderThickness
End Sub
Private Sub WDRowBorderThickness_Change()
    If IsNumeric(Me.WDRowBorderThickness) = False And Me.WDRowBorderThickness <> "" Then
        Me.WDRowBorderThickness = WDRowBorderThicknessIsNumeric
        Call MustBeNumeric
    End If
End Sub



Private Sub WDMartricsTextSize1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize1IsNumeric = Me.WDMartricsTextSize1
End Sub

Private Sub WDMartricsTextSize1_Change()
If IsNumeric(Me.WDMartricsTextSize1) = False And Me.WDMartricsTextSize1 <> "" Then
Me.WDMartricsTextSize1 = WDMartricsTextSize1IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize2_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize2IsNumeric = Me.WDMartricsTextSize2
End Sub

Private Sub WDMartricsTextSize2_Change()
If IsNumeric(Me.WDMartricsTextSize2) = False And Me.WDMartricsTextSize2 <> "" Then
Me.WDMartricsTextSize2 = WDMartricsTextSize2IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize3_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize3IsNumeric = Me.WDMartricsTextSize3
End Sub

Private Sub WDMartricsTextSize3_Change()
If IsNumeric(Me.WDMartricsTextSize3) = False And Me.WDMartricsTextSize3 <> "" Then
Me.WDMartricsTextSize3 = WDMartricsTextSize3IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize4_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize4IsNumeric = Me.WDMartricsTextSize4
End Sub

Private Sub WDMartricsTextSize4_Change()
If IsNumeric(Me.WDMartricsTextSize4) = False And Me.WDMartricsTextSize4 <> "" Then
Me.WDMartricsTextSize4 = WDMartricsTextSize4IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize5_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize5IsNumeric = Me.WDMartricsTextSize5
End Sub

Private Sub WDMartricsTextSize5_Change()
If IsNumeric(Me.WDMartricsTextSize5) = False And Me.WDMartricsTextSize5 <> "" Then
Me.WDMartricsTextSize5 = WDMartricsTextSize5IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize6_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize6IsNumeric = Me.WDMartricsTextSize6
End Sub

Private Sub WDMartricsTextSize6_Change()
If IsNumeric(Me.WDMartricsTextSize6) = False And Me.WDMartricsTextSize6 <> "" Then
Me.WDMartricsTextSize6 = WDMartricsTextSize6IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize7_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize7IsNumeric = Me.WDMartricsTextSize7
End Sub

Private Sub WDMartricsTextSize7_Change()
If IsNumeric(Me.WDMartricsTextSize7) = False And Me.WDMartricsTextSize7 <> "" Then
Me.WDMartricsTextSize7 = WDMartricsTextSize7IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize8_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize8IsNumeric = Me.WDMartricsTextSize8
End Sub

Private Sub WDMartricsTextSize8_Change()
If IsNumeric(Me.WDMartricsTextSize8) = False And Me.WDMartricsTextSize8 <> "" Then
Me.WDMartricsTextSize8 = WDMartricsTextSize8IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize9_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize9IsNumeric = Me.WDMartricsTextSize9
End Sub

Private Sub WDMartricsTextSize9_Change()
If IsNumeric(Me.WDMartricsTextSize9) = False And Me.WDMartricsTextSize9 <> "" Then
Me.WDMartricsTextSize9 = WDMartricsTextSize9IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsTextSize10_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsTextSize10IsNumeric = Me.WDMartricsTextSize10
End Sub

Private Sub WDMartricsTextSize10_Change()
If IsNumeric(Me.WDMartricsTextSize10) = False And Me.WDMartricsTextSize10 <> "" Then
Me.WDMartricsTextSize10 = WDMartricsTextSize10IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth1IsNumeric = Me.WDMartricsWidth1
End Sub

Private Sub WDMartricsWidth1_Change()
If IsNumeric(Me.WDMartricsWidth1) = False And Me.WDMartricsWidth1 <> "" Then
Me.WDMartricsWidth1 = WDMartricsWidth1IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth2_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth2IsNumeric = Me.WDMartricsWidth2
End Sub

Private Sub WDMartricsWidth2_Change()
If IsNumeric(Me.WDMartricsWidth2) = False And Me.WDMartricsWidth2 <> "" Then
Me.WDMartricsWidth2 = WDMartricsWidth2IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth3_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth3IsNumeric = Me.WDMartricsWidth3
End Sub

Private Sub WDMartricsWidth3_Change()
If IsNumeric(Me.WDMartricsWidth3) = False And Me.WDMartricsWidth3 <> "" Then
Me.WDMartricsWidth3 = WDMartricsWidth3IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth4_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth4IsNumeric = Me.WDMartricsWidth4
End Sub

Private Sub WDMartricsWidth4_Change()
If IsNumeric(Me.WDMartricsWidth4) = False And Me.WDMartricsWidth4 <> "" Then
Me.WDMartricsWidth4 = WDMartricsWidth4IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth5_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth5IsNumeric = Me.WDMartricsWidth5
End Sub

Private Sub WDMartricsWidth5_Change()
If IsNumeric(Me.WDMartricsWidth5) = False And Me.WDMartricsWidth5 <> "" Then
Me.WDMartricsWidth5 = WDMartricsWidth5IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth6_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth6IsNumeric = Me.WDMartricsWidth6
End Sub

Private Sub WDMartricsWidth6_Change()
If IsNumeric(Me.WDMartricsWidth6) = False And Me.WDMartricsWidth6 <> "" Then
Me.WDMartricsWidth6 = WDMartricsWidth6IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth7_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth7IsNumeric = Me.WDMartricsWidth7
End Sub

Private Sub WDMartricsWidth7_Change()
If IsNumeric(Me.WDMartricsWidth7) = False And Me.WDMartricsWidth7 <> "" Then
Me.WDMartricsWidth7 = WDMartricsWidth7IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth8_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth8IsNumeric = Me.WDMartricsWidth8
End Sub

Private Sub WDMartricsWidth8_Change()
If IsNumeric(Me.WDMartricsWidth8) = False And Me.WDMartricsWidth8 <> "" Then
Me.WDMartricsWidth8 = WDMartricsWidth8IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth9_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth9IsNumeric = Me.WDMartricsWidth9
End Sub

Private Sub WDMartricsWidth9_Change()
If IsNumeric(Me.WDMartricsWidth9) = False And Me.WDMartricsWidth9 <> "" Then
Me.WDMartricsWidth9 = WDMartricsWidth9IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsWidth10_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsWidth10IsNumeric = Me.WDMartricsWidth10
End Sub

Private Sub WDMartricsWidth10_Change()
If IsNumeric(Me.WDMartricsWidth10) = False And Me.WDMartricsWidth10 <> "" Then
Me.WDMartricsWidth10 = WDMartricsWidth10IsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsColumnPaddingLeft_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsColumnPaddingLeftIsNumeric = Me.WDMartricsColumnPaddingLeft
End Sub

Private Sub WDMartricsColumnPaddingLeft_Change()
If IsNumeric(Me.WDMartricsColumnPaddingLeft) = False And Me.WDMartricsColumnPaddingLeft <> "" Then
Me.WDMartricsColumnPaddingLeft = WDMartricsColumnPaddingLeftIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsColumnPaddingRight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsColumnPaddingRightIsNumeric = Me.WDMartricsColumnPaddingRight
End Sub

Private Sub WDMartricsColumnPaddingRight_Change()
If IsNumeric(Me.WDMartricsColumnPaddingRight) = False And Me.WDMartricsColumnPaddingRight <> "" Then
Me.WDMartricsColumnPaddingRight = WDMartricsColumnPaddingRightIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsColumnMarginLeft_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsColumnMarginLeftIsNumeric = Me.WDMartricsColumnMarginLeft
End Sub

Private Sub WDMartricsColumnMarginLeft_Change()
If IsNumeric(Me.WDMartricsColumnMarginLeft) = False And Me.WDMartricsColumnMarginLeft <> "" Then
Me.WDMartricsColumnMarginLeft = WDMartricsColumnMarginLeftIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsColumnMarginRight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsColumnMarginRightIsNumeric = Me.WDMartricsColumnMarginRight
End Sub

Private Sub WDMartricsColumnMarginRight_Change()
If IsNumeric(Me.WDMartricsColumnMarginRight) = False And Me.WDMartricsColumnMarginRight <> "" Then
Me.WDMartricsColumnMarginRight = WDMartricsColumnMarginRightIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsColumnHeight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsColumnHeightIsNumeric = Me.WDMartricsColumnHeight
End Sub

Private Sub WDMartricsColumnHeight_Change()
If IsNumeric(Me.WDMartricsColumnHeight) = False And Me.WDMartricsColumnHeight <> "" Then
Me.WDMartricsColumnHeight = WDMartricsColumnHeightIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDMartricsColumnWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDMartricsColumnWidthIsNumeric = Me.WDMartricsColumnWidth
End Sub

Private Sub WDMartricsColumnWidth_Change()
If IsNumeric(Me.WDMartricsColumnWidth) = False And Me.WDMartricsColumnWidth <> "" Then
Me.WDMartricsColumnWidth = WDMartricsColumnWidthIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTableTextSize_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTableTextSizeIsNumeric = Me.WDColumnTableTextSize
End Sub

Private Sub WDColumnTableTextSize_Change()
If IsNumeric(Me.WDColumnTableTextSize) = False And Me.WDColumnTableTextSize <> "" Then
Me.WDColumnTableTextSize = WDColumnTableTextSizeIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTablePaddingLeft_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTablePaddingLeftIsNumeric = Me.WDColumnTablePaddingLeft
End Sub

Private Sub WDColumnTablePaddingLeft_Change()
If IsNumeric(Me.WDColumnTablePaddingLeft) = False And Me.WDColumnTablePaddingLeft <> "" Then
Me.WDColumnTablePaddingLeft = WDColumnTablePaddingLeftIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTablePaddingRight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTablePaddingRightIsNumeric = Me.WDColumnTablePaddingRight
End Sub

Private Sub WDColumnTablePaddingRight_Change()
If IsNumeric(Me.WDColumnTablePaddingRight) = False And Me.WDColumnTablePaddingRight <> "" Then
Me.WDColumnTablePaddingRight = WDColumnTablePaddingRightIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTableMarginLeft_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTableMarginLeftIsNumeric = Me.WDColumnTableMarginLeft
End Sub

Private Sub WDColumnTableMarginLeft_Change()
If IsNumeric(Me.WDColumnTableMarginLeft) = False And Me.WDColumnTableMarginLeft <> "" Then
Me.WDColumnTableMarginLeft = WDColumnTableMarginLeftIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTableMarginRight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTableMarginRightIsNumeric = Me.WDColumnTableMarginRight
End Sub

Private Sub WDColumnTableMarginRight_Change()
If IsNumeric(Me.WDColumnTableMarginRight) = False And Me.WDColumnTableMarginRight <> "" Then
Me.WDColumnTableMarginRight = WDColumnTableMarginRightIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTableWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTableWidthIsNumeric = Me.WDColumnTableWidth
End Sub

Private Sub WDColumnTableWidth_Change()
If IsNumeric(Me.WDColumnTableWidth) = False And Me.WDColumnTableWidth <> "" Then
Me.WDColumnTableWidth = WDColumnTableWidthIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTextHeight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTextHeightIsNumeric = Me.WDColumnTextHeight
End Sub

Private Sub WDColumnTextHeight_Change()
If IsNumeric(Me.WDColumnTextHeight) = False And Me.WDColumnTextHeight <> "" Then
Me.WDColumnTextHeight = WDColumnTextHeightIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDColumnTextWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDColumnTextWidthIsNumeric = Me.WDColumnTextWidth
End Sub

Private Sub WDColumnTextWidth_Change()
If IsNumeric(Me.WDColumnTextWidth) = False And Me.WDColumnTextWidth <> "" Then
Me.WDColumnTextWidth = WDColumnTextWidthIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartTitleTextSize_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartTitleTextSizeIsNumeric = Me.WDChartTitleTextSize
End Sub

Private Sub WDChartTitleTextSize_Change()
If IsNumeric(Me.WDChartTitleTextSize) = False And Me.WDChartTitleTextSize <> "" Then
Me.WDChartTitleTextSize = WDChartTitleTextSizeIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartSubtitleTextSize_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartSubtitleTextSizeIsNumeric = Me.WDChartSubtitleTextSize
End Sub

Private Sub WDChartSubtitleTextSize_Change()
If IsNumeric(Me.WDChartSubtitleTextSize) = False And Me.WDChartSubtitleTextSize <> "" Then
Me.WDChartSubtitleTextSize = WDChartSubtitleTextSizeIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartHeight_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartHeightIsNumeric = Me.WDChartHeight
End Sub

Private Sub WDChartHeight_Change()
If IsNumeric(Me.WDChartHeight) = False And Me.WDChartHeight <> "" Then
Me.WDChartHeight = WDChartHeightIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartWidth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartWidthIsNumeric = Me.WDChartWidth
End Sub

Private Sub WDChartWidth_Change()
If IsNumeric(Me.WDChartWidth) = False And Me.WDChartWidth <> "" Then
Me.WDChartWidth = WDChartWidthIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartAnimationLenght_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartAnimationLenghtIsNumeric = Me.WDChartAnimationLenght
End Sub

Private Sub WDChartAnimationLenght_Change()
If IsNumeric(Me.WDChartAnimationLenght) = False And Me.WDChartAnimationLenght <> "" Then
Me.WDChartAnimationLenght = WDChartAnimationLenghtIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartTooltipTextSize_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartTooltipTextSizeIsNumeric = Me.WDChartTooltipTextSize
End Sub

Private Sub WDChartTooltipTextSize_Change()
If IsNumeric(Me.WDChartTooltipTextSize) = False And Me.WDChartTooltipTextSize <> "" Then
Me.WDChartTooltipTextSize = WDChartTooltipTextSizeIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartLegendTextSize_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartLegendTextSizeIsNumeric = Me.WDChartLegendTextSize
End Sub

Private Sub WDChartLegendTextSize_Change()
If IsNumeric(Me.WDChartLegendTextSize) = False And Me.WDChartLegendTextSize <> "" Then
Me.WDChartLegendTextSize = WDChartLegendTextSizeIsNumeric
Call MustBeNumeric
End If
End Sub

Private Sub WDChartTooltipBorderRadious_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
WDChartTooltipBorderRadiousIsNumeric = Me.WDChartTooltipBorderRadious
End Sub

Private Sub WDChartTooltipBorderRadious_Change()
If IsNumeric(Me.WDChartTooltipBorderRadious) = False And Me.WDChartTooltipBorderRadious <> "" Then
Me.WDChartTooltipBorderRadious = WDChartTooltipBorderRadiousIsNumeric
Call MustBeNumeric
End If
End Sub

'List_Box_Load_Values################################################################################################################################################
Sub ListBoxesLoad()
    Dim WD As WebDev
    Set WD = New WebDev
    
    With WDChartTitleFont
        .AddItem "arial"
        .AddItem "tahoma"
        .AddItem "verdana"
        .AddItem "Calibri"
        .AddItem "Optima"
        .AddItem "Candara"
        .AddItem "Verdana"
        .AddItem "Geneva"
        .AddItem "sans-serif"
    End With

    With WDChartTooltipFont
        .AddItem "arial"
        .AddItem "tahoma"
        .AddItem "verdana"
        .AddItem "Calibri"
        .AddItem "Optima"
        .AddItem "Candara"
        .AddItem "Verdana"
        .AddItem "Geneva"
        .AddItem "sans-serif"
    End With
    
    With WDChartSubtitleFont
        .AddItem "arial"
        .AddItem "tahoma"
        .AddItem "verdana"
        .AddItem "Calibri"
        .AddItem "Optima"
        .AddItem "Candara"
        .AddItem "Verdana"
        .AddItem "Geneva"
        .AddItem "sans-serif"
    End With
    
    With WDChartLegendFont
        .AddItem "arial"
        .AddItem "tahoma"
        .AddItem "verdana"
        .AddItem "Calibri"
        .AddItem "Optima"
        .AddItem "Candara"
        .AddItem "Verdana"
        .AddItem "Geneva"
        .AddItem "sans-serif"
    End With
    
    With WDColumnTableFont
        .AddItem "arial"
        .AddItem "tahoma"
        .AddItem "verdana"
        .AddItem "Calibri"
        .AddItem "Optima"
        .AddItem "Candara"
        .AddItem "Verdana"
        .AddItem "Geneva"
        .AddItem "sans-serif"
    End With
    
    With WDChartTitleFontWeight
        .AddItem "lighter"
        .AddItem "normal"
        .AddItem "bold"
        .AddItem "bolder"
    End With
    
    With WDChartSubtitleFontWeight
        .AddItem "lighter"
        .AddItem "normal"
        .AddItem "bold"
        .AddItem "bolder"
    End With
    
    With WDChartTitleFontStyle
        .AddItem "normal"
        .AddItem "italic"
        .AddItem "oblique"
    End With
    
    With WDChartSubtitleFontStyle
        .AddItem "normal"
        .AddItem "italic"
        .AddItem "oblique"
    End With
    
    With WDChartTitleVerticleAlign
        .AddItem "top"
        .AddItem "center"
        .AddItem "bottom"
    End With
    
    With WDChartSubtitleVerticleAlign
        .AddItem "top"
        .AddItem "center"
        .AddItem "bottom"
    End With
    
    With WDChartLegendVerticleAlign
        .AddItem "top"
        .AddItem "center"
        .AddItem "bottom"
    End With
    
    With WDChartTitleHorizontalAlign
        .AddItem "left"
        .AddItem "right"
        .AddItem "center"
    End With
    
    With WDChartSubtitleHorizontalAlign
        .AddItem "left"
        .AddItem "right"
        .AddItem "center"
    End With
    
    With WDChartLegendHorizontalAlign
        .AddItem "left"
        .AddItem "right"
        .AddItem "center"
    End With
    
    With WDRowTextAlign
        .AddItem "left"
        .AddItem "right"
        .AddItem "center"
    End With
    
    With WDChartTheme
        .AddItem "light1"
        .AddItem "light2"
        .AddItem "dark1"
        .AddItem "dark2"
    End With
    
    With WDChartColorsPredefigned
        .AddItem "colorSet1"
        .AddItem "colorSet2"
        .AddItem "colorSet3"
    End With
    
    With WDChartLayerChartType1
        .AddItem "line"
        .AddItem "column"
        .AddItem "bar"
        .AddItem "area"
        .AddItem "spline"
        .AddItem "splineArea"
        .AddItem "stepLine"
        .AddItem "scatter"
        .AddItem "bubble"
        .AddItem "stackedColumn"
        .AddItem "stackedBar"
        .AddItem "stackedArea"
        .AddItem "stackedColumn100"
        .AddItem "stackedBar100"
        .AddItem "stackedArea100"
        .AddItem "pie"
        .AddItem "doughnut"
    End With
    
    With WDChartLayerChartType2
        .AddItem "line"
        .AddItem "column"
        .AddItem "bar"
        .AddItem "area"
        .AddItem "spline"
        .AddItem "splineArea"
        .AddItem "stepLine"
        .AddItem "scatter"
        .AddItem "bubble"
        .AddItem "stackedColumn"
        .AddItem "stackedBar"
        .AddItem "stackedArea"
        .AddItem "stackedColumn100"
        .AddItem "stackedBar100"
        .AddItem "stackedArea100"
        .AddItem "pie"
        .AddItem "doughnut"
    End With
    
    With WDChartLayerChartType3
        .AddItem "line"
        .AddItem "column"
        .AddItem "bar"
        .AddItem "area"
        .AddItem "spline"
        .AddItem "splineArea"
        .AddItem "stepLine"
        .AddItem "scatter"
        .AddItem "bubble"
        .AddItem "stackedColumn"
        .AddItem "stackedBar"
        .AddItem "stackedArea"
        .AddItem "stackedColumn100"
        .AddItem "stackedBar100"
        .AddItem "stackedArea100"
        .AddItem "pie"
        .AddItem "doughnut"
    End With
    
    With WDChartLayerChartType4
        .AddItem "line"
        .AddItem "column"
        .AddItem "bar"
        .AddItem "area"
        .AddItem "spline"
        .AddItem "splineArea"
        .AddItem "stepLine"
        .AddItem "scatter"
        .AddItem "bubble"
        .AddItem "stackedColumn"
        .AddItem "stackedBar"
        .AddItem "stackedArea"
        .AddItem "stackedColumn100"
        .AddItem "stackedBar100"
        .AddItem "stackedArea100"
        .AddItem "pie"
        .AddItem "doughnut"
    End With
    
    With WDChartLayerChartType5
        .AddItem "line"
        .AddItem "column"
        .AddItem "bar"
        .AddItem "area"
        .AddItem "spline"
        .AddItem "splineArea"
        .AddItem "stepLine"
        .AddItem "scatter"
        .AddItem "bubble"
        .AddItem "stackedColumn"
        .AddItem "stackedBar"
        .AddItem "stackedArea"
        .AddItem "stackedColumn100"
        .AddItem "stackedBar100"
        .AddItem "stackedArea100"
        .AddItem "pie"
        .AddItem "doughnut"
    End With
    
    With WDChartLayerAggregation1
        .AddItem "sum"
        .AddItem "average"
        .AddItem "count"
    End With
    
    With WDChartLayerAggregation2
        .AddItem "sum"
        .AddItem "average"
        .AddItem "count"
    End With
    
    With WDChartLayerAggregation3
        .AddItem "sum"
        .AddItem "average"
        .AddItem "count"
    End With
    
    With WDChartLayerAggregation4
        .AddItem "sum"
        .AddItem "average"
        .AddItem "count"
    End With
    
    With WDChartLayerAggregation5
        .AddItem "sum"
        .AddItem "average"
        .AddItem "count"
    End With
    
    With WDColumnTableAggregation
        .AddItem "sum"
        .AddItem "average"
        .AddItem "count"
    End With
    
    With WDChartLayerYAxisType1
        .AddItem "Primary"
        .AddItem "Secondary"
    End With
    
    With WDChartLayerYAxisType2
        .AddItem "Primary"
        .AddItem "Secondary"
    End With
    
    With WDChartLayerYAxisType3
        .AddItem "Primary"
        .AddItem "Secondary"
    End With
    
    With WDChartLayerYAxisType4
        .AddItem "Primary"
        .AddItem "Secondary"
    End With
    
    With WDChartLayerYAxisType5
        .AddItem "Primary"
        .AddItem "Secondary"
    End With
    
    With WDChartLayerShowInLegend1
        .AddItem "true"
        .AddItem "false"
    End With
    
    With WDChartLayerShowInLegend2
        .AddItem "true"
        .AddItem "false"
    End With
    
    With WDChartLayerShowInLegend3
        .AddItem "true"
        .AddItem "false"
    End With
    
    With WDChartLayerShowInLegend4
        .AddItem "true"
        .AddItem "false"
    End With
    
    With WDChartLayerShowInLegend5
        .AddItem "true"
        .AddItem "false"
    End With
    
    With WDColumnTableHeader
        .AddItem "yes"
        .AddItem "no"
    End With
    
    With Me.WDRowStyle
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With

    With Me.WDColumnTableStyle
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle1
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle2
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle3
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle4
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle5
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle6
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle7
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle8
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle9
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
    With Me.WDMartricsStyle10
        .AddItem WD.BScontainerDanger
        .AddItem WD.BScontainerInfo
        .AddItem WD.BScontainerWarning
        .AddItem WD.BScontainerSecondary
        .AddItem WD.BStextWhite
        .AddItem WD.BStextPrimary
        .AddItem WD.BStextDark
        .AddItem WD.BStextDanger
        .AddItem WD.BStextInfo
        .AddItem WD.BStextWarning
        .AddItem WD.BStextSecondary
        .AddItem WD.BStablePlane
        .AddItem WD.BStableStripped
        .AddItem WD.BStableBorered
        .AddItem WD.BStableHover
        .AddItem WD.BStableDark
        .AddItem WD.BStableDarkStripped
        .AddItem WD.BStableDarkHover
        .AddItem WD.BStableborderless
        .AddItem WD.BStabletheadDark
        .AddItem WD.BStabletheadLight
        .AddItem WD.BStableSmall
        .AddItem WD.BStableColorPrimary
        .AddItem WD.BStableColorSuccess
        .AddItem WD.BStableColorDanger
        .AddItem WD.BStableColorInfo
        .AddItem WD.BStableColorWarning
        .AddItem WD.BStableColorActive
        .AddItem WD.BStableColorSecondary
        .AddItem WD.BStableColorLight
        .AddItem WD.BStableColorDark
    End With
    
End Sub


Sub WriteHTMLDocument()
    Dim WD As WebDev
    Set WD = New WebDev

    HTML = ""
    HTML = HTML & "<html>" & vbNewLine
    HTML = HTML & "<head>" & vbNewLine
    
    
    
'Title######################################################################################
    HTML = HTML & "<title>" & Me.WDdata(Me.WDPage, 0, 0, 1) & "</title>" & vbNewLine
        
        
        
'CSS########################################################################################
       '<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css'>
    
    HTML = HTML & "<link rel='stylesheet' href='" & WD.Bootstrap4CSS & "' href='styles.css'>" & vbNewLine
    HTML = HTML & "<link rel='stylesheet' href='" & "https://objective/id:BQ13784116" & "' href='styles.css'>" & vbNewLine
    HTML = HTML & "<link rel='stylesheet' href='" & "https://objective/id:BQ13784109" & "' href='styles.css'>" & vbNewLine
    
    
'User CSS###################################################################################
    HTML = HTML & "<link rel='stylesheet' href='" & WDdata(Me.WDPage, 0, 0, 7) & "' href='styles.css'>" & vbNewLine

    
    

'Custom CSS#################################################################################
    HTML = HTML & "<style>" & vbNewLine
    
        HTML = HTML & ".body{" & vbNewLine
            HTML = HTML & "background-color:" & WDdata(Me.WDPage, 0, 0, 3) & ";" & vbNewLine
            HTML = HTML & "background-image: url('" & WDdata(Me.WDPage, 0, 0, 4) & "');" & vbNewLine
            HTML = HTML & "background-position: center;" & vbNewLine
            HTML = HTML & "background-repeat: no-repeat;" & vbNewLine
            HTML = HTML & "background-size: cover;" & vbNewLine
        HTML = HTML & "}" & vbNewLine
        
        HTML = HTML & ".wrapper{" & vbNewLine
            HTML = HTML & "display: grid;" & vbNewLine
            HTML = HTML & "grid-gap: 0px;" & vbNewLine
        HTML = HTML & "}" & vbNewLine
        
    HTML = HTML & "</style>" & vbNewLine

    
    
'Body#######################################################################################
    HTML = HTML & "</head>" & vbNewLine
    HTML = HTML & "<body class='body'>" & vbNewLine
    
    
    For R = 1 To AD2
        
        If WDdata(Me.WDPage, R, 0, 14) = False Then Exit For
        
        'Column Counter
        For i = 1 To AD3
            If WDdata(Me.WDPage, R, i, 195) = False Then Exit For
        Next i
        i = i - 1
        
        HTML = HTML & "<div style='grid-template-columns: repeat(auto-fit, minmax(" & 100 / i & "%, " & 100 / i & "%));' class='wrapper'>" & vbNewLine
        
        


        For c = 1 To AD3
            If WDdata(Me.WDPage, R, c, 195) = False Then Exit For
            
                'Text
                If WDdata(Me.WDPage, R, c, 191) = True Then
                    HTML = HTML & WDdata(Me.WDPage, R, c, 1) & vbNewLine
                End If
                
                'Metrics
                If WDdata(Me.WDPage, R, c, 192) = True Then
                    
                End If
                
                'Charts
                If WDdata(Me.WDPage, R, c, 193) = True Then
                
                
                End If
                
                'Table
                If WDdata(Me.WDPage, R, c, 194) = True Then
                
                
                End If
                
        Next c
        HTML = HTML & "</div>" & vbNewLine
    Next R
    
    
    
    
'JS#########################################################################################
    HTML = HTML & "<script src='" & WD.Bootstrap4JS & "'></script>" & vbNewLine
    HTML = HTML & "<script src='" & WD.popperJS & "'></script>" & vbNewLine
    HTML = HTML & "<script src='" & WD.jQueryJS & "'></script>" & vbNewLine
    HTML = HTML & "<script src='" & "https://objective/id:BQ13784116" & "'></script>" & vbNewLine
    HTML = HTML & "<script src='https://canvasjs.com/assets/script/canvasjs.min.js'></script>" & vbNewLine
    
    
'User JS####################################################################################
    HTML = HTML & "<script src='" & WDdata(Me.WDPage, 0, 0, 5) & "'></script>" & vbNewLine
    HTML = HTML & "<script src='" & WDdata(Me.WDPage, 0, 0, 6) & "'></script>" & vbNewLine
    

    
'Custom JS##################################################################################
    HTML = HTML & "<script>" & vbNewLine
        HTML = HTML & WDdata(Me.WDPage, 0, 0, 18) & vbNewLine
    HTML = HTML & "</script>" & vbNewLine
    
    HTML = HTML & "</body>" & vbNewLine
    HTML = HTML & "</html>" & vbNewLine
    
    
    
'Canvas JS (Charts)#########################################################################




    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\" & Environ("UserName") & "\Desktop\" & Me.WDdata(Me.WDPage, 0, 0, 1) & ".html", True)
    a.WriteLine HTML
    a.Close
End Sub
