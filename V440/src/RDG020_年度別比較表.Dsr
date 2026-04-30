VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDG020_”N“x•Ê”äŠr•\ 
   Caption         =   "”N“x•Ê”äŠr•\"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   Icon            =   "RDG020_”N“x•Ê”äŠr•\.dsx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
   WindowState     =   2  'Å‘å‰»
   _ExtentX        =   19685
   _ExtentY        =   12594
   SectionData     =   "RDG020_”N“x•Ê”äŠr•\.dsx":0ECA
End
Attribute VB_Name = "RDG020_”N“x•Ê”äŠr•\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "”N“x•Ê”äŠr•\"
'
Dim wRs As ADODB.Recordset

Dim wstr As String, wWhere As String

Dim w”Ô† As String, wsTbl As String
Dim w•ª•ê As Integer
Dim wML As Integer

Dim G1_Yusi1(1) As Double, G1_Yusi2(1) As Double, G1_Yusi3(1) As Double
Dim G1_TYusi1(1) As Double, G1_TYusi2(1) As Double, G1_TYusi3(1) As Double
Dim G1_Gankin1(1) As Double, G1_Gankin2(1) As Double, G1_Gankin3(1) As Double
Dim G1_Risoku1(1) As Double, G1_Risoku2(1) As Double, G1_Risoku3(1) As Double
Dim G1_Hensai1(1) As Double, G1_Hensai2(1) As Double, G1_Hensai3(1) As Double
Dim G1_Yusizan1(1) As Double, G1_Yusizan2(1) As Double, G1_Yusizan3(1) As Double

Dim G2_Yusi1(1) As Double, G2_Yusi2(1) As Double, G2_Yusi3(1) As Double
Dim G2_TYusi1(1) As Double, G2_TYusi2(1) As Double, G2_TYusi3(1) As Double
Dim G2_Gankin1(1) As Double, G2_Gankin2(1) As Double, G2_Gankin3(1) As Double
Dim G2_Risoku1(1) As Double, G2_Risoku2(1) As Double, G2_Risoku3(1) As Double
Dim G2_Hensai1(1) As Double, G2_Hensai2(1) As Double, G2_Hensai3(1) As Double
Dim G2_Yusizan1(1) As Double, G2_Yusizan2(1) As Double, G2_Yusizan3(1) As Double

Dim GT_Yusi1(1) As Double, GT_Yusi2(1) As Double, GT_Yusi3(1) As Double
Dim GT_TYusi1(1) As Double, GT_TYusi2(1) As Double, GT_TYusi3(1) As Double
Dim GT_Gankin1(1) As Double, GT_Gankin2(1) As Double, GT_Gankin3(1) As Double
Dim GT_Risoku1(1) As Double, GT_Risoku2(1) As Double, GT_Risoku3(1) As Double
Dim GT_Hensai1(1) As Double, GT_Hensai2(1) As Double, GT_Hensai3(1) As Double
Dim GT_Yusizan1(1) As Double, GT_Yusizan2(1) As Double, GT_Yusizan3(1) As Double
'
Dim w„ˆÚ•\ƒ^ƒCƒgƒ‹ As MAA910_„ˆÚ•\ƒ^ƒCƒgƒ‹

'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim j As Integer, wIndex As Integer
    Dim ws01 As String, wsS As String
    Dim ws_Ginko As String, wOrder As String
    Dim FLG_Order As Boolean
    
    Dim wŠJn”NŒ“ú As Date, wdate As Date
    Dim w„ˆÚ•\‹æ•ª As String
    Dim wsNengetu As String
'
    On Error GoTo ActiveReport_ReportStart_ERR
'
    '----------------------------------------------------------------
    '                         ** ‰Šúİ’è **
    '----------------------------------------------------------------
    'Connection
    Me.DataControl1.Connection = GDb
   
    '—p†ƒZƒbƒg
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = ddOLandscape
    
    wML = 12
'
    '----------------------------------------------------------------
    '                           ** Œ©o‚µ **
    '----------------------------------------------------------------
    o—Í“ú = Now
    Šé‹Æ–¼ = GCoName
'
    'H00_ÅIÀÑ”NŒ = ""
    'If GƒRƒ“ƒgƒ[ƒ‹.ÅIÀÑ”NŒ > CDate("2001/01/01") Then
    '    H00_ÅIÀÑ”NŒ = Format(GƒRƒ“ƒgƒ[ƒ‹.ÅIÀÑ”NŒ, Gfmt”NŒ)
    'End If

    Me.PageHeader.Controls("L_‘OŠú") = CInt(GRpt.ƒeƒLƒXƒg_01) - 1 & "”N“xi‘OŠúj"
    Me.PageHeader.Controls("L_“–Šú") = CInt(GRpt.ƒeƒLƒXƒg_01) & "”N“xi“–Šúj"
    Me.PageHeader.Controls("L_—ˆŠú") = CInt(GRpt.ƒeƒLƒXƒg_01) + 1 & "”N“xi—ˆŠúj"

    H00_‹à—ZƒŠƒXƒgƒ‰”Ô† = GRpt.‹à—Z
    
    GRpt.„ˆÚ = "”NŸ"
    GRpt.ƒeƒLƒXƒg_01 = CInt(GRpt.ƒeƒLƒXƒg_01) - 1
    
    w•ª•ê = 1000
    L_’PˆÊ.Caption = "iç‰~’PˆÊj"

    Me.Detail.Height = 0
    
    If GRpt.WŒv = "‹à—˜í•ÊŒv" Then
    '‹à—˜í•Ê
        Me.PageHeader.Controls("L_WŒv‹æ•ª") = "‹à—˜í•Ê"
        Me.GroupFooter1.Controls("G11_Œv–¼") = "•Ï“®Œv"
        Me.GroupFooter1.Controls("G12_Œv–¼") = "ŒÅ’èŒv"
        Me.GroupFooter2.Controls("G21_Œv–¼") = "•Ï“®Œv"
        Me.GroupFooter2.Controls("G22_Œv–¼") = "ŒÅ’èŒv"
        Me.ReportFooter.Controls("G91_Œv–¼") = "•Ï“®Œv"
        Me.ReportFooter.Controls("G92_Œv–¼") = "ŒÅ’èŒv"
    ElseIf GRpt.WŒv = "’S•Û‹æ•ªŒv" Then
    '’S•Û‹æ•ª
        Me.PageHeader.Controls("L_WŒv‹æ•ª") = "’S•Û‹æ•ª"
        Me.GroupFooter1.Controls("G11_Œv–¼") = "—L’SŒv"
        Me.GroupFooter1.Controls("G12_Œv–¼") = "–³’SŒv"
        Me.GroupFooter2.Controls("G21_Œv–¼") = "—L’SŒv"
        Me.GroupFooter2.Controls("G22_Œv–¼") = "–³’SŒv"
        Me.ReportFooter.Controls("G91_Œv–¼") = "—L’SŒv"
        Me.ReportFooter.Controls("G92_Œv–¼") = "–³’SŒv"
    ElseIf GRpt.WŒv = "’·’Z‹æ•ªŒv" Then
    '’·’Z‹æ•ª
        Me.PageHeader.Controls("L_WŒv‹æ•ª") = "’·’Z‹æ•ª"
        Me.GroupFooter1.Controls("G11_Œv–¼") = "’·ŠúŒv"
        Me.GroupFooter1.Controls("G12_Œv–¼") = "’ZŠúŒv"
        Me.GroupFooter2.Controls("G21_Œv–¼") = "’·ŠúŒv"
        Me.GroupFooter2.Controls("G22_Œv–¼") = "’ZŠúŒv"
        Me.ReportFooter.Controls("G91_Œv–¼") = "’·ŠúŒv"
        Me.ReportFooter.Controls("G92_Œv–¼") = "’ZŠúŒv"
    End If
'
    If GRpt.WŒv = "WŒv•\¦‚µ‚È‚¢" Then
        Me.PageHeader.Controls("L_WŒv‹æ•ª").Visible = False
        Me.ReportFooter.Controls("LineRF").Y1 = 200
        Me.ReportFooter.Controls("LineRF").Y2 = 200
        
        Me.GroupFooter1.Height = 210
        Me.GroupFooter2.Height = 210
        Me.ReportFooter.Height = 205
    End If
'
    '----------------------------------------------------------------
    '                       ** ˆóš—pƒtƒ@ƒCƒ‹ì¬ **
    '----------------------------------------------------------------
    For j = 0 To 1
        G1_Yusi1(j) = 0: G1_Yusi2(j) = 0: G1_Yusi3(j) = 0
        G1_TYusi1(j) = 0: G1_TYusi2(j) = 0: G1_TYusi3(j) = 0
        G1_Gankin1(j) = 0: G1_Gankin2(j) = 0: G1_Gankin3(j) = 0
        G1_Risoku1(j) = 0: G1_Risoku2(j) = 0: G1_Risoku3(j) = 0
        G1_Hensai1(j) = 0: G1_Hensai2(j) = 0: G1_Hensai3(j) = 0
        G1_Yusizan1(j) = 0: G1_Yusizan2(j) = 0: G1_Yusizan3(j) = 0
        
        G2_Yusi1(j) = 0: G2_Yusi2(j) = 0: G2_Yusi3(j) = 0
        G2_TYusi1(j) = 0: G2_TYusi2(j) = 0: G2_TYusi3(j) = 0
        G2_Gankin1(j) = 0: G2_Gankin2(j) = 0: G2_Gankin3(j) = 0
        G2_Risoku1(j) = 0: G2_Risoku2(j) = 0: G2_Risoku3(j) = 0
        G2_Hensai1(j) = 0: G2_Hensai2(j) = 0: G2_Hensai3(j) = 0
        G2_Yusizan1(j) = 0: G2_Yusizan2(j) = 0: G2_Yusizan3(j) = 0
        
        GT_Yusi1(j) = 0: GT_Yusi2(j) = 0: GT_Yusi3(j) = 0
        GT_TYusi1(j) = 0: GT_TYusi2(j) = 0: GT_TYusi3(j) = 0
        GT_Gankin1(j) = 0: GT_Gankin2(j) = 0: GT_Gankin3(j) = 0
        GT_Risoku1(j) = 0: GT_Risoku2(j) = 0: GT_Risoku3(j) = 0
        GT_Hensai1(j) = 0: GT_Hensai2(j) = 0: GT_Hensai3(j) = 0
        GT_Yusizan1(j) = 0: GT_Yusizan2(j) = 0: GT_Yusizan3(j) = 0
    Next j
'
    Call MBD020_Ø“ü‹àƒ[ƒNƒe[ƒuƒ‹ì¬("DCIA010_Ø“ü‹àƒ[ƒN")
    Call MRB010_•W€“ü—ÍØ“üc‚•\("DCIA010_Ø“ü‹àƒ[ƒN")       '07/02/18 V180
    Call MRB010_è“ü—ÍØ“üc‚•\("DCIA010_Ø“ü‹àƒ[ƒN")         '07/02/09 V180
'
    'wŠJn”NŒ“ú = C”NŒ“ú.”N“xŠJn”NŒ“ú(GRpt.ƒeƒLƒXƒg_01, "•½¬")
    '2019/01/15 “ú•t“ü—Í‹æ•ª d—l•ÏX
    If GŠî–{î•ñ.“ú•t“ü—Í‹æ•ª = "0" Then
    '˜a—ï
        If Len(GRpt.ƒeƒLƒXƒg_01) <= 2 Then
            wŠJn”NŒ“ú = C”NŒ“ú.”N“xŠJn”NŒ“ú(GRpt.ƒeƒLƒXƒg_01, "•½¬")
        Else
        wŠJn”NŒ“ú = C”NŒ“ú.”N“xŠJn”NŒ“ú(GRpt.ƒeƒLƒXƒg_01, "¼—ï")
        End If
    Else
    '¼—ï
        wŠJn”NŒ“ú = C”NŒ“ú.”N“xŠJn”NŒ“ú(GRpt.ƒeƒLƒXƒg_01, "¼—ï")
    End If
    
    w„ˆÚ•\ƒ^ƒCƒgƒ‹ = MUA010_„ˆÚ•\ƒtƒ@ƒCƒ‹ì¬("", "", wŠJn”NŒ“ú, GRpt.„ˆÚ, wML)
'
    '----------------------------------------------------------------
    '                          ** ‡Œv ŒvZ **
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    '                       ** ƒŒƒR[ƒh@ƒ\[ƒX **
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    '          ** ƒOƒ‹[ƒv E Where E Order‚a‚™ **
    '----------------------------------------------------------------
    'GroupHeader1.DataField = "GrpFld_Ginko"
    'GroupHeader2.DataField = "GrpFld_KShubetu"
    'GroupHeader2.DataField = "GrpFld_Ginko"
    'GroupHeader1.DataField = "GrpFld_KShubetu"
    
    If GRpt.S_í•Ê = "•ª—Ş1" Then
        GroupHeader1.DataField = "GrpFld_KShubetu"
    ElseIf GRpt.S_í•Ê = "•ª—Ş2" Then
        GroupHeader2.DataField = "GrpFld_KShubetu"
    End If
    
    If GRpt.S_•”–å = "•ª—Ş1" Then
        GroupHeader1.DataField = "GrpFld_Bumon"
    ElseIf GRpt.S_•”–å = "•ª—Ş2" Then
        GroupHeader2.DataField = "GrpFld_Bumon"
    End If
    
    If GRpt.S_‹à—Z = "•ª—Ş1" Then
        GroupHeader1.DataField = "GrpFld_Kinyu"
    ElseIf GRpt.S_‹à—Z = "•ª—Ş2" Then
        GroupHeader2.DataField = "GrpFld_Kinyu"
    End If
    
    If GRpt.S_‹âs = "•ª—Ş1" Then
        GroupHeader1.DataField = "GrpFld_Ginko"
    ElseIf GRpt.S_‹âs = "•ª—Ş2" Then
        GroupHeader2.DataField = "GrpFld_Ginko"
    End If

    '’ •[w¦
    wsS = ""
    If GRpt.Ø“ü‹àŠÇ—‹æ•ª = P8.FCDbl(XMXA020_‹æ•ª("Ø“ü‹àŠÇ—‹æ•ª", "ŒˆZ—p")) Then
        wsS = wsS & "’ •[w¦:ŒˆZ—p "
    Else
        wsS = wsS & "’ •[w¦:ŠÇ——p "
    End If
    
    'Œv–¼ƒZƒbƒgAShapeƒJƒ‰[
    If GroupHeader1.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "•ª—Ş1:Ø“ü‹àí•Ê "
        Me.GroupFooter1.Controls("G10_Œv–¼").DataField = "I_Ø“ü‹àí•Ê–¼"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HFFFFC0
    ElseIf GroupHeader1.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "•ª—Ş1:•”–å "
        Me.GroupFooter1.Controls("G10_Œv–¼").DataField = "I_•”–å–¼"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HC0FFC0
    ElseIf GroupHeader1.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "•ª—Ş1:‹à—Z‹@ŠÖ "
        Me.GroupFooter1.Controls("G10_Œv–¼").DataField = "I_‹à—Z‹@ŠÖ–¼"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HE0E0E0
    ElseIf GroupHeader1.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "•ª—Ş1:‹âs "
        Me.GroupFooter1.Controls("G10_Œv–¼").DataField = "I_‹âs–¼"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HC0FFFF
    End If
    
    If GroupHeader2.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "•ª—Ş2:Ø“ü‹àí•Ê "
        Me.GroupFooter2.Controls("G20_Œv–¼").DataField = "I_Ø“ü‹àí•Ê–¼"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HFFFFC0
    ElseIf GroupHeader2.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "•ª—Ş2:•”–å "
        Me.GroupFooter2.Controls("G20_Œv–¼").DataField = "I_•”–å–¼"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HC0FFC0
    ElseIf GroupHeader2.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "•ª—Ş2:‹à—Z‹@ŠÖ "
        Me.GroupFooter2.Controls("G20_Œv–¼").DataField = "I_‹à—Z‹@ŠÖ–¼"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HE0E0E0
    ElseIf GroupHeader2.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "•ª—Ş2:‹âs "
        Me.GroupFooter2.Controls("G20_Œv–¼").DataField = "I_‹âs–¼"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HC0FFFF
    End If
'
    GroupFooter1.Visible = True
    GroupFooter2.Visible = True
    If GroupHeader1.DataField = "" Then
        GroupFooter1.Visible = False
    End If
    If GroupHeader2.DataField = "" Then
        GroupFooter2.Visible = False
    End If
'
    '’ •[w¦
    Me.PageHeader.Controls("L_’ •[w¦").Caption = wsS
'
    '‰üƒy[ƒW
    Me.GroupFooter1.NewPage = GRpt.NewPage1
    Me.GroupFooter2.NewPage = GRpt.NewPage2
    
    '----------------------------------------------------------------
    '                       ** ƒŒƒR[ƒh@ƒ\[ƒX **
    '----------------------------------------------------------------
    wstr = ""
    wstr = wstr & "SELECT "
    'wstr = wstr & " G.‹âs”Ô† As GrpFld_Ginko,"
    'wstr = wstr & "G.‹âs–¼ As I_‹æ•ª–¼1,"
    'wstr = wstr & "G.‹âs–¼ As I_‹æ•ª–¼2,"
    
    'wstr = wstr & " K.Ø“ü‹àí•Ê‹æ•ª As GrpFld_KShubetu,"
    'wstr = wstr & "S.Ø“ü‹àí•Ê–¼ As I_‹æ•ª–¼2,"
    'wstr = wstr & "S.Ø“ü‹àí•Ê–¼ As I_‹æ•ª–¼1,"
    
    'ƒZƒNƒVƒ‡ƒ“GR
    wstr = wstr & "K.‹âs”Ô† As GrpFld_Ginko,"
    wstr = wstr & "G.‹à—Z‹@ŠÖ”Ô† As GrpFld_Kinyu,"
    wstr = wstr & "B.•”–å”Ô† As GrpFld_Bumon,"
    wstr = wstr & "K.Ø“ü‹àí•Ê‹æ•ª As GrpFld_KShubetu,"
    
    wstr = wstr & "G.‹âs–¼ As I_‹âs–¼,"
    wstr = wstr & "G.‹à—Z‹@ŠÖ–¼ As I_‹à—Z‹@ŠÖ–¼,"
    wstr = wstr & "B.•”–å–¼ As I_•”–å–¼,"
    wstr = wstr & "S.Ø“ü‹àí•Ê–¼ As I_Ø“ü‹àí•Ê–¼,"
    
    'WŒv‹æ•ª
    If GRpt.WŒv = "‹à—˜í•ÊŒv" Then
    '‹à—˜í•Ê
        wstr = wstr & " IIF(K.‹à—˜í•Ê = " & P8.FCDbl(XMXA020_‹æ•ª("‹à—˜í•Ê", "•Ï“®‹à—˜")) & ",'•Ï“®','ŒÅ’è') As I_WŒv‹æ•ª," 'As I_‹à—˜í•Ê,"
    ElseIf GRpt.WŒv = "’S•Û‹æ•ªŒv" Then
    '’S•Û‹æ•ª
        wstr = wstr & " IIF(K.—L’S•Ûƒtƒ‰ƒO = " & P8.FCDbl(XMXA020_‹æ•ª("—L’Sƒtƒ‰ƒO", "–³’S•Û")) & ",'–³’S','—L’S') As I_WŒv‹æ•ª," 'As I_’S•Û,"
    ElseIf GRpt.WŒv = "’·’Z‹æ•ªŒv" Then
    '’·’Z‹æ•ª
        wstr = wstr & " IIF(K.’·’Z‹æ•ª = " & P8.FCDbl(XMXA020_‹æ•ª("’·’Z‹æ•ª", "’ZŠúØ“ü‹à")) & ",'’ZŠú','’·Šú') AS I_WŒv‹æ•ª," 'AS I_’·’Z‹æ•ª,"
    Else
        wstr = wstr & "'' AS I_WŒv‹æ•ª,"
    End If
    
    wstr = wstr & "IIF(—Z‘_01<>0 Or Œ³‹à_01<>0 Or —˜‘§_01<>0 Or •ÔÏ_01<>0 Or ‰ğ–ñ_01<>0 Or c‚_01<>0,K.—Z‘‹àŠz,0) AS I_—Z‘‹àŠz,"
    wstr = wstr & "IIF(—Z‘_02<>0 Or Œ³‹à_02<>0 Or —˜‘§_02<>0 Or •ÔÏ_02<>0 Or ‰ğ–ñ_02<>0 Or c‚_02<>0,K.—Z‘‹àŠz,0) AS I_—Z‘‹àŠz2,"
    wstr = wstr & "IIF(—Z‘_03<>0 Or Œ³‹à_03<>0 Or —˜‘§_03<>0 Or •ÔÏ_03<>0 Or ‰ğ–ñ_03<>0 Or c‚_03<>0,K.—Z‘‹àŠz,0) AS I_—Z‘‹àŠz3,"
    'wstr = wstr & "IIF(Z.—Z‘_01<>0,Z.—Z‘_01,IIF(Z.c‚_01<>0,K.—Z‘‹àŠz,IIF(Z.‰ğ–ñ_01<>0,K.—Z‘‹àŠz,0))) AS I_—Z‘‹àŠz,"
    'wstr = wstr & "IIF(Z.—Z‘_02<>0,Z.—Z‘_02,IIF(Z.c‚_02<>0,K.—Z‘‹àŠz,IIF(Z.‰ğ–ñ_02<>0,K.—Z‘‹àŠz,0))) AS I_—Z‘‹àŠz2,"
    'wstr = wstr & "IIF(Z.—Z‘_03<>0,Z.—Z‘_02,IIF(Z.c‚_03<>0,K.—Z‘‹àŠz,IIF(Z.‰ğ–ñ_03<>0,K.—Z‘‹àŠz,0))) AS I_—Z‘‹àŠz3,"
    wstr = wstr & "Z.—Z‘_01 AS I_“–Šú—Z‘‹àŠz,"
    wstr = wstr & "Z.—Z‘_02 AS I_“–Šú—Z‘‹àŠz2,"
    wstr = wstr & "Z.—Z‘_03 AS I_“–Šú—Z‘‹àŠz3,"
    wstr = wstr & "Z.Œ³‹à_01 AS I_Œ³‹àŠz,"
    wstr = wstr & "Z.Œ³‹à_02 AS I_Œ³‹àŠz2,"
    wstr = wstr & "Z.Œ³‹à_03 AS I_Œ³‹àŠz3,"
    wstr = wstr & "Z.—˜‘§_01 AS I_—˜‘§Šz,"
    wstr = wstr & "Z.—˜‘§_02 AS I_—˜‘§Šz2,"
    wstr = wstr & "Z.—˜‘§_03 AS I_—˜‘§Šz3,"
    wstr = wstr & "Z.•ÔÏ_01 AS I_•ÔÏ‹àŠz,"
    wstr = wstr & "Z.•ÔÏ_02 AS I_•ÔÏ‹àŠz2,"
    wstr = wstr & "Z.•ÔÏ_03 AS I_•ÔÏ‹àŠz3,"
    wstr = wstr & "Z.c‚_01 AS I_—Z‘c‚,"
    wstr = wstr & "Z.c‚_02 AS I_—Z‘c‚2,"
    wstr = wstr & "Z.c‚_03 AS I_—Z‘c‚3"
    wstr = wstr & " FROM (((DCDA010_Ø“üc‚„ˆÚ•\Œ‹‰Ê As Z"
    wstr = wstr & " INNER JOIN DCIA010_Ø“ü‹àƒ[ƒN As K"
    wstr = wstr & " ON Z.Ø“ü”Ô†=K.Ø“ü”Ô†)"
    wstr = wstr & " INNER JOIN DAAA040_‹âsƒ}ƒXƒ^ As G"
    wstr = wstr & " ON K.‹âs”Ô†=G.‹âs”Ô†)"
    wstr = wstr & " LEFT JOIN DAAA116_Ø“ü‹àí•Ê As S"
    wstr = wstr & " ON K.Ø“ü‹àí•Ê‹æ•ª = S.Ø“ü‹àí•Ê‹æ•ª)"
    wstr = wstr & " LEFT JOIN DAAA200_•”–åƒ}ƒXƒ^ As B"
    wstr = wstr & " ON K.ƒvƒƒWƒFƒNƒg”Ô† = B.•”–å”Ô†"

    wstr = wstr & " Where (—Z‘_01<>0"
    wstr = wstr & " Or Œ³‹à_01<>0"
    wstr = wstr & " Or —˜‘§_01<>0"
    wstr = wstr & " Or •ÔÏ_01<>0"
    wstr = wstr & " Or ‰ğ–ñ_01<>0"
    wstr = wstr & " Or c‚_01<>0"
    wstr = wstr & " Or —Z‘_02<>0"
    wstr = wstr & " Or Œ³‹à_02<>0"
    wstr = wstr & " Or —˜‘§_02<>0"
    wstr = wstr & " Or •ÔÏ_02<>0"
    wstr = wstr & " Or ‰ğ–ñ_02<>0"
    wstr = wstr & " Or c‚_02<>0"
    wstr = wstr & " Or —Z‘_03<>0"
    wstr = wstr & " Or Œ³‹à_03<>0"
    wstr = wstr & " Or —˜‘§_03<>0"
    wstr = wstr & " Or •ÔÏ_03<>0"
    wstr = wstr & " Or ‰ğ–ñ_03<>0"
    wstr = wstr & " Or c‚_03<>0)"
    
    'If GRpt.w’è <> "" Then
    '    wstr = wstr & " and S.Ø“ü‹àí•Ê–¼='" & GRpt.w’è & "'"
    'End If

    'wOrder
    wOrder = "": FLG_Order = False
    For j = 1 To 2
        If (GRpt.S_‹à—Z = "•ª—Ş" & CStr(j) Or GRpt.S_‹âs = "•ª—Ş" & CStr(j)) _
        And FLG_Order = False Then
            wOrder = wOrder & "K.‹âs”Ô†,"
            FLG_Order = True
        ElseIf GRpt.S_í•Ê = "•ª—Ş" & CStr(j) Then
            wOrder = wOrder & "K.Ø“ü‹àí•Ê‹æ•ª,"
        ElseIf GRpt.S_•”–å = "•ª—Ş" & CStr(j) Then
            wOrder = wOrder & "B.•”–å”Ô†,"
        ElseIf GRpt.S_‹à—˜ = "•ª—Ş" & CStr(j) Then
            If GStr <> "‹à—˜GR" Then
                wOrder = wOrder & "K.Ø“ü‹àí•Ê‹æ•ª,"
            Else
            '‹à—˜SM
                wOrder = wOrder & "IIF(K.‹à—˜ƒOƒ‹[ƒv‹æ•ª<>'',K.‹à—˜ƒOƒ‹[ƒv‹æ•ª,'99999'),"
            End If
        End If
    Next j
    wOrder = " Order by " & wOrder & "K.Ø“ü”Ô†"
    wstr = wstr & wOrder
    
    Me.DataControl1.Source = wstr
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
ActiveReport_ReportStart_ERR:
    pERR_MES = pPROGRAM_ID + "/ ActiveReport_ReportStart() ‚ÅƒGƒ‰[" + vbCrLf + vbCrLf + _
                "ƒGƒ‰[”Ô†@@F" + CStr(Err.Number) + vbCrLf + _
                "ƒvƒƒWƒFƒNƒg–¼F" + Err.Source + vbCrLf + _
                "ƒGƒ‰[“à—e@@F" + Err.Description + vbCrLf + vbCrLf + _
                    GProduct + "‚ğI—¹‚µ‚Ü‚·"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' ActiveReport_ReportEnd
'------------------------------------------------
Private Sub ActiveReport_ReportEnd()
'
    'FBA010_’ •[”ÍˆÍw’è.ƒƒbƒZ[ƒW = ""
    'FBA010_’ •[”ÍˆÍw’è.ƒƒbƒZ[ƒW.Refresh
'
    ' =========================================
    '           @ ƒ{ƒ^ƒ“§Œä
    ' =========================================
    'FBA010_’ •[”ÍˆÍw’è.Às.Enabled = True
    'FBA010_’ •[”ÍˆÍw’è.•Â‚¶‚é.Enabled = True
'
    'FBA010_’ •[”ÍˆÍw’è.Šg’£.SetFocus
    
    ' =========================================
    '  ØŠ·‚½‚ë‚¤I‚¨‚µ”Å’ •[o—Í‰ñ”ƒ`ƒFƒbƒN
    ' =========================================
    If GSys.Sys = "Ø“ü‹à ‚¨‚µ”Å" Then
        Call MAA001_KARIKAETAROU_CNT
    End If
'
End Sub

'------------------------------------------------
' ActiveReport_NoData
'------------------------------------------------
Private Sub ActiveReport_NoData()
'
    'FBA010_’ •[”ÍˆÍw’è.ƒƒbƒZ[ƒW = "o—Í‚·‚×‚«ƒf[ƒ^‚Í‚ ‚è‚Ü‚¹‚ñ"
    'FBA010_’ •[”ÍˆÍw’è.ƒƒbƒZ[ƒW.Refresh
    GSstrt’ •[Msg = "o—Í‚·‚×‚«ƒf[ƒ^‚Í‚ ‚è‚Ü‚¹‚ñ"
'
    Me.Cancel
    DoEvents
'
    ' =========================================
    '           @ ƒ{ƒ^ƒ“§Œä
    ' =========================================
    'FBA010_’ •[”ÍˆÍw’è.Às.Enabled = True
    'FBA010_’ •[”ÍˆÍw’è.•Â‚¶‚é.Enabled = True
'
    'FBA010_’ •[”ÍˆÍw’è.Šg’£.SetFocus
'
    Unload Me
'
End Sub

'------------------------------------------------
' ActiveReport_Error
'------------------------------------------------
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
'
    'FBA010_’ •[”ÍˆÍw’è.ƒƒbƒZ[ƒW = "o—Í‚Å‚«‚Ü‚¹‚ñ‚Å‚µ‚½"
    'FBA010_’ •[”ÍˆÍw’è.ƒƒbƒZ[ƒW.Refresh
    GSstrt’ •[Msg = "o—Í‚Å‚«‚Ü‚¹‚ñ‚Å‚µ‚½"
'
    Me.Cancel
    DoEvents

    ' =========================================
    '           @ ƒ{ƒ^ƒ“§Œä
    ' =========================================
    'FBA010_’ •[”ÍˆÍw’è.Às.Enabled = True
    'FBA010_’ •[”ÍˆÍw’è.•Â‚¶‚é.Enabled = True
'
    'FBA010_’ •[”ÍˆÍw’è.Šg’£.SetFocus
'
    Unload Me
'
End Sub

'------------------------------------------------
' Detail_BeforePrint
'------------------------------------------------
Private Sub Detail_BeforePrint()
'
    Dim wiKubun As Integer
    Dim wsKubun As String
    
    Dim Yusi1 As Double, Yusi2 As Double, Yusi3 As Double
    Dim TYusi1 As Double, TYusi2 As Double, TYusi3 As Double
    Dim Gankin1 As Double, Gankin2 As Double, Gankin3 As Double
    Dim Risoku1 As Double, Risoku2 As Double, Risoku3 As Double
    Dim Hensai1 As Double, Hensai2 As Double, Hensai3 As Double
    Dim Yusizan1 As Double, Yusizan2 As Double, Yusizan3 As Double
'
    wsKubun = P8.FCStr(Me.Detail.Controls("I_WŒv‹æ•ª"))
    If GRpt.WŒv = "‹à—˜í•ÊŒv" Then
    '‹à—˜í•Ê
        If wsKubun = "•Ï“®" Then
            wiKubun = 0
        Else
            wiKubun = 1
        End If
    ElseIf GRpt.WŒv = "’S•Û‹æ•ªŒv" Then
    '’S•Û‹æ•ª
        If wsKubun = "—L’S" Then
            wiKubun = 0
        Else
            wiKubun = 1
        End If
    ElseIf GRpt.WŒv = "’·’Z‹æ•ªŒv" Then
    '’·’Z‹æ•ª
        If wsKubun = "’·Šú" Then
            wiKubun = 0
        Else
            wiKubun = 1
        End If
    End If
'
    Yusi1 = P8.FCDbl(Me.Detail.Controls("I_—Z‘‹àŠz"))
    Yusi2 = P8.FCDbl(Me.Detail.Controls("I_—Z‘‹àŠz2"))
    Yusi3 = P8.FCDbl(Me.Detail.Controls("I_—Z‘‹àŠz3"))
    TYusi1 = P8.FCDbl(Me.Detail.Controls("I_“–Šú—Z‘‹àŠz"))
    TYusi2 = P8.FCDbl(Me.Detail.Controls("I_“–Šú—Z‘‹àŠz2"))
    TYusi3 = P8.FCDbl(Me.Detail.Controls("I_“–Šú—Z‘‹àŠz3"))
    Gankin1 = P8.FCDbl(Me.Detail.Controls("I_Œ³‹àŠz"))
    Gankin2 = P8.FCDbl(Me.Detail.Controls("I_Œ³‹àŠz2"))
    Gankin3 = P8.FCDbl(Me.Detail.Controls("I_Œ³‹àŠz3"))
    Risoku1 = P8.FCDbl(Me.Detail.Controls("I_—˜‘§Šz"))
    Risoku2 = P8.FCDbl(Me.Detail.Controls("I_—˜‘§Šz2"))
    Risoku3 = P8.FCDbl(Me.Detail.Controls("I_—˜‘§Šz3"))
    Hensai1 = P8.FCDbl(Me.Detail.Controls("I_•ÔÏ‹àŠz"))
    Hensai2 = P8.FCDbl(Me.Detail.Controls("I_•ÔÏ‹àŠz2"))
    Hensai3 = P8.FCDbl(Me.Detail.Controls("I_•ÔÏ‹àŠz3"))
    Yusizan1 = P8.FCDbl(Me.Detail.Controls("I_—Z‘c‚"))
    Yusizan2 = P8.FCDbl(Me.Detail.Controls("I_—Z‘c‚2"))
    Yusizan3 = P8.FCDbl(Me.Detail.Controls("I_—Z‘c‚3"))
    
    Me.Detail.Controls("I_•ÔÏ—¦") = Format(Round(P8.FCDiv(Yusi1 - Yusizan1, Yusi1) * 100, 3), "#,##0.00")
    Me.Detail.Controls("I_•ÔÏ—¦2") = Format(Round(P8.FCDiv(Yusi2 - Yusizan2, Yusi2) * 100, 3), "#,##0.00")
    Me.Detail.Controls("I_•ÔÏ—¦3") = Format(Round(P8.FCDiv(Yusi3 - Yusizan3, Yusi3) * 100, 3), "#,##0.00")
    
    G1_Yusi1(wiKubun) = G1_Yusi1(wiKubun) + Yusi1
    G1_Yusi2(wiKubun) = G1_Yusi2(wiKubun) + Yusi2
    G1_Yusi3(wiKubun) = G1_Yusi3(wiKubun) + Yusi3
    G1_TYusi1(wiKubun) = G1_TYusi1(wiKubun) + TYusi1
    G1_TYusi2(wiKubun) = G1_TYusi2(wiKubun) + TYusi2
    G1_TYusi3(wiKubun) = G1_TYusi3(wiKubun) + TYusi3
    G1_Gankin1(wiKubun) = G1_Gankin1(wiKubun) + Gankin1
    G1_Gankin2(wiKubun) = G1_Gankin2(wiKubun) + Gankin2
    G1_Gankin3(wiKubun) = G1_Gankin3(wiKubun) + Gankin3
    G1_Risoku1(wiKubun) = G1_Risoku1(wiKubun) + Risoku1
    G1_Risoku2(wiKubun) = G1_Risoku2(wiKubun) + Risoku2
    G1_Risoku3(wiKubun) = G1_Risoku3(wiKubun) + Risoku3
    G1_Hensai1(wiKubun) = G1_Hensai1(wiKubun) + Hensai1
    G1_Hensai2(wiKubun) = G1_Hensai2(wiKubun) + Hensai2
    G1_Hensai3(wiKubun) = G1_Hensai3(wiKubun) + Hensai3
    G1_Yusizan1(wiKubun) = G1_Yusizan1(wiKubun) + Yusizan1
    G1_Yusizan2(wiKubun) = G1_Yusizan2(wiKubun) + Yusizan2
    G1_Yusizan3(wiKubun) = G1_Yusizan3(wiKubun) + Yusizan3
'
    Me.Detail.Controls("I_—Z‘‹àŠz") = Format(P8.FCDblRD(Yusi1 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—Z‘‹àŠz2") = Format(P8.FCDblRD(Yusi2 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—Z‘‹àŠz3") = Format(P8.FCDblRD(Yusi3 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(TYusi1 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(TYusi2 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(TYusi3 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_•ÔÏ‹àŠz") = Format(P8.FCDblRD(Hensai1 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(Hensai2 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(Hensai3 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_Œ³‹àŠz") = Format(P8.FCDblRD(Gankin1 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_Œ³‹àŠz2") = Format(P8.FCDblRD(Gankin2 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_Œ³‹àŠz3") = Format(P8.FCDblRD(Gankin3 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—˜‘§Šz") = Format(P8.FCDblRD(Risoku1 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—˜‘§Šz2") = Format(P8.FCDblRD(Risoku2 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—˜‘§Šz3") = Format(P8.FCDblRD(Risoku3 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—Z‘c‚") = Format(P8.FCDblRD(Yusizan1 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—Z‘c‚2") = Format(P8.FCDblRD(Yusizan2 / w•ª•ê), "#,##0")
    Me.Detail.Controls("I_—Z‘c‚3") = Format(P8.FCDblRD(Yusizan3 / w•ª•ê), "#,##0")
'
    G2_Yusi1(wiKubun) = G2_Yusi1(wiKubun) + Yusi1
    G2_Yusi2(wiKubun) = G2_Yusi2(wiKubun) + Yusi2
    G2_Yusi3(wiKubun) = G2_Yusi3(wiKubun) + Yusi3
    G2_TYusi1(wiKubun) = G2_TYusi1(wiKubun) + TYusi1
    G2_TYusi2(wiKubun) = G2_TYusi2(wiKubun) + TYusi2
    G2_TYusi3(wiKubun) = G2_TYusi3(wiKubun) + TYusi3
    G2_Gankin1(wiKubun) = G2_Gankin1(wiKubun) + Gankin1
    G2_Gankin2(wiKubun) = G2_Gankin2(wiKubun) + Gankin2
    G2_Gankin3(wiKubun) = G2_Gankin3(wiKubun) + Gankin3
    G2_Risoku1(wiKubun) = G2_Risoku1(wiKubun) + Risoku1
    G2_Risoku2(wiKubun) = G2_Risoku2(wiKubun) + Risoku2
    G2_Risoku3(wiKubun) = G2_Risoku3(wiKubun) + Risoku3
    G2_Hensai1(wiKubun) = G2_Hensai1(wiKubun) + Hensai1
    G2_Hensai2(wiKubun) = G2_Hensai2(wiKubun) + Hensai2
    G2_Hensai3(wiKubun) = G2_Hensai3(wiKubun) + Hensai3
    G2_Yusizan1(wiKubun) = G2_Yusizan1(wiKubun) + Yusizan1
    G2_Yusizan2(wiKubun) = G2_Yusizan2(wiKubun) + Yusizan2
    G2_Yusizan3(wiKubun) = G2_Yusizan3(wiKubun) + Yusizan3

    GT_Yusi1(wiKubun) = GT_Yusi1(wiKubun) + Yusi1
    GT_Yusi2(wiKubun) = GT_Yusi2(wiKubun) + Yusi2
    GT_Yusi3(wiKubun) = GT_Yusi3(wiKubun) + Yusi3
    GT_TYusi1(wiKubun) = GT_TYusi1(wiKubun) + TYusi1
    GT_TYusi2(wiKubun) = GT_TYusi2(wiKubun) + TYusi2
    GT_TYusi3(wiKubun) = GT_TYusi3(wiKubun) + TYusi3
    GT_Gankin1(wiKubun) = GT_Gankin1(wiKubun) + Gankin1
    GT_Gankin2(wiKubun) = GT_Gankin2(wiKubun) + Gankin2
    GT_Gankin3(wiKubun) = GT_Gankin3(wiKubun) + Gankin3
    GT_Risoku1(wiKubun) = GT_Risoku1(wiKubun) + Risoku1
    GT_Risoku2(wiKubun) = GT_Risoku2(wiKubun) + Risoku2
    GT_Risoku3(wiKubun) = GT_Risoku3(wiKubun) + Risoku3
    GT_Hensai1(wiKubun) = GT_Hensai1(wiKubun) + Hensai1
    GT_Hensai2(wiKubun) = GT_Hensai2(wiKubun) + Hensai2
    GT_Hensai3(wiKubun) = GT_Hensai3(wiKubun) + Hensai3
    GT_Yusizan1(wiKubun) = GT_Yusizan1(wiKubun) + Yusizan1
    GT_Yusizan2(wiKubun) = GT_Yusizan2(wiKubun) + Yusizan2
    GT_Yusizan3(wiKubun) = GT_Yusizan3(wiKubun) + Yusizan3
'
End Sub

'------------------------------------------------
' GroupFooter1_BeforePrint
'------------------------------------------------
Private Sub GroupFooter1_BeforePrint()
'
    Dim Yusi1 As Double, Yusi2 As Double, Yusi3 As Double
    Dim Yusizan1 As Double, Yusizan2 As Double, Yusizan3 As Double
'
    Yusi1 = P8.FCDbl(Me.GroupFooter1.Controls("G10_—Z‘‹àŠz"))
    Yusi2 = P8.FCDbl(Me.GroupFooter1.Controls("G10_—Z‘‹àŠz2"))
    Yusi3 = P8.FCDbl(Me.GroupFooter1.Controls("G10_—Z‘‹àŠz3"))
    Yusizan1 = P8.FCDbl(Me.GroupFooter1.Controls("G10_—Z‘c‚"))
    Yusizan2 = P8.FCDbl(Me.GroupFooter1.Controls("G10_—Z‘c‚2"))
    Yusizan3 = P8.FCDbl(Me.GroupFooter1.Controls("G10_—Z‘c‚3"))
'
    Me.GroupFooter1.Controls("G10_—Z‘‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—Z‘‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_—Z‘‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—Z‘‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_—Z‘‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—Z‘‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter1.Controls("G11_—Z‘‹àŠz") = Format(P8.FCDblRD(G1_Yusi1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—Z‘‹àŠz") = Format(P8.FCDblRD(G1_Yusi1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_—Z‘‹àŠz2") = Format(P8.FCDblRD(G1_Yusi2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—Z‘‹àŠz2") = Format(P8.FCDblRD(G1_Yusi2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_—Z‘‹àŠz3") = Format(P8.FCDblRD(G1_Yusi3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—Z‘‹àŠz3") = Format(P8.FCDblRD(G1_Yusi3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter1.Controls("G10_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_“–Šú—Z‘‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_“–Šú—Z‘‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_“–Šú—Z‘‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter1.Controls("G11_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(G1_TYusi1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(G1_TYusi1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(G1_TYusi2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(G1_TYusi2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(G1_TYusi3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(G1_TYusi3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter1.Controls("G10_Œ³‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_Œ³‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_Œ³‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_Œ³‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_Œ³‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_Œ³‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter1.Controls("G11_Œ³‹àŠz") = Format(P8.FCDblRD(G1_Gankin1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_Œ³‹àŠz") = Format(P8.FCDblRD(G1_Gankin1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_Œ³‹àŠz2") = Format(P8.FCDblRD(G1_Gankin2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_Œ³‹àŠz2") = Format(P8.FCDblRD(G1_Gankin2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_Œ³‹àŠz3") = Format(P8.FCDblRD(G1_Gankin3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_Œ³‹àŠz3") = Format(P8.FCDblRD(G1_Gankin3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter1.Controls("G10_—˜‘§Šz") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—˜‘§Šz") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_—˜‘§Šz2") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—˜‘§Šz2") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_—˜‘§Šz3") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—˜‘§Šz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter1.Controls("G11_—˜‘§Šz") = Format(P8.FCDblRD(G1_Risoku1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—˜‘§Šz") = Format(P8.FCDblRD(G1_Risoku1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_—˜‘§Šz2") = Format(P8.FCDblRD(G1_Risoku2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—˜‘§Šz2") = Format(P8.FCDblRD(G1_Risoku2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_—˜‘§Šz3") = Format(P8.FCDblRD(G1_Risoku3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—˜‘§Šz3") = Format(P8.FCDblRD(G1_Risoku3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter1.Controls("G10_•ÔÏ‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_•ÔÏ‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_•ÔÏ‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_•ÔÏ‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter1.Controls("G11_•ÔÏ‹àŠz") = Format(P8.FCDblRD(G1_Hensai1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_•ÔÏ‹àŠz") = Format(P8.FCDblRD(G1_Hensai1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(G1_Hensai2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(G1_Hensai2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(G1_Hensai3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(G1_Hensai3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter1.Controls("G10_—Z‘c‚") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—Z‘c‚") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_—Z‘c‚2") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—Z‘c‚2") / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G10_—Z‘c‚3") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_—Z‘c‚3") / w•ª•ê), "#,##0")

    Me.GroupFooter1.Controls("G11_—Z‘c‚") = Format(P8.FCDblRD(G1_Yusizan1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—Z‘c‚") = Format(P8.FCDblRD(G1_Yusizan1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_—Z‘c‚2") = Format(P8.FCDblRD(G1_Yusizan2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—Z‘c‚2") = Format(P8.FCDblRD(G1_Yusizan2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G11_—Z‘c‚3") = Format(P8.FCDblRD(G1_Yusizan3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter1.Controls("G12_—Z‘c‚3") = Format(P8.FCDblRD(G1_Yusizan3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter1.Controls("G10_•ÔÏ—¦") = Format(Round(P8.FCDiv(Yusi1 - Yusizan1, Yusi1) * 100, 3), "#,##0.00")
    Me.GroupFooter1.Controls("G10_•ÔÏ—¦2") = Format(Round(P8.FCDiv(Yusi2 - Yusizan2, Yusi2) * 100, 3), "#,##0.00")
    Me.GroupFooter1.Controls("G10_•ÔÏ—¦3") = Format(Round(P8.FCDiv(Yusi3 - Yusizan3, Yusi3) * 100, 3), "#,##0.00")

    Me.GroupFooter1.Controls("G11_•ÔÏ—¦") = Format(Round(P8.FCDiv(G1_Yusi1(0) - G1_Yusizan1(0), G1_Yusi1(0)) * 100, 3), "#,##0.00")
    Me.GroupFooter1.Controls("G12_•ÔÏ—¦") = Format(Round(P8.FCDiv(G1_Yusi1(1) - G1_Yusizan1(1), G1_Yusi1(1)) * 100, 3), "#,##0.00")
    Me.GroupFooter1.Controls("G11_•ÔÏ—¦2") = Format(Round(P8.FCDiv(G1_Yusi2(0) - G1_Yusizan2(0), G1_Yusi2(0)) * 100, 3), "#,##0.00")
    Me.GroupFooter1.Controls("G12_•ÔÏ—¦2") = Format(Round(P8.FCDiv(G1_Yusi2(1) - G1_Yusizan2(1), G1_Yusi2(1)) * 100, 3), "#,##0.00")
    Me.GroupFooter1.Controls("G11_•ÔÏ—¦3") = Format(Round(P8.FCDiv(G1_Yusi3(0) - G1_Yusizan3(0), G1_Yusi3(0)) * 100, 3), "#,##0.00")
    Me.GroupFooter1.Controls("G12_•ÔÏ—¦3") = Format(Round(P8.FCDiv(G1_Yusi3(1) - G1_Yusizan3(1), G1_Yusi3(1)) * 100, 3), "#,##0.00")
'
End Sub

'------------------------------------------------
' GroupFooter1_AfterPrint
'------------------------------------------------
Private Sub GroupFooter1_AfterPrint()
'
    Dim j As Integer
'
    For j = 0 To 1
        G1_Yusi1(j) = 0: G1_Yusi2(j) = 0: G1_Yusi3(j) = 0
        G1_TYusi1(j) = 0: G1_TYusi2(j) = 0: G1_TYusi3(j) = 0
        G1_Gankin1(j) = 0: G1_Gankin2(j) = 0: G1_Gankin3(j) = 0
        G1_Risoku1(j) = 0: G1_Risoku2(j) = 0: G1_Risoku3(j) = 0
        G1_Hensai1(j) = 0: G1_Hensai2(j) = 0: G1_Hensai3(j) = 0
        G1_Yusizan1(j) = 0: G1_Yusizan2(j) = 0: G1_Yusizan3(j) = 0
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter2_BeforePrint
'------------------------------------------------
Private Sub GroupFooter2_BeforePrint()
'
    Dim Yusi1 As Double, Yusi2 As Double, Yusi3 As Double
    Dim Yusizan1 As Double, Yusizan2 As Double, Yusizan3 As Double
'
    Yusi1 = P8.FCDbl(Me.GroupFooter2.Controls("G20_—Z‘‹àŠz"))
    Yusi2 = P8.FCDbl(Me.GroupFooter2.Controls("G20_—Z‘‹àŠz2"))
    Yusi3 = P8.FCDbl(Me.GroupFooter2.Controls("G20_—Z‘‹àŠz3"))
    Yusizan1 = P8.FCDbl(Me.GroupFooter2.Controls("G20_—Z‘c‚"))
    Yusizan2 = P8.FCDbl(Me.GroupFooter2.Controls("G20_—Z‘c‚2"))
    Yusizan3 = P8.FCDbl(Me.GroupFooter2.Controls("G20_—Z‘c‚3"))
'
    Me.GroupFooter2.Controls("G20_—Z‘‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—Z‘‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_—Z‘‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—Z‘‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_—Z‘‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—Z‘‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter2.Controls("G21_—Z‘‹àŠz") = Format(P8.FCDblRD(G2_Yusi1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—Z‘‹àŠz") = Format(P8.FCDblRD(G2_Yusi1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_—Z‘‹àŠz2") = Format(P8.FCDblRD(G2_Yusi2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—Z‘‹àŠz2") = Format(P8.FCDblRD(G2_Yusi2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_—Z‘‹àŠz3") = Format(P8.FCDblRD(G2_Yusi3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—Z‘‹àŠz3") = Format(P8.FCDblRD(G2_Yusi3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter2.Controls("G20_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_“–Šú—Z‘‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_“–Šú—Z‘‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_“–Šú—Z‘‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter2.Controls("G21_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(G2_TYusi1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(G2_TYusi1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(G2_TYusi2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(G2_TYusi2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(G2_TYusi3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(G2_TYusi3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter2.Controls("G20_Œ³‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_Œ³‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_Œ³‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_Œ³‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_Œ³‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_Œ³‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter2.Controls("G21_Œ³‹àŠz") = Format(P8.FCDblRD(G2_Gankin1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_Œ³‹àŠz") = Format(P8.FCDblRD(G2_Gankin1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_Œ³‹àŠz2") = Format(P8.FCDblRD(G2_Gankin2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_Œ³‹àŠz2") = Format(P8.FCDblRD(G2_Gankin2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_Œ³‹àŠz3") = Format(P8.FCDblRD(G2_Gankin3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_Œ³‹àŠz3") = Format(P8.FCDblRD(G2_Gankin3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter2.Controls("G20_—˜‘§Šz") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—˜‘§Šz") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_—˜‘§Šz2") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—˜‘§Šz2") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_—˜‘§Šz3") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—˜‘§Šz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter2.Controls("G21_—˜‘§Šz") = Format(P8.FCDblRD(G2_Risoku1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—˜‘§Šz") = Format(P8.FCDblRD(G2_Risoku1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_—˜‘§Šz2") = Format(P8.FCDblRD(G2_Risoku2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—˜‘§Šz2") = Format(P8.FCDblRD(G2_Risoku2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_—˜‘§Šz3") = Format(P8.FCDblRD(G2_Risoku3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—˜‘§Šz3") = Format(P8.FCDblRD(G2_Risoku3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter2.Controls("G20_•ÔÏ‹àŠz") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_•ÔÏ‹àŠz") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_•ÔÏ‹àŠz2") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_•ÔÏ‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.GroupFooter2.Controls("G21_•ÔÏ‹àŠz") = Format(P8.FCDblRD(G2_Hensai1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_•ÔÏ‹àŠz") = Format(P8.FCDblRD(G2_Hensai1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(G2_Hensai2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(G2_Hensai2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(G2_Hensai3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(G2_Hensai3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter2.Controls("G20_—Z‘c‚") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—Z‘c‚") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_—Z‘c‚2") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—Z‘c‚2") / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G20_—Z‘c‚3") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_—Z‘c‚3") / w•ª•ê), "#,##0")

    Me.GroupFooter2.Controls("G21_—Z‘c‚") = Format(P8.FCDblRD(G2_Yusizan1(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—Z‘c‚") = Format(P8.FCDblRD(G2_Yusizan1(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_—Z‘c‚2") = Format(P8.FCDblRD(G2_Yusizan2(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—Z‘c‚2") = Format(P8.FCDblRD(G2_Yusizan2(1) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G21_—Z‘c‚3") = Format(P8.FCDblRD(G2_Yusizan3(0) / w•ª•ê), "#,##0")
    Me.GroupFooter2.Controls("G22_—Z‘c‚3") = Format(P8.FCDblRD(G2_Yusizan3(1) / w•ª•ê), "#,##0")
    '
    Me.GroupFooter2.Controls("G20_•ÔÏ—¦") = Format(Round(P8.FCDiv(Yusi1 - Yusizan1, Yusi1) * 100, 3), "#,##0.00")
    Me.GroupFooter2.Controls("G20_•ÔÏ—¦2") = Format(Round(P8.FCDiv(Yusi2 - Yusizan2, Yusi2) * 100, 3), "#,##0.00")
    Me.GroupFooter2.Controls("G20_•ÔÏ—¦3") = Format(Round(P8.FCDiv(Yusi3 - Yusizan3, Yusi3) * 100, 3), "#,##0.00")

    Me.GroupFooter2.Controls("G21_•ÔÏ—¦") = Format(Round(P8.FCDiv(G2_Yusi1(0) - G2_Yusizan1(0), G2_Yusi1(0)) * 100, 3), "#,##0.00")
    Me.GroupFooter2.Controls("G22_•ÔÏ—¦") = Format(Round(P8.FCDiv(G2_Yusi1(1) - G2_Yusizan1(1), G2_Yusi1(1)) * 100, 3), "#,##0.00")
    Me.GroupFooter2.Controls("G21_•ÔÏ—¦2") = Format(Round(P8.FCDiv(G2_Yusi2(0) - G2_Yusizan2(0), G2_Yusi2(0)) * 100, 3), "#,##0.00")
    Me.GroupFooter2.Controls("G22_•ÔÏ—¦2") = Format(Round(P8.FCDiv(G2_Yusi2(1) - G2_Yusizan2(1), G2_Yusi2(1)) * 100, 3), "#,##0.00")
    Me.GroupFooter2.Controls("G21_•ÔÏ—¦3") = Format(Round(P8.FCDiv(G2_Yusi3(0) - G2_Yusizan3(0), G2_Yusi3(0)) * 100, 3), "#,##0.00")
    Me.GroupFooter2.Controls("G22_•ÔÏ—¦3") = Format(Round(P8.FCDiv(G2_Yusi3(1) - G2_Yusizan3(1), G2_Yusi3(1)) * 100, 3), "#,##0.00")
'
End Sub

'------------------------------------------------
' GroupFooter2_AfterPrint
'------------------------------------------------
Private Sub GroupFooter2_AfterPrint()
'
    Dim j As Integer
'
    For j = 0 To 1
        G2_Yusi1(j) = 0: G2_Yusi2(j) = 0: G2_Yusi3(j) = 0
        G2_TYusi1(j) = 0: G2_TYusi2(j) = 0: G2_TYusi3(j) = 0
        G2_Gankin1(j) = 0: G2_Gankin2(j) = 0: G2_Gankin3(j) = 0
        G2_Risoku1(j) = 0: G2_Risoku2(j) = 0: G2_Risoku3(j) = 0
        G2_Hensai1(j) = 0: G2_Hensai2(j) = 0: G2_Hensai3(j) = 0
        G2_Yusizan1(j) = 0: G2_Yusizan2(j) = 0: G2_Yusizan3(j) = 0
    Next j
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Dim Yusi1 As Double, Yusi2 As Double, Yusi3 As Double
    Dim Yusizan1 As Double, Yusizan2 As Double, Yusizan3 As Double
'
    Yusi1 = P8.FCDbl(Me.ReportFooter.Controls("G90_—Z‘‹àŠz"))
    Yusi2 = P8.FCDbl(Me.ReportFooter.Controls("G90_—Z‘‹àŠz2"))
    Yusi3 = P8.FCDbl(Me.ReportFooter.Controls("G90_—Z‘‹àŠz3"))
    Yusizan1 = P8.FCDbl(Me.ReportFooter.Controls("G90_—Z‘c‚"))
    Yusizan2 = P8.FCDbl(Me.ReportFooter.Controls("G90_—Z‘c‚2"))
    Yusizan3 = P8.FCDbl(Me.ReportFooter.Controls("G90_—Z‘c‚3"))
'
    Me.ReportFooter.Controls("G90_—Z‘‹àŠz") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—Z‘‹àŠz") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_—Z‘‹àŠz2") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—Z‘‹àŠz2") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_—Z‘‹àŠz3") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—Z‘‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.ReportFooter.Controls("G91_—Z‘‹àŠz") = Format(P8.FCDblRD(GT_Yusi1(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—Z‘‹àŠz") = Format(P8.FCDblRD(GT_Yusi1(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_—Z‘‹àŠz2") = Format(P8.FCDblRD(GT_Yusi2(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—Z‘‹àŠz2") = Format(P8.FCDblRD(GT_Yusi2(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_—Z‘‹àŠz3") = Format(P8.FCDblRD(GT_Yusi3(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—Z‘‹àŠz3") = Format(P8.FCDblRD(GT_Yusi3(1) / w•ª•ê), "#,##0")
    '
    Me.ReportFooter.Controls("G90_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_“–Šú—Z‘‹àŠz") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_“–Šú—Z‘‹àŠz2") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_“–Šú—Z‘‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.ReportFooter.Controls("G91_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(GT_TYusi1(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_“–Šú—Z‘‹àŠz") = Format(P8.FCDblRD(GT_TYusi1(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(GT_TYusi2(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_“–Šú—Z‘‹àŠz2") = Format(P8.FCDblRD(GT_TYusi2(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(GT_TYusi3(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_“–Šú—Z‘‹àŠz3") = Format(P8.FCDblRD(GT_TYusi3(1) / w•ª•ê), "#,##0")
    '
    Me.ReportFooter.Controls("G90_Œ³‹àŠz") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_Œ³‹àŠz") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_Œ³‹àŠz2") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_Œ³‹àŠz2") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_Œ³‹àŠz3") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_Œ³‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.ReportFooter.Controls("G91_Œ³‹àŠz") = Format(P8.FCDblRD(GT_Gankin1(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_Œ³‹àŠz") = Format(P8.FCDblRD(GT_Gankin1(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_Œ³‹àŠz2") = Format(P8.FCDblRD(GT_Gankin2(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_Œ³‹àŠz2") = Format(P8.FCDblRD(GT_Gankin2(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_Œ³‹àŠz3") = Format(P8.FCDblRD(GT_Gankin3(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_Œ³‹àŠz3") = Format(P8.FCDblRD(GT_Gankin3(1) / w•ª•ê), "#,##0")
    '
    Me.ReportFooter.Controls("G90_—˜‘§Šz") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—˜‘§Šz") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_—˜‘§Šz2") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—˜‘§Šz2") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_—˜‘§Šz3") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—˜‘§Šz3") / w•ª•ê), "#,##0")

    Me.ReportFooter.Controls("G91_—˜‘§Šz") = Format(P8.FCDblRD(GT_Risoku1(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—˜‘§Šz") = Format(P8.FCDblRD(GT_Risoku1(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_—˜‘§Šz2") = Format(P8.FCDblRD(GT_Risoku2(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—˜‘§Šz2") = Format(P8.FCDblRD(GT_Risoku2(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_—˜‘§Šz3") = Format(P8.FCDblRD(GT_Risoku3(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—˜‘§Šz3") = Format(P8.FCDblRD(GT_Risoku3(1) / w•ª•ê), "#,##0")
    '
    Me.ReportFooter.Controls("G90_•ÔÏ‹àŠz") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_•ÔÏ‹àŠz") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_•ÔÏ‹àŠz2") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_•ÔÏ‹àŠz3") / w•ª•ê), "#,##0")
    
    Me.ReportFooter.Controls("G91_•ÔÏ‹àŠz") = Format(P8.FCDblRD(GT_Hensai1(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_•ÔÏ‹àŠz") = Format(P8.FCDblRD(GT_Hensai1(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(GT_Hensai2(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_•ÔÏ‹àŠz2") = Format(P8.FCDblRD(GT_Hensai2(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(GT_Hensai3(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_•ÔÏ‹àŠz3") = Format(P8.FCDblRD(GT_Hensai3(1) / w•ª•ê), "#,##0")
    '
    Me.ReportFooter.Controls("G90_—Z‘c‚") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—Z‘c‚") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_—Z‘c‚2") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—Z‘c‚2") / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G90_—Z‘c‚3") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_—Z‘c‚3") / w•ª•ê), "#,##0")

    Me.ReportFooter.Controls("G91_—Z‘c‚") = Format(P8.FCDblRD(GT_Yusizan1(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—Z‘c‚") = Format(P8.FCDblRD(GT_Yusizan1(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_—Z‘c‚2") = Format(P8.FCDblRD(GT_Yusizan2(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—Z‘c‚2") = Format(P8.FCDblRD(GT_Yusizan2(1) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G91_—Z‘c‚3") = Format(P8.FCDblRD(GT_Yusizan3(0) / w•ª•ê), "#,##0")
    Me.ReportFooter.Controls("G92_—Z‘c‚3") = Format(P8.FCDblRD(GT_Yusizan3(1) / w•ª•ê), "#,##0")
    '
    Me.ReportFooter.Controls("G90_•ÔÏ—¦") = Format(Round(P8.FCDiv(Yusi1 - Yusizan1, Yusi1) * 100, 3), "#,##0.00")
    Me.ReportFooter.Controls("G90_•ÔÏ—¦2") = Format(Round(P8.FCDiv(Yusi2 - Yusizan2, Yusi2) * 100, 3), "#,##0.00")
    Me.ReportFooter.Controls("G90_•ÔÏ—¦3") = Format(Round(P8.FCDiv(Yusi3 - Yusizan3, Yusi3) * 100, 3), "#,##0.00")

    Me.ReportFooter.Controls("G91_•ÔÏ—¦") = Format(Round(P8.FCDiv(GT_Yusi1(0) - GT_Yusizan1(0), GT_Yusi1(0)) * 100, 3), "#,##0.00")
    Me.ReportFooter.Controls("G92_•ÔÏ—¦") = Format(Round(P8.FCDiv(GT_Yusi1(1) - GT_Yusizan1(1), GT_Yusi1(1)) * 100, 3), "#,##0.00")
    Me.ReportFooter.Controls("G91_•ÔÏ—¦2") = Format(Round(P8.FCDiv(GT_Yusi2(0) - GT_Yusizan2(0), GT_Yusi2(0)) * 100, 3), "#,##0.00")
    Me.ReportFooter.Controls("G92_•ÔÏ—¦2") = Format(Round(P8.FCDiv(GT_Yusi2(1) - GT_Yusizan2(1), GT_Yusi2(1)) * 100, 3), "#,##0.00")
    Me.ReportFooter.Controls("G91_•ÔÏ—¦3") = Format(Round(P8.FCDiv(GT_Yusi3(0) - GT_Yusizan3(0), GT_Yusi3(0)) * 100, 3), "#,##0.00")
    Me.ReportFooter.Controls("G92_•ÔÏ—¦3") = Format(Round(P8.FCDiv(GT_Yusi3(1) - GT_Yusizan3(1), GT_Yusi3(1)) * 100, 3), "#,##0.00")
'
End Sub

