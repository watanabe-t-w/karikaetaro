VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_I借入金登録_金利変更 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入金登録 金利変更"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   Icon            =   "frm_I借入金登録_金利変更.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton 閉じる 
      Caption         =   "閉じる(F12)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton 登録 
      Caption         =   "登録（F11)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton 削除 
      Caption         =   "削除"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox 金利変更利率 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   2
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox 金利変更年月 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2160
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4965
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8758
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "金利変更"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label L_金利変更年月日 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000000&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label L_番号 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label L_番号1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "借入番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label79 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "金利変更年月"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label26 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "金利変更利率(%)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "金利変更年月日"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   2055
   End
End
Attribute VB_Name = "frm_I借入金登録_金利変更"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "借入金登録_金利変更"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim wi利息支払 As Integer, wi支払日 As Integer, wi営業日区分 As Integer, wi金利種別 As Integer
Dim wi利息計算日数区分 As Integer, wi返済単位 As Integer
Dim ws利息区分 As String
Dim wFname As String, wsTbl As String
Dim wsBango As String

Dim wv初回返済実行日 As Variant, wv最終返済実行日 As Variant
Dim wv実行日 As Variant, wv初回返済年月 As Variant, wv最終返済年月 As Variant, wv初回金利年月 As Variant
Dim FLG_MAX As Boolean
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    Dim j As Integer
'
'    Me.Caption = GFcap
    Me.Left = G_LEFT
    Me.Top = G_TOP
'
    wFname = GStr
    'ZU050_Button1.Caption = wFname & Space(1) & "登録"
    
    wsBango = GStr_2
    
    L_番号.Caption = wsBango
    
    'L_返済方法.Caption = "元金均等返済"
    
    L_番号1.Caption = "借入番号"
    Select Case wFname
    Case "借入金登録"
        L_番号1.Caption = " 借入番号"
        
        wsTbl = "DBDA010_借入金"
    Case "貸付登録"
        L_番号1.Caption = " 貸付番号"
    
        wsTbl = "DBDA010_貸付金"
    End Select
    
    GStr = "": GStr_1 = "": GStr_2 = ""
'
    ' =========================================
    '                 初期設定
    ' =========================================
    FLG_MAX = False
    
    金利変更年月 = ""
    L_金利変更年月日.Caption = ""
    金利変更利率 = 0
    '取消 = 0
    
    wv初回返済実行日 = Null
    wv最終返済実行日 = Null
    
    'ワークテーブル作成とワークデータセット
    Call 金利ワークテーブル作成
    
    Call 画面セット
'
End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    DoEvents
    
    'Call MXA010_検索用データクリア
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
    If KeyCode = vbKeyF11 Then
        Call 登録_Click
    End If
    
    If KeyCode = vbKeyF12 Then
        Call 閉じる_Click
    End If
'
End Sub

'------------------------------------------------
' Form_KeyPress
'------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
'
    KeyAscii = CEkey.X020_EnterKey(Me, KeyAscii, True)
'
End Sub

'------------------------------------------------
' 金利ワークテーブル作成
'------------------------------------------------
Private Sub 金利ワークテーブル作成()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer
    Dim ws01 As String
'
    On Error GoTo 金利ワークテーブル作成_ERR
'
    '----------< ワークテーブル削除 >------------------------------------------
    wstr = "Delete * from DCHA010_Gridワーク"
    GDb.Execute wstr
'
    If wsBango = "" Then
        Exit Sub
    End If
'
    wi支払日 = 0
    wi営業日区分 = 0
    wi金利種別 = 0
    
    wv初回返済実行日 = Null
    wv最終返済実行日 = Null

    wi利息支払 = 0
    wi返済単位 = 1
    wi利息計算日数区分 = 0
    ws利息区分 = ""
    wv実行日 = Null
    wv初回返済年月 = Null
    wv最終返済年月 = Null
    wv初回金利年月 = Null
    
    '----------< テーブル Write >----------------------------------------------
    wstr = "Select * from DCHA010_Gridワーク"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr1 = "Select * from " & wsTbl
        wstr1 = wstr1 & " Where " & P8.FCStr(L_番号1.Caption) & "='" & wsBango & "'"
        Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        If Not wRs1.EOF Then
            
            wi支払日 = P8.FCDbl(wRs1("支払日"))
            wi営業日区分 = P8.FCDbl(wRs1("営業日区分"))
            wi金利種別 = P8.FCDbl(wRs1("金利種別"))
            wv初回返済実行日 = wRs1("初回返済実行日")
            wv最終返済実行日 = wRs1("最終返済実行日")
            
            wi利息支払 = P8.FCDbl(wRs1("利息支払方法"))
            wi返済単位 = P8.FCDbl(wRs1("返済単位月数"))
            wi利息計算日数区分 = P8.FCDbl(wRs1("利息計算日数区分"))
            ws利息区分 = P8.FCStr(wRs1("利息区分"))
            wv実行日 = wRs1("実行日")
            wv初回返済年月 = wRs1("初回返済年月")
            wv最終返済年月 = wRs1("最終返済年月")
            wv初回金利年月 = wRs1("金利初回年月")
            
            For j = 2 To 100
                
                ws01 = "金利変更" & CStr(j) & "回目年月"
                If Not IsNull(P8.FCDate(wRs1(ws01))) Then
                    
                    wRs.AddNew
                    
                    wRs("テキスト1") = wsBango
                    wRs("テキスト2") = j
                    
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs("年月日1") = P8.FCDate(wRs1(ws01))
                    
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs("数値1") = P8.FCDbl(wRs1(ws01))
                
                    wRs.Update
                    
                End If
                
            Next
            
            If Not IsNull(P8.FCDate(wRs1("金利変更１００回目年月"))) Then
                FLG_MAX = True
            End If
        
        End If
        wRs1.Close
        Set wRs1 = Nothing

    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利ワークテーブル作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利ワークテーブル作成() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' AdodcRefresh
'------------------------------------------------
Private Sub AdodcRefresh()
'
    On Error GoTo AdodcRefresh_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
    Call MXA030_DataGridInit(DataGrid1)
    Set DataGrid1.DataSource = Adodc1
  
    ' =========================================
    '              ConnectionString
    ' =========================================
    Call AdodcSet(Adodc1, GDb)
  
    ' =========================================
    '              メインクエリ
    ' =========================================
    GWhere = ""
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " テキスト2 As Grd回,"
    wstr = wstr + " format(年月日1,'" & Gfmt年月 & "') As Grd年月,"
    wstr = wstr + " format(数値1,'#,##0.00000') As Grd利率"
    'wstr = wstr + " IIF(取消フラグ = 0,'','×') As Grd取消"
    wstr = wstr + " From DCHA010_Gridワーク"
    wstr = wstr + GWhere
    wstr = wstr + " Order By 年月日1"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("回", "", 600, "L")
        Call XZMA010_DataGrid_Set("年月", "金利変更年月", 2000, "R")
        Call XZMA010_DataGrid_Set("利率", "金利変更利率", 2000, "R")
        'Call XZMA010_DataGrid_Set("取消", "", 550, "C")
    Call XZMA010_DataGrid_Action(DataGrid1)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' DataGrid1_Click
'------------------------------------------------
Private Sub DataGrid1_Click()
'
    Call CEkey.SetFs(金利変更年月, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd年月")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        金利変更年月 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd年月"))
    On Error GoTo 0
    
    Call 画面セット
   
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(金利変更年月, True)

Exit_Sub:
    Exit Sub
    '---------------------------------------------------
Err_Hundle:
    If Err.Number = 91 Then Resume Next
    If Err.Number = 94 Then Resume Next
    MsgBox CStr(Err.Number) + ":" + Err.Description
    Resume Exit_Sub
End Sub

'------------------------------------------------
' B_SET_Click
'------------------------------------------------
Private Sub B_SET_Click()
    Call 画面セット
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット() As Boolean
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
'
    金利変更利率 = 0
    '取消 = 0
    
    ' =========================================
    '                画面クリア
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    If Not IsNull(GVar1) And GVar1 <> "" Then
        'GRet = 金利変更年月CHECK(Format(GVar1, "yyyy/mm/dd"))
        'If GRet <> True Then
        '    GRet = MsgBox("金利変更年月を確認してください", vbOKOnly + vbCritical)
        '
        '    金利変更年月 = ""
        '    L_金利変更年月日.Caption = ""
        '    金利変更利率 = 0
        '
        '    Call CEkey.SetFs(金利変更年月, True)
        '        Exit Function
        'End If
    End If
    
    wstr = ""
    wstr = wstr + "Select * From  DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.EOF Then
        If Not IsNull(GVar1) Then
            GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
            If GRet = vbNo Then
                wRs.Close
                Set wRs = Nothing
                
                Exit Function
            End If
    
            GRet = 金利変更年月CHECK(Format(GVar1, "yyyy/mm/dd"))
            If GRet <> True Then
                GRet = MsgBox("金利変更年月を確認してください", vbOKOnly + vbCritical)
                
                金利変更年月 = ""
                L_金利変更年月日.Caption = ""
                金利変更利率 = 0
                        
                Call CEkey.SetFs(金利変更年月, True)
            
                wRs.Close
                Set wRs = Nothing
                
                Exit Function
            End If
            
            If FLG_MAX = True Then
                GRet = MsgBox("金利変更100回を越えると登録できません。", vbOKOnly)
                wRs.Close
                Set wRs = Nothing
                
                Exit Function
            End If
            
            Call CEkey.SetFs(金利変更利率, True)
        End If
    Else
    
        金利変更年月 = Format(wRs("年月日1"), Gfmt年月)
        金利変更利率 = P8.FFormat(P8.FCDbl(wRs("数値1")), "#,##0.00000")
        
        Call 金利変更年月日_セット
    
    End If
    wRs.Close
    Set wRs = Nothing
    
    ' =========================================
    '            Grid セット
    ' =========================================
    Call AdodcRefresh

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月 = '" + 金利変更年月 + "'")
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
画面セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 画面セット() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

Private Function 金利変更年月CHECK(pDate As Variant) As Boolean
'
    Dim wi01 As Integer
    Dim wd01 As Date
    Dim wvStr As Variant, wvEnd As Variant, wv01 As Variant
'
    On Error GoTo 金利変更年月CHECK_ERR
'
    金利変更年月CHECK = False
    
    If ws利息区分 = XMXA020_区分("利息区分", "利息先払") Then
        If CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
            wvStr = wv初回金利年月
            wvEnd = DateAdd("m", -1, CDate(wv最終返済年月))
            
            If Format(wvStr, "yyyy/mm/01") <= Format(pDate, "yyyy/mm/01") _
            And Format(wvEnd, "yyyy/mm/01") >= Format(pDate, "yyyy/mm/01") Then
                金利変更年月CHECK = True
            End If
        
        ElseIf CStr(wi利息支払) = XMXA020_区分("利息支払", "一括") Then
            If Format(wv初回返済年月, "yyyy/mm/01") > Format(pDate, "yyyy/mm/01") Then
                wvStr = wv初回金利年月
                wvEnd = wv初回返済年月
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") < Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            
            Else
                wvStr = wv初回返済年月
                wvEnd = DateAdd("m", -wi返済単位, CDate(wv最終返済年月))
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") <= Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            End If
        End If

    ElseIf ws利息区分 = XMXA020_区分("利息区分", "利息後払") Then
        If CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
            wvStr = DateAdd("m", 1, CDate(wv初回金利年月))
            wvEnd = wv最終返済年月
                
            If Format(wvStr, "yyyy/mm/01") <= Format(pDate, "yyyy/mm/01") _
            And Format(wvEnd, "yyyy/mm/01") >= Format(pDate, "yyyy/mm/01") Then
                金利変更年月CHECK = True
            End If
        
        ElseIf CStr(wi利息支払) = XMXA020_区分("利息支払", "一括") Then
        
            If Format(wv初回返済年月, "yyyy/mm/01") > Format(pDate, "yyyy/mm/01") Then
                wvStr = DateAdd("m", wi返済単位, CDate(wv初回金利年月))
                wvEnd = wv初回返済年月
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") < Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            Else
                wvStr = DateAdd("m", wi返済単位, CDate(wv初回金利年月))
                wvEnd = wv最終返済年月
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") <= Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            End If
            
        End If
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利変更年月CHECK_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利変更年月CHECK() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
    
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
'
End Sub

Private Sub 金利変更年月_LostFocus()
    L_金利変更年月日.Caption = ""

    金利変更年月 = C年月日.FormatDate("年月", 金利変更年月)
'
    If P8.FCStr(金利変更年月) <> "" Then
        Call 金利変更年月日_セット
    End If
End Sub

Private Sub 金利変更利率_LostFocus()
    金利変更利率 = P8.FFormat(金利変更利率, "#,##0.00000")
End Sub

'------------------------------------------------
' 金利変更年月日_セット
'------------------------------------------------
Private Sub 金利変更年月日_セット()
'
    Dim wv01 As Variant, wv02 As Variant
'
    On Error GoTo 金利変更年月日_セット_ERR
'
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    'GVar1 = MXA030_翌営業年月日計算(CDate(GVar1), wi支払日, wi営業日区分)
    wv01 = MBD010_利息計算年月日(CDate(GVar1), wi支払日, wi営業日区分, wi利息計算日数区分)
'
    If Format(GVar1, "yyyy/mm/01") = Format(wv初回返済年月, "yyyy/mm/01") Then
        If Format(wv01, "yyyy/mm/dd") = Format(wv初回返済実行日, "yyyy/mm/dd") Then
            wv01 = wv初回返済実行日
        End If
    End If
'
    If Format(GVar1, "yyyy/mm/01") = Format(wv最終返済年月, "yyyy/mm/01") Then
        If Format(wv01, "yyyy/mm/dd") = Format(wv最終返済実行日, "yyyy/mm/dd") Then
            wv01 = wv最終返済実行日
        End If
    End If
'
    L_金利変更年月日.Caption = Format(wv01, Gfmt年月日)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利変更年月日_セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利変更年月日_セット() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Private Sub 削除_Click()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer
    Dim FLG_DEL As Boolean
    Dim w金利変更年月日 As Variant, wv01 As Variant
    Dim ws01 As String
'
    On Error GoTo 削除_Click_ERR
'
    ' =========================================
    '           権限チェック
    ' =========================================
    Select Case GUserKen
        Case "0"
            '入力権限
        Case "1"
            '照会権限
            MsgBox "権限がありません", vbExclamation
            Exit Sub
        Case "5"
            '管理者権限
        Case Else
            MsgBox "権限がありません", vbExclamation
            Exit Sub
    End Select

    FLG_DEL = False
    
    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
    
    If C年月日.平成To西暦("年月", 金利変更年月, True) = 0 Then
        MsgBox "年月日が違います"
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
'
    GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
    If GRet = vbNo Then
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    FLG_DEL = True
    
    '----------< 取消データ削除 >------------------------------------------
    wstr = ""
    wstr = wstr + "Delete *"
    wstr = wstr + " From DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    GDb.Execute wstr
'
    '----------< テーブル Write >----------------------------------------------
    wstr1 = "Select * from " & wsTbl
    wstr1 = wstr1 & " Where " & P8.FCStr(L_番号1.Caption) & "='" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
    If Not wRs1.EOF Then

        j = 2 '2回目から始まる
        
        wstr = "Select * from DCHA010_Gridワーク"
        wstr = wstr & " Where テキスト1='" & wsBango & "'"
        wstr = wstr & " Order by 年月日1"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
            Do Until wRs.EOF
            
                ws01 = "金利変更" & CStr(j) & "回目年月"
                wRs1(ws01) = P8.FCDate(wRs("年月日1"))
    
                ws01 = "金利" & CStr(j) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値1"))
    
                j = j + 1
                
                wRs.MoveNext
            Loop
            
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        Else
        
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        End If
        
        wRs.Close
        Set wRs = Nothing
        
    End If
    wRs1.Close
    Set wRs1 = Nothing
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "削除しました。", vbInformation
'
    ' =========================================
    '                 初期設定
    ' =========================================
    If FLG_DEL = True Then
        金利変更年月 = ""
        金利変更利率 = 0
    
        L_金利変更年月日.Caption = ""
    End If
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号=" & wsBango & ","
    GLogStr = "年月日=" & Format(GVar1, Gfmt年月日)
    Call MXA030_LOG_WRITE("借入金金利入力登録", "削除", GLogStr)
'
    '取消 = 0
    
    'ワークテーブル作成とワークデータセット
    Call 金利ワークテーブル作成
    
    Call 画面セット
        
    Call CEkey.SetFs(金利変更年月, False)
'
    ' =========================================
    '               メッセージ
    ' =========================================
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
削除_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 削除_Click() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' 登録_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer
    Dim FLG_DEL As Boolean
    Dim w金利変更年月日 As Variant, wv01 As Variant
    Dim ws01 As String
'
    On Error GoTo 登録_Click_ERR
'
    ' =========================================
    '           権限チェック
    ' =========================================
    Select Case GUserKen
        Case "0"
            '入力権限
        Case "1"
            '照会権限
            MsgBox "権限がありません", vbExclamation
            Exit Sub
        Case "5"
            '管理者権限
        Case Else
            MsgBox "権限がありません", vbExclamation
            Exit Sub
    End Select

    FLG_DEL = False
    
    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
    
    GRet = 金利変更年月CHECK(Format(GVar1, "yyyy/mm/dd"))
    If GRet <> True Then
        MsgBox "年月日が違います"
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
    
    If FLG_MAX = True Then
        GRet = MsgBox("金利変更100回を越えると登録できません。", vbOKOnly)
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
'
    If C年月日.平成To西暦("年月", 金利変更年月, True) = 0 Then
        MsgBox "年月日が違います"
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
'
    If (Not IsNumeric(金利変更利率) And 金利変更利率 <> "") _
    Or P8.FCDbl(金利変更利率) >= 100 Or P8.FCDbl(金利変更利率) < 0 Then
        MsgBox "入力を確認してください"
        Call CEkey.SetFs(金利変更利率, True)
        Exit Sub
    End If
'
    '固定金利で金利変更Ｘ回目年月等があればエラー
    If P8.FCDbl(XMXA020_区分("金利種別", "固定金利")) = wi金利種別 Then
        If 金利変更年月 <> "" Then
            MsgBox "固定金利では設定できません"
            Call CEkey.SetFs(金利変更年月, True)
            Exit Sub
        End If
        If P8.FCDbl(金利変更利率) <> 0 Then
            MsgBox "固定金利では設定できません"
            Call CEkey.SetFs(金利変更利率, True)
            Exit Sub
        End If
    End If
    
    If 金利変更年月 = "" Then
        If P8.FCDbl(金利変更利率) <> 0 Then
            MsgBox "金利変更利率が違います"
            Call CEkey.SetFs(金利変更年月, True)
            Exit Sub
        End If
    Else
        If IsNumeric(金利変更利率) Then
            If CInt(金利変更利率) > 100 Then
                MsgBox "金利変更利率が大きい"
                Call CEkey.SetFs(金利変更利率, True)
                Exit Sub
            End If
        End If
    End If
'
    ' =========================================
    '             金利変更年月整合性check
    ' =========================================
'    w金利変更年月日 = C年月日.平成To西暦("年月日", P8.FCStr(L_金利変更年月日.Caption), True)
'    If Not IsNull(w金利変更年月日) Then
'        If CDate(w金利変更年月日) < CDate(wv初回返済実行日) Or CDate(w金利変更年月日) > CDate(wv最終返済実行日) Then
'            MsgBox "金利変更年月が誤りです"
'            Call CEkey.SetFs(金利変更年月, True)
'            Exit Sub
'        End If
'    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.EOF Then
        wRs.AddNew
    End If
        
        wRs("テキスト1") = wsBango
        
        wRs("年月日1") = C年月日.平成To西暦("年月", 金利変更年月, True)
        wRs("数値1") = P8.FCDbl(金利変更利率)
        
        wRs("取消フラグ") = 0 'P8.FCDbl(取消)
        
        'If P8.FCDbl(取消) = 1 Then
        '    FLG_DEL = True
        'End If
        
        wRs.Update
    
    wRs.Close
    Set wRs = Nothing
'
    '----------< 取消データ削除 >------------------------------------------
'    wstr = "Delete * from DCHA010_Gridワーク"
'    wstr = wstr & " Where 取消フラグ=1"
'    GDb.Execute wstr
'
    '----------< テーブル Write >----------------------------------------------
    wstr1 = "Select * from " & wsTbl
    wstr1 = wstr1 & " Where " & P8.FCStr(L_番号1.Caption) & "='" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
    If Not wRs1.EOF Then

        j = 2 '2回目から始まる
        
        wstr = "Select * from DCHA010_Gridワーク"
        wstr = wstr & " Where テキスト1='" & wsBango & "'"
        wstr = wstr & " Order by 年月日1"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
            Do Until wRs.EOF
            
                ws01 = "金利変更" & CStr(j) & "回目年月"
                wRs1(ws01) = P8.FCDate(wRs("年月日1"))
    
                ws01 = "金利" & CStr(j) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値1"))
    
                j = j + 1
                
                wRs.MoveNext
            Loop
            
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        Else
        
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        End If
        
        wRs.Close
        Set wRs = Nothing
        
    End If
    wRs1.Close
    Set wRs1 = Nothing
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号=" & wsBango & ","
    GLogStr = "年月日=" & Format(GVar1, "yyyy/mm/dd") & ","
    GLogStr = "利率=" & P8.FCDbl(金利変更利率)
    Call MXA030_LOG_WRITE("借入金金利入力登録", "更新", GLogStr)
'
    ' =========================================
    '                 初期設定
    ' =========================================
    If FLG_DEL = True Then
        金利変更年月 = ""
        金利変更利率 = 0
    
        L_金利変更年月日.Caption = ""
    End If
    
    '取消 = 0
    
    'ワークテーブル作成とワークデータセット
    Call 金利ワークテーブル作成
    
    Call 画面セット
        
    Call CEkey.SetFs(金利変更年月, False)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました。", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
登録_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 登録_Click() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    '----------< DataGrid Close >----------------------------------------------
    If Not DataGrid1.DataSource Is Nothing Then
        Set DataGrid1.DataSource = Nothing
    End If
    Adodc1.Recordset.Close
'
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
    
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
'
End Sub

