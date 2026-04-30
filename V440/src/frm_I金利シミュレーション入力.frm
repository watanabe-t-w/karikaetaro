VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_I金利シミュレーション入力 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "金利シミュレーション入力"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "frm_I金利シミュレーション入力.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   5400
      Top             =   6000
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin 借換たろう.ZU020_ComboBox 金利シミュレーションGP 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      ForeColor       =   -2147483640
      ForeColor       =   -2147483640
      TextWidth       =   600
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      P8_ListBoxMax   =   0
      P8_KeySort      =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "登録"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5400
      TabIndex        =   4
      Top             =   1440
      Width           =   4215
      Begin VB.CheckBox 削除 
         Caption         =   "削除"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   855
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
         Left            =   2640
         TabIndex        =   3
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox 年月日 
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox 増減利率 
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
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor_Shape1=   8454016
         BackColor_Shape2=   8421504
         BorderColor_Shape1=   49152
         BorderColor_Shape2=   4210752
         ForeColor       =   255
         Caption         =   "新規変更"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label L_金利シミュレーションGP 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   12
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "金利ｼﾐｭﾚｰｼｮﾝGP"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "年月日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "増減利率"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
   End
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
      Left            =   7680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4965
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   5175
      _ExtentX        =   9128
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
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "金利ｼﾐｭﾚｰｼｮﾝ入力"
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
   Begin VB.Label Label45 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "金利ｼﾐｭﾚｰｼｮﾝGP"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frm_I金利シミュレーション入力"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "金利シミュレーション入力"

Dim wDb As New ADODB.Connection
Dim wRs As ADODB.Recordset
Dim wRs1 As ADODB.Recordset
Dim wstr As String

Dim wslog As String

Dim FLG_New As Boolean
Dim FLG_MAX As Boolean



'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    ' =========================================
    '                 初期設定
    ' =========================================
    Me.Left = G_LEFT
    Me.Top = G_TOP
    '----------< GCurDir GTemp Open >-----------------------------------------------
    
    Call 登録後初期セット

End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    DoEvents
'
    ' =========================================
    '             コンボボックス
    ' =========================================
    Call コンボセット
'
    Call CEkey.AllSelect
'
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
' AdodcRefresh
'------------------------------------------------
Private Sub AdodcRefresh()

    Dim I As Long
    Dim j As Long
    Dim p金利シミュレーションGP As String
    
    Dim wsql As String
    
    Dim w年月日(100) As String
    Dim w増減率(100) As String
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
    '              メインクエリ DAAA117_金利シミュレーション利率
    ' =========================================
    
    Call 金利ワークテーブル作成
    
    
    GWhere = ""
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " テキスト2 As Grd回,"
    wstr = wstr + " format(年月日1,'" & Gfmt年月日 & "') As Grd年月日,"
    wstr = wstr + " format(数値1,'#,##0.00000') As Grd増減利率"
    'wstr = wstr + " IIF(取消フラグ = 0,'','×') As Grd取消"
    wstr = wstr + " From DCHA010_Gridワーク"
    wstr = wstr + GWhere
    wstr = wstr + " Order By 年月日1"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("年月日", "年月日", 2000, "R")
        Call XZMA010_DataGrid_Set("増減利率", "増減利率", 2000, "R")
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
    Call CEkey.SetFs(年月日, True)
    
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd年月日")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        年月日 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd年月日"))
    On Error GoTo 0
    
    年月日 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd年月日"))
    
    Call 画面セット(True)
   
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(年月日, True)

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
    Call 画面セット(False)
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット(pGridClick As Boolean) As Boolean
'
    Dim ws01 As String
    Dim w金利シミュレーションGP As String
    Dim w年月日 As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    FLG_New = True
    増減利率 = ""
    削除.Value = 0
'
'    ' =========================================
'    '                パラメータ
'    ' =========================================
    w金利シミュレーションGP = P8.FCStr(L_金利シミュレーションGP)
    w年月日 = C年月日.平成To西暦("年月日", 年月日)
'
'    ' =========================================
'    '            マスタ セット
'    ' =========================================
    wstr = ""
    wstr = wstr + "Select * From  DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(w年月日, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.EOF Then
    '新規登録
        削除.Enabled = False
        新規変更.Caption = "新規"
        Me.増減利率.Text = ""
        Me.削除.Value = 0
    Else
    '変更登録
        削除.Enabled = True
        新規変更.Caption = "変更"
        Me.年月日.Text = Format(wRs("年月日1"), Gfmt年月日)
        Me.増減利率.Text = Format(wRs("数値1"), "#0.00000")
        Me.削除.Value = 0
    End If
    
    wRs.Close
    Set wRs = Nothing
'
''
''
'    '------------------------------------------
'    '          ** グリッドコントロール **
'    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If
'
    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
''
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

'------------------------------------------------
' USERID_GotFocus
'------------------------------------------------
Private Sub USERID_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' USERID_LostFocus
'------------------------------------------------
Private Sub USERID_LostFocus()
'
    On Error GoTo USERID_LostFocus_ERR
'
'    Call P8.FCControlLeft(ユーザーID, 10)
'
'    Select Case Screen.ActiveControl.Name
'        Case "閉じる", "DataGrid1", "USERID", "入力クリア"
'            Exit Sub
'    End Select
'
'    If ユーザーID = "" Then
'        MsgBox "コードを入力してください"
'        Call CEkey.SetFs(ユーザーID, True)
'        Exit Sub
'    End If
''
'    Select Case Screen.ActiveControl.Name
'        Case "保存"
'            Call CEkey.SetFs(ユーザー名, True)
'            MsgBox "該当データをセットします。保存処理は行いません。"
'            Exit Sub
'    End Select
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
USERID_LostFocus_ERR:
    pERR_MES = pPROGRAM_ID + "/ USERID_LostFocus() でエラー" + vbCrLf + vbCrLf + _
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
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w金利シミュレーションGP As String
    Dim w年月日 As String
'
    w金利シミュレーションGP = L_金利シミュレーションGP
    w年月日 = ""
    
    増減利率 = ""
    Call 画面セット(False)
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + w年月日 + "'")
'    Call CEkey.SetFs(Me.ユーザーID, True)
'
End Sub

Private Sub 金利シミュレーションGP_Change()
    
    L_金利シミュレーションGP = 金利シミュレーションGP.Text
    
'    Call 金利ワークテーブル作成
    
    Call AdodcRefresh
    
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim ws01 As String
    Dim wstr1 As String
    
    Dim j As Long
    
    Dim p金利シミュレーションGP As String
    Dim p年月日 As String
    Dim p増減利率 As Double
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

    ' =========================================
    '               パラメータセット
    ' =========================================
    p金利シミュレーションGP = P8.FCStr(Me.金利シミュレーションGP.Text)
    p年月日 = P8.FCStr(Me.年月日.Text)
    p増減利率 = P8.FCDbl(Me.増減利率.Text)

'
    ' =========================================
    '               入力チェック
    ' =========================================
'
    If 金利シミュレーションGP.Text = "" Then
        MsgBox "金利シミュレーションGPが未入力です", vbExclamation
        Call CEkey.SetFs(金利シミュレーションGP, True)
        Exit Sub
    End If

    If 金利シミュレーションGP.Text <> "" And 金利シミュレーションGP.P8_Name = "" Then
        MsgBox "金利シミュレーションGPが不正です", vbExclamation
        Call CEkey.SetFs(金利シミュレーションGP, True)
        Exit Sub
    End If
    
    If C年月日.平成To西暦("年月日", p年月日) = 0 Then
        MsgBox "年月日が不正です", vbExclamation
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
    
    If p増減利率 >= 100 Then
        MsgBox "増減利率が大きすぎます", vbExclamation
        Call CEkey.SetFs(増減利率, True)
        Exit Sub
    End If
    
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    If 削除.Value = 1 Then
    '削除
        If MsgBox("削除します。よろしいですか？", vbYesNo) = vbYes Then
            wstr = ""
            wstr = wstr + "DELETE"
            wstr = wstr + " From DCHA010_Gridワーク"
            wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
            GDb.Execute wstr
        Else
            Exit Sub
        End If
    Else
    '新規、変更
        wstr = ""
        wstr = wstr + "Select *"
        wstr = wstr + " From DCHA010_Gridワーク"
        wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
        
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            wRs.AddNew
        End If
            
            wRs("テキスト1") = p金利シミュレーションGP
            
            wRs("年月日1") = C年月日.平成To西暦("年月日", 年月日)
            wRs("数値1") = P8.FCDbl(増減利率)
            
            wRs.Update
        
        wRs.Close
        Set wRs = Nothing
        
    End If

    
    '----------< テーブル Write >----------------------------------------------
    wstr1 = "Select * from DAAA117_金利シミュレーション利率"
    wstr1 = wstr1 & " Where 金利グループ区分 = '" & p金利シミュレーションGP & "'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
    If Not wRs1.EOF Then

        j = 1 '2回目から始まる
        
        wstr = "Select * from DCHA010_Gridワーク"
        wstr = wstr & " Where テキスト1='" & p金利シミュレーションGP & "'"
        wstr = wstr & " Order by 年月日1"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
            Do Until wRs.EOF
            
                ws01 = "年月日" & CStr(j)
                wRs1(ws01) = P8.FCDate(wRs("年月日1"))
    
                ws01 = "利率増減率" & CStr(j)
                wRs1(ws01) = P8.FCDbl(wRs("数値1"))
    
                j = j + 1
                
                wRs.MoveNext
            Loop
            
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "年月日" & CStr(j)
                    wRs1(ws01) = Null
        
                    ws01 = "利率増減率" & CStr(j)
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        Else
        
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "年月日" & CStr(j)
                    wRs1(ws01) = Null
        
                    ws01 = "利率増減率" & CStr(j)
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        End If
        
        wRs.Close
        Set wRs = Nothing
        
    Else
        wRs1.AddNew
        wRs1("金利グループ区分") = p金利シミュレーションGP
        wRs1("年月日1") = C年月日.平成To西暦("年月日", p年月日)
        wRs1("利率増減率1") = p増減利率
        
        wRs1.Update
        
    End If
    wRs1.Close
    Set wRs1 = Nothing
    
    
    If 削除.Value = 1 Then
        MsgBox "削除しました", vbInformation
        Me.年月日 = ""
        Me.増減利率 = ""
    Else
        MsgBox "登録しました", vbInformation
    End If
    
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    If 新規変更.Caption = "新規" Then
        wslog = "追加"
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 0 Then
        wslog = "更新"
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 1 Then
        wslog = "削除"
    End If
    GLogStr = "金利ｼﾐｭﾚｰｼｮﾝGP=" & p金利シミュレーションGP & ","
    GLogStr = GLogStr & "年月日=" & p年月日 & ","
    GLogStr = GLogStr & "増減利率=" & p増減利率
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
    
'
    ' =========================================
    '                テーブル変更
    ' =========================================
    Call MAA070_金利グループ設定
    Call MAA070_金利SM率設定

'    Adodc1.Refresh
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(Me.年月日, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
登録_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 保存_Click() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Private Sub コンボセット()
    
    With 金利シミュレーションGP
        .P8_Db = GDb
        
        wstr = "Select * From DAAA115_金利シミュレーショングループ"
        wstr = wstr + " Where 取消フラグ = 0"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 2
        .P8_ListBoxMax = 500
        .P8_KeyName = "金利グループ区分"
        .P8_ItemName = "金利グループ名"
    End With
    金利シミュレーションGP.CreateCombo

End Sub

'------------------------------------------------
' 金利ワークテーブル作成
'------------------------------------------------
Private Sub 金利ワークテーブル作成()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim wsql As String
    
    Dim wSMG As String
    
    Dim j As Integer
    Dim ws01 As String
'
    On Error GoTo 金利ワークテーブル作成_ERR
    
    wSMG = P8.FCStr(L_金利シミュレーションGP)
'
    '----------< ワークテーブル削除 >------------------------------------------
    wstr = "Delete * from DCHA010_Gridワーク"
    GDb.Execute wstr
'
'
    '----------< テーブル Write >----------------------------------------------
'    wstr = "Select * from DCHA010_Gridワーク"
'    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr1 = "Select * from DAAA117_金利シミュレーション利率"
        wstr1 = wstr1 & " Where 金利グループ区分 ='" & wSMG & "'"
        Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        If Not wRs1.EOF Then
            
            For j = 1 To 100
                
                ws01 = "年月日" & CStr(j)
                If Not IsNull(P8.FCDate(wRs1(ws01))) Then
                    
'                    wRs.AddNew
'
'                    wRs("テキスト1") = wSMG
'                    wRs("テキスト2") = j
'
'                    ws01 = "年月日" & CStr(j)
'                    wRs("年月日1") = P8.FCDate(wRs1(ws01))
'
'                    ws01 = "利率増減率" & CStr(j)
'                    wRs("数値1") = P8.FCDbl(wRs1(ws01))
'
'                    wRs.Update
                    
                    wsql = ""
                    wsql = wsql & "INSERT INTO DCHA010_Gridワーク"
                    wsql = wsql & "(テキスト1,テキスト2,年月日1,数値1)"
                    wsql = wsql & " VALUES("
                    wsql = wsql & " '" & wSMG & "',"
                    wsql = wsql & " '" & j & "',"
                    wsql = wsql & " '" & P8.FCDate(wRs1("年月日" & CStr(j))) & "',"
                    wsql = wsql & " '" & P8.FCDbl(wRs1("利率増減率" & CStr(j))) & "')"
                    
                    GDb.Execute wsql
                    
                    wstr = "SELECT * FROM DCHA010_Gridワーク"
                    Call AdoRecordsetOpen(GDb, wRs, wstr)
                    wRs.Close
                    Set wRs = Nothing
                    
                End If
                
            Next
            
            If Not IsNull(P8.FCDate(wRs1("年月日100"))) Then
                FLG_MAX = True
            End If
        
        End If
        wRs1.Close
        Set wRs1 = Nothing

'    wRs.Close
'    Set wRs = Nothing
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

Private Sub ユーザーID_LostFocus()

    Call 画面セット(False)
    
End Sub


'------------------------------------------------
' 入力クリア_Click
'------------------------------------------------
Private Sub 入力クリア_Click()
    Call 登録後初期セット

End Sub

Private Sub 削除データを表示_Click()
    
    Call AdodcRefresh
    
End Sub

Private Sub 年月日_LostFocus()

    年月日 = C年月日.FormatDate("年月日", 年月日)
    Call 画面セット(True)
    
End Sub

Private Sub 年月日_GotFocus()
    
    Call CEkey.AllSelect
    
End Sub

Private Sub 増減利率_LostFocus()

    増減利率 = Format(増減利率, "#0.00000")

End Sub

Private Sub 増減利率_GotFocus()

    Call CEkey.AllSelect

End Sub

'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    Unload Me

End Sub






