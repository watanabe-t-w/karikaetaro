VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_M祝日マスタ 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "祝日マスタ"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "カレンダー"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   6480
      Width           =   4935
      Begin VB.TextBox 印刷年 
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   8
         ToolTipText     =   "YYYY"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton 印刷 
         Caption         =   "表示"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "西暦 年"
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
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "CSVデータ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   4935
      Begin VB.CommandButton 出力 
         Caption         =   "出力"
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
         Left            =   2880
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox 取込年 
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "YYYY"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton 取込 
         Caption         =   "取込"
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
         Left            =   2880
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '実線
         Caption         =   "西暦 年"
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
         TabIndex        =   21
         Top             =   360
         Width           =   1215
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
      Left            =   3120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "検索"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   4935
      Begin VB.CommandButton ログ出力 
         Caption         =   "日付チェック"
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
         Left            =   2880
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton 検索 
         Caption         =   "検索(F1)"
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
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox 検索年 
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   0
         ToolTipText     =   "YYYY"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "西暦 年"
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
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   4935
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
         IMEMode         =   2  'ｵﾌ
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton 登録 
         Caption         =   "登録(F11)"
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
         Left            =   2880
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox 削除 
         Caption         =   "削除"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox 名称 
         Height          =   330
         IMEMode         =   4  '全角ひらがな
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1080
         Width           =   3495
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   120
         TabIndex        =   17
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
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "名称"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
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
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8685
      Left            =   5280
      TabIndex        =   11
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   15319
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "祝日マスタ"
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
End
Attribute VB_Name = "frm_M祝日マスタ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2017/12/01 Add 祝日マスタ
Option Explicit
'
Private Const pPROGRAM_ID As String = "祝日マスタ"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim pcnt As Integer
Dim wslog As String
Dim FLG_New As Boolean

''------------------------------------------------
'' Form_Initialize
''------------------------------------------------
'Private Sub Form_Initialize()
''
'    ' =========================================
'    '             MAA100_SERIAL
'    ' =========================================
'    GRet = MAA100_SERIAL()
'    If GRet <> True Then
'        GRet = MsgBox("シリアル情報が正しくありません。" + Chr(13) + vbCrLf + GProduct + "を終了します", vbOKOnly + vbCritical)
'        GDb.Close
'        Set GDb = Nothing
'
'        End
'    End If
'
'End Sub
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
    
    pcnt = 0
    ReDim GCal(pcnt)
    
    Call 登録後初期セット

End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    DoEvents
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
    wstr = ""
    wstr = wstr & "SELECT"
    wstr = wstr + " Format(年月日,'yyyy/mm/dd') As Grd年月日,"
    wstr = wstr & " 名称 AS Grd名称"
    wstr = wstr & " FROM DACA010_祝日マスタ"
    wstr = wstr & " WHERE (0=0)"
    If Me.検索年.Text <> "" Then
        wstr = wstr & " And Format(年月日,'yyyymmdd') >= '" & Me.検索年.Text & "0101'"
        wstr = wstr & " And Format(年月日,'yyyymmdd') <= '" & Me.検索年.Text & "1231'"
    End If
    wstr = wstr & " ORDER BY Format(年月日,'yyyy') desc,年月日"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("年月日", "年月日", 1400, "L")
        Call XZMA010_DataGrid_Set("名称", "名称", 4000, "L")
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
    
    Call 画面セット(True)
   
'    If DataGrid1.Splits.Count <> 1 Then
'        DataGrid1.Splits.Remove 1
'    End If

    Call CEkey.SetFs(名称, True)

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
' 画面セット
'------------------------------------------------
Private Function 画面セット(pGridClick As Boolean) As Boolean
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    Me.名称.Text = ""
    Me.削除.Value = 0
    
    ' =========================================
    '                パラメータ
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    ' =========================================
    '            ユーザーマスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "* "
    wstr = wstr & "FROM DACA010_祝日マスタ "
    wstr = wstr & "WHERE Format(年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
    '新規登録
        新規変更.Caption = "新規"
        Me.名称.Text = ""
    Else
    '変更登録
        新規変更.Caption = "変更"
        Me.名称.Text = P8.FCStr(wRs("名称"))
    End If
'
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
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


'------------------------------------------------
' 年月日_GotFocus
'------------------------------------------------
Private Sub 年月日_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w年月日 As String
'
    検索年.Text = ""
    
    w年月日 = 年月日
    
    年月日 = ""
    Call 画面セット(False)
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + w年月日 + "'")
    Call CEkey.SetFs(Me.年月日, True)
'
End Sub

Private Sub 名称_LostFocus()
    Call P8.FCControlLeft(名称, 50)
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim FlgDell As Boolean
    Dim ws01 As String
    Dim p名称 As String
'
    On Error GoTo 保存_Click_ERR
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
'
    ' =========================================
    '               入力チェック
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
'
    ' =========================================
    '            祝日マスタ 更新処理
    ' =========================================
    p名称 = P8.FCStr(Me.名称.Text)
    FlgDell = False
    
    If Me.新規変更.Caption = "新規" Then
    '新規登録
        
        p名称 = LTrim(p名称)
        名称.Text = p名称
            
        wstr = ""
        wstr = wstr & "Select *"
        wstr = wstr & " From DACA010_祝日マスタ"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            wRs.AddNew
                        
            wRs("年月日") = CDate(GVar1)
            wRs("名称") = p名称
            wRs("区分") = 1
    
            wRs.Update
            
            GRet = MsgBox("登録しました。", vbOKOnly)
    
        wRs.Close
        Set wRs = Nothing
    Else
    '更新
        If Me.削除.Value = 0 Then
            wstr = ""
            wstr = wstr & "Select *"
            wstr = wstr & " From DACA010_祝日マスタ"
            wstr = wstr & " WHERE Format(年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
            Call AdoRecordsetOpen(GDb, wRs, wstr)
                wRs("名称") = p名称
                wRs("区分") = 1
                
                wRs.Update
        
                GRet = MsgBox("登録しました。", vbOKOnly)
            
            wRs.Close
            Set wRs = Nothing
        Else
        '削除
            
            GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
            If GRet = vbNo Then
                Exit Sub
            End If
            
            FlgDell = True
    
            wstr = ""
            wstr = wstr & "Delete * "
            wstr = wstr & " From DACA010_祝日マスタ"
            wstr = wstr & " WHERE Format(年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
            GDb.Execute wstr
            
            Me.年月日.Text = ""
            Me.削除.Value = 0
        End If
        
    End If
'
    ' =========================================
    '                テーブル変更
    ' =========================================
    Call MAA090_祝日マスタ設定
'
    Adodc1.Refresh
'
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
    GLogStr = "年月日=" & 年月日 & ","
    GLogStr = GLogStr & "名称=" & p名称
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
    
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(名称, True)
'
    ' =========================================
    '   実行日、初回返済年月日、最終返済年月日、内入年月日 CHECk
    ' =========================================
    If 年月日.Text <> "" Then
        If FlgDell = False Then
            pcnt = 0
            ReDim GCal(pcnt)
            GRet = 営業年月日CHECK(年月日.Text)
            If GRet = True Then
                GRpt.帳票名 = "祝日マスタログ一覧"
                RCA020_祝日マスタログ一覧.Show
                GRet = MsgBox("登録した祝日を含む借入金登録データがあります。" & vbCr & vbLf & "祝日マスタログ一覧を確認してください。", vbOKOnly + vbExclamation, "祝日マスタログ一覧を確認してください")
            End If
        End If
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
保存_Click_ERR:
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

Private Sub 年月日_LostFocus()
'
    Dim ws01 As String
    Dim w年月日 As Date
'
    Call P8.FCControlLeft(年月日, 30)
    
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1" ', "年月日"
            Exit Sub
    End Select
   
    If 年月日 = "" Then
'        MsgBox "コードを入力してください"
'        Call CEkey.SetFs(年月日, True)
        Exit Sub
    Else
        If InStrRev(年月日, "年") Then
            GVar1 = C年月日.平成To西暦("", 年月日)
            If GVar1 = 0 Then
                MsgBox "年月日を入力してください", vbExclamation
                年月日 = ""
                名称 = ""
                Call CEkey.SetFs(年月日, True)
                Exit Sub
            End If
        Else

        End If
    End If
       
    年月日 = C年月日.FormatDate("年月日", 年月日)
    If C年月日.平成To西暦("年月", 年月日) = 0 Then
        MsgBox "年月日が違います", vbExclamation
        年月日 = ""
        名称 = ""
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    Call 画面セット(False)
    
End Sub

Private Sub 検索_Click()
    
    If Not P8.FIsInt(検索年) Then
        検索年 = ""
        Call CEkey.SetFs(検索年, True)
    End If
    
    If P8.FCDbl(検索年) <= 0 Then
        検索年 = ""
        Call CEkey.SetFs(検索年, True)
    End If
    
    If Len(P8.FCStr(検索年)) < 4 Then
        検索年 = ""
        Call CEkey.SetFs(検索年, True)
    End If
        
    Call AdodcRefresh
    
End Sub

'------------------------------------------------
' CSVファイル取込
'------------------------------------------------
Private Sub 取込_Click()

    Dim wsRet As String
    Dim wNen As Integer
    Dim wRet As Boolean
    
    If Not P8.FIsInt(取込年) Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(取込年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
    
    If P8.FCDbl(取込年) <= 0 Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(取込年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
    
    If Len(P8.FCStr(取込年)) < 4 Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(取込年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
'
    GRet = MsgBox("CSVファイルをインポートします。" _
                    & vbCrLf & vbCrLf & "既存のデータは上書きされます。" & vbCrLf & _
                    "よろしいですか？", vbExclamation + vbYesNo, "取込")
    If GRet = vbNo Then
        Exit Sub
    End If
'
    wsRet = MXA040_COMDLG(CommonDialog1, "CSVファイル選択", "", _
                            "テキストファイル(*.csv)|*.csv", "祝日データ.csv")
    If wsRet = "" Then
        Exit Sub
    ElseIf wsRet = "キャンセル" Then
        Exit Sub
    End If
'
    wNen = P8.FCDbl(取込年)
    GRet = MXA040_祝日データ取込(wsRet, wNen)
    If GRet <> True Then
        MsgBox "CSVファイルをインポートできませんでした", vbInformation
        
        Exit Sub
    End If
'
    ' =========================================
    '           Csv File Drive
    ' =========================================
    Call MX040_CsvPath
'
    ' =========================================
    '                テーブル変更
    ' =========================================
    Call MAA090_祝日マスタ設定
'
    Me.検索年.Text = Me.取込年.Text
    Call 画面セット(False)
'
    ' =========================================
    '   実行日、初回返済年月日、最終返済年月日、内入年月日 CHECk
    ' =========================================
    pcnt = 0
    ReDim GCal(pcnt)
    ' =========================================
    '            ユーザーマスタ セット
    ' =========================================
    wRet = False
    GRet = False
    
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "* "
    wstr = wstr & "FROM DACA010_祝日マスタ "
    wstr = wstr & " WHERE Format(年月日,'yyyymmdd') >= '" & P8.FCStr(wNen) & "0101" & "'"
    wstr = wstr & " and Format(年月日,'yyyymmdd') <= '" & P8.FCStr(wNen) & "1231" & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            
            wRet = 営業年月日CHECK(wRs("年月日"))
            If wRet = True Then
                GRet = wRet
            End If
            
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing

    If GRet = True Then
        GRpt.帳票名 = "祝日マスタログ一覧"
        RCA020_祝日マスタログ一覧.Show
        GRet = MsgBox("取り込んだ祝日を含む借入金登録データがあります。" & vbCr & vbLf & "祝日マスタログ一覧を確認してください。", vbOKOnly + vbExclamation, "祝日マスタログ一覧を確認してください")
    End If
'
    ' =========================================
    '               メッセージ
    ' =========================================
    Me.取込年.Text = ""
    MsgBox "CSVファイルをインポートしました", vbInformation
'
End Sub

'------------------------------------------------
' 営業年月日CHECK
'------------------------------------------------
Private Function 営業年月日CHECK(pDate As String) As Boolean
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
        
    Dim l As Integer
    Dim wd01 As Date
'
    営業年月日CHECK = False
'
    wd01 = CDate(pDate)

    '実行日、初回返済年月日、最終返済年月日
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.実行日"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    wstr1 = wstr1 + ",K.初回返済実行日"
    wstr1 = wstr1 + ",K.最終返済実行日"
    wstr1 = wstr1 + " From DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号"
    wstr1 = wstr1 + " Where format(K.実行日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "'"
    wstr1 = wstr1 + " OR format(K.初回返済実行日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "'"
    wstr1 = wstr1 + " OR format(K.最終返済実行日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
            If Format(wRs1("実行日"), "yyyymmdd") = Format(wd01, "yyyymmdd") Then
                pcnt = pcnt + 1
                ReDim Preserve GCal(pcnt)
                GCal(pcnt).借入番号 = wRs1("借入番号")
                GCal(pcnt).銀行番号 = wRs1("銀行番号")
                GCal(pcnt).銀行名 = wRs1("銀行名")
                GCal(pcnt).確認年月日 = "実行日:" & wRs1("実行日")
            
                営業年月日CHECK = True
            End If
            
            If Format(wRs1("初回返済実行日"), "yyyymmdd") = Format(wd01, "yyyymmdd") Then
                pcnt = pcnt + 1
                ReDim Preserve GCal(pcnt)
                GCal(pcnt).借入番号 = wRs1("借入番号")
                GCal(pcnt).銀行番号 = wRs1("銀行番号")
                GCal(pcnt).銀行名 = wRs1("銀行名")
                GCal(pcnt).確認年月日 = "初回返済年月日:" & wRs1("初回返済実行日")
            
                営業年月日CHECK = True
            End If
            
            If Format(wRs1("最終返済実行日"), "yyyymmdd") = Format(wd01, "yyyymmdd") Then
                pcnt = pcnt + 1
                ReDim Preserve GCal(pcnt)
                GCal(pcnt).借入番号 = wRs1("借入番号")
                GCal(pcnt).銀行番号 = wRs1("銀行番号")
                GCal(pcnt).銀行名 = wRs1("銀行名")
                GCal(pcnt).確認年月日 = "最終返済年月日:" & wRs1("最終返済実行日")
            
                営業年月日CHECK = True
            End If
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing
'
    '実際年月日、利息計算年月日
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.実行日"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    wstr1 = wstr1 + ",KT.実際年月日"
    wstr1 = wstr1 + ",KT.利息計算年月日"
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金明細TR As KT"
    wstr1 = wstr1 + " ON K.借入番号 = KT.借入番号"
    wstr1 = wstr1 + " Where K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
    wstr1 = wstr1 + " And (format(KT.実際年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "'"
    wstr1 = wstr1 + " OR format(KT.利息計算年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
            If Format(wRs1("実際年月日"), "yyyymmdd") = Format(wd01, "yyyymmdd") Then
                pcnt = pcnt + 1
                ReDim Preserve GCal(pcnt)
                GCal(pcnt).借入番号 = wRs1("借入番号")
                GCal(pcnt).銀行番号 = wRs1("銀行番号")
                GCal(pcnt).銀行名 = wRs1("銀行名")
                GCal(pcnt).確認年月日 = "返済年月日:" & wRs1("実際年月日")
            
                営業年月日CHECK = True
            End If
            
            If Format(wRs1("利息計算年月日"), "yyyymmdd") = Format(wd01, "yyyymmdd") Then
                pcnt = pcnt + 1
                ReDim Preserve GCal(pcnt)
                GCal(pcnt).借入番号 = wRs1("借入番号")
                GCal(pcnt).銀行番号 = wRs1("銀行番号")
                GCal(pcnt).銀行名 = wRs1("銀行名")
                GCal(pcnt).確認年月日 = "利息計算年月日:" & wRs1("利息計算年月日")
            
                営業年月日CHECK = True
            End If
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing
'
    'DBDA010_借入金内入1
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    For l = 1 To 80
        wstr1 = wstr1 + ",KU.内入" & CStr(l) & "回目年月日"
    Next l
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金内入1 As KU"
    wstr1 = wstr1 + " ON K.借入番号 = KU.借入番号"
    wstr1 = wstr1 + " Where (format(KU.内入1回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    For l = 2 To 80
        wstr1 = wstr1 + " OR (format(KU.内入" & CStr(l) & "回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Next l
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
        
            For l = 1 To 80
                If P8.FCIsDate(wRs1("内入" & CStr(l) & "回目年月日")) Then
                If CDate(wRs1("内入" & CStr(l) & "回目年月日")) = wd01 Then
                    pcnt = pcnt + 1
                    ReDim Preserve GCal(pcnt)
                    GCal(pcnt).借入番号 = wRs1("借入番号")
                    GCal(pcnt).銀行番号 = wRs1("銀行番号")
                    GCal(pcnt).銀行名 = wRs1("銀行名")
                    GCal(pcnt).確認年月日 = "内入" & CStr(l) & "回目年月日:" & wRs1("内入" & CStr(l) & "回目年月日")
            
                    営業年月日CHECK = True
                End If
                End If
            Next l
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing

    'DBDA010_借入金内入2
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    For l = 81 To 160
        wstr1 = wstr1 + ",KU.内入" & CStr(l) & "回目年月日"
    Next l
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金内入2 As KU"
    wstr1 = wstr1 + " ON K.借入番号 = KU.借入番号"
    wstr1 = wstr1 + " Where (format(KU.内入81回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    For l = 82 To 160
        wstr1 = wstr1 + " OR (format(KU.内入" & CStr(l) & "回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Next l
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
        
            For l = 81 To 160
                If P8.FCIsDate(wRs1("内入" & CStr(l) & "回目年月日")) Then
                    If CDate(wRs1("内入" & CStr(l) & "回目年月日")) = wd01 Then
                        pcnt = pcnt + 1
                        ReDim Preserve GCal(pcnt)
                        GCal(pcnt).借入番号 = wRs1("借入番号")
                        GCal(pcnt).銀行番号 = wRs1("銀行番号")
                        GCal(pcnt).銀行名 = wRs1("銀行名")
                        GCal(pcnt).確認年月日 = "内入" & CStr(l) & "回目年月日:" & wRs1("内入" & CStr(l) & "回目年月日")
                
                        営業年月日CHECK = True
                    End If
                End If
            Next l
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing

    'DBDA010_借入金内入3
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    For l = 161 To 240
        wstr1 = wstr1 + ",KU.内入" & CStr(l) & "回目年月日"
    Next l
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金内入3 As KU"
    wstr1 = wstr1 + " ON K.借入番号 = KU.借入番号"
    wstr1 = wstr1 + " Where (format(KU.内入161回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    For l = 162 To 240
        wstr1 = wstr1 + " OR (format(KU.内入" & CStr(l) & "回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Next l
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
        
            For l = 161 To 240
                If P8.FCIsDate(wRs1("内入" & CStr(l) & "回目年月日")) Then
                    If CDate(wRs1("内入" & CStr(l) & "回目年月日")) = wd01 Then
                        pcnt = pcnt + 1
                        ReDim Preserve GCal(pcnt)
                        GCal(pcnt).借入番号 = wRs1("借入番号")
                        GCal(pcnt).銀行番号 = wRs1("銀行番号")
                        GCal(pcnt).銀行名 = wRs1("銀行名")
                        GCal(pcnt).確認年月日 = "内入" & CStr(l) & "回目年月日:" & wRs1("内入" & CStr(l) & "回目年月日")
                
                        営業年月日CHECK = True
                    End If
                End If
            Next l
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing

    'DBDA010_借入金内入4
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    For l = 241 To 320
        wstr1 = wstr1 + ",KU.内入" & CStr(l) & "回目年月日"
    Next l
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金内入4 As KU"
    wstr1 = wstr1 + " ON K.借入番号 = KU.借入番号"
    wstr1 = wstr1 + " Where (format(KU.内入241回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    For l = 242 To 320
        wstr1 = wstr1 + " OR (format(KU.内入" & CStr(l) & "回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Next l
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
        
            For l = 241 To 320
                If P8.FCIsDate(wRs1("内入" & CStr(l) & "回目年月日")) Then
                    If CDate(wRs1("内入" & CStr(l) & "回目年月日")) = wd01 Then
                        pcnt = pcnt + 1
                        ReDim Preserve GCal(pcnt)
                        GCal(pcnt).借入番号 = wRs1("借入番号")
                        GCal(pcnt).銀行番号 = wRs1("銀行番号")
                        GCal(pcnt).銀行名 = wRs1("銀行名")
                        GCal(pcnt).確認年月日 = "内入" & CStr(l) & "回目年月日:" & wRs1("内入" & CStr(l) & "回目年月日")
                
                        営業年月日CHECK = True
                    End If
                End If
            Next l
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing

    'DBDA010_借入金内入5
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    For l = 321 To 400
        wstr1 = wstr1 + ",KU.内入" & CStr(l) & "回目年月日"
    Next l
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金内入5 As KU"
    wstr1 = wstr1 + " ON K.借入番号 = KU.借入番号"
    wstr1 = wstr1 + " Where (format(KU.内入321回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    For l = 322 To 400
        wstr1 = wstr1 + " OR (format(KU.内入" & CStr(l) & "回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Next l
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
        
            For l = 321 To 400
                If CDate(wRs1("内入" & CStr(l) & "回目年月日")) = wd01 Then
                    If P8.FCIsDate(wRs1("内入" & CStr(l) & "回目年月日")) Then
                        pcnt = pcnt + 1
                        ReDim Preserve GCal(pcnt)
                        GCal(pcnt).借入番号 = wRs1("借入番号")
                        GCal(pcnt).銀行番号 = wRs1("銀行番号")
                        GCal(pcnt).銀行名 = wRs1("銀行名")
                        GCal(pcnt).確認年月日 = "内入" & CStr(l) & "回目年月日:" & wRs1("内入" & CStr(l) & "回目年月日")
                
                        営業年月日CHECK = True
                    End If
                End If
            Next l
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing

    'DBDA010_借入金内入6
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    For l = 401 To 480
        wstr1 = wstr1 + ",KU.内入" & CStr(l) & "回目年月日"
    Next l
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金内入6 As KU"
    wstr1 = wstr1 + " ON K.借入番号 = KU.借入番号"
    wstr1 = wstr1 + " Where (format(KU.内入401回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    For l = 402 To 480
        wstr1 = wstr1 + " OR (format(KU.内入" & CStr(l) & "回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Next l
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
        
            For l = 401 To 480
                If P8.FCIsDate(wRs1("内入" & CStr(l) & "回目年月日")) Then
                    If CDate(wRs1("内入" & CStr(l) & "回目年月日")) = wd01 Then
                        pcnt = pcnt + 1
                        ReDim Preserve GCal(pcnt)
                        GCal(pcnt).借入番号 = wRs1("借入番号")
                        GCal(pcnt).銀行番号 = wRs1("銀行番号")
                        GCal(pcnt).銀行名 = wRs1("銀行名")
                        GCal(pcnt).確認年月日 = "内入" & CStr(l) & "回目年月日:" & wRs1("内入" & CStr(l) & "回目年月日")
                
                        営業年月日CHECK = True
                    End If
                End If
            Next l
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing

    'DBDA010_借入金内入7
    wstr1 = ""
    wstr1 = wstr1 + "Select "
    wstr1 = wstr1 + " K.借入番号"
    wstr1 = wstr1 + ",K.銀行番号"
    wstr1 = wstr1 + ",G.銀行名"
    For l = 481 To 560
        wstr1 = wstr1 + ",KU.内入" & CStr(l) & "回目年月日"
    Next l
    wstr1 = wstr1 + " From (DBDA010_借入金 As K"
    wstr1 = wstr1 + " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr1 = wstr1 + " ON K.銀行番号 = G.銀行番号)"
    wstr1 = wstr1 + " Left Join DBDA010_借入金内入7 As KU"
    wstr1 = wstr1 + " ON K.借入番号 = KU.借入番号"
    wstr1 = wstr1 + " Where (format(KU.内入481回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    For l = 482 To 560
        wstr1 = wstr1 + " OR (format(KU.内入" & CStr(l) & "回目年月日,'yyyymmdd') = '" & Format(wd01, "yyyymmdd") + "')"
    Next l
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        Do Until wRs1.eof
        
            For l = 481 To 560
                If P8.FCIsDate(wRs1("内入" & CStr(l) & "回目年月日")) Then
                    If CDate(wRs1("内入" & CStr(l) & "回目年月日")) = wd01 Then
                        pcnt = pcnt + 1
                        ReDim Preserve GCal(pcnt)
                        GCal(pcnt).借入番号 = wRs1("借入番号")
                        GCal(pcnt).銀行番号 = wRs1("銀行番号")
                        GCal(pcnt).銀行名 = wRs1("銀行名")
                        GCal(pcnt).確認年月日 = "内入" & CStr(l) & "回目年月日:" & wRs1("内入" & CStr(l) & "回目年月日")
                
                        営業年月日CHECK = True
                    End If
                End If
            Next l
            
            wRs1.MoveNext
        Loop
    wRs1.Close
    Set wRs1 = Nothing
'
End Function

'------------------------------------------------
' 出力_Click
'------------------------------------------------
Private Sub 出力_Click()

    If Not P8.FIsInt(取込年) Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(取込年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
    
    If P8.FCDbl(取込年) <= 0 Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(取込年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
    
    If Len(P8.FCStr(取込年)) < 4 Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(取込年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
'
    Call MX040_祝日マスタ(GKeyName & "_" & 取込年 & "祝日マスタ.csv", 取込年)
'
End Sub

'------------------------------------------------
' 印刷_Click
'------------------------------------------------
Private Sub 印刷_Click()
'
    Dim rpt As Object
'
    If Not P8.FIsInt(印刷年) Then
        検索年 = ""
        Call CEkey.SetFs(印刷年, True)
        Exit Sub
    End If
    
    If P8.FCDbl(印刷年) <= 0 Then
        印刷年 = ""
        Call CEkey.SetFs(印刷年, True)
        Exit Sub
    End If
    
    If Len(P8.FCStr(印刷年)) < 4 Then
        印刷年 = ""
        Call CEkey.SetFs(印刷年, True)
        Exit Sub
    End If
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    検索.Enabled = False
    ログ出力.Enabled = False
    取込.Enabled = False
    登録.Enabled = False
    出力.Enabled = False
    印刷.Enabled = False
'
    ' =========================================
    '              レポート表示
    ' =========================================
    GRpt.帳票名 = "営業カレンダー"
    GRpt.テキスト_01 = P8.FCStr(印刷年)
    Set rpt = New RCA010_営業日カレンダー
'
    rpt.Show vbModal
'
    Set rpt = Nothing
'
    検索.Enabled = True
    ログ出力.Enabled = True
    取込.Enabled = True
    登録.Enabled = True
    出力.Enabled = True
    印刷.Enabled = True
'
End Sub

'------------------------------------------------
' ログ出力_Click
'------------------------------------------------
Private Sub ログ出力_Click()
'
    Dim rpt As Object
    
    Dim wNen As Integer
    Dim wRet As Boolean
'
    If Not P8.FIsInt(検索年) Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(検索年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
    
    If P8.FCDbl(検索年) <= 0 Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(検索年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
    
    If Len(P8.FCStr(検索年)) < 4 Then
        MsgBox "西暦 年を入力してください", vbExclamation
        Call CEkey.SetFs(検索年, True)
        Call CEkey.AllSelect
        
        Exit Sub
    End If
'
    ' =========================================
    '   実行日、初回返済年月日、最終返済年月日、内入年月日 CHECk
    ' =========================================
    pcnt = 0
    ReDim GCal(pcnt)
    
    wNen = P8.FCDbl(検索年)
    
    ' =========================================
    '            ユーザーマスタ セット
    ' =========================================
    wRet = False
    GRet = False
    
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "* "
    wstr = wstr & "FROM DACA010_祝日マスタ "
    wstr = wstr & " WHERE Format(年月日,'yyyymmdd') >= '" & P8.FCStr(wNen) & "0101" & "'"
    wstr = wstr & " and Format(年月日,'yyyymmdd') <= '" & P8.FCStr(wNen) & "1231" & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            
            wRet = 営業年月日CHECK(wRs("年月日"))
            If wRet = True Then
                GRet = wRet
            End If
            
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing

    If GRet = True Then
        ' =========================================
        '           　 ボタン制御
        ' =========================================
        検索.Enabled = False
        ログ出力.Enabled = False
        取込.Enabled = False
        登録.Enabled = False
        出力.Enabled = False
        印刷.Enabled = False
'
        ' =========================================
        '              レポート表示
        ' =========================================
        GRpt.帳票名 = "祝日マスタログ一覧"
        Set rpt = New RCA020_祝日マスタログ一覧
'
        rpt.Show vbModal
'
        Set rpt = Nothing
'
        検索.Enabled = True
        ログ出力.Enabled = True
        取込.Enabled = True
        登録.Enabled = True
        出力.Enabled = True
        印刷.Enabled = True
'
    Else
        GRet = MsgBox("出力するデータがありません。", vbOKOnly)
    End If
'
End Sub

'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    Unload Me
    
End Sub
