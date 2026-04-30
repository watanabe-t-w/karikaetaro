VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_M長期プライムレート 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "基準金利レート"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   Icon            =   "frm_M長期プライムレート.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CSV取込 
      Caption         =   "CSV取込"
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
      Left            =   2520
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton CSV出力 
      Caption         =   "CSV出力"
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
      Left            =   4320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   7680
      Width           =   7575
      Begin VB.TextBox 長期プライムレート 
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
         Left            =   2400
         MaxLength       =   7
         TabIndex        =   10
         Top             =   1200
         Width           =   2535
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
         IMEMode         =   2  'ｵﾌ
         Left            =   2400
         TabIndex        =   9
         Top             =   720
         Width           =   2535
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
         Left            =   5640
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
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
         Left            =   5640
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   120
         TabIndex        =   11
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
      Begin VB.Label Label57 
         Alignment       =   2  '中央揃え
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         TabIndex        =   14
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 基準金利レート"
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
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 年月日"
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
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.CommandButton 表示 
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
      Left            =   6120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
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
      Left            =   6120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   6360
      Top             =   120
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5565
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9816
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "基準金利レート"
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
   Begin 借換たろう.ZU020_ComboBox 基準金利 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      ForeColor       =   -2147483640
      ForeColor       =   -2147483640
      IMEMode         =   3
      TextWidth       =   615
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.Label Label27 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '実線
      Caption         =   " 基準金利名"
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
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frm_M長期プライムレート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "基準金利レート"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim ws基準金利 As String

Dim FLG_DELL As Boolean
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    ' =========================================
    '                 初期設定
    ' =========================================
'    Me.Caption = GFcap
    
    Me.Left = G_LEFT
    Me.Top = G_TOP

    With 基準金利
        .P8_Db = GDb
        
        wstr = "Select * From DAAA116_基準金利"
        wstr = wstr + " Where 取消フラグ = 0"
        wstr = wstr + " Order By 基準金利区分"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 5
        .P8_ListBoxMax = 500
        .P8_KeyName = "基準金利区分"
        .P8_ItemName = "基準金利名"
    End With
    基準金利.CreateCombo

    基準金利.Text = ""
'
    ' =========================================
    '                 初期設定
    ' =========================================
    'Call 登録後初期セット
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
'    Call MXA030_DataGridInit(DataGrid1)
    DataGrid1.AllowRowSizing = False
    DataGrid1.HeadFont.Size = 9
    DataGrid1.HeadFont.Bold = True
    DataGrid1.Font.Size = 9
    DataGrid1.BackColor = C_Yellow
    DataGrid1.ForeColor = RGB(0, 0, 160)
    
    Set DataGrid1.DataSource = Adodc1
  
    ' =========================================
    '              ConnectionString
    ' =========================================
    Call AdodcSet(Adodc1, GDb)
  
    ' =========================================
    '              メインクエリ
    ' =========================================
    GWhere = ""
    GWhere = " Where (1=1) "
    GWhere = GWhere + "  And 基準金利区分='" & P8.FCStr(基準金利.Text) & "'"
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " Format(年月日,'" & Gfmt年月日 & "') As Grd年月日,"
    wstr = wstr + " Format(長期プライムレート,'#,##0.00000') As Grd基準金利レート"
    wstr = wstr + " From  DBDA010_借入金長期プライムレート"
    wstr = wstr + GWhere
    wstr = wstr + " Order By 年月日 Desc"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("年月日", "", 2000, "L")
        Call XZMA010_DataGrid_Set("基準金利レート", "", 2000, "R")
    Call XZMA010_DataGrid_Action(DataGrid1)
  
'    メッセージ = ""
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
    長期プライムレート.Text = ""
    
    ' =========================================
    '            マスタ セット
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    wstr = ""
    wstr = wstr & "Select *"
    wstr = wstr & " From DBDA010_借入金長期プライムレート "
    wstr = wstr & " Where 基準金利区分='" & P8.FCStr(基準金利.Text) & "'"
    wstr = wstr & " And Format(年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            If 年月日 <> "" Then
                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
                If GRet = vbNo Then
                    新規変更.Caption = ""
                    wRs.Close
                    Set wRs = Nothing

                    Exit Function
                End If
                
                新規変更.Caption = "新規登録"
    
            End If
        Else
            画面セット = True
            新規変更.Caption = "変更"
            
            Call CEkey.SetFs(長期プライムレート, True)
            
            年月日 = Format(P8.FCStr(wRs("年月日")), Gfmt年月日)
            長期プライムレート = P8.FFormat(wRs("長期プライムレート"), "#,##0.00000")
            
        End If
    
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
    ws基準金利 = P8.FCStr(基準金利.Text)
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
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w年月日 As String
'
    w年月日 = ""
    
    Call 画面セット(False)
    新規変更.Caption = ""
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
    Call CEkey.SetFs(年月日, True)
'
End Sub

'------------------------------------------------
' B_SET_Click
'------------------------------------------------
Private Sub B_SET_Click()
'
    Call 画面セット(False)
'
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' 年月日_LostFocus
'------------------------------------------------
Private Sub 年月日_LostFocus()
'
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1", "年月日"
            Exit Sub
    End Select
'
    年月日 = C年月日.FormatDate("年月日", 年月日)
'
    Select Case Screen.ActiveControl.Name
        Case "登録", "削除", "CSV出力"
            Exit Sub
    End Select
'
    Call B_SET_Click
'
End Sub

'------------------------------------------------
' 年月日_GotFocus
'------------------------------------------------
Private Sub 年月日_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 長期プライムレート_LostFocus
'------------------------------------------------
Private Sub 長期プライムレート_LostFocus()
    長期プライムレート = P8.FFormat(長期プライムレート, "#,##0.00000")
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim wslog As String
'
    On Error GoTo 保存_Click_ERR
'
    ' =========================================
    '           入力チェック
    ' =========================================
    If P8.FCStr(基準金利.Text) = "" Then
        GRet = MsgBox("基準金利区分を選択してください。", vbOKOnly)
        Exit Sub
    End If

    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "年月日が違います", vbExclamation
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
    
    If 長期プライムレート = "" Then
        MsgBox "基準金利レートが未入力です。", vbExclamation
        Call CEkey.SetFs(長期プライムレート, True)
        Exit Sub
    End If
    If Not IsNumeric(長期プライムレート) And 長期プライムレート <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(長期プライムレート, True)
        Exit Sub
    End If
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
    '            更新処理
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DBDA010_借入金長期プライムレート"
    wstr = wstr + " Where 基準金利区分='" & P8.FCStr(基準金利.Text) & "'"
    wstr = wstr + " and Format(年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            wRs.AddNew
            
            年月日 = GVar1
            長期プライムレート = LTrim(P8.FCStr(長期プライムレート))
            
            wRs("基準金利区分") = P8.FCStr(基準金利.Text)
            wRs("年月日") = 年月日
            
            年月日 = Format(P8.FCStr(wRs("年月日")), Gfmt年月日)

            wslog = "追加"
        End If

        wRs("長期プライムレート") = LTrim(P8.FCStr(長期プライムレート))
        
        wRs.Update
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    If 新規変更.Caption = "新規" Then
        wslog = "追加"
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 0 Then
        wslog = "更新"
    End If
    GLogStr = "基準金利区分=" & P8.FCStr(基準金利.Text) & ","
    GLogStr = GLogStr & "年月日=" & Format(GVar1, "yyyy/mm/dd") & ","
    GLogStr = GLogStr & "基準金利レート=" & P8.FCStr(長期プライムレート.Text)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)

'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(年月日, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました", vbInformation
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

'------------------------------------------------
' 表示_Click
'------------------------------------------------
Private Sub 表示_Click()
    
    If P8.FCStr(基準金利.Text) = "" Then
        GRet = MsgBox("基準金利区分を選択してください。", vbOKOnly)
        Exit Sub
    End If
    
    ' =========================================
    '                 初期設定
    ' =========================================
    Call 登録後初期セット

End Sub

Private Sub 基準金利_Change()
    If P8.FCStr(基準金利.Text) <> "" And P8.FCStr(基準金利.Text) <> ws基準金利 Then
        年月日.Text = ""
        長期プライムレート.Text = ""
        
        If Not DataGrid1.DataSource Is Nothing Then
            Set DataGrid1.DataSource = Nothing
        End If
    End If
End Sub

'------------------------------------------------
' CSV取込_Click
'------------------------------------------------
Private Sub CSV取込_Click()
'
    Dim ws01 As String, wsRet As String
'
    If P8.FCStr(基準金利.Text) = "" Then
        GRet = MsgBox("基準金利区分を選択してください。", vbOKOnly)
        Exit Sub
    End If
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

    GRet = MsgBox("CSVファイルをインポートします。" _
                    & vbCrLf & vbCrLf & "既存のデータは上書きされます。" & vbCrLf & _
                    "よろしいですか？", vbExclamation + vbYesNo, "取込")
    If GRet = vbNo Then
        Exit Sub
    End If
'
    ws01 = "基準金利レート.csv"
    wsRet = MXA040_COMDLG(CommonDialog1, "CSVファイル選択", "", _
                            "テキストファイル(*.csv)|*.csv", ws01)
    If wsRet = "" Then
        Exit Sub
    ElseIf wsRet = "キャンセル" Then
        Exit Sub
    End If

    GRet = MXA040_基準金利レート取込(wsRet, P8.FCStr(基準金利.Text))
    If GRet <> True Then
        MsgBox "CSVファイルをインポートできませんでした", vbInformation
        
        Exit Sub
    End If
    
    MsgBox "CSVファイルをインポートしました", vbOKOnly
'
    ' =========================================
    '                 初期設定
    ' =========================================
    Call 登録後初期セット
'
End Sub

'------------------------------------------------
' CSV出力_Click
'------------------------------------------------
Private Sub CSV出力_Click()
'
    Dim wsRet As String, wsFileName As String
'
    If P8.FCStr(基準金利.Text) = "" Then
        GRet = MsgBox("基準金利区分を選択してください。", vbOKOnly)
        Exit Sub
    End If
'
    GRpt.帳票名 = "基準金利レート"
    GRpt.テキスト_01 = P8.FCStr(基準金利.Text)
    GRpt.CSV = 1
'
    ' =========================================
    '           　 CsvFile 作成
    ' =========================================
    Call MX040_CsvOut_KARI
'
    ' =========================================
    '           Csv File Drive
    ' =========================================
    Call MX040_CsvPath
'
End Sub

'------------------------------------------------
' 削除_Click
'------------------------------------------------
Private Sub 削除_Click()
'
    Dim wi01 As Integer
    Dim wd01 As Double, wd02 As Double
    Dim wDate1 As Date, wDate2 As Date
'
    On Error GoTo 削除_Click_ERR
'
    ' =========================================
    '           入力チェック
    ' =========================================
    If P8.FCStr(基準金利.Text) = "" Then
        GRet = MsgBox("基準金利区分を選択してください。", vbOKOnly)
        Exit Sub
    End If

    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "年月日が違います", vbExclamation
        Exit Sub
    End If
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

    GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
    If GRet = vbNo Then
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    ' =========================================
    '            明細TR
    ' =========================================
    wstr = ""
    wstr = wstr + "Delete * From DBDA010_借入金長期プライムレート"
    wstr = wstr + " Where 基準金利区分='" & P8.FCStr(基準金利.Text) & "'"
    wstr = wstr + " And Format(年月日,'yyyy/mm/dd') = '" & Format(GVar1, "yyyy/mm/dd") & "'"
    GDb.Execute wstr
'
    '----------< DataGrid Close >----------------------------------------------
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "年月日=" & Format(GVar1, "yyyy/mm/dd") & ","
    GLogStr = GLogStr & "長期プライムレート=" & P8.FCStr(長期プライムレート.Text)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, "削除", GLogStr)
'
    ' =========================================
    '               画面セット
    ' =========================================
    年月日.Text = ""
    
    Call 画面セット(False)
'    Call 登録後初期セット
    Call CEkey.SetFs(年月日, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "削除しました。", vbInformation
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
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    '----------< DataGrid Close >----------------------------------------------
    If Not DataGrid1.DataSource Is Nothing Then
        Set DataGrid1.DataSource = Nothing
    End If

    If Not Adodc1.Recordset Is Nothing Then
        Adodc1.Recordset.Close
    End If
'
    Unload Me
End Sub
