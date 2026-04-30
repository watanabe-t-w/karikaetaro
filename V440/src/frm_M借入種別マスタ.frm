VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_M借入種別マスタ 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入種別マスタ"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "frm_M借入種別マスタ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox 削除データを表示 
      Caption         =   "削除データを表示"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   6840
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
      Left            =   4200
      TabIndex        =   5
      Top             =   6840
      Width           =   1815
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1815
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
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   7815
      Begin VB.CheckBox 利子補給金フラグ 
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox 社債フラグ 
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox 削除 
         Caption         =   "削除"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox 借入種別名 
         Height          =   330
         IMEMode         =   4  '全角ひらがな
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox 借入種別コード 
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   240
         TabIndex        =   13
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
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "利子補給金フラグ"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "社債フラグ"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "借入種別名"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "借入種別コード"
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4948
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "借入種別マスタ"
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
Attribute VB_Name = "frm_M借入種別マスタ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "借入種別マスタ"

Dim wRs As ADODB.Recordset
Dim wstr As String

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
    wstr = wstr & " IIF(取消フラグ=0,'','*') AS Grd削除,"
    wstr = wstr & " 借入金種別区分 AS Grd借入金種別区分,"
    wstr = wstr & " 借入金種別名 AS Grd借入金種別名,"
    wstr = wstr & " IIF(社債フラグ=0,'×','○') AS Grd社債フラグ"
    
    '16/03/26 利子補給に伴う変更
    'wstr = wstr & " IIF(利子補給金フラグ=0,'×','○') AS Grd利子補給フラグ"
    
    wstr = wstr & " FROM DAAA116_借入金種別"
    wstr = wstr & " WHERE (0=0)"
    If Me.削除データを表示.Value = 0 Then
        wstr = wstr & " AND 取消フラグ = 0"
    End If
    wstr = wstr & " ORDER BY 借入金種別区分"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        If Me.削除データを表示.Value = 1 Then
            Call XZMA010_DataGrid_Set("削除", "削", 300, "C")
        End If
        Call XZMA010_DataGrid_Set("借入金種別区分", "コード", 1400, "L")
        Call XZMA010_DataGrid_Set("借入金種別名", "借入種別名", 4000, "L")
        Call XZMA010_DataGrid_Set("社債フラグ", "社債", 600, "C")
        '16/03/26 利子補給に伴う変更
        'Call XZMA010_DataGrid_Set("利子補給フラグ", "利子補給", 1200, "C")
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
    Call CEkey.SetFs(借入種別コード, True)
    
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd借入金種別区分")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        借入種別コード = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd借入金種別区分"))
    On Error GoTo 0
    
    Call 画面セット(True)
   
'    If DataGrid1.Splits.Count <> 1 Then
'        DataGrid1.Splits.Remove 1
'    End If

    Call CEkey.SetFs(借入種別名, True)

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
    Dim p借入種別コード As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    Me.借入種別名.Text = ""
    社債フラグ.Value = 0
    削除.Value = 0
    
    ' =========================================
    '                パラメータ
    ' =========================================
    p借入種別コード = P8.FCStr(借入種別コード.Text)
    
    ' =========================================
    '            ユーザーマスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "* "
    wstr = wstr & "FROM DAAA116_借入金種別 "
    wstr = wstr & "WHERE 借入金種別区分 = '" & p借入種別コード & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
    '新規登録
        削除.Enabled = False
        新規変更.Caption = "新規"
        Me.借入種別名.Text = ""
    Else
    '変更登録
        削除.Enabled = True
        新規変更.Caption = "変更"
        Me.借入種別名.Text = P8.FCStr(wRs("借入金種別名"))
        Me.社債フラグ.Value = P8.FCDbl(wRs("社債フラグ"))
        Me.利子補給金フラグ.Value = 0
        
        '16/03/26 利子補給に伴う変更
        'Me.利子補給金フラグ.Value = P8.FCDbl(wRs("利子補給金フラグ"))
        
        Me.削除.Value = P8.FCDbl(wRs("取消フラグ"))
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
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd借入金種別区分 = '" + 借入種別コード + "'")
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
' 借入種別コード_GotFocus
'------------------------------------------------
Private Sub 借入種別コード_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w借入種別コード As String
'
    w借入種別コード = 借入種別コード
    
    借入種別コード = ""
    Call 画面セット(False)
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd借入金種別区分 = '" + w借入種別コード + "'")
    Call CEkey.SetFs(Me.借入種別コード, True)
'
End Sub

Private Sub 借入種別名_LostFocus()
    Call P8.FCControlLeft(借入種別名, 30)
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim ws01 As String
    
    Dim p借入種別コード As String
    Dim p借入種別名 As String
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

    ' =========================================
    '               パラメータセット
    ' =========================================
    p借入種別コード = P8.FCStr(Me.借入種別コード.Text)
    p借入種別名 = P8.FCStr(Me.借入種別名.Text)
'
    ' =========================================
    '               入力チェック
    ' =========================================
    If p借入種別コード = "" Then
        MsgBox "借入種別コードが未入力です。", vbExclamation
        Call CEkey.SetFs(借入種別コード, True)
        Exit Sub
    End If
    
    If Len(p借入種別コード) <> 2 Then
        MsgBox "借入種別コードは2桁で入力してください。", vbExclamation
        Call CEkey.SetFs(借入種別コード, True)
        Exit Sub
    End If
'
    ' =========================================
    '            借入種別マスタ 更新処理
    ' =========================================
    If Me.新規変更.Caption = "新規" Then
    '新規登録
        
        p借入種別コード = LTrim(p借入種別コード)
        p借入種別名 = LTrim(p借入種別名)
        借入種別コード.Text = p借入種別コード
        借入種別名.Text = p借入種別名
            
        wstr = ""
        wstr = wstr & "Select *"
        wstr = wstr & " From DAAA116_借入金種別"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            wRs.AddNew
                        
            wRs("借入金種別区分") = p借入種別コード
            wRs("借入金種別名") = p借入種別名
            
            If 社債フラグ.Value = 1 Then
                 wRs("社債フラグ") = 1
            Else
                 wRs("社債フラグ") = 0
            End If
            
            '16/03/26 利子補給に伴う変更
            If 利子補給金フラグ.Value = 1 Then
                 wRs("利子補給金フラグ") = 1
            Else
                 wRs("利子補給金フラグ") = 0
            End If
            
            wRs("取消フラグ") = 0
    
            wRs.Update
    
        wRs.Close
        Set wRs = Nothing
    Else
    '更新
        wstr = ""
        wstr = wstr & "Select *"
        wstr = wstr & " From DAAA116_借入金種別"
        wstr = wstr & " WHERE 借入金種別区分 = '" & p借入種別コード & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        
            wRs("借入金種別区分") = p借入種別コード
            wRs("借入金種別名") = p借入種別名
           
            If 社債フラグ.Value = 1 Then
                 wRs("社債フラグ") = 1
            Else
                 wRs("社債フラグ") = 0
            End If
            
            '16/03/26 利子補給に伴う変更
            If 利子補給金フラグ.Value = 1 Then
                 wRs("利子補給金フラグ") = 1
            Else
                 wRs("利子補給金フラグ") = 0
            End If
            
            If 削除.Value = 1 Then
                 wRs("取消フラグ") = 1
            Else
                 wRs("取消フラグ") = 0
            End If
            
            wRs.Update
    
        wRs.Close
        Set wRs = Nothing
    End If
'
    Call UNLOAD_借入金FRM
'
    ' =========================================
    '                テーブル変更
    ' =========================================
    Call MAA070_借入金種別設定
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
    GLogStr = "借入種別コード=" & p借入種別コード & ","
    GLogStr = GLogStr & "借入種別名=" & p借入種別名
    GLogStr = GLogStr & "社債フラグ=" & P8.FCDbl(社債フラグ)
    
    '16/03/26 利子補給に伴う変更
    'GLogStr = GLogStr & "利子補給金フラグ=" & P8.FCDbl(利子補給金フラグ)
    
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(借入種別コード, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
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

Private Sub 借入種別コード_LostFocus()

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

'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    Unload Me
    
End Sub






