VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_M基準金利マスタ 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "基準金利マスタ"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "frm_M基準金利マスタ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox 削除データを表示 
      Caption         =   "削除データを表示"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Left            =   3000
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton 閉じる 
      Caption         =   "閉じる（F12)"
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
      Left            =   4920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5640
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
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   6615
      Begin VB.TextBox 基準金利コード 
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox 基準金利名 
         Height          =   330
         IMEMode         =   4  '全角ひらがな
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "あああああいいいいいうううううえええええ"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.CheckBox 削除 
         Caption         =   "削除"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   240
         TabIndex        =   10
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
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "基準金利コード"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "基準金利名"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4524
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "基準金利マスタ"
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
Attribute VB_Name = "frm_M基準金利マスタ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "基準金利マスタ"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim wslog As String

Dim FLG_New As Boolean
'------------------------------------------------
' Form_Initialize
'------------------------------------------------
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
''
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
    wstr = wstr & " 基準金利区分 AS Grd基準金利区分,"
    wstr = wstr & " 基準金利名 AS Grd基準金利名"
    wstr = wstr & " FROM DAAA116_基準金利"
    wstr = wstr & " WHERE (0=0)"
    If Me.削除データを表示.Value = 0 Then
        wstr = wstr & " AND 取消フラグ = 0"
    End If
    wstr = wstr & " ORDER BY 基準金利区分"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        If Me.削除データを表示.Value = 1 Then
            Call XZMA010_DataGrid_Set("削除", "削", 300, "C")
        End If
        Call XZMA010_DataGrid_Set("基準金利区分", "コード", 1400, "L")
        Call XZMA010_DataGrid_Set("基準金利名", "基準金利名", 4000, "L")
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
    Call CEkey.SetFs(基準金利コード, True)
    
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd基準金利区分")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        基準金利コード = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd基準金利区分"))
    On Error GoTo 0
    
    Call 画面セット(True)
   
'    If DataGrid1.Splits.Count <> 1 Then
'        DataGrid1.Splits.Remove 1
'    End If

    Call CEkey.SetFs(基準金利名, True)

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
    Dim p基準金利コード As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    Me.基準金利名.Text = ""
    削除.Value = 0
    
    ' =========================================
    '                パラメータ
    ' =========================================
    p基準金利コード = P8.FCStr(基準金利コード.Text)
    
    ' =========================================
    '            ユーザーマスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "* "
    wstr = wstr & "FROM DAAA116_基準金利 "
    wstr = wstr & "WHERE 基準金利区分 = '" & p基準金利コード & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.EOF Then
    '新規登録
        削除.Enabled = False
        新規変更.Caption = "新規"
        Me.基準金利名.Text = ""
    Else
    '変更登録
        削除.Enabled = True
        新規変更.Caption = "変更"
        Me.基準金利名.Text = P8.FCStr(wRs("基準金利名"))
        Me.削除.Value = P8.FCDbl(wRs("取消フラグ"))
    End If
    
'
'
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd基準金利区分 = '" + 基準金利コード + "'")
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
' 基準金利コード_GotFocus
'------------------------------------------------
Private Sub 基準金利コード_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w基準金利コード As String
'
    w基準金利コード = 基準金利コード
    
    基準金利コード = ""
    Call 画面セット(False)
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd基準金利区分 = '" + w基準金利コード + "'")
    Call CEkey.SetFs(Me.基準金利コード, True)
'
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim ws01 As String
    
    Dim p基準金利コード As String
    Dim p基準金利名 As String
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
    p基準金利コード = P8.FCStr(Me.基準金利コード.Text)
    p基準金利名 = P8.FCStr(Me.基準金利名.Text)
'
    ' =========================================
    '               入力チェック
    ' =========================================
    If p基準金利コード = "" Then
        MsgBox "基準金利コードが未入力です。", vbExclamation
        Call CEkey.SetFs(基準金利コード, True)
        Exit Sub
    End If
    
    If Len(p基準金利コード) <> 2 Then
        MsgBox "基準金利コードは2桁で入力してください。", vbExclamation
        Call CEkey.SetFs(基準金利コード, True)
        Exit Sub
    End If
        
'
    ' =========================================
    '            借入種別マスタ 更新処理
    ' =========================================
    If Me.新規変更.Caption = "新規" Then
    '新規登録
        wstr = ""
        wstr = wstr & "Select *"
        wstr = wstr & " From DAAA116_基準金利"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            wRs.AddNew
            
            wRs("基準金利区分") = p基準金利コード
            wRs("基準金利名") = p基準金利名
            wRs("取消フラグ") = 0
    
            wRs.Update
    
        wRs.Close
        Set wRs = Nothing
    Else
    '更新
        wstr = ""
        wstr = wstr & "Select *"
        wstr = wstr & " From DAAA116_基準金利"
        wstr = wstr & " WHERE 基準金利区分 = '" & p基準金利コード & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        
            wRs("基準金利区分") = p基準金利コード
            wRs("基準金利名") = p基準金利名
           
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
    GLogStr = "基準金利コード=" & p基準金利コード & ","
    GLogStr = GLogStr & "基準金利名=" & p基準金利名
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)

    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(基準金利コード, True)
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

Private Sub 基準金利コード_LostFocus()

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








