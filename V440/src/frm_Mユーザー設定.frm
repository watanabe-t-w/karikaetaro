VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Mユーザー設定 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ユーザー設定"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Mユーザー設定.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox 削除データを表示 
      Caption         =   "削除データを表示"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   120
      Top             =   7560
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
   Begin VB.CommandButton 登録 
      Caption         =   "登録（F11)"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton 閉じる 
      Caption         =   "閉じる(F12)"
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7320
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
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   7695
      Begin 借換たろう.ZU020_ComboBox ユーザー権限 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1920
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
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
      End
      Begin VB.TextBox パスワード 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   2040
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox 所属 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   4  '全角ひらがな
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CheckBox 削除 
         Caption         =   "削除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox ユーザー名 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   4  '全角ひらがな
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox ユーザーID 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   0
         Top             =   720
         Width           =   1455
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   240
         TabIndex        =   16
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
      Begin VB.Label Label5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "パスワード"
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
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "ユーザー権限"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "所属"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblユーザー名 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "ユーザー名"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblユーザーID 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "ユーザーID"
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
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2805
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   7695
      _ExtentX        =   13573
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
      Caption         =   "ユーザー設定"
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
Attribute VB_Name = "frm_Mユーザー設定"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "ユーザー設定"

Dim wDb As New ADODB.Connection
Dim wRs3 As ADODB.Recordset
Dim wstr As String

Dim wslog As String

Dim FLG_New As Boolean

'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    ' =========================================
    '                 初期設定
    ' =========================================
    '----------< GCurDir GTemp Open >-----------------------------------------------
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)
    
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
    With Me.ユーザー権限
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 2
        
        Call .AddItem(0, "入力")
        Call .AddItem(1, "照会")
        Call .AddItem(5, "管理者")
'        Call .AddItem(9, "導入作成者")
    End With
    ユーザー権限.CreateCombo
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
    Call AdodcSet(Adodc1, wDb)
  
    ' =========================================
    '              メインクエリ
    ' =========================================
    GWhere = ""
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " IIF(停止日 IS NOT NULL,'*','') As Grd停止日,"
    wstr = wstr + " USERID As GrdUSERID,"
    wstr = wstr + " 社員名 As Grd社員名,"
    wstr = wstr + " 所属 As Grd所属,"
    wstr = wstr + " IIF(権限=0,'入力',IIF(権限=1,'照会',IIF(権限=5,'管理者','導入作成者'))) As Grd権限,"
    wstr = wstr + " Format(PW変更日,'yyyy/mm/dd') As GrdPW変更日"
    wstr = wstr + " From DBUA001_ユーザーマスタ"
    wstr = wstr + " WHERE (0=0)"
    If Me.削除データを表示.Value = 0 Then
        wstr = wstr + " AND 停止日 IS NULL"
    End If
    wstr = wstr + " AND USERID <> 'INST' "
    wstr = wstr + " Order By USERID"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        If Me.削除データを表示.Value = 1 Then
            Call XZMA010_DataGrid_Set("停止日", "削", 300, "C")
        End If
        Call XZMA010_DataGrid_Set("USERID", "ユーザーID", 1400, "L")
        Call XZMA010_DataGrid_Set("社員名", "ユーザー名", 2000, "L")
        Call XZMA010_DataGrid_Set("所属", "", 2000, "L")
        Call XZMA010_DataGrid_Set("権限", "", 1500, "L")
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
    Call CEkey.SetFs(ユーザーID, True)
    
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("GrdUSERID")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        ユーザーID = P8.FCStr(Adodc1.Recordset.Fields.Item("GrdUSERID"))
    On Error GoTo 0
    
    Call 画面セット(True)
   
'    If DataGrid1.Splits.Count <> 1 Then
'        DataGrid1.Splits.Remove 1
'    End If

    Call CEkey.SetFs(ユーザー名, True)

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
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    FLG_New = True
    ユーザー名 = ""
    所属 = ""
    ユーザー権限.Text = ""
    削除.Value = 0
    
    ' =========================================
    '                パラメータ
    ' =========================================
    ws01 = P8.FCStr(ユーザーID)
    
    ' =========================================
    '            ユーザーマスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "IIF(停止日 IS NOT NULL,'1','0') AS 削除,"
    wstr = wstr & "* "
    wstr = wstr & "FROM DBUA001_ユーザーマスタ "
    wstr = wstr & "WHERE USERID = '" & ws01 & "'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If wRs3.EOF Then
    '新規登録
        削除.Enabled = False
        新規変更.Caption = "新規"
        Me.ユーザー名.Text = ""
        Me.所属.Text = ""
        Me.パスワード.Text = ""
        Me.削除.Value = 0
    Else
    '変更登録
        削除.Enabled = True
        新規変更.Caption = "変更"
        Me.ユーザー名.Text = P8.FCStr(wRs3("社員名"))
        Me.所属.Text = P8.FCStr(wRs3("所属"))
        Me.ユーザー権限.Text = P8.FCDbl(wRs3("権限"))
        Me.パスワード.Text = P8.FCStr(wRs3("PASSWORD"))
        Me.削除.Value = P8.FCDbl(wRs3("削除"))
    End If
    
    wRs3.Close
    Set wRs3 = Nothing
    
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
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "GrdUSERID = '" + ユーザーID + "'")
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
    Call P8.FCControlLeft(ユーザーID, 10)
    
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1", "USERID", "入力クリア"
            Exit Sub
    End Select
   
    If ユーザーID = "" Then
        MsgBox "コードを入力してください"
        Call CEkey.SetFs(ユーザーID, True)
        Exit Sub
    End If
'
    Select Case Screen.ActiveControl.Name
        Case "保存"
            Call CEkey.SetFs(ユーザー名, True)
            MsgBox "該当データをセットします。保存処理は行いません。"
            Exit Sub
    End Select
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
    Dim wUSERID As String
'
    wUSERID = ユーザーID
    
    ユーザーID = ""
    Call 画面セット(False)
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "GrdUSERID = '" + wUSERID + "'")
    Call CEkey.SetFs(Me.ユーザーID, True)
'
End Sub

'------------------------------------------------
' Ch_PASSWORD_Click
'------------------------------------------------
Private Sub Ch_PASSWORD_Click()
'
'
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim ws01 As String
    
    Dim pユーザーID As String
    Dim pユーザー名 As String
    Dim p所属 As String
    Dim pユーザー権限 As String
    Dim pパスワード As String
'
    On Error GoTo 保存_Click_ERR
'
    ' =========================================
    '               パラメータセット
    ' =========================================
    pユーザーID = P8.FCStr(Me.ユーザーID.Text)
    pユーザー名 = P8.FCStr(Me.ユーザー名.Text)
    p所属 = P8.FCStr(Me.所属.Text)
    pユーザー権限 = P8.FCDbl(ユーザー権限.Text)
    pパスワード = P8.FCStr(Me.パスワード.Text)
'
    ' =========================================
    '               入力チェック
    ' =========================================
    If pユーザーID = "" Then
        MsgBox "ユーザーIDが未入力です", vbExclamation
        Call CEkey.SetFs(ユーザーID, True)
        Exit Sub
    End If

    If pユーザー名 = "" Then
        MsgBox "ユーザー名が未入力です", vbExclamation
        Call CEkey.SetFs(ユーザー名, True)
        Exit Sub
    End If
'
    If pユーザー権限 = "" Or (P8.FCDbl(ユーザー権限.Text) <> 0 And P8.FCDbl(ユーザー権限.Text) <> 1 And P8.FCDbl(ユーザー権限.Text) <> 5) Then
        MsgBox "ユーザー権限を選択してください", vbExclamation
        Call CEkey.SetFs(ユーザー権限, True)
        Exit Sub
    End If
    
    If pパスワード = "" Then
        MsgBox "パスワードが未入力です", vbExclamation
        Call CEkey.SetFs(パスワード, True)
        Exit Sub
    End If
'

'
    ' =========================================
    '            ユーザーマスタ 更新処理
    ' =========================================
    ws01 = P8.FCStr(ユーザーID.Text)
    If Me.新規変更.Caption = "新規" Then
    '新規登録
        wstr = ""
        wstr = wstr & "Select *"
        wstr = wstr & " From DBUA001_ユーザーマスタ"
        Call AdoRecordsetOpen(wDb, wRs3, wstr)
            wRs3.AddNew
            
            wRs3("USERID") = pユーザーID
            
            wRs3("登録USERID") = GUserID
            wRs3("登録日時") = Format(Now, "yyyy/mm/dd hh:mm:ss")
     
            wRs3("社員名") = pユーザー名
            wRs3("所属") = p所属
            wRs3("権限") = pユーザー権限
            wRs3("PASSWORD") = pパスワード
    
            wRs3.Update
    
        wRs3.Close
        Set wRs3 = Nothing
    Else
    '更新
        wstr = ""
        wstr = wstr & "Select *"
        wstr = wstr & " From DBUA001_ユーザーマスタ"
        wstr = wstr & " WHERE USERID = '" & pユーザーID & "'"
        Call AdoRecordsetOpen(wDb, wRs3, wstr)
        
            wRs3("社員名") = pユーザー名
            wRs3("所属") = p所属
            wRs3("権限") = pユーザー権限
            wRs3("PASSWORD") = pパスワード
    
            wRs3("更新USERID") = GUserID
            wRs3("更新日時") = Format(Now, "yyyy/mm/dd hh:mm:ss")
            If 削除.Value = 1 Then
                wRs3("停止日") = Format(Now, "yyyy/mm/dd hh:mm:ss")
            Else
                wRs3("停止日") = Null
            End If
            
            wRs3.Update
    
        wRs3.Close
        Set wRs3 = Nothing
    End If
'
    Adodc1.Refresh
    
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
    GLogStr = "ユーザーID=" & pユーザーID & ","
    GLogStr = GLogStr & "ユーザー名=" & pユーザー名 & ","
    GLogStr = GLogStr & "所属=" & p所属 & ","
    GLogStr = GLogStr & "ユーザー権限=" & pユーザー権限
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
    
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(ユーザーID, True)
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
    wDb.Close
    Set wDb = Nothing
'
    Unload Me
End Sub




