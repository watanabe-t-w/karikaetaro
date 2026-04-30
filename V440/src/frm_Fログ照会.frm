VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Fログ照会 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ログ照会"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   Icon            =   "frm_Fログ照会.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ログDB 
      Caption         =   "過去ログ参照"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox 番号 
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   17
      Top             =   8880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox 内容 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      IMEMode         =   4  '全角ひらがな
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直
      TabIndex        =   16
      Top             =   7560
      Width           =   12375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   120
      Top             =   8880
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
      Left            =   10800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "検索"
      Height          =   1935
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   12375
      Begin VB.ComboBox Co_KUBUN 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1800
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox Co_PROGRAMID 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox Co_ID 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox 終了年月日 
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   5400
         MaxLength       =   30
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox 開始年月日 
         Height          =   330
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   1
         Top             =   360
         Width           =   1695
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
         Left            =   10440
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "操作区分"
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
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "フォーム名"
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
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "～"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "年月日To"
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
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "年月日From"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4365
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   7699
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
      Caption         =   "ログ照会"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_Fログ照会"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "ログ照会"

Dim wDb As New ADODB.Connection, wDb2 As New ADODB.Connection
Dim wRs3 As ADODB.Recordset
Dim wstr As String

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
    
    '----------< LOG.mdb Open >------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GCurDir + "\LOG.mdb", "", , GPwd)
'
    Call 検索項目_作成
    
    開始年月日 = Format(Format(Now, "yyyy/mm/01"), Gfmt年月日)
    終了年月日 = Format(DateAdd("d", -1, DateAdd("m", 1, DateValue(Format(Now, "yyyy/mm/01")))), Gfmt年月日)
    
    Call AdodcRefresh
'
End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
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
' 検索項目_作成
'------------------------------------------------
Private Sub 検索項目_作成()
'
    開始年月日 = ""
    終了年月日 = ""
'
    Co_ID.Clear
    wstr = ""
    wstr = wstr & "SELECT USERID"
    wstr = wstr & " From T_Log"
    wstr = wstr & " GROUP BY USERID"
    wstr = wstr & " ORDER BY USERID"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
        Do Until wRs3.EOF
            Co_ID.AddItem (P8.FCStr(wRs3("USERID")))

            wRs3.MoveNext
        Loop
    wRs3.Close
    Set wRs3 = Nothing
'
    Co_PROGRAMID.Clear
    wstr = ""
    wstr = wstr & "SELECT PROGRAMID"
    wstr = wstr & " From T_Log"
    wstr = wstr & " GROUP BY PROGRAMID"
    wstr = wstr & " ORDER BY PROGRAMID"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
        Do Until wRs3.EOF
            Co_PROGRAMID.AddItem (P8.FCStr(wRs3("PROGRAMID")))

            wRs3.MoveNext
        Loop
    wRs3.Close
    Set wRs3 = Nothing
'
    Co_KUBUN.Clear
    Co_KUBUN.AddItem ""
    Co_KUBUN.AddItem "ログイン"
    Co_KUBUN.AddItem "ログアウト"
    Co_KUBUN.AddItem "追加"
    Co_KUBUN.AddItem "更新"
    Co_KUBUN.AddItem "削除"
    Co_KUBUN.AddItem "帳票"
'
End Sub

'------------------------------------------------
' DataGrid1_Click
'------------------------------------------------
Private Sub DataGrid1_Click()
'
    Call CEkey.SetFs(内容, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("番号")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        番号 = P8.FCStr(Adodc1.Recordset.Fields.Item("番号"))
        内容 = P8.FCStr(Adodc1.Recordset.Fields.Item("内容"))
        内容 = 内容 & P8.FCStr(Adodc1.Recordset.Fields.Item("内容2"))
        内容 = 内容 & P8.FCStr(Adodc1.Recordset.Fields.Item("内容3"))
        
    On Error GoTo 0
    
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "番号 = '" + 番号 + "'")
    
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(内容, True)

Exit_Sub:
    Exit Sub
    '---------------------------------------------------
Err_Hundle:
    If Err.Number = 91 Then Resume Next
    If Err.Number = 94 Then Resume Next
    MsgBox CStr(Err.Number) + ":" + Err.Description
    Resume Exit_Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '----------< DataGrid Close >----------------------------------------------
    If Not DataGrid1.DataSource Is Nothing Then
        Set DataGrid1.DataSource = Nothing
    End If
    
    Adodc1.Recordset.Close
'
    wDb.Close
    Set wDb = Nothing
End Sub

'------------------------------------------------
' 検索_Click
'------------------------------------------------
Private Sub 検索_Click()
'
    Call AdodcRefresh
'
End Sub

'------------------------------------------------
' AdodcRefresh
'------------------------------------------------
Private Sub AdodcRefresh()
'
    Dim wv01 As Variant
'
    On Error GoTo AdodcRefresh_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
'    Call MXA030_DataGridInit(DataGrid1)
    DataGrid1.AllowRowSizing = False
    DataGrid1.HeadFont.Size = 10
    DataGrid1.HeadFont.Bold = True
    DataGrid1.Font.Size = 10
    DataGrid1.BackColor = C_Yellow
    DataGrid1.ForeColor = RGB(0, 0, 160)
    Set DataGrid1.DataSource = Adodc1
  
    ' =========================================
    '              ConnectionString
    ' =========================================
    Call AdodcSet(Adodc1, wDb)
  
    ' =========================================
    '              メインクエリ
    ' =========================================
    GWhere = ""
    
    GVar1 = C年月日.平成To西暦("年月日", 開始年月日.Text)
    GVar2 = C年月日.平成To西暦("年月日", 終了年月日.Text)
    If P8.FCStr(開始年月日.Text) <> "" And P8.FCStr(終了年月日.Text) = "" Then
        GWhere = GWhere & " And Format(ログ日付,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
    ElseIf P8.FCStr(開始年月日.Text) = "" And P8.FCStr(終了年月日.Text) <> "" Then
        GWhere = GWhere & " And Format(ログ日付,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    ElseIf P8.FCStr(開始年月日.Text) <> "" And P8.FCStr(終了年月日.Text) <> "" Then
        GWhere = GWhere & " And Format(ログ日付,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
        GWhere = GWhere & " And Format(ログ日付,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    End If
    
    If P8.FCStr(Co_ID.Text) <> "" Then
        GWhere = GWhere & " And USERID = '" & P8.FCStr(Co_ID.Text) & "'"
    End If
    If P8.FCStr(Co_PROGRAMID.Text) <> "" Then
        GWhere = GWhere & " And PROGRAMID = '" & P8.FCStr(Co_PROGRAMID.Text) & "'"
    End If
    If P8.FCStr(Co_KUBUN.Text) <> "" Then
        If P8.FCStr(Co_KUBUN.Text) = "ログイン" Then
            GWhere = GWhere & " And ログ区分=0"
        ElseIf P8.FCStr(Co_KUBUN.Text) = "追加" Then
            GWhere = GWhere & " And ログ区分=1"
        ElseIf P8.FCStr(Co_KUBUN.Text) = "更新" Then
            GWhere = GWhere & " And ログ区分=2"
        ElseIf P8.FCStr(Co_KUBUN.Text) = "削除" Then
            GWhere = GWhere & " And ログ区分=3"
        ElseIf P8.FCStr(Co_KUBUN.Text) = "帳票" Then
            GWhere = GWhere & " And ログ区分=4"
        ElseIf P8.FCStr(Co_KUBUN.Text) = "ログアウト" Then
            GWhere = GWhere & " And ログ区分=5"
        End If
    End If
    
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " Format(ログ日付,'eemmdd')&Format(ログ時刻,'hhmmss') As 番号,"
    wstr = wstr + " Format(ログ日付,'" & Gfmt年月日 & "') As Grd日付,"
    wstr = wstr + " Format(ログ時刻,'hh:mm:ss') As Grd時刻,"
    wstr = wstr + " USERID As GrdUSERID,"
    wstr = wstr + " PROGRAMID As GrdPROGRAMID,"
    wstr = wstr + " IIF(ログ区分=0,'ログイン',IIF(ログ区分=1,'追加',IIF(ログ区分=2,'更新',"
    wstr = wstr + " IIF(ログ区分=3,'削除',IIF(ログ区分=4,'帳票',IIF(ログ区分=5,'ログアウト','')))))) As Grd区分,"
    wstr = wstr + " 内容,"
    wstr = wstr + " 内容2,"
    wstr = wstr + " 内容3,"
    wstr = wstr + " 内容&内容2&内容3 As Grd内容"
    wstr = wstr + " From T_Log"
    wstr = wstr + GWhere
    wstr = wstr + " Order By Format(ログ日付,'eemmdd')&Format(ログ時刻,'hhmmss') DESC"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("日付", "", 1500, "L")
        Call XZMA010_DataGrid_Set("時刻", "", 1100, "L")
        Call XZMA010_DataGrid_Set("USERID", "ユーザーID", 1200, "L")
        Call XZMA010_DataGrid_Set("PROGRAMID", "フォーム名", 2300, "L")
        Call XZMA010_DataGrid_Set("区分", "", 1400, "L")
        Call XZMA010_DataGrid_Set("内容", "", 4300, "L")
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

Private Sub 開始年月日_LostFocus()
    開始年月日 = C年月日.FormatDate("年月日", 開始年月日)
End Sub

Private Sub 終了年月日_LostFocus()
    終了年月日 = C年月日.FormatDate("年月日", 終了年月日)
End Sub

'------------------------------------------------
' ログDBClick
'------------------------------------------------
Private Sub ログDB_Click()
'
    Dim wDb2 As New ADODB.Connection
    
    Dim wlCnt As Long
    Dim wretfn As String, wSDate As String
    Dim wsDB名 As String
    Dim ws01 As String
'
    On Error GoTo ログDB_Click_ERR
'
    番号 = ""
    内容 = ""
    
    'ボタン名判断 過去ログ or 現在ログ
     If ログDB.Caption = "過去ログ参照" Then
        GRet = MsgBox("過去ログDBを選択してください", vbOKCancel + vbExclamation)
        If GRet = vbCancel Then
            Exit Sub
        End If
        
        '----------< COMDLG >-----------------------------------------------------------
        wretfn = COMDLG("過去ログDBを選択してください", "", "AccessMdbファイル(*.mdb)|*.mdb", wsDB名)
        If wretfn = "" Then
            Exit Sub
        End If
        '
        '----------< AdoDbOpen_Check >----------------------------------------------
        GRet = ADODBOPEN_CHECK("Jet", wDb2, wretfn, "", , GPwd, "排他")
        If GRet <> True Then
            GRet = MsgBox("使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
            
            wDb2.Close
            Set wDb2 = Nothing
            
            Exit Sub
        End If
        
        'DB確認
        '----------< mdb Open >-----------------------------------------------------
        Call AdoDbOpen("Jet", wDb2, wretfn, "", , GPwd)
        Set wRs3 = New ADODB.Recordset
        
        wstr = ""
        wstr = wstr + "Select "
        wstr = wstr + " count(Name) As Cnt "
        wstr = wstr + "From "
        wstr = wstr + " MSysObjects "
        wstr = wstr + "Where "
        wstr = wstr + " name = 'DAAA000_バージョン'"
        wstr = wstr + " and type = 1"
        On Error Resume Next
            wRs3.CursorType = adOpenKeyset
            wRs3.LockType = adLockOptimistic
            wRs3.Open wstr, wDb2, , , adCmdText
        
            If Err.Number = "-2147217911" Then
                MsgBox "このMDBは選択できません", vbCritical
                wRs3.Close
                Set wRs3 = Nothing
                
                wDb2.Close
                Set wDb2 = Nothing
                Err.Clear
                
                Exit Sub
            End If
        
        If wRs3("Cnt") = 0 Then
            MsgBox "このMDBは選択できません", vbCritical
            
            wRs3.Close
            Set wRs3 = Nothing
            
            '----------< XXX.mdb Close >------------------------------------------------
            wDb2.Close
            Set wDb2 = Nothing
            
            Err.Clear
            
            Exit Sub
        End If
        wRs3.Close
        
        wDb2.Close
        Set wDb2 = Nothing
        '
        
        '
        '----------< DataGrid Close >----------------------------------------------
        If Not DataGrid1.DataSource Is Nothing Then
            Set DataGrid1.DataSource = Nothing
        End If
        
        Adodc1.Recordset.Close
        '
        wDb.Close
        Set wDb = Nothing
        
        ログDB.Caption = "現在ログに戻る"
     
        '----------< LOG.mdb Open >------------------------------------------------
        Call AdoDbOpen("Jet", wDb, wretfn, "", , GPwd)
     
        Call 検索項目_作成
        Call AdodcRefresh
     Else
        '----------< DataGrid Close >----------------------------------------------
        If Not DataGrid1.DataSource Is Nothing Then
            Set DataGrid1.DataSource = Nothing
        End If
        
        Adodc1.Recordset.Close
        '
        wDb.Close
        Set wDb = Nothing
     
        ログDB.Caption = "過去ログ参照"
        
        '----------< LOG.mdb Open >------------------------------------------------
        Call AdoDbOpen("Jet", wDb, GCurDir + "\LOG.mdb", "", , GPwd)
        
        Call 検索項目_作成
        Call AdodcRefresh
     End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
ログDB_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ ログDB_Click() でエラー" + vbCrLf + vbCrLf + _
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
    Unload Me
    
End Sub

'------------------------------------------------
' COMDLG
'------------------------------------------------
Private Function COMDLG(wTitle As String, wDir As String, wFilter As String, wFile) As String
'
On Error GoTo ComCancel
    CommonDialog1.DialogTitle = wTitle
    CommonDialog1.InitDir = wDir
    CommonDialog1.Filter = wFilter
    CommonDialog1.FileName = wFile
    CommonDialog1.CancelError = True
    
    CommonDialog1.ShowSave
    COMDLG = CommonDialog1.FileName

    Exit Function
ComCancel:
    COMDLG = ""
End Function

'------------------------------------------------
' ADODBOPEN_CHECK
'------------------------------------------------
Private Function ADODBOPEN_CHECK _
                   (pProvider As String, _
                    pAdoDb As ADODB.Connection, _
                    pDbName As String, _
                    Optional pSource As String = "", _
                    Optional pUID As String = "", _
                    Optional pPassword As String = "", _
                    Optional pMode As String = "") As Boolean
'
    Dim wstr As String
'
    ADODBOPEN_CHECK = False
'
    On Error GoTo Err_Hundle
        Select Case LCase(pProvider)
        Case "jet"
            wstr = "Provider=Microsoft.Jet.OLEDB.4.0"
            wstr = wstr & ";Data Source=" & pDbName
            wstr = wstr & ";Persist Security Info=False"
            wstr = wstr & ";Jet OLEDB:Database Password=" & pPassword
        
            pAdoDb.ConnectionString = wstr
    
            If pMode = "排他" Then
                pAdoDb.Mode = adModeShareExclusive
            Else
                pAdoDb.Mode = adModeUnknown
            End If
            
        End Select

        pAdoDb.Open
        
        pAdoDb.Close
        Set pAdoDb = Nothing
    On Error GoTo 0
'
    ADODBOPEN_CHECK = True
'
Exit Function
'----------< ERROR ROUTINE >--------------------------------------------------------
Err_Hundle:
    Resume Err_Hundle_END
Err_Hundle_END:
    Exit Function
End Function


