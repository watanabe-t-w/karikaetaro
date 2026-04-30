VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_M個別補助科目マスタ 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "個別補助科目マスタ"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11235
   Icon            =   "frm_M個別補助科目マスタ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
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
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7200
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
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   8
      Top             =   4800
      Width           =   10935
      Begin VB.TextBox 個別補助科目 
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
         IMEMode         =   1  'ｵﾝ
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox 個別補助科目名 
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
         IMEMode         =   1  'ｵﾝ
         Left            =   1920
         MaxLength       =   80
         TabIndex        =   1
         Top             =   1800
         Width           =   7695
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   120
         TabIndex        =   9
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
      Begin VB.Label L_銀行名 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   16
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label L_銀行番号 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   15
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label L_借入番号 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   14
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 借入番号"
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
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 銀行番号"
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
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 個別補助科目"
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
         TabIndex        =   11
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 個別補助科目名"
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
         TabIndex        =   10
         Top             =   1800
         Width           =   1815
      End
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
      Left            =   7320
      TabIndex        =   2
      Top             =   7200
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
      Left            =   9240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1815
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
      Caption         =   "個別補助科目マスタ"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3885
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6853
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
      Left            =   2280
      Top             =   7320
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
End
Attribute VB_Name = "frm_M個別補助科目マスタ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "個別補助科目マスタ"

Dim wRs As ADODB.Recordset
Dim wstr As String
'
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

'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    Dim j As Integer
'
    ' =========================================
    '                 初期設定
    ' =========================================
'    Me.Caption = GFcap
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
'
    ' =========================================
    '                 初期設定
    ' =========================================
    Call 登録後初期セット
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
    GWhere = " Where (1=1) " & GWhere
    
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "K.借入番号 As Grd借入番号,"
    wstr = wstr & "K.銀行番号 As Grd銀行番号,"
    wstr = wstr & "G.銀行名 As Grd銀行名,"
    wstr = wstr & "H.個別補助科目 As Grd個別補助科目,"
    wstr = wstr & "H.個別補助科目名 As Grd個別補助科目名"
    wstr = wstr & " FROM (DBDA010_借入金 As K"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr & " LEFT JOIN DABA030_個別補助科目マスタ As H"
    wstr = wstr & " ON K.借入番号 = H.借入番号"
    wstr = wstr & GWhere
    wstr = wstr + " Order By K.借入番号"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("借入番号", "", 2000, "L")
        Call XZMA010_DataGrid_Set("銀行番号", "", 1600, "L")
        Call XZMA010_DataGrid_Set("銀行名", "", 3000, "L")
        Call XZMA010_DataGrid_Set("個別補助科目", "", 1050, "L")
        Call XZMA010_DataGrid_Set("個別補助科目名", "", 3000, "L")
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
    Call CEkey.SetFs(個別補助科目, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd借入番号")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        L_借入番号.Caption = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd借入番号"))
        L_銀行番号.Caption = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd銀行番号"))
        L_銀行名.Caption = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd銀行名"))
        個別補助科目 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd個別補助科目"))
    On Error GoTo 0
    
    Call 画面セット(True)
   
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(個別補助科目, True)

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
    個別補助科目名.Text = ""
    削除 = 0
    
    ' =========================================
    '            補助科目マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DABA030_個別補助科目マスタ"
    wstr = wstr + " Where 借入番号 = '" & L_借入番号.Caption & "'"
    wstr = wstr + " And 銀行番号 = '" & L_銀行番号.Caption & "'"
    wstr = wstr + " And 個別補助科目 = '" & 個別補助科目.Text & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            If (L_借入番号.Caption <> "" And L_銀行番号.Caption <> "") And 個別補助科目 <> "" Then
                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
                If GRet = vbNo Then
                    新規変更.Caption = ""
                    wRs.Close
                    Set wRs = Nothing

                    Exit Function
                End If
                
                新規変更.Caption = "新規登録"
                Call CEkey.SetFs(個別補助科目名, True)
    
            End If
        Else
            画面セット = True
            
            Call CEkey.SetFs(個別補助科目名, True)
            新規変更.Caption = "変更"
            
            個別補助科目 = P8.FCStr(wRs("個別補助科目"))
            個別補助科目名 = P8.FCStr(wRs("個別補助科目名"))
            
        End If
    wRs.Close
    Set wRs = Nothing
    
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd借入番号 = '" + L_借入番号.Caption + "'")
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
' 検索_Click
'------------------------------------------------
Private Sub 検索_Click()
    Call 登録後初期セット
End Sub

'------------------------------------------------
' 個別補助科目_LostFocus
'------------------------------------------------
Private Sub 個別補助科目_LostFocus()
'
    On Error GoTo 個別補助科目_LostFocus_ERR
'
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1", "借入番号", "銀行番号", "削除", "CSV出力"
            Exit Sub
'        Case Else
'            Exit Sub
    End Select
   
    Call 画面セット(False)
    Call CEkey.AllSelect

    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
個別補助科目_LostFocus_ERR:
    pERR_MES = pPROGRAM_ID + "/ 個別補助科目_LostFocus() でエラー" + vbCrLf + vbCrLf + _
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
    L_借入番号.Caption = ""
    L_銀行番号.Caption = ""
    L_銀行名.Caption = ""
    Call 画面セット(False)
    新規変更.Caption = ""
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd借入番号 = '" + L_借入番号.Caption + "'")
    Call CEkey.SetFs(個別補助科目名, True)
'
End Sub

'------------------------------------------------
' 登録_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim wslog As String
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
'
    If P8.FCStr(個別補助科目) = "" Then
        MsgBox "個別補助科目が未入力です。", vbExclamation
        Call CEkey.SetFs(個別補助科目, True)
        Exit Sub
    End If

    If 個別補助科目名 = "" Then
        MsgBox "個別補助科目名が未入力です。", vbExclamation
        Call CEkey.SetFs(個別補助科目名, True)
        Exit Sub
    End If
'
    ' =========================================
    '            個別補助科目マスタ 更新処理
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DABA030_個別補助科目マスタ"
    wstr = wstr + " Where 借入番号 = '" & P8.FCStr(L_借入番号.Caption) & "'"
    wstr = wstr + " And 銀行番号 = '" & P8.FCStr(L_銀行番号.Caption) & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            wRs.AddNew
            
            wRs("借入番号") = P8.FCStr(L_借入番号.Caption)
            wRs("銀行番号") = P8.FCStr(L_銀行番号.Caption)
            
            wslog = "追加"
        End If
     
        wRs("個別補助科目") = P8.FCStr(個別補助科目.Text)
        wRs("個別補助科目名") = P8.FCStr(個別補助科目名.Text)
        
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
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 1 Then
        wslog = "削除"
    End If
    GLogStr = "借入番号=" & P8.FCStr(L_借入番号.Caption) & ","
    GLogStr = GLogStr & "銀行番号=" & P8.FCStr(L_銀行番号.Caption) & ","
    GLogStr = GLogStr & "個別補助科目=" & P8.FCStr(個別補助科目.Text) & ","
    GLogStr = GLogStr & "個別補助科目名=" & P8.FCStr(個別補助科目名.Text)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
'
    Adodc1.Refresh
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(個別補助科目名, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました", vbInformation
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
' 削除_Click
'------------------------------------------------
Private Sub 削除_Click()
'
    Dim wslog As String
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

    If P8.FCStr(L_借入番号.Caption) = "" Then
        Exit Sub
    End If
    
    If P8.FCStr(L_銀行番号.Caption) = "" Then
        Exit Sub
    End If
    
    If P8.FCStr(個別補助科目.Text) = "" Then
        Exit Sub
    End If
'
    GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
    If GRet = vbNo Then
        Exit Sub
    End If
'
    wstr = ""
    wstr = wstr & "Delete * From DABA030_個別補助科目マスタ"
    wstr = wstr + " Where 借入番号 = '" & L_借入番号.Caption & "'"
    wstr = wstr + " And 銀行番号 = '" & L_銀行番号.Caption & "'"
    wstr = wstr + " And 個別補助科目 = '" & 個別補助科目.Text & "'"
    GDb.Execute wstr
    
    DoEvents
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    wslog = "削除"
    GLogStr = "借入番号=" & P8.FCStr(L_借入番号.Caption) & ","
    GLogStr = GLogStr & "銀行番号=" & P8.FCStr(L_銀行番号.Caption) & ","
    GLogStr = GLogStr & "個別補助科目=" & P8.FCStr(個別補助科目.Text) & ","
    GLogStr = GLogStr & "個別補助科目名=" & P8.FCStr(個別補助科目名.Text)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
'
    Adodc1.Refresh
'
    ' =========================================
    '               画面セット
    ' =========================================
    L_借入番号.Caption = ""
    L_銀行番号.Caption = ""
    個別補助科目 = ""

    Call 画面セット(False)
    Call CEkey.SetFs(個別補助科目, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "削除しました。", vbInformation
'
End Sub

'------------------------------------------------
' CSV出力_Click
'------------------------------------------------
Private Sub CSV出力_Click()
'
    Call MX040_個別補助科目(GKeyName & "_" & "個別補助科目.csv")
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
    Unload Me
End Sub


