VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_K借入金検索 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入金検索"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13500
   Icon            =   "frm_K借入金検索.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13500
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
      Left            =   11280
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton 選択 
      Caption         =   "選択（F11)"
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
      Left            =   9360
      TabIndex        =   13
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "検索"
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   13215
      Begin VB.CheckBox 完済データ非表示 
         Caption         =   "完済データ非表示(前期まで)"
         Height          =   255
         Left            =   9960
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox Co_8 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1560
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.ComboBox Co_7 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   6360
         TabIndex        =   9
         Top             =   1800
         Width           =   3375
      End
      Begin VB.ComboBox Co_6 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   6360
         TabIndex        =   8
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox 番号 
         Height          =   285
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox 名称 
         Height          =   285
         IMEMode         =   4  '全角ひらがな
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox Co_1 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ComboBox Co_2 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   3375
      End
      Begin VB.ComboBox Co_3 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   6360
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
      Begin VB.ComboBox Co_4 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   6360
         TabIndex        =   6
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox Co_5 
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   6360
         TabIndex        =   7
         Top             =   1080
         Width           =   3375
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
         Left            =   11160
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label L_8 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "基準金利区分"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label L_7 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "利息区分"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   26
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label L_6 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "金利種別"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label L_RecCnt 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10680
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label L_名称 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   " 借入内容"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label L_番号 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   " 借入番号"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label L_5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "登録方法"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L_1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "銀行名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L_2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "借入金種別"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label L_3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "長短区分"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label L_4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "金利グループ名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "借入金　検索"
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
      Height          =   5805
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10239
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
      Left            =   120
      Top             =   9120
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
Attribute VB_Name = "frm_K借入金検索"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_K借入金検索"

Dim wRs As ADODB.Recordset
Dim wstr As String, GWhere As String

Dim wFname As String
Dim wsName As String, wsTbl As String

Dim pIndex As Integer
Dim wRecCnt As Long
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    Dim j As Integer
'
    Me.Caption = GFcap
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    ' =========================================
    '                 初期設定
    ' =========================================
    wFname = GStr

    GStr = "": GStr_1 = "": GStr_2 = ""
'
    L_1.Visible = False
    L_2.Visible = False
    L_3.Visible = False
    L_4.Visible = False
    L_5.Visible = False
    L_6.Visible = False
    L_7.Visible = False
    L_8.Visible = False

    Co_1.Visible = False
    Co_2.Visible = False
    Co_3.Visible = False
    Co_4.Visible = False
    Co_5.Visible = False
    Co_6.Visible = False
    Co_7.Visible = False
    Co_8.Visible = False

    L_RecCnt.Caption = ""
    
    '2017/09/29 watanabe 完済データ
    完済データ非表示.Visible = False
'
    G借現.保証会社区分 = ""
    G借現.保証会社区分名 = ""
    G借現.融資区分 = ""
    G借現.融資区分名 = ""
    G借現.制度融資区分 = 0
    G借現.銀行番号 = ""
    G借現.銀行名 = ""
    G借現.有担保フラグ = 0
    G借現.借入番号 = ""
    G借現.残回数 = 0
    G借現.残据置 = 0
    G借現.利率 = 0
    G借現.保証料率 = 0
    G借現.設備フラグ = 0
    G借現.融資金額 = 0
    G借現.毎月返済額 = 0
    G借現.融資残高 = 0
'
    Call 検索項目_作成
    Call AdodcRefresh
'
End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    DoEvents

    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
    If KeyCode = vbKeyF11 Then
        Call 選択_Click
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
' DataGrid1_Click
'------------------------------------------------
Private Sub DataGrid1_Click()
'
    Call CEkey.SetFs(番号, True)
End Sub
'------------------------------------------------
' DataGrid1_DblClick
'------------------------------------------------
Private Sub DataGrid1_DblClick()
    
    Call DataGrid1_LostFocus
    Call 選択_Click
    
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
        
        If wFname <> "借換現状追加" Then
            名称 = P8.FCStr(Adodc1.Recordset.Fields.Item("名称"))
        End If
        
        Select Case wFname
            Case "借換現状追加"
                G借現.借入番号 = 番号
                G借現.保証会社区分 = P8.FCStr(Adodc1.Recordset.Fields.Item("保証会社区分"))
                G借現.保証会社区分名 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd保証会社名"))
                G借現.融資区分 = P8.FCStr(Adodc1.Recordset.Fields.Item("融資区分"))
                G借現.融資区分名 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd融資区分名"))
                G借現.銀行番号 = P8.FCStr(Adodc1.Recordset.Fields.Item("銀行番号"))
                G借現.銀行名 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd銀行名"))
                G借現.制度融資区分 = P8.FCDbl(Adodc1.Recordset.Fields.Item("制度融資区分"))
                G借現.有担保フラグ = P8.FCDbl(Adodc1.Recordset.Fields.Item("有担保フラグ"))
                G借現.初回返済年月 = P8.FCStr(Adodc1.Recordset.Fields.Item("初回返済年月"))
                G借現.最終返済年月 = P8.FCStr(Adodc1.Recordset.Fields.Item("最終返済年月"))
                G借現.返済単位月数 = P8.FCDbl(Adodc1.Recordset.Fields.Item("返済単位月数"))
                G借現.残据置 = P8.FCDbl(Adodc1.Recordset.Fields.Item("残据置"))
                G借現.利率 = P8.FCDbl(Adodc1.Recordset.Fields.Item("利率"))
                G借現.保証料率 = P8.FCDbl(Adodc1.Recordset.Fields.Item("保証料率"))
                G借現.設備フラグ = P8.FCDbl(Adodc1.Recordset.Fields.Item("設備フラグ"))
                G借現.融資金額 = P8.FCDbl(Adodc1.Recordset.Fields.Item("Grd融資金額"))
                G借現.毎月返済額 = P8.FCDbl(Adodc1.Recordset.Fields.Item("Grd毎月返済額"))
                G借現.融資残高 = P8.FCDbl(Adodc1.Recordset.Fields.Item("Grd融資残高"))
        
        End Select
        
    On Error GoTo 0
    
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "番号 = '" + 番号 + "'")
    
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(番号, True)

Exit_Sub:
    Exit Sub
    '---------------------------------------------------
Err_Hundle:
    If Err.Number = 91 Then Resume Next
    If Err.Number = 94 Then Resume Next
    MsgBox CStr(Err.Number) + ":" + Err.Description
    Resume Exit_Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
    Unload Me
'
    Select Case wFname
        'Case "設備計画登録"
        '    FCC010_設備計画登録.Show vbModal
        Case "借入金登録", "貸付登録"
            GStr = wFname
'            Unload frm_I借入金登録
            frm_I借入金登録.Enabled = True
        Case "借入金台帳"
            frm_R借入金台帳.Enabled = True
        Case "借入明細表", "貸付明細表", "社債明細表"
            'FBA010_帳票範囲指定.Show vbModal
            frm_R借入金明細表.Enabled = True
        Case "借入金時価評価明細表"
'            frm_R借入金時価評価明細表.Enabled = True
        Case "利息明細表"
            frm_R利息前払未払明細表.Enabled = True
        
        Case "借換現状追加"
    End Select
'
End Sub

Private Sub 完済データ非表示_Click()
    Call AdodcRefresh
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
' 選択_Click
'------------------------------------------------
Private Sub 選択_Click()
'
    GStr_1 = P8.FCStr(番号)
'
    '----------< DataGrid Close >----------------------------------------------
    If Not DataGrid1.DataSource Is Nothing Then
        Set DataGrid1.DataSource = Nothing
    End If
    Adodc1.Recordset.Close
'

'
    Select Case wFname
        'Case "設備計画登録"
        '    FCC010_設備計画登録.Show vbModal
        Case "借入金登録", "貸付登録"
            GStr = wFname
'            Unload frm_I借入金登録
'            frm_I借入金登録.Show
            frm_I借入金登録.Enabled = True
            frm_I借入金登録.借入番号 = GStr_1
            Call frm_I借入金登録.画面セット呼出
        Case "借入金台帳"
            frm_R借入金台帳.借入番号 = GStr_1
            frm_R借入金台帳.Enabled = True
        Case "借入明細表", "貸付明細表", "社債明細表"
            'FBA010_帳票範囲指定.Show vbModal
            frm_R借入金明細表.Enabled = True
            frm_R借入金明細表.借入番号 = GStr_1
        Case "利息明細表"
            frm_R利息前払未払明細表.Enabled = True
            frm_R利息前払未払明細表.借入番号 = GStr_1
            
        Case "借入金時価評価明細表"
'            frm_R借入金時価評価明細表.Enabled = True
'            frm_R借入金時価評価明細表.借入番号 = GStr_1
            
        Case "借換現状追加"
    End Select
'
    Unload Me
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
'
    Select Case wFname
        'Case "設備計画登録"
        '    FCC010_設備計画登録.Show vbModal
        Case "借入金登録", "貸付登録"
            GStr = wFname
'            Unload frm_I借入金登録
            frm_I借入金登録.Enabled = True
        Case "借入金台帳"
            frm_R借入金台帳.Enabled = True
        Case "借入明細表", "貸付明細表", "社債明細表"
            'FBA010_帳票範囲指定.Show vbModal
            frm_R借入金明細表.Enabled = True
        Case "借入金時価評価明細表"
'            frm_R借入金時価評価明細表.Enabled = True
        Case "利息明細表"
            frm_R利息前払未払明細表.Enabled = True
        
        Case "借換現状追加"
    End Select
'
End Sub

'------------------------------------------------
' 検索項目_作成
'------------------------------------------------
Private Sub 検索項目_作成()
'
    wsName = ""
    wsTbl = ""
'
    Select Case wFname
        Case "設備計画登録", "設備計画明細表"
            Call 検索項目_作成_設備計画
        
        Case "借入金登録", "借入明細表", "借入金台帳", "利息明細表", "社債明細表"
            
            wsName = "借入"
            wsTbl = "DBDA010_借入金"
            
            Call 検索項目_作成_借入金
            
        Case "借入金時価評価明細表"
            
            wsName = "借入"
            wsTbl = "DBDA010_借入金"
            
            Call 検索項目_作成_借入金時価評価
        
        Case "貸付登録", "貸付明細表"
            wsName = "貸付"
            wsTbl = "DBDA010_貸付金"
            
            Call 検索項目_作成_借入金
            
        Case "借換現状追加"
            Call 検索項目_作成_借換現状追加
            
    End Select
'
End Sub

'------------------------------------------------
' AdodcRefresh
'------------------------------------------------
Private Sub AdodcRefresh()
'
    wRecCnt = 0
    Select Case wFname
        Case "設備計画登録", "設備計画明細表"
            Call AdodcRefresh_設備計画
        Case "借入金登録", "貸付登録", "借入明細表", "貸付明細表", "借入金台帳", "利息明細表", "社債明細表"
            Call AdodcRefresh_借入金
        Case "借入金時価評価明細表"
            Call AdodcRefresh_借入金時価評価
        Case "借換現状追加"
            Call AdodcRefresh_借換現状追加
    End Select
'
    wRecCnt = Adodc1.Recordset.RecordCount
    L_RecCnt.Caption = wRecCnt & " 件"
'
End Sub

'------------------------------------------------
' 検索項目_作成
'------------------------------------------------
Private Sub 検索項目_作成_設備計画()
'
    L_1.Caption = "設備計画番号"
    L_2.Caption = "設備リストラ番号"
    L_3.Caption = "部門"
    L_4.Caption = "勘定科目"
    L_5.Caption = ""
    
    L_番号.Caption = "設備計画番号"
    L_名称.Caption = "設備名"
'
    L_1.Visible = True
    L_2.Visible = True
    L_3.Visible = True
    L_4.Visible = True
    L_5.Visible = False
    
    Co_1.Visible = True
    Co_2.Visible = True
    Co_3.Visible = True
    Co_4.Visible = True
    Co_5.Visible = False
'
    Co_1.Clear
    wstr = ""
    wstr = wstr & "Select 設備計画番号"
    wstr = wstr & " From DBCA010_設備計画"
    wstr = wstr & " Group By 設備計画番号"
    wstr = wstr & " Order By 設備計画番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_1.AddItem (P8.FCStr(wRs("設備計画番号")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Co_2.Clear
    wstr = ""
    wstr = wstr & "Select 設備リストラ番号"
    wstr = wstr & " From DBCA010_設備計画"
    wstr = wstr & " Group By 設備リストラ番号"
    wstr = wstr & " Order By 設備リストラ番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_2.AddItem (P8.FCStr(wRs("設備リストラ番号")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Co_3.Clear
    wstr = ""
    wstr = wstr & "Select 部門番号,部門名"
    wstr = wstr & " From DAAC020_固定資産部門マスタ"
    wstr = wstr & " Group By 部門番号,部門名"
    wstr = wstr & " Order By 部門番号,部門名"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_3.AddItem (P8.FCStr(wRs("部門名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Co_4.Clear
    wstr = ""
    wstr = wstr & "Select 勘定科目番号,勘定科目名"
    wstr = wstr & " From DAAC010_固定資産勘定科目マスタ"
    wstr = wstr & " Group By 勘定科目番号,勘定科目名"
    wstr = wstr & " Order By 勘定科目番号,勘定科目名"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_4.AddItem (P8.FCStr(wRs("勘定科目名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
End Sub

Private Sub 検索項目_作成_借入金()
'
    Dim wdDate As Date
'
    L_1.Caption = "銀行名"
    L_2.Caption = "借入金種別"
    L_3.Caption = "長短区分"
    L_4.Caption = "金利グループ名"
    L_5.Caption = "登録方法"
    L_6.Caption = "金利種別"
    L_7.Caption = "利息区分"
    L_8.Caption = "基準金利区分"

    L_番号.Caption = wsName & "番号"
    L_名称.Caption = wsName & "内容"
'
    L_1.Visible = True
    L_2.Visible = True
    L_3.Visible = True
    L_4.Visible = True
    L_5.Visible = True
    L_6.Visible = True
    L_7.Visible = True
    L_8.Visible = True
    
    Co_1.Visible = True
    Co_2.Visible = True
    Co_3.Visible = True
    Co_4.Visible = True
    Co_5.Visible = True
    Co_6.Visible = True
    Co_7.Visible = True
    Co_8.Visible = True
'
    Co_1.Clear
    wstr = ""
    wstr = wstr & "Select 銀行番号,銀行名"
    wstr = wstr & " From DAAA040_銀行マスタ"
    wstr = wstr & " Where 取消フラグ = 0"
    wstr = wstr & " Order By 銀行番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_1.AddItem (P8.FCStr(wRs("銀行名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Co_2.Clear
    wstr = ""
    wstr = wstr & "Select 借入金種別区分,借入金種別名"
    wstr = wstr & " From DAAA116_借入金種別"
    wstr = wstr & " Where 取消フラグ = 0"
    wstr = wstr & " Order By 借入金種別区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_2.AddItem (P8.FCStr(wRs("借入金種別名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    With Co_3
        .Clear
        .AddItem ""
        .AddItem "短期"
        .AddItem "長期"
    End With
'
    Co_4.Clear
    wstr = ""
    wstr = wstr & "Select 金利グループ区分,金利グループ名"
    wstr = wstr & " From DAAA115_金利シミュレーショングループ"
    wstr = wstr & " Where 取消フラグ = 0"
    wstr = wstr & " Order By 金利グループ区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_4.AddItem (P8.FCStr(wRs("金利グループ名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    With Co_5
        .Clear
        .AddItem ""
        .AddItem "標準登録"
        .AddItem "入力登録"
    End With
'
    With Co_6
        .Clear
        .AddItem ""
        .AddItem "固定金利"
        .AddItem "変動金利"
    End With
'
    With Co_7
        .Clear
        .AddItem ""
        .AddItem "利息先払"
        .AddItem "利息後払"
    End With
'
    Co_8.Clear
    wstr = ""
    wstr = wstr & "Select 基準金利区分,基準金利名"
    wstr = wstr & " From DAAA116_基準金利"
    wstr = wstr & " Where 取消フラグ = 0"
    wstr = wstr & " Order By 基準金利区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_8.AddItem (P8.FCStr(wRs("基準金利名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    '2017/09/29 watanabe 完済データ
    完済データ非表示.Visible = True
    完済データ非表示.Value = 1
'
End Sub

Private Sub 検索項目_作成_借入金時価評価()
'
    L_1.Caption = "銀行名"
    L_2.Caption = "借入金種別"
    L_3.Caption = "長短区分"
    L_4.Caption = "金利グループ名"
    L_5.Caption = "登録方法"
    L_6.Caption = "金利種別"
    L_7.Caption = "利息区分"
    L_8.Caption = "基準金利区分"

    L_番号.Caption = wsName & "番号"
    L_名称.Caption = wsName & "内容"
'
    L_1.Visible = True
    L_2.Visible = True
    L_3.Visible = True
    L_4.Visible = True
    L_5.Visible = True
    L_6.Visible = True
    L_7.Visible = True
    L_8.Visible = True
    
    Co_1.Visible = True
    Co_2.Visible = True
    Co_3.Visible = True
    Co_4.Visible = True
    Co_5.Visible = True
    Co_6.Visible = True
    Co_7.Visible = True
    Co_8.Visible = True
'
    Co_1.Clear
    wstr = ""
    wstr = wstr & "Select 銀行番号,銀行名"
    wstr = wstr & " From DAAA040_銀行マスタ"
    wstr = wstr & " Where 取消フラグ = 0"
    wstr = wstr & " Order By 銀行番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_1.AddItem (P8.FCStr(wRs("銀行名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
'    Co_2.Clear
'    wstr = ""
'    wstr = wstr & "Select 借入金種別区分,借入金種別名"
'    wstr = wstr & " From DAAA116_借入金種別"
'    wstr = wstr & " Where 取消フラグ = 0"
'    wstr = wstr & " Order By 借入金種別区分"
'    Call AdoRecordsetOpen(GDb, wRs, wstr)
'        Do Until wRs.EOF
'            Co_2.AddItem (P8.FCStr(wRs("借入金種別名")))
'
'            wRs.MoveNext
'        Loop
'    wRs.Close
'    Set wRs = Nothing
    With Co_2
        .Clear
        .AddItem "借入金"
    End With
    Co_2 = Co_2.List(0)
'
'
    With Co_3
        .Clear
        .AddItem "長期"
    End With
    Co_3 = Co_3.List(0)
'
    Co_4.Clear
    wstr = ""
    wstr = wstr & "Select 金利グループ区分,金利グループ名"
    wstr = wstr & " From DAAA115_金利シミュレーショングループ"
    wstr = wstr & " Where 取消フラグ = 0"
    wstr = wstr & " Order By 金利グループ区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_4.AddItem (P8.FCStr(wRs("金利グループ名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    With Co_5
        .Clear
        .AddItem ""
        .AddItem "標準登録"
        .AddItem "入力登録"
    End With
'
    With Co_6
        .Clear
        .AddItem "固定金利"
    End With
    Co_6 = Co_6.List(0)
'
    With Co_7
        .Clear
        .AddItem ""
        .AddItem "利息先払"
        .AddItem "利息後払"
    End With
'
    Co_8.Clear
    wstr = ""
    wstr = wstr & "Select 基準金利区分,基準金利名"
    wstr = wstr & " From DAAA116_基準金利"
    wstr = wstr & " Where 取消フラグ = 0"
    wstr = wstr & " Order By 基準金利区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_8.AddItem (P8.FCStr(wRs("基準金利名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
End Sub

Private Sub 検索項目_作成_借換現状追加()
'
    L_1.Caption = "保証会社名"
    L_2.Caption = "銀行名"
    L_3.Caption = "融資区分名"

    L_番号.Caption = "借入番号"
    L_名称.Caption = "借入内容"
'
    L_1.Visible = True
    L_2.Visible = True
    L_3.Visible = True
    
    Co_1.Visible = True
    Co_2.Visible = True
    Co_3.Visible = True
'
    Co_1.Clear
    wstr = ""
    wstr = wstr & "SELECT 保証会社区分,保証会社区分名"
    wstr = wstr & " From DAAA100_保証会社区分"
    wstr = wstr & " GROUP BY 保証会社区分,保証会社区分名,代表区分"
    wstr = wstr & " Having 代表区分=0"
    wstr = wstr & " ORDER BY 保証会社区分,保証会社区分名"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_1.AddItem (P8.FCStr(wRs("保証会社区分名")))

            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Co_2.Clear
    wstr = ""
    wstr = wstr & "Select 銀行番号,銀行名"
    wstr = wstr & " From DAAA040_銀行マスタ"
    wstr = wstr & " Group By 銀行番号,銀行名"
    wstr = wstr & " Order By 銀行番号,銀行名"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_2.AddItem (P8.FCStr(wRs("銀行名")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Co_3.Clear
    wstr = ""
    wstr = wstr & "Select 融資区分,融資区分名"
    wstr = wstr & " From DAAA110_融資区分"
    wstr = wstr & " Group By 融資区分,融資区分名"
    wstr = wstr & " Order By 融資区分,融資区分名"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Co_3.AddItem (P8.FCStr(wRs("融資区分名")))

            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
End Sub

'------------------------------------------------
' AdodcRefresh
'------------------------------------------------
Private Sub AdodcRefresh_設備計画()
'
    On Error GoTo AdodcRefresh_設備計画_ERR
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
    
    If P8.FCStr(番号) <> "" Then
        GWhere = GWhere & " And S.設備番号 Like '%" & P8.FCStr(番号) & "%'"
    End If
    If P8.FCStr(Co_1.Text) <> "" Then
        GWhere = GWhere & " And 設備計画番号 = '" & P8.FCStr(Co_1.Text) & "'"
    End If
    If P8.FCStr(Co_2.Text) <> "" Then
        GWhere = GWhere & " And 設備リストラ番号 = '" & P8.FCStr(Co_2.Text) & "'"
    End If
    If P8.FCStr(Co_3.Text) <> "" Then
        GWhere = GWhere & " And B.部門名 = '" & P8.FCStr(Co_3.Text) & "'"
    End If
    If P8.FCStr(Co_4.Text) <> "" Then
        GWhere = GWhere & " And K.勘定科目名 = '" & P8.FCStr(Co_4.Text) & "'"
    End If
    
    GWhere = " Where (1=1) " & GWhere
    
    wstr = ""
    wstr = wstr & "SELECT"
    wstr = wstr & " S.設備番号 As 番号,"
    wstr = wstr & " S.設備名 As 名称,"
    
    wstr = wstr & " S.設備番号 As Grd設備番号,"
    wstr = wstr & " S.設備名 As Grd設備名,"
    wstr = wstr & " B.部門名 As Grd部門,"
    wstr = wstr & " K.勘定科目名 As Grd勘定科目,"
    wstr = wstr & " IIF(S.Sm区分 = 1,'○','') As Grdシミュレーション,"
    wstr = wstr & " S.設備計画番号 As Grd設備計画番号,"
    wstr = wstr & " Format(S.設備年月,'" & Gfmt年月 & "') As Grd設備年月,"
    wstr = wstr & " Format(S.設備金額,'#,##0') As Grd設備金額,"
    wstr = wstr & " IIF(S.手入力フラグ = 0,'○','') As Grd取込,"
    wstr = wstr & " IIF(S.修正不可F = 1,'○','') As Grd売却,"
    wstr = wstr & " IIF(S.取消フラグ = 0,'','×') As Grd取消"
    
    wstr = wstr & " FROM (DBCA010_設備計画 As S"
    wstr = wstr & " LEFT JOIN DAAC020_固定資産部門マスタ As B"
    wstr = wstr & " ON S.部門番号 = B.部門番号)"
    wstr = wstr & " LEFT JOIN DAAC010_固定資産勘定科目マスタ As K"
    wstr = wstr & " ON S.勘定科目番号 = K.勘定科目番号"
    wstr = wstr & GWhere
    wstr = wstr & " ORDER BY B.部門番号, S.設備番号"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("設備番号", "", 2100, "L")
        Call XZMA010_DataGrid_Set("設備名", "", 2100, "L")
        Call XZMA010_DataGrid_Set("部門", "", 2100, "L")
        Call XZMA010_DataGrid_Set("勘定科目", "", 2100, "L")
        Call XZMA010_DataGrid_Set("シミュレーション", "sm", 500, "C")
        Call XZMA010_DataGrid_Set("設備計画番号", "計画番号", 1300, "L")
        Call XZMA010_DataGrid_Set("設備年月", "", 1100, "L")
        Call XZMA010_DataGrid_Set("設備金額", "", 2000, "R")
        Call XZMA010_DataGrid_Set("取込", "", 500, "C")
        Call XZMA010_DataGrid_Set("売却", "", 500, "C")
        Call XZMA010_DataGrid_Set("取消", "", 500, "C")
    Call XZMA010_DataGrid_Action(DataGrid1)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_設備計画_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh_設備計画() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Private Sub AdodcRefresh_借入金()
'
    Dim wdDate As Date
'
    On Error GoTo AdodcRefresh_借入金_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
'    Call MXA030_DataGridInit(DataGrid1)
    DataGrid1.AllowRowSizing = False
    DataGrid1.HeadFont.Size = 9
    DataGrid1.HeadFont.Bold = True
    DataGrid1.Font.Size = 9
    DataGrid1.BackColor = C_White
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
    
    If P8.FCStr(番号) <> "" Then
        GWhere = GWhere & " And K.借入番号 Like '%" & P8.FCStr(番号) & "%'"
    End If
    If P8.FCStr(名称) <> "" Then
        GWhere = GWhere & " And K.借入内容 Like '%" & P8.FCStr(名称) & "%'"
    End If
    If P8.FCStr(Co_1.Text) <> "" Then
        GWhere = GWhere & " And G.銀行名 = '" + P8.FCStr(Co_1.Text) + "'"
    End If
    If P8.FCStr(Co_2.Text) <> "" Then
        GWhere = GWhere & " And S.借入金種別名 = '" + P8.FCStr(Co_2.Text) + "'"
    End If
    If P8.FCStr(Co_3.Text) <> "" Then
        If P8.FCStr(Co_3.Text) = "短期" Then
            GWhere = GWhere & " And K.長短区分 =" & P8.FCDbl(XMXA020_区分("長短区分", "短期"))
        ElseIf P8.FCStr(Co_3.Text) = "長期" Then
            GWhere = GWhere & " And K.長短区分 =" & P8.FCDbl(XMXA020_区分("長短区分", "長期"))
        End If
    End If
    If P8.FCStr(Co_4.Text) <> "" Then
        GWhere = GWhere & " And GR.金利グループ名 = '" + P8.FCStr(Co_4.Text) + "'"
    End If
    If P8.FCStr(Co_5.Text) <> "" Then
        If P8.FCStr(Co_5.Text) = "標準登録" Then
            GWhere = GWhere & " And K.手入力区分 =" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録"))
        ElseIf P8.FCStr(Co_5.Text) = "入力登録" Then
            GWhere = GWhere & " And K.手入力区分 =" & P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
        End If
    End If
    If P8.FCStr(Co_6.Text) <> "" Then
        If P8.FCStr(Co_6.Text) = "固定金利" Then
            GWhere = GWhere & " And K.金利種別 =" & P8.FCDbl(XMXA020_区分("金利種別", "固定金利"))
        ElseIf P8.FCStr(Co_6.Text) = "変動金利" Then
            GWhere = GWhere & " And K.金利種別 =" & P8.FCDbl(XMXA020_区分("金利種別", "変動金利"))
        End If
    End If
    If P8.FCStr(Co_7.Text) <> "" Then
        If P8.FCStr(Co_7.Text) = "利息先払" Then
            GWhere = GWhere & " And K.利息区分 ='" & P8.FCStr(XMXA020_区分("利息区分", "利息先払")) & "'"
        ElseIf P8.FCStr(Co_7.Text) = "利息後払" Then
            GWhere = GWhere & " And K.利息区分 ='" & P8.FCStr(XMXA020_区分("利息区分", "利息後払")) & "'"
        End If
    End If
    If P8.FCStr(Co_8.Text) <> "" Then
        GWhere = GWhere & " And KK.基準金利名 = '" + P8.FCStr(Co_8.Text) + "'"
    End If
    
    '2017/09/28 watanabe 完済データ
    If 完済データ非表示.Value = 1 Then
        wdDate = DateAdd("yyyy", -1, CDate(C年月日.年度開始年月日(Format(Now, "yyyy"), "西暦")))
        GWhere = GWhere & " And ("
        GWhere = GWhere & "     (解約実行日 is null And format(K.最終返済実行日,'yyyy/mm/dd')>'" + Format(wdDate, "yyyy/mm/dd") + "')"
        GWhere = GWhere & "  Or (解約実行日 is not null And format(K.解約実行日,'yyyy/mm/dd')>'" + Format(wdDate, "yyyy/mm/dd") + "')"
        GWhere = GWhere & " )"
    End If
    
    GWhere = " Where (1=1) " & GWhere
    
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " K.借入番号 As 番号,"
    wstr = wstr & " K.借入内容 As 名称,"
    
    wstr = wstr & " K.借入番号 As Grd借入番号,"
    wstr = wstr & " G.銀行名 As Grd銀行名,"
    wstr = wstr & " Format(K.融資金額,'#,##0') As Grd融資金額,"
    wstr = wstr & " S.借入金種別名 As Grd借入金種別名,"
    wstr = wstr & " IIF(K.長短区分=" & P8.FCDbl(XMXA020_区分("長短区分", "短期")) & ",'短期','長期') As Grd長短区分,"
    wstr = wstr + " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動','固定') As Grd金利種別,"
    wstr = wstr + " IIF(K.利息区分 = '" & P8.FCStr(XMXA020_区分("利息区分", "利息先払")) & "','先払','後払') As Grd利息区分,"
    wstr = wstr + " IIF(K.手入力区分 = 0,'標準','入力') As Grd登録方法,"
    wstr = wstr & " K.借入内容 As Grd借入内容,"
    wstr = wstr & " KK.基準金利名 As Grd基準金利名,"
    wstr = wstr & " GR.金利グループ名 As Grd金利グループ名,"
    
    '2017/09/29 完済データ
    wstr = wstr + " IIf(isnull(K.解約実行日),Format(K.最終返済実行日,'" & Gfmt年月日 & "') ,Format(K.解約実行日,'" & Gfmt年月日 & "')) AS Grd最終返済日,"

    wstr = wstr & " K.金融リストラ番号 As Grd金融リストラ番号,"
    wstr = wstr & " IIF(K.取消フラグ = 0,'','×') As Grd取消"
    
    wstr = wstr & " FROM (((" & wsTbl & " As K"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr = wstr & " LEFT JOIN DAAA116_基準金利 As KK"
    wstr = wstr & " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As GR"
    wstr = wstr & " ON K.金利グループ区分 = GR.金利グループ区分"

    wstr = wstr + GWhere
    wstr = wstr + " Order By K.借入番号"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("借入番号", "", 1900, "L")
        Call XZMA010_DataGrid_Set("銀行名", "", 1800, "L")
        Call XZMA010_DataGrid_Set("融資金額", "", 1500, "R")
        Call XZMA010_DataGrid_Set("借入金種別名", "借入種別", 1400, "L")
        Call XZMA010_DataGrid_Set("長短区分", "長短", 700, "C")
        Call XZMA010_DataGrid_Set("金利種別", "金利", 700, "C")
        Call XZMA010_DataGrid_Set("利息区分", "利息", 700, "C")
        Call XZMA010_DataGrid_Set("登録方法", "登録", 700, "C")
        Call XZMA010_DataGrid_Set("借入内容", "", 1400, "L")
        Call XZMA010_DataGrid_Set("基準金利名", "", 1400, "L")
        Call XZMA010_DataGrid_Set("金利グループ名", "", 1400, "L")
        '2017/09/28 watanabe 完済データ
        Call XZMA010_DataGrid_Set("最終返済日", "最終返済日", 1200, "L")
        Call XZMA010_DataGrid_Set("金融リストラ番号", "借入SM番号", 1300, "L")
'        Call XZMA010_DataGrid_Set("取消", "", 550, "C")
    Call XZMA010_DataGrid_Action(DataGrid1)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_借入金_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh_借入金() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Private Sub AdodcRefresh_借入金時価評価()
'
    On Error GoTo AdodcRefresh_借入金時価評価_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
'    Call MXA030_DataGridInit(DataGrid1)
    DataGrid1.AllowRowSizing = False
    DataGrid1.HeadFont.Size = 9
    DataGrid1.HeadFont.Bold = True
    DataGrid1.Font.Size = 9
    DataGrid1.BackColor = C_White
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
    
    If P8.FCStr(番号) <> "" Then
        GWhere = GWhere & " And K.借入番号 Like '%" & P8.FCStr(番号) & "%'"
    End If
    If P8.FCStr(名称) <> "" Then
        GWhere = GWhere & " And K.借入内容 Like '%" & P8.FCStr(名称) & "%'"
    End If
    If P8.FCStr(Co_1.Text) <> "" Then
        GWhere = GWhere & " And G.銀行名 = '" + P8.FCStr(Co_1.Text) + "'"
    End If
    If P8.FCStr(Co_2.Text) <> "" Then
        GWhere = GWhere & " And S.借入金種別名 = '" + P8.FCStr(Co_2.Text) + "'"
    End If
    If P8.FCStr(Co_3.Text) <> "" Then
        If P8.FCStr(Co_3.Text) = "短期" Then
            GWhere = GWhere & " And K.長短区分 =" & P8.FCDbl(XMXA020_区分("長短区分", "短期"))
        ElseIf P8.FCStr(Co_3.Text) = "長期" Then
            GWhere = GWhere & " And K.長短区分 =" & P8.FCDbl(XMXA020_区分("長短区分", "長期"))
        End If
    End If
    If P8.FCStr(Co_4.Text) <> "" Then
        GWhere = GWhere & " And GR.金利グループ名 = '" + P8.FCStr(Co_4.Text) + "'"
    End If
    If P8.FCStr(Co_5.Text) <> "" Then
        If P8.FCStr(Co_5.Text) = "標準登録" Then
            GWhere = GWhere & " And K.手入力区分 =" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録"))
        ElseIf P8.FCStr(Co_5.Text) = "入力登録" Then
            GWhere = GWhere & " And K.手入力区分 =" & P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
        End If
    End If
    If P8.FCStr(Co_6.Text) <> "" Then
        If P8.FCStr(Co_6.Text) = "固定金利" Then
            GWhere = GWhere & " And K.金利種別 =" & P8.FCDbl(XMXA020_区分("金利種別", "固定金利"))
        ElseIf P8.FCStr(Co_6.Text) = "変動金利" Then
            GWhere = GWhere & " And K.金利種別 =" & P8.FCDbl(XMXA020_区分("金利種別", "変動金利"))
        End If
    End If
    If P8.FCStr(Co_7.Text) <> "" Then
        If P8.FCStr(Co_7.Text) = "利息先払" Then
            GWhere = GWhere & " And K.利息区分 ='" & P8.FCStr(XMXA020_区分("利息区分", "利息先払")) & "'"
        ElseIf P8.FCStr(Co_7.Text) = "利息後払" Then
            GWhere = GWhere & " And K.利息区分 ='" & P8.FCStr(XMXA020_区分("利息区分", "利息後払")) & "'"
        End If
    End If
    If P8.FCStr(Co_8.Text) <> "" Then
        GWhere = GWhere & " And KK.基準金利名 = '" + P8.FCStr(Co_8.Text) + "'"
    End If
    
    GWhere = " Where (1=1) " & GWhere
    
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " K.借入番号 As 番号,"
    wstr = wstr & " K.借入内容 As 名称,"
    
    wstr = wstr & " K.借入番号 As Grd借入番号,"
    wstr = wstr & " G.銀行名 As Grd銀行名,"
    wstr = wstr & " Format(K.融資金額,'#,##0') As Grd融資金額,"
    wstr = wstr & " S.借入金種別名 As Grd借入金種別名,"
    wstr = wstr & " IIF(K.長短区分=" & P8.FCDbl(XMXA020_区分("長短区分", "短期")) & ",'短期','長期') As Grd長短区分,"
    wstr = wstr + " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動','固定') As Grd金利種別,"
    wstr = wstr + " IIF(K.利息区分 = '" & P8.FCStr(XMXA020_区分("利息区分", "利息先払")) & "','先払','後払') As Grd利息区分,"
    wstr = wstr + " IIF(K.手入力区分 = 0,'標準','入力') As Grd登録方法,"
    wstr = wstr & " K.借入内容 As Grd借入内容,"
    wstr = wstr & " KK.基準金利名 As Grd基準金利名,"
    wstr = wstr & " GR.金利グループ名 As Grd金利グループ名,"
    wstr = wstr & " K.金融リストラ番号 As Grd金融リストラ番号,"

    wstr = wstr & " IIF(K.取消フラグ = 0,'','×') As Grd取消"
    
    wstr = wstr & " FROM (((" & wsTbl & " As K"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr = wstr & " LEFT JOIN DAAA116_基準金利 As KK"
    wstr = wstr & " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As GR"
    wstr = wstr & " ON K.金利グループ区分 = GR.金利グループ区分"

    wstr = wstr + GWhere
    wstr = wstr + " Order By K.借入番号"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("借入番号", "", 1900, "L")
        Call XZMA010_DataGrid_Set("銀行名", "", 1800, "L")
        Call XZMA010_DataGrid_Set("融資金額", "", 1500, "R")
        Call XZMA010_DataGrid_Set("借入金種別名", "借入種別", 1400, "L")
        Call XZMA010_DataGrid_Set("長短区分", "長短", 700, "C")
        Call XZMA010_DataGrid_Set("金利種別", "金利", 700, "C")
        Call XZMA010_DataGrid_Set("利息区分", "利息", 700, "C")
        Call XZMA010_DataGrid_Set("登録方法", "登録", 700, "C")
        Call XZMA010_DataGrid_Set("借入内容", "", 1400, "L")
        Call XZMA010_DataGrid_Set("基準金利名", "", 1400, "L")
        Call XZMA010_DataGrid_Set("金利グループ名", "", 1400, "L")
        Call XZMA010_DataGrid_Set("金融リストラ番号", "借入SM番号", 1300, "L")
'        Call XZMA010_DataGrid_Set("取消", "", 550, "C")
    Call XZMA010_DataGrid_Action(DataGrid1)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_借入金時価評価_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh_借入金時価評価() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Private Sub AdodcRefresh_借換現状追加()
'
    On Error GoTo AdodcRefresh_借入金_ERR
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
    GWhere = GWhere & " And KH.融資残高<>0"
    GWhere = GWhere & " And (((KH.代表借入番号 Is Null"
    GWhere = GWhere & " Or KH.代表借入番号='')"
    GWhere = GWhere & " Or (KH.代表借入番号<>''"
    GWhere = GWhere & " And KH.代表借入連番=0)"
    GWhere = GWhere & " Or (KH.代表借入番号<>''"
    GWhere = GWhere & " And KH.代表借入連番=1)))"
    
    If P8.FCStr(番号) <> "" Then
        GWhere = GWhere & " And KH.借入番号 = '" & P8.FCStr(番号) & "'"
    End If
    If P8.FCStr(Co_1.Text) <> "" Then
        GWhere = GWhere & " And H.保証会社区分名 = '" + P8.FCStr(Co_1.Text) + "'"
    End If
    If P8.FCStr(Co_2.Text) <> "" Then
        GWhere = GWhere & " And G.銀行名 = '" + P8.FCStr(Co_2.Text) + "'"
    End If
    If P8.FCStr(Co_3.Text) <> "" Then
        GWhere = GWhere & " And Y.融資区分名 = '" + P8.FCStr(Co_3.Text) + "'"
    End If
    
    GWhere = " Where (1=1) " & GWhere
    
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " KH.借入番号 As 番号,"
    
    wstr = wstr & " KH.借入番号 As Grd借入番号,"
    wstr = wstr & " Format(KH.融資金額,'#,##0') As Grd融資金額,"
    wstr = wstr & " Format(KH.毎月返済額,'#,##0') As Grd毎月返済額,"
    wstr = wstr & " Format(KH.融資残高,'#,##0') As Grd融資残高,"
    
    wstr = wstr & " H.保証会社区分名 As Grd保証会社名,"
    wstr = wstr & " G.銀行名 As Grd銀行名,"
    wstr = wstr & " Y.融資区分名 As Grd融資区分名,"
    wstr = wstr & " KH.保証会社区分,"
    wstr = wstr & " KH.銀行番号,"
    wstr = wstr & " KH.融資区分,"
    wstr = wstr & " Y.制度融資区分,"
    wstr = wstr & " K.初回返済年月,"
    wstr = wstr & " K.最終返済年月,"
    wstr = wstr & " K.返済単位月数,"
    wstr = wstr & " KH.残据置,"
    wstr = wstr & " KH.有担保フラグ,"
    wstr = wstr & " KH.利率,"
    wstr = wstr & " KH.保証料率,"
    wstr = wstr & " KH.設備フラグ,"
    wstr = wstr & " IIF(KH.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "有担保")) & ",'有','無') As Grd担保区分"
    
    wstr = wstr & " FROM (((DBEA010_借換表 As KH"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON KH.借入番号 = K.借入番号)"
    wstr = wstr & " LEFT JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON KH.銀行番号 = G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA100_保証会社区分 As H"
    wstr = wstr & " ON KH.保証会社区分 = H.保証会社区分)"
    wstr = wstr & " LEFT JOIN DAAA110_融資区分 As Y"
    wstr = wstr & " ON KH.融資区分 = Y.融資区分"

    wstr = wstr + GWhere
    wstr = wstr + " Order By KH.保証会社区分,KH.銀行番号,KH.有担保フラグ,KH.融資区分,KH.借入番号"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("借入番号", "", 1800, "L")
        Call XZMA010_DataGrid_Set("融資金額", "", 1700, "R")
        Call XZMA010_DataGrid_Set("毎月返済額", "", 1700, "R")
        Call XZMA010_DataGrid_Set("融資残高", "", 1700, "R")
        Call XZMA010_DataGrid_Set("保証会社名", "", 2000, "L")
        Call XZMA010_DataGrid_Set("銀行名", "", 2000, "L")
        Call XZMA010_DataGrid_Set("融資区分名", "", 1800, "L")
        Call XZMA010_DataGrid_Set("担保区分", "", 700, "L")
    Call XZMA010_DataGrid_Action(DataGrid1)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_借入金_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh_借入金() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

