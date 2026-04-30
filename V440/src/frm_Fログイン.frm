VERSION 5.00
Begin VB.Form frm_Fログイン 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ログイン"
   ClientHeight    =   2190
   ClientLeft      =   7965
   ClientTop       =   6645
   ClientWidth     =   4290
   Icon            =   "frm_Fログイン.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4290
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton 閉じる 
      Caption         =   "閉じる"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton ログイン 
      Caption         =   "ログイン"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox PASSWORD 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "abcdefgh"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox USERID 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "ancdefgh"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "PASSWORD"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "ID"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Fログイン"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "ログイン画面"

Dim wRs As ADODB.Recordset
Dim wstr As String
Dim wslog As String
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    ' =========================================
    '                 初期設定
    ' =========================================
    Me.USERID.Text = ""
    Me.PASSWORD.Text = ""
    
    
End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    DoEvents
    
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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
' USERID_GotFocus
'------------------------------------------------
Private Sub USERID_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' ログイン_Click
'------------------------------------------------
Private Sub ログイン_Click()
'
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    
    Dim FLG_USR As Boolean, FLG_PAS As Boolean
    
    Dim wi01 As Integer
    Dim ws01 As String, ws02 As String
    
    wslog = "ログイン"
'
    FLG_USR = False
    FLG_PAS = False
    
    GUserID = ""
    GUserKen = 0
'
    '----------< GCurDir GTemp Open >-----------------------------------------------
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)
    
    ws01 = P8.FCStr(USERID)
    ws02 = P8.FCStr(PASSWORD)
    
    wstr = ""
    wstr = wstr & "Select *"
    wstr = wstr & " From DBUA001_ユーザーマスタ"
    wstr = wstr & " Where isnull(停止日)"
    wstr = wstr & " Order by USERID"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    Do Until wRs3.EOF
        If wRs3("USERID") = ws01 Then
            FLG_USR = True
            
            If wRs3("PASSWORD") = ws02 Then
                FLG_PAS = True
            End If
            
            wi01 = P8.FCDbl(wRs3("権限"))
            
            Exit Do
            
        End If
        wRs3.MoveNext
    Loop
    wRs3.Close
    Set wRs3 = Nothing
    
    wDb.Close
    Set wDb = Nothing
'
    ' =========================================
    '               メッセージ
    ' =========================================
    If FLG_USR = False Then
        GRet = MsgBox("IDを確認してください", vbOKOnly + vbExclamation)
        Call CEkey.SetFs(USERID, True)
        Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, "ログイン失敗 ユーザーID=" & USERID)
        Exit Sub
    End If
    
    If FLG_PAS = False Then
        GRet = MsgBox("PASSWORDを確認してください", vbOKOnly + vbExclamation)
        Call CEkey.SetFs(PASSWORD, True)
        Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, "ログイン失敗 ユーザーID=" & USERID)
        Exit Sub
    End If
'
    GUserID = ws01
    GUserKen = wi01
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "ユーザーID=" & P8.FCStr(USERID.Text)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
'

    Call 実行_Click
'
End Sub

'------------------------------------------------
' 実行_Click
'------------------------------------------------
Private Sub 実行_Click()
'
    Unload Me
    
    frm_Parent.Show
'
End Sub


'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
    End
End Sub

