VERSION 5.00
Begin VB.Form frm_R残高一覧表 
   Caption         =   "残高一覧表　出力"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   Icon            =   "frm_R残高一覧表.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   495
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
      Left            =   8160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton 実行 
      Caption         =   "実行（F11)"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.ComboBox 金融リストラ番号 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   3
      Text            =   "HH年MM月DD日"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox 実行日 
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   2
      Text            =   "HH年MM月DD日"
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   495
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
      Caption         =   "残高一覧表　出力"
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
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "CSV出力"
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
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "集計"
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
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label45 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "銀行名"
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
      TabIndex        =   13
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "総合計を表示"
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
      TabIndex        =   12
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "～"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "返済年月日To"
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
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "返済年月日From"
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
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R残高一覧表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
