VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_I借入金登録_銀行 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入金登録 銀行登録"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frm_I借入金登録_銀行.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
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
      Left            =   4560
      TabIndex        =   9
      Top             =   4800
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
      Left            =   6480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton B_SET 
      Caption         =   "SET"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   10320
      Width           =   495
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
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   8175
      Begin 借換たろう.ZU020_ComboBox 金利計算 
         Height          =   300
         Left            =   1920
         TabIndex        =   8
         Top             =   3360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 利息控除 
         Height          =   300
         Left            =   1920
         TabIndex        =   7
         Top             =   3000
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   529
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 利息支払 
         Height          =   300
         Left            =   1920
         TabIndex        =   6
         Top             =   2640
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 利息日数 
         Height          =   300
         Left            =   1920
         TabIndex        =   5
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 利息区分 
         Height          =   300
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 営業日 
         Height          =   300
         Left            =   1920
         TabIndex        =   3
         Top             =   1560
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 支払日 
         Height          =   300
         Left            =   1920
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 銀行 
         Height          =   300
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   529
         Enabled         =   0   'False
         Enabled         =   0   'False
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   2000
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin VB.Label L_番号 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label L_番号1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "借入番号"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "金利計算日数"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "利息控除区分"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "利息支払方法"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "利息計算日数"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
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
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "支払日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "営業日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
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
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9120
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   120
      Top             =   9960
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
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "銀行詳細　登録"
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
   Begin VB.Label メッセージ 
      BackColor       =   &H00C0C000&
      Caption         =   "メッセージ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   8640
      Width           =   15015
   End
End
Attribute VB_Name = "frm_I借入金登録_銀行"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "銀行詳細登録"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim wFname As String, wsTbl As String, wsTblTR As String
Dim wsBango As String
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
'
    wFname = GStr
    'ZU050_Button1.Caption = wFname & Space(1) & "登録"
    
    wsBango = GStr_2
    
    L_番号.Caption = wsBango
    
    Select Case wFname
    Case "借入金登録"
        L_番号1.Caption = " 借入番号"
        
        wsTbl = "DBDA010_借入金"
        wsTblTR = "DBDA010_借入金明細TR"
    Case "貸付登録"
        L_番号1.Caption = " 貸付番号"
    
        wsTbl = "DBDA010_貸付金"
        wsTblTR = "DBDA010_貸付金明細TR"
    End Select
    
    GStr = "": GStr_1 = "": GStr_2 = ""
'
    ' =========================================
    '             コンボボックス
    ' =========================================
    With 銀行
        .P8_Db = GDb
        
        wstr = "Select * From DAAA040_銀行マスタ "
        wstr = wstr + " Where 取消フラグ = 0"
        
        If GSys.Sit = True Then
            For j = 2 To UBound(G独算)
                If G基本情報.支店コード = G独算(j).支店コード Then
                    wstr = wstr + " AND 銀行番号 = 'SS'"
                    Exit For
                End If
            Next j
        End If
        
        wstr = wstr + " Order By 銀行番号 "
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 10
        .P8_ListBoxMax = 500
        .P8_KeyName = "銀行番号"
        .P8_ItemName = "銀行名"
    End With
    銀行.CreateCombo
'
    With 支払日
        .P8_Db = GDb
        
        wstr = "Select * From DAAB020_支払区分マスタ "
        wstr = wstr + " Order By 支払日"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 2
        .P8_ListBoxMax = 500
        .P8_KeyName = "支払日"
        .P8_ItemName = "支払区分名"
    End With
    支払日.CreateCombo
'
    With 営業日
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("営業日", "翌営業日"), "翌営業日")
        Call .AddItem(XMXA020_区分("営業日", "前営業日"), "前営業日")
    End With
    営業日.CreateCombo
'
    With 利息区分
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("利息区分", "利息先払"), "利息先払")
        Call .AddItem(XMXA020_区分("利息区分", "利息後払"), "利息後払")
    End With
    利息区分.CreateCombo
'
    With 利息日数
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("利息日数", "営業日数"), "営業日数")
        Call .AddItem(XMXA020_区分("利息日数", "固定日数"), "固定日数")
    End With
    利息日数.CreateCombo
'
    With 利息支払
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("利息支払", "毎月"), "毎月")
        Call .AddItem(XMXA020_区分("利息支払", "一括"), "一括")
    End With
    利息支払.CreateCombo
'
    With 利息控除
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("利息控除", "控除無し"), "控除無し")
        Call .AddItem(XMXA020_区分("利息控除", "実行日控除"), "実行日控除")
        Call .AddItem(XMXA020_区分("利息控除", "最終返済日控除"), "最終返済日控除")
        Call .AddItem(XMXA020_区分("利息控除", "実行日及び最終返済日控除"), "実行日及び最終返済日控除")
        Call .AddItem(XMXA020_区分("利息控除", "中間利払最終日控除"), "中間利払最終日控除")
    End With
    利息控除.CreateCombo
'
    With 金利計算
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("金利計算", "365日"), "365日")
        Call .AddItem(XMXA020_区分("金利計算", "360日"), "360日")
    End With
    金利計算.CreateCombo
'
    ' =========================================
    '                 初期設定
    ' =========================================
    Call 登録後初期セット
    メッセージ = ""
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
    メッセージ = ""
'
End Sub

'------------------------------------------------
' B_SET_Click
'------------------------------------------------
Private Sub B_SET_Click()
    Call 銀行項目_セット
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット() As Boolean
'
    Dim j As Integer
    Dim ws01 As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
'
    ' =========================================
    '                画面クリア
    ' =========================================
    銀行.Text = ""
    支払日.Text = ""
    営業日.Text = ""
    利息区分.Text = ""
    利息日数.Text = ""
    利息支払.Text = ""
    利息控除.Text = ""
    金利計算.Text = ""
    
    ' =========================================
    '            借入金マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & " K.借入番号,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " K.支払日,"
    wstr = wstr & " K.営業日区分,"
    wstr = wstr & " K.利息区分,"
    wstr = wstr & " K.利息計算日数区分,"
    wstr = wstr & " K.利息支払方法,"
    wstr = wstr & " K.利息控除区分,"
    wstr = wstr & " K.金利計算年間日数,"
    wstr = wstr & " K.保証料率,"
    wstr = wstr & " K.自己資金フラグ,"
    wstr = wstr & " K.保証料分割フラグ"
    wstr = wstr & " From " & wsTbl & " As K"
    wstr = wstr + " Where K.借入番号 = '" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
    
        画面セット = True
        Call CEkey.SetFs(銀行, True)
        
        銀行.Text = P8.FCStr(wRs("銀行番号"))
        
        支払日.Text = P8.FCStr(wRs("支払日")) 'V180
        営業日.Text = P8.FCStr(wRs("営業日区分")) 'V180
        利息区分.Text = P8.FCStr(wRs("利息区分")) 'V180
        利息日数.Text = P8.FCStr(wRs("利息計算日数区分")) 'V180
        利息支払.Text = P8.FCStr(wRs("利息支払方法")) 'V180
        利息控除.Text = P8.FCStr(wRs("利息控除区分")) 'V180
        金利計算.Text = P8.FCStr(wRs("金利計算年間日数")) 'V180
        
       End If
    wRs.Close
    Set wRs = Nothing
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
' 銀行項目_セット
'------------------------------------------------
Private Sub 銀行項目_セット()
'
    ' =========================================
    '            銀行マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA040_銀行マスタ"
    wstr = wstr + " Where 銀行番号 = '" & P8.FCStr(銀行.Text) & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
    
        支払日.Text = P8.FCStr(wRs("支払日"))
        営業日.Text = P8.FCStr(wRs("営業日区分"))
        利息区分.Text = P8.FCStr(wRs("利息区分"))
        利息日数.Text = P8.FCStr(wRs("利息計算日数区分"))
        利息支払.Text = P8.FCStr(wRs("利息支払方法"))
        利息控除.Text = P8.FCStr(wRs("利息控除区分"))
        金利計算.Text = P8.FCStr(wRs("金利計算年間日数"))
        
    End If
    wRs.Close
    Set wRs = Nothing
'
End Sub

'------------------------------------------------
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Call 画面セット
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
    
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
    
'    Unload frm_I借入金登録
'    frm_I借入金登録.Show
'
End Sub

'------------------------------------------------
' 登録_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim w借入データ As MAA910_借入金
    
    Dim j As Integer
    Dim wv01 As Variant, wv02 As Variant
    Dim wvJikou2 As Variant
    Dim ws01 As String
    Dim FLG_AUTO As Boolean
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

    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    If 銀行.P8_Name = "" Then
        MsgBox "コードが違います"
        'SSTab1.Tab = 0
        Call CEkey.SetFs(銀行, True)
        Exit Sub
    End If
 
     If Not IsNumeric(支払日.Text) Or 支払日.Text = "" Then
        MsgBox "支払日を選択してください"
        Call CEkey.SetFs(支払日, True)
        Exit Sub
    End If

    If 営業日.Text = "" Or (営業日.Text < "0" Or 営業日.Text > "1") Then
        MsgBox "営業日を選択してください"
        Call CEkey.SetFs(営業日, True)
        Exit Sub
    End If

    If 利息区分.Text = "" Or (利息区分.Text < "1" Or 利息区分.Text > "2") Then
        MsgBox "利息区分を選択してください"
        Call CEkey.SetFs(利息区分, True)
        Exit Sub
    End If

    If 利息日数.Text = "" Or (利息日数.Text < "0" Or 利息日数.Text > "1") Then
        MsgBox "利息計算日数を選択してください"
        Call CEkey.SetFs(利息日数, True)
        Exit Sub
    End If

    If 利息支払.Text = "" Or (利息支払.Text < "0" Or 利息支払.Text > "1") Then
        MsgBox "利息支払方法を選択してください"
        Call CEkey.SetFs(利息支払, True)
        Exit Sub
    End If

    If 利息控除.Text = "" Or (利息控除.Text < "0" Or 利息控除.Text > "4") Then
        MsgBox "利息控除区分を選択してください"
        Call CEkey.SetFs(利息控除, True)
        Exit Sub
    End If
    
    If 金利計算.Text = "" Or (金利計算.Text < "0" Or 金利計算.Text > "1") Then
        MsgBox "金利計算年間日数を選択してください"
        Call CEkey.SetFs(金利計算, True)
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    FLG_AUTO = False
    w借入データ.借入番号 = wsBango
    
    w借入データ.実行日 = Null
    w借入データ.初回返済実行日 = Null
    w借入データ.最終返済実行日 = Null
    
    w借入データ.銀行番号 = ""
    w借入データ.支払日 = 31
    w借入データ.営業日区分 = 0
    w借入データ.利息区分 = "1"
    w借入データ.利息計算日数区分 = 0
    w借入データ.利息支払方法 = 0
    w借入データ.利息控除区分 = 0
    w借入データ.金利計算年間日数 = 0
'
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
    
        wRs("銀行番号") = P8.FCStr(銀行.Text)
        wRs("支払日") = P8.FCDbl(支払日.Text)
        wRs("営業日区分") = P8.FCDbl(営業日.Text)
        wRs("利息区分") = P8.FCStr(利息区分.Text)
        wRs("利息計算日数区分") = P8.FCDbl(利息日数.Text)
        wRs("利息支払方法") = P8.FCDbl(利息支払.Text)
        wRs("利息控除区分") = P8.FCDbl(利息控除.Text)
        wRs("金利計算年間日数") = P8.FCDbl(金利計算.Text)
        
        wvJikou2 = MBD010_実行日支払年月算出(P8.FCDate(wRs("実行日")), wRs("営業日区分"), wRs("支払日"))

        wv01 = MXA030_金利初回年月(wRs("利息区分"), wRs("利息支払方法"), wRs("支払日"), wRs("営業日区分"), wvJikou2, P8.FCDate(wRs("初回返済年月")), wRs("返済単位月数"))
        wRs("金利初回年月") = P8.FCDate(wv01)
        
        If P8.FCDbl(wRs("手入力区分")) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
            wRs("金利初回年月") = Null
        End If
        
        wv02 = wRs("金利初回年月")
        
        If P8.FCDbl(wRs("手入力区分")) = P8.FCDbl(XMXA020_区分("登録方法", "入力登録")) _
        And P8.FCDbl(wRs("日割計算区分")) = P8.FCDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
            FLG_AUTO = True
        End If
            
        wRs.Update
        
        w借入データ.実行日 = wRs("実行日")
        w借入データ.初回返済実行日 = wRs("初回返済実行日")
        w借入データ.最終返済実行日 = wRs("最終返済実行日")
        
        w借入データ.銀行番号 = wRs("銀行番号")
        w借入データ.支払日 = wRs("支払日")
        w借入データ.営業日区分 = wRs("営業日区分")
        w借入データ.利息区分 = wRs("利息区分")
        w借入データ.利息計算日数区分 = wRs("利息計算日数区分")
        w借入データ.利息支払方法 = wRs("利息支払方法")
        w借入データ.利息控除区分 = wRs("利息控除区分")
        w借入データ.金利計算年間日数 = wRs("金利計算年間日数")
    
    End If
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               明細TR再計算
    ' =========================================
     If FLG_AUTO = True Then
        Call MBD010_借入金入力明細作成_日割日数再計算(w借入データ, wsTblTR)
     End If
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号=" & P8.FCStr(L_番号.Caption) & ","
    GLogStr = GLogStr & "銀行番号=" & P8.FCStr(銀行.Text) & ","
    GLogStr = GLogStr & "支払日=" & P8.FCDbl(支払日.Text) & ","
    GLogStr = GLogStr & "営業日区分=" & P8.FCDbl(営業日.Text) & ","
    GLogStr = GLogStr & "利息区分=" & P8.FCStr(利息区分.Text) & ","
    GLogStr = GLogStr & "利息計算日数区分=" & P8.FCDbl(利息日数.Text) & ","
    GLogStr = GLogStr & "利息支払方法=" & P8.FCDbl(利息支払.Text) & ","
    GLogStr = GLogStr & "利息控除区分=" & P8.FCDbl(利息控除.Text) & ","
    GLogStr = GLogStr & "金利計算年間日数=" & P8.FCDbl(金利計算.Text) & ","
    GLogStr = GLogStr & "金利初回年月=" & wv02
    Call MXA030_LOG_WRITE("借入金銀行詳細登録", "更新", GLogStr)
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット
'    Call 登録後初期セット
    Call CEkey.SetFs(銀行, False)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました。", vbInformation
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
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
    
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
'    Unload frm_I借入金登録
'    frm_I借入金登録.Show
'
End Sub

