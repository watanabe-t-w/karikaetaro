VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_I借入金登録 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入金登録"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   Icon            =   "frm_I借入金登録.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   75
      Top             =   2760
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "借入内容"
      TabPicture(0)   =   "frm_I借入金登録.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label27"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label29"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label32"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "L_返済単位月数"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "L_金利初回年月"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label57"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label67"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label31"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label26"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label24"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "L_返済方法"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label13"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label15"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label16"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label17"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label19"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label22"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label58"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label59"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label65"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label66"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label25"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "L_最終返済額"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "L_初回返済額"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "L_毎月返済額"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "L_解約年月日"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "支払日"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "基準金利"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "銀行"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "金利種別"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "金利条件"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "最終返済年月"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "初回返済年月"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "実行日"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "利率"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "返済単位月数"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "初回返済実行日"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "最終返済実行日"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "C_金利初回年月"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "設備区分"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "資金用途"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "長短区分"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "担保区分"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "担保名"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "解約年月日"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "融資金額"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "毎月返済額"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "初回返済額"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "最終返済額"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).ControlCount=   60
      TabCaption(1)   =   "シミュレーション"
      TabPicture(1)   =   "frm_I借入金登録.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label30"
      Tab(1).Control(1)=   "Label46"
      Tab(1).Control(2)=   "Label45"
      Tab(1).Control(3)=   "LblSM区分"
      Tab(1).Control(4)=   "L_金融解約日"
      Tab(1).Control(5)=   "金利グループ区分"
      Tab(1).Control(6)=   "金融解約日"
      Tab(1).Control(7)=   "金融リストラ番号"
      Tab(1).Control(8)=   "SM区分"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CheckBox SM区分 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72720
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.ComboBox 金融リストラ番号 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   1  'ｵﾝ
         Left            =   -72720
         TabIndex        =   30
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox 金融解約日 
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
         Left            =   -72720
         TabIndex        =   32
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox 最終返済額 
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
         Left            =   9840
         MaxLength       =   16
         TabIndex        =   23
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox 初回返済額 
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
         Left            =   9840
         MaxLength       =   16
         TabIndex        =   22
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox 毎月返済額 
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
         Left            =   9840
         MaxLength       =   16
         TabIndex        =   21
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox 融資金額 
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
         Left            =   9840
         MaxLength       =   16
         TabIndex        =   20
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox 解約年月日 
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
         Left            =   9840
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox 担保名 
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
         IMEMode         =   4  '全角ひらがな
         Left            =   9840
         TabIndex        =   27
         Top             =   3960
         Width           =   2415
      End
      Begin VB.ComboBox 担保区分 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9840
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   26
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox 長短区分 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9840
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   25
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox 資金用途 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   1  'ｵﾝ
         Left            =   7680
         TabIndex        =   29
         Top             =   5040
         Width           =   4575
      End
      Begin VB.ComboBox 設備区分 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9840
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   28
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox C_金利初回年月 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   12
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox 最終返済実行日 
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
         Left            =   2280
         TabIndex        =   14
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox 初回返済実行日 
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
         Left            =   2280
         TabIndex        =   11
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox 返済単位月数 
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
         IMEMode         =   2  'ｵﾌ
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox 利率 
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
         Left            =   2280
         MaxLength       =   7
         TabIndex        =   17
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox 実行日 
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
         Left            =   2280
         TabIndex        =   9
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox 初回返済年月 
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
         Left            =   2280
         TabIndex        =   10
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox 最終返済年月 
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
         Left            =   2280
         TabIndex        =   13
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox 金利条件 
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
         Left            =   2280
         TabIndex        =   18
         Top             =   5040
         Width           =   4575
      End
      Begin VB.ComboBox 金利種別 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   15
         Top             =   3960
         Width           =   2535
      End
      Begin 借換たろう.ZU020_ComboBox 銀行 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   2000
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
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 基準金利 
         Height          =   315
         Left            =   2280
         TabIndex        =   16
         Top             =   4320
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
      Begin 借換たろう.ZU020_ComboBox 支払日 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
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
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin 借換たろう.ZU020_ComboBox 金利グループ区分 
         Height          =   315
         Left            =   -72720
         TabIndex        =   33
         Top             =   1560
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   975
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
      Begin VB.Label L_金融解約日 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72720
         TabIndex        =   116
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label LblSM区分 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  '実線
         Caption         =   " SM区分"
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
         Left            =   -74880
         TabIndex        =   115
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label45 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  '実線
         Caption         =   " 借入ｼﾐｭﾚｰｼｮﾝ番号"
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
         Left            =   -74880
         TabIndex        =   114
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label46 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  '実線
         Caption         =   " ｼﾐｭﾚｰｼｮﾝ解約日"
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
         Left            =   -74880
         TabIndex        =   113
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label30 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  '実線
         Caption         =   "金利ｼﾐｭﾚｰｼｮﾝGP"
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
         Left            =   -74880
         TabIndex        =   112
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label L_解約年月日 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9840
         TabIndex        =   111
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label L_毎月返済額 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9840
         TabIndex        =   110
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label L_初回返済額 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9840
         TabIndex        =   109
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label L_最終返済額 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9840
         TabIndex        =   108
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label25 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 返済方法"
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
         Left            =   7680
         TabIndex        =   107
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label66 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         TabIndex        =   106
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label65 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         TabIndex        =   105
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label59 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         TabIndex        =   104
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label58 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         TabIndex        =   103
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label22 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "担保区分"
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
         Left            =   7680
         TabIndex        =   102
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label19 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 最終返済額"
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
         Left            =   7680
         TabIndex        =   101
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label17 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 初回返済額"
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
         Left            =   7680
         TabIndex        =   100
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label16 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 毎回返済元金"
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
         Left            =   7680
         TabIndex        =   99
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 融資金額"
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
         Left            =   7680
         TabIndex        =   98
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 解約年月日"
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
         Left            =   7680
         TabIndex        =   97
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label L_返済方法 
         BorderStyle     =   1  '実線
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
         Left            =   9840
         TabIndex        =   19
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label24 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 担保内容"
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
         Left            =   7680
         TabIndex        =   96
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label26 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 資金用途"
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
         Left            =   7680
         TabIndex        =   95
         Top             =   4740
         Width           =   4575
      End
      Begin VB.Label Label31 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "長短区分"
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
         Left            =   7680
         TabIndex        =   94
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label67 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 設備"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         TabIndex        =   93
         Top             =   4320
         Width           =   2175
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
         Left            =   4560
         TabIndex        =   92
         Top             =   4800
         Width           =   375
      End
      Begin VB.Label L_金利初回年月 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   91
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label L_返済単位月数 
         BorderStyle     =   1  '実線
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   90
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 最終返済年月日"
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
         TabIndex        =   89
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 初回返済年月日"
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
         TabIndex        =   88
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label14 
         Alignment       =   1  '右揃え
         Caption         =   "ｶ月 "
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   87
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 返済単位月数"
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
         TabIndex        =   86
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 利率"
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
         TabIndex        =   85
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 銀行名"
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
         TabIndex        =   84
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 実行日"
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
         TabIndex        =   83
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 初回返済年月  "
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
         TabIndex        =   82
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 最終返済年月  "
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
         TabIndex        =   81
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label32 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 金利初回年月  "
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
         TabIndex        =   80
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label29 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 支払日"
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
         TabIndex        =   79
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 金利備考"
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
         TabIndex        =   78
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label18 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 金利種別"
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
         TabIndex        =   77
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label27 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
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
         TabIndex        =   76
         Top             =   4320
         Width           =   2175
      End
   End
   Begin VB.CommandButton 銀行詳細 
      Caption         =   "銀行詳細登録"
      Height          =   495
      Left            =   120
      TabIndex        =   74
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton 金利変更 
      Caption         =   "金利変更"
      Height          =   495
      Left            =   1800
      TabIndex        =   73
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton 明細入力 
      Caption         =   "明細入力"
      Height          =   495
      Left            =   3480
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton 内入入力 
      Caption         =   "内入入力"
      Height          =   495
      Left            =   5160
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton 削除 
      Caption         =   "削除"
      Height          =   495
      Left            =   6840
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton 登録 
      Caption         =   "登録"
      Height          =   495
      Left            =   8520
      TabIndex        =   34
      Top             =   9000
      Width           =   1695
   End
   Begin VB.CommandButton CSV出力 
      Caption         =   "CSV出力"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton CSV取込 
      Caption         =   "CSV取込"
      Height          =   495
      Left            =   5160
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton 明細書表示 
      Caption         =   "明細書表示"
      Height          =   495
      Left            =   1800
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Frame fra借入金データ 
      Caption         =   "借入金データ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   62
      Top             =   720
      Width           =   6855
      Begin VB.CommandButton Copy 
         Caption         =   "Copy"
         Height          =   330
         Left            =   6000
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton 検索 
         Caption         =   "..."
         Height          =   330
         Left            =   5520
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox 借入内容 
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
         IMEMode         =   4  '全角ひらがな
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox 借入番号 
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
         MaxLength       =   30
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin 借換たろう.ZU020_ComboBox 借入金種別 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
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
      Begin 借換たろう.ZU020_ComboBox 部門 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VB.Label Label21 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 部門"
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
         Left            =   240
         TabIndex        =   117
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label28 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 借入金種別"
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
         Left            =   240
         TabIndex        =   66
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label L_借入番号 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 借入番号"
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
         Left            =   240
         TabIndex        =   64
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label L_借入内容 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 借入内容"
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
         Left            =   240
         TabIndex        =   63
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "登録方法"
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
      Left            =   7080
      TabIndex        =   56
      Top             =   720
      Width           =   5655
      Begin 借換たろう.ZU020_ComboBox 登録方法 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
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
      Begin 借換たろう.ZU020_ComboBox 日割計算区分 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.Label L_登録方法 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '実線
         Caption         =   "未完成"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4680
         TabIndex        =   61
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label73 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   " 登録方法"
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
         TabIndex        =   60
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label23 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   " 日割計算区分"
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
         TabIndex        =   59
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label L_日割計算区分 
         BorderStyle     =   1  '実線
         Caption         =   " 自動計算"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   58
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label L_登録方法2 
         BorderStyle     =   1  '実線
         Caption         =   " 登録方法"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   57
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.ComboBox 借入計画番号 
      Height          =   300
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2400
      TabIndex        =   42
      Text            =   "借入計画番号"
      Top             =   12600
      Visible         =   0   'False
      Width           =   3135
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
      Left            =   8520
      TabIndex        =   41
      Top             =   11640
      Width           =   495
   End
   Begin VB.ComboBox プロジェクト名 
      Height          =   300
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2400
      TabIndex        =   40
      Top             =   12360
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.CheckBox 保証料分割 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10560
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   13440
      Width           =   255
   End
   Begin VB.CheckBox 自己資金 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10560
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   13800
      Width           =   255
   End
   Begin VB.TextBox 保証料率 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   10560
      MaxLength       =   7
      TabIndex        =   37
      Top             =   13080
      Width           =   2295
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
      Left            =   11040
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   120
      Top             =   11760
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
   Begin 借換たろう.ZU070_Label 新規変更 
      Height          =   375
      Left            =   3720
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin 借換たろう.ZU020_ComboBox 融資区分 
      Height          =   345
      Left            =   10560
      TabIndex        =   44
      Top             =   12720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   609
      ForeColor       =   -2147483640
      ForeColor       =   -2147483640
      IMEMode         =   3
      TextWidth       =   615
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      P8_ListBoxMax   =   0
      P8_KeySort      =   0   'False
   End
   Begin 借換たろう.ZU020_ComboBox 保証会社区分 
      Height          =   345
      Left            =   10560
      TabIndex        =   45
      Top             =   12360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   609
      ForeColor       =   -2147483640
      ForeColor       =   -2147483640
      IMEMode         =   3
      TextWidth       =   615
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      P8_ListBoxMax   =   0
      P8_KeySort      =   0   'False
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   375
      Left            =   120
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "借入金登録"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label L_借入計画番号 
      BackColor       =   &H00D6DBBD&
      Caption         =   " 借入計画番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   55
      Top             =   12000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00D6DBBD&
      Caption         =   " ﾌﾟﾛｼﾞｪｸﾄ名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   54
      Top             =   12360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label L_保証料率 
      BackColor       =   &H00D6DBBD&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10560
      TabIndex        =   53
      Top             =   13080
      Width           =   2295
   End
   Begin VB.Label Label60 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00D6DBBD&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12960
      TabIndex        =   52
      Top             =   13080
      Width           =   375
   End
   Begin VB.Label Label41 
      BackColor       =   &H00D6DBBD&
      Caption         =   " 保証料分割"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   51
      Top             =   13440
      Width           =   2175
   End
   Begin VB.Label Label40 
      BackColor       =   &H00D6DBBD&
      Caption         =   " 自己資金"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   50
      Top             =   13800
      Width           =   2175
   End
   Begin VB.Label Label39 
      BackColor       =   &H00D6DBBD&
      Caption         =   " 保証料率"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   49
      Top             =   13080
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00D6DBBD&
      Caption         =   " 融資区分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   48
      Top             =   12720
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00D6DBBD&
      Caption         =   " 保証会社区分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   47
      Top             =   12360
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFFF&
      Caption         =   " 保証会社を利用する場合"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   46
      Top             =   12600
      Width           =   6615
   End
End
Attribute VB_Name = "frm_I借入金登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_I借入金登録"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim wTmp借入番号 As String

Dim wSyusi As Double, wSshokai As Double, wSsaishu As Double, wSmaituki As Double
Dim wi単位 As Integer
Dim wv実行日 As Variant, wv初回返済年月 As Variant
Dim wi利息支払 As Integer, wi支払日 As Integer, wi営業日 As Integer
Dim wi利息日数 As Integer, wi利息控除 As Integer, wi金利計算 As Integer
Dim w初回変更年月 As Variant, w最終変更年月 As Variant
Dim FLG_New As String, FLG_GSET As Boolean
Dim FLG_Src As String

Dim wi登録方法 As Integer, wi登録方法_変更 As Integer
Dim wiRet As Integer

Dim wFname As String, wsTbl As String, wsTbl2 As String
Dim ws利息区分 As String, wiTblNo As Integer
Dim wv最終返済年月日 As Variant
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
'    Me.Caption = GFcap
    Me.Left = G_LEFT
    Me.Top = G_TOP
    Me.SSTab1.Tab = 0
'
    GStr = "借入金登録"
    
    wFname = GStr
    ZU050_Button1.Caption = wFname & Space(1) & "登録"
    
    L_借入番号.Caption = " 借入番号"
    L_借入内容.Caption = " 借入内容"
    L_借入計画番号.Caption = " 借入計画番号"
    
    wsTbl = "DBDA010_借入金"
    wsTbl2 = "DBDA010_借入金明細TR"
    wiTblNo = P8.FCDbl(XMXA020_区分("借入貸付", "借入"))
    Select Case wFname
    Case "借入金登録"
        L_借入番号.Caption = " 借入番号"
        L_借入内容.Caption = " 借入内容"
        L_借入計画番号.Caption = " 借入計画番号"
        
        wsTbl = "DBDA010_借入金"
        wsTbl2 = "DBDA010_借入金明細TR"
        wiTblNo = P8.FCDbl(XMXA020_区分("借入貸付", "借入"))
    Case "貸付登録"
        L_借入番号.Caption = " 貸付番号"
        L_借入内容.Caption = " 貸付内容"
        L_借入計画番号.Caption = " 貸付計画番号"
    
        wsTbl = "DBDA010_貸付金"
        wsTbl2 = "DBDA010_貸付金明細TR"
        wiTblNo = P8.FCDbl(XMXA020_区分("借入貸付", "貸付"))
    End Select
    GStr = ""
'
    If GProduct <> "金剛石" Then
        L_借入計画番号.Visible = False
        借入計画番号.Visible = False
    End If
'
    登録.Enabled = True
    If GSys.Sit = True Then
'        If (G基本情報.支店コード = G独算(0).支店コード And G基本情報.企業区分 <> "単独企業") Then
        If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結親会社" Then
            登録.Enabled = False
        End If
    End If
'
    ' =========================================
    '             コンボボックス
    ' =========================================
    With 銀行
        .P8_Db = GDb
        
        wstr = "Select * From DAAA040_銀行マスタ"
        wstr = wstr + " Where 取消フラグ = 0"
        
        If GSys.Sit = True Then
            For j = 2 To UBound(G独算)
                If G基本情報.支店コード = G独算(j).支店コード Then
                    wstr = wstr + " AND 銀行番号 = 'SS'"
                    Exit For
                End If
            Next j
        End If
        
        wstr = wstr + " Order By 銀行番号"
        
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
        
        wstr = "Select * From DAAB020_支払区分マスタ"
        wstr = wstr + " Order By 支払日"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 2
        .P8_ListBoxMax = 500
        .P8_KeyName = "支払日"
        .P8_ItemName = "支払区分名"
    End With
    支払日.CreateCombo
'
    With 借入金種別
        .P8_Db = GDb
        
        wstr = "Select * From DAAA116_借入金種別"
        wstr = wstr + " Where 取消フラグ = 0"
        wstr = wstr + " Order By 借入金種別区分"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 2
        .P8_ListBoxMax = 500
        .P8_KeyName = "借入金種別区分"
        .P8_ItemName = "借入金種別名"
    End With
    借入金種別.CreateCombo
'
    With 部門
        .P8_Db = GDb
        
        wstr = "Select * From DAAA200_部門マスタ"
        wstr = wstr + " Where 取消フラグ = 0"
        wstr = wstr + " Order By 部門番号"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 10
        .P8_ListBoxMax = 500
        .P8_KeyName = "部門番号"
        .P8_ItemName = "部門名"
    End With
    部門.CreateCombo
'
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
'
    With 金利グループ区分
        .P8_Db = GDb
        
        wstr = "Select * From DAAA115_金利シミュレーショングループ"
        wstr = wstr + " Where (取消フラグ = 0 or 取消フラグ is null)"
        wstr = wstr + " Order By 金利グループ区分"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 5
        .P8_ListBoxMax = 500
        .P8_KeyName = "金利グループ区分"
        .P8_ItemName = "金利グループ名"
    End With
    金利グループ区分.CreateCombo
'
    With 登録方法
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("登録方法", "標準登録"), "標準登録")
        Call .AddItem(XMXA020_区分("登録方法", "入力登録"), "入力登録")
    End With
    登録方法.CreateCombo
'
    With 日割計算区分
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("日割計算区分", "自動計算"), "自動計算")
        Call .AddItem(XMXA020_区分("日割計算区分", "入力登録"), "入力登録")
    End With
    日割計算区分.CreateCombo
'
    With 長短区分
        .Clear
        
        .AddItem "短期借入金"
        .ItemData(長短区分.NewIndex) = XMXA020_区分("長短区分", "短期借入金")
        .AddItem "長期借入金"
        .ItemData(長短区分.NewIndex) = XMXA020_区分("長短区分", "長期借入金")
    End With
'
    With 金利種別
        .Clear
        
        .AddItem "変動金利"
        .ItemData(金利種別.NewIndex) = XMXA020_区分("金利種別", "変動金利")
        .AddItem "固定金利"
        .ItemData(金利種別.NewIndex) = XMXA020_区分("金利種別", "固定金利")
    End With
'
    With 担保区分
        .Clear
        
        .AddItem "無担保"
        .ItemData(担保区分.NewIndex) = XMXA020_区分("有担フラグ", "無担保")
        .AddItem "有担保"
        .ItemData(担保区分.NewIndex) = XMXA020_区分("有担フラグ", "有担保")
    End With
'
    With 設備区分
        .Clear
        
        .AddItem "運転資金"
        .ItemData(設備区分.NewIndex) = XMXA020_区分("設備区分", "運転資金")
        .AddItem "設備"
        .ItemData(設備区分.NewIndex) = XMXA020_区分("設備区分", "設備")
    End With
'
    '保証協会
    With 保証会社区分
        .P8_Db = GDb
        
        wstr = "Select * From DAAA100_保証会社区分"
        wstr = wstr + " Where 取消フラグ = 0"
        wstr = wstr + " And 代表区分=0"
        wstr = wstr + " Order By 保証会社区分"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 3
        .P8_ListBoxMax = 500
        .P8_KeyName = "保証会社区分"
        .P8_ItemName = "保証会社区分名"
    End With
    保証会社区分.CreateCombo
'
    With 融資区分
        .P8_Db = GDb
        
        wstr = "Select * From DAAA110_融資区分"
        wstr = wstr + " Where 取消フラグ = 0"
        wstr = wstr + " Order By 融資区分"
        
        .P8_SqlString = wstr
        .P8_KeyLeng = 2
        .P8_ListBoxMax = 500
        .P8_KeyName = "融資区分"
        .P8_ItemName = "融資区分名"
    End With
    融資区分.CreateCombo
'
    '----------------------------------------
    '                借入金セット
    '----------------------------------------
    借入計画番号.Clear
    借入計画番号.AddItem ""
    
    wstr = ""
    wstr = wstr + "Select 借入計画番号"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入計画番号 <> '' "
    wstr = wstr + " Group By 借入計画番号"
    wstr = wstr + " Order By 借入計画番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.EOF
            借入計画番号.AddItem (P8.FCStr(wRs("借入計画番号")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    '----------------------------------------
    '            金融リストラ番号セット
    '----------------------------------------
    金融リストラ番号.Clear
    金融リストラ番号.AddItem ""
    
    wstr = ""
    wstr = wstr + "Select 金融リストラ番号"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 金融リストラ番号 <> '' "
    wstr = wstr + " Group By 金融リストラ番号"
    wstr = wstr + " Order By 金融リストラ番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.EOF
            金融リストラ番号.AddItem (P8.FCStr(wRs("金融リストラ番号")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    '----------------------------------------
    '            資金用途セット
    '----------------------------------------
    資金用途.Clear
    資金用途.AddItem ""
    
    wstr = ""
    wstr = wstr + "Select 資金用途"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 資金用途 <> ''"
    wstr = wstr + " Group By 資金用途"
    wstr = wstr + " Order By 資金用途"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.EOF
            資金用途.AddItem (P8.FCStr(wRs("資金用途")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    C_金利初回年月.Clear
'
    '金利種別.Visible = True
    '金利条件.Visible = True
    '利率.Visible = True
    '保証料率.Visible = True
    
    '金利初回年月.Visible = True
    C_金利初回年月.Visible = True
    L_金利初回年月.Visible = False
    返済単位月数.Visible = True
    解約年月日.Visible = True
    金融解約日.Visible = True
    毎月返済額.Visible = True
    初回返済額.Visible = True
    最終返済額.Visible = True
    
    L_金利初回年月.Caption = ""
    L_返済単位月数.Caption = ""
    L_解約年月日.Caption = ""
    L_金融解約日.Caption = ""
    L_返済方法.Caption = ""
    L_毎月返済額.Caption = ""
    L_初回返済額.Caption = ""
    L_最終返済額.Caption = ""
'
    登録方法.Visible = True
    L_登録方法.Visible = False
    L_登録方法2.Caption = ""
    L_登録方法2.Visible = False
    
    明細入力.Caption = "明細入力"
    明細入力.Enabled = True
    
    内入入力.Caption = "内入入力"
    内入入力.Enabled = True
    
    If GSys.Sys = "借入金 Lite" Then
    '借入金 Lite
        登録方法.Text = "0"
        登録方法.Visible = False
        L_登録方法.Visible = False
        L_登録方法2.Caption = " 標準入力"
        L_登録方法2.Visible = True
            
        明細入力.Caption = ""
        明細入力.Enabled = False
        内入入力.Caption = ""
        内入入力.Enabled = False
    End If
'
    ' =========================================
    '                 初期設定
    ' =========================================
'    FLG_Src = False
'    If GStr_1 = "" Then
'        Call 登録後初期セット
'    Else
'        借入番号 = GStr_1
'        Call 画面セット(False)
'
'        FLG_Src = True
'    End If

    GStr = "": GStr_1 = "": GStr_2 = ""
    GStr_3 = ""
'
'    Call 登録後初期セット
'    メッセージ = ""
End Sub

''------------------------------------------------
'' Form_Activate
''------------------------------------------------
'Private Sub Form_Activate()
''
'    DoEvents
'    Call MXA010_検索用データクリア
'    Call CEkey.AllSelect
'
'    検索結果 セット
'    If FLG_Src = True Then
'        Call 画面セット(False)
'    End If
'
'     =========================================
'                     画面セット
'     =========================================
'    FLG_Src = False
'    If GStr <> "" And GStr_1 <> "" Then
'        借入番号 = GStr_1
'        Call 画面セット(False)
'
'        FLG_Src = True
'    End If
'
'    GStr = "": GStr_1 = "": GStr_2 = ""
'    GStr_3 = ""
''
'End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF11 Then
'        Call 登録_Click
'    End If
    
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
    Dim w借入金 As MAA910_借入金
    Dim w融資残高 As Double
    Dim w借入金管理区分 As String
    Dim ws01 As String
    Dim wi01 As Integer
    Dim j As Integer
    
    Dim wvShokai As Variant, wvJikou As Variant, wv01 As Variant
    
    Dim wsMsg As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
'
    wTmp借入番号 = P8.FCStr(借入番号)

    ' =========================================
    '                画面クリア
    ' =========================================
    FLG_GSET = False
    FLG_GSET = True
    
    '金利初回年月.Text = ""
    L_金利初回年月.Caption = ""
    
    '金利初回年月.Visible = True
    C_金利初回年月.Visible = True
    L_金利初回年月.Visible = False
    
    L_登録方法.Visible = False
    日割計算区分.Visible = False
    日割計算区分.Text = CDbl(XMXA020_区分("日割計算区分", "自動計算"))
    
    'L_金利種別.Caption = ""
    'L_金利条件.Caption = ""
    'L_利率.Caption = ""
    'L_保証料率.Caption = ""
    
    L_返済方法.Caption = ""
    
    wi支払日 = 0
    wi営業日 = 0
    ws利息区分 = "0"
    wi利息日数 = 0
    wi利息支払 = 0
    wi利息控除 = 0
    wi金利計算 = 0
    w初回変更年月 = Null
    w最終変更年月 = Null
    wv最終返済年月日 = Null
    
    借入内容 = ""
    借入金種別.Text = "01"
    登録方法.Text = XMXA020_区分("登録方法", "標準登録")
    借入計画番号 = ""
    銀行.Text = ""
    支払日.Text = ""
    返済単位月数 = 1
    
    実行日 = ""
    初回返済年月 = ""
    初回返済実行日 = ""
    最終返済年月 = ""
    最終返済実行日 = ""
    解約年月日 = ""
    
    融資金額 = 0
    毎月返済額 = 0
    初回返済額 = 0
    最終返済額 = 0
    
    金利種別.ListIndex = -1
    基準金利.Text = ""
    利率 = 0
    金利条件 = ""
    
    長短区分.ListIndex = -1
    担保区分.ListIndex = -1
    担保名 = ""
    資金用途 = ""
    設備区分.ListIndex = -1
    
    金融リストラ番号 = ""
    金融解約日 = ""
    金利グループ区分.Text = ""
    
    'LblSM区分.Visible = True
    'SM区分.Visible = True
    SM区分 = 0 'False
    
    '保証会社
    保証料率 = 0
    自己資金 = 0
    保証料分割 = 0
    保証会社区分.Text = ""
    融資区分.Text = ""
    
    '取消 = 0
        
    '入力画面
    明細入力.Enabled = False
    内入入力.Enabled = True
    
    FLG_New = False
    
    '金利種別.Visible = True
    '金利条件.Visible = True
    '利率.Visible = True
    '保証料率.Visible = True
    
    '金利初回年月.Visible = True
    C_金利初回年月.Visible = True
    返済単位月数.Visible = True
    解約年月日.Visible = True
    金融解約日.Visible = True
    毎月返済額.Visible = True
    初回返済額.Visible = True
    最終返済額.Visible = True
    
    L_金利初回年月.Caption = ""
    L_返済単位月数.Caption = ""
    L_解約年月日.Caption = ""
    L_金融解約日.Caption = ""
    L_返済方法.Caption = ""
    L_毎月返済額.Caption = ""
    L_初回返済額.Caption = ""
    L_最終返済額.Caption = ""
    
    wsMsg = ""
    wi登録方法 = 0
    wi登録方法_変更 = 0
    
    新規変更.Caption = ""
    
    ' =========================================
    '            借入金マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr + "Select k.*"
    wstr = wstr + " From " & wsTbl & " As K"
    wstr = wstr + " Where K.借入番号 = '" & 借入番号 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            If 借入番号 <> "" Then
                '
                If GSys.Sit = True Then
'                        If G基本情報.支店コード = G独算(0).支店コード And G基本情報.企業区分 <> "単独企業" Then
                    If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結親会社" Then
                        新規変更.Caption = ""
                        wRs.Close
                        Set wRs = Nothing
                        
                        借入番号 = ""
                        Call CEkey.SetFs(借入番号, True)
                        
                        Exit Function
                    End If
                End If
                '
'                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
'                If GRet = vbNo Then
'                    新規変更.Caption = ""
'                    wRs.Close
'                    Set wRs = Nothing
'
'                    Exit Function
'                End If
                
'                If GSys.Sys = "借入金 お試し版" Then
'                '借入金 お試し版
'                    GRet = 登録件数CHECK
'                    If GRet >= 3 Then
'                        GRet = MsgBox("おそれいりますが、お試し版では3件以上の借入金登録はできません。", vbOKOnly + vbExclamation)
'                        新規変更.Caption = ""
'                        wRs.Close
'                        Set wRs = Nothing
'
'                        借入番号 = ""
'                        Call CEkey.SetFs(借入番号, True)
'
'                        Exit Function
'                    End If
'
'                ElseIf GSys.Sys = "借入金 Lite" Then
'                '借入金 Lite
'                    GRet = 登録件数CHECK
'                    If GRet >= 10 Then
'                        GRet = MsgBox("おそれいりますが、Lite版では10件以上の借入金登録はできません。", vbOKOnly + vbExclamation)
'                        新規変更.Caption = ""
'                        wRs.Close
'                        Set wRs = Nothing
'
'                        借入番号 = ""
'                        Call CEkey.SetFs(借入番号, True)
'
'                        Exit Function
'                    End If
'                End If
                
                新規変更.Caption = "新規登録"
'                Call CEkey.SetFs(借入内容, True)
    
                L_返済方法.Caption = "元金均等返済"
                
                登録方法.Text = XMXA020_区分("登録方法", "標準登録")
                
                FLG_New = True
            
            End If
        Else
            画面セット = True
'            Call CEkey.SetFs(借入内容, True)
            新規変更.Caption = "変更"
                        
            L_返済方法.Caption = "元金均等返済"
            
            借入内容 = P8.FCStr(wRs("借入内容"))
            
            'V180
            wi01 = P8.FCDbl(wRs("手入力区分"))
            日割計算区分.Text = CDbl(XMXA020_区分("日割計算区分", "自動計算"))
            Select Case wi01
            Case P8.FCDbl(XMXA020_区分("登録方法", "標準登録"))
                登録方法.Text = wi01
            Case P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
                登録方法.Text = wi01
                日割計算区分.Text = P8.FCDbl(wRs("日割計算区分"))
            Case Else
                登録方法.Text = XMXA020_区分("登録方法", "入力登録")
                L_登録方法.Visible = True
                日割計算区分.Text = P8.FCDbl(wRs("日割計算区分"))
            End Select
            
            wi登録方法 = wi01
            wi登録方法_変更 = wi01
                    
            借入金種別.Text = P8.FCStr(wRs("借入金種別区分"))
            借入計画番号 = P8.FCStr(wRs("借入計画番号"))
            部門.Text = P8.FCStr(wRs("プロジェクト番号"))
            銀行.Text = P8.FCStr(wRs("銀行番号"))
            支払日.Text = P8.FCDbl(wRs("支払日"))
            返済単位月数 = P8.FFormat(wRs("返済単位月数"), "#,##0")
            
            wi支払日 = 支払日.Text
            wi営業日 = P8.FCDbl(wRs("営業日区分"))
            ws利息区分 = P8.FCStr(wRs("利息区分"))
            wi利息日数 = P8.FCDbl(wRs("利息計算日数区分"))
            wi利息支払 = P8.FCDbl(wRs("利息支払方法"))
            wi利息控除 = P8.FCDbl(wRs("利息控除区分"))
            wi金利計算 = P8.FCDbl(wRs("金利計算年間日数"))
        
            実行日 = Format(P8.FCStr(wRs("実行日")), Gfmt年月日)
            初回返済年月 = Format(P8.FCStr(wRs("初回返済年月")), Gfmt年月)
            初回返済実行日 = Format(P8.FCStr(wRs("初回返済実行日")), Gfmt年月日)
            L_金利初回年月.Caption = Format(P8.FCStr(wRs("金利初回年月")), Gfmt年月) 'V180
            最終返済年月 = Format(P8.FCStr(wRs("最終返済年月")), Gfmt年月)
            最終返済実行日 = Format(P8.FCStr(wRs("最終返済実行日")), Gfmt年月日)
            解約年月日 = Format(P8.FCStr(wRs("解約実行日")), Gfmt年月日)
            wv最終返済年月日 = wRs("最終返済実行日")
            
            融資金額 = P8.FFormat(wRs("融資金額"), "#,##0")
            毎月返済額 = P8.FFormat(wRs("毎月返済額"), "#,##0")
            初回返済額 = P8.FFormat(wRs("初回返済額"), "#,##0")
            最終返済額 = P8.FFormat(wRs("最終返済額"), "#,##0")
            
            金利種別.ListIndex = SET_LISTCOMBO(金利種別, "金利種別", P8.FCStr(wRs("金利種別")))
            基準金利.Text = P8.FCStr(wRs("基準金利区分"))
            利率 = P8.FFormat(wRs("利率"), "#,##0.00000")
            金利条件 = P8.FCStr(wRs("金利条件"))
            
            長短区分.ListIndex = SET_LISTCOMBO(長短区分, "長短区分", P8.FCStr(wRs("長短区分")))
            担保区分.ListIndex = SET_LISTCOMBO(担保区分, "有担フラグ", P8.FCStr(wRs("有担保フラグ")))
            担保名 = P8.FCStr(wRs("担保名"))
            設備区分.ListIndex = SET_LISTCOMBO(設備区分, "設備区分", P8.FCStr(wRs("設備フラグ")))
            資金用途 = P8.FCStr(wRs("資金用途"))
            
            '金利シミュレーション
            金融リストラ番号 = P8.FCStr(wRs("金融リストラ番号"))
            SM区分 = wRs("Sm区分")
            金融解約日 = Format(P8.FCStr(wRs("金融解約実行日")), Gfmt年月日)
            金利グループ区分.Text = P8.FCStr(wRs("金利グループ区分"))
            
            '保証会社
            保証料率 = P8.FFormat(wRs("保証料率"), "#,##0.00000")
            自己資金 = wRs("自己資金フラグ")
            保証料分割 = wRs("保証料分割フラグ")
            保証会社区分.Text = P8.FCStr(wRs("保証会社区分"))
            融資区分.Text = P8.FCStr(wRs("融資区分"))
            
            '取消 = wRs("取消フラグ")
            
            'If 借入計画番号 = "" _
            'And 金融リストラ番号 <> "" Then
            '    LblSM区分.Visible = True
            '    SM区分.Visible = True
            'Else
            '    LblSM区分.Visible = False
            '    SM区分.Visible = False
            'End If

            '金利初回年月
            'GRet = 金利初回年月_修正不可_区分支払(ws利息区分, wi利息支払)
            'If GRet = True Then
            '    L_金利初回年月.Caption = 金利初回年月.Text
            '
            '    '金利初回年月.Visible = False
            '    'L_金利初回年月.Visible = True
            'End If

            If P8.FCStr(金利種別.Text) = P8.FCStr(XMXA020_区分("金利種別", "1")) Then
                金利変更.Enabled = False
            Else
                金利変更.Enabled = True
            End If

            '明細入力画面 入力登録
            If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
                明細入力.Enabled = True
                内入入力.Enabled = False
                金利変更.Enabled = False
            End If
                        
            '金利変更年月
            For j = 2 To 100
                ws01 = "金利変更" & CStr(j) & "回目年月"
                GVar1 = P8.FCDate(wRs(ws01))
                
                If Not IsNull(GVar1) Then
                    If j = 2 Then
                        w初回変更年月 = GVar1
                    End If
                    
                    w最終変更年月 = GVar1
                Else
                    Exit For
                End If
                
            Next j
            
            
'            ''プロジェクト以外の時、融資残高<>0のチェック
'            w借入金 = MBD010_借入データセット(wRs)
'
'            'If w借入金.手入力区分 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) _
'            'And プロジェクト名 = "" Then
'            If w借入金.手入力区分 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
'
'                '** 借入金テーブル セット **
'                Call MBD010_借入金テーブル作成(w借入金.金融リストラ番号, w借入金)
'
'                w融資残高 = MBD010_借入最終融資残高(w借入金, CDate(w借入金.最終返済実行日))
'                If w融資残高 <> 0 Then
'                    wsMsg = "融資残高を確認してください"
'                End If
'            End If
            
        End If
    wRs.Close
    Set wRs = Nothing
    
    wSyusi = P8.FCDbl(融資金額)
    wSshokai = P8.FCDbl(初回返済額)
    wSsaishu = P8.FCDbl(最終返済額)
    wSmaituki = P8.FCDbl(毎月返済額)
    
    FLG_GSET = False
'
    '金利初回年月
    wi単位 = P8.FCDbl(返済単位月数)
    wv実行日 = 実行日
    wv初回返済年月 = 初回返済年月
    If P8.FCDbl(登録方法.Text) = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        GVar1 = 金利初回年月_セット
        Call 金利年月_リスト作成(GVar1)
        Call 金利年月_リストセット(L_金利初回年月.Caption)
    Else
        L_金利初回年月.Caption = ""
    End If
'
    '2010/04/30 変更　check
    wvJikou = C年月日.平成To西暦("年月日", P8.FCStr(実行日.Text))
    If wvJikou = 0 Then
        Exit Function
    End If
    
    wvShokai = C年月日.平成To西暦("年月日", P8.FCStr(初回返済実行日.Text))
    If wvShokai = 0 Then
        Exit Function
    End If
    
    If Format(CDate(wvJikou), "yyyy/mm/dd") = Format(CDate(wvShokai), "yyyy/mm/dd") Then
        Call 金利初回年月_実行日初回返済同一(wvJikou)
    End If
'
    '融資残高メッセージがある場合は表示
    'メッセージ = wsMsg
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
' 登録件数CHECK
'------------------------------------------------
Private Function 登録件数CHECK() As Integer
'
    Dim wRs2 As ADODB.Recordset
    Dim wstr2 As String
'
    wstr2 = "Select Count(*) As カウント From " & wsTbl
    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        登録件数CHECK = wRs2("カウント")
    
    wRs2.Close
    Set wRs2 = Nothing
'
    Exit Function
'
End Function

Private Sub SM区分_Click()
    
    If SM区分.Value = 0 Then
        fra借入金データ.ForeColor = vbBlack
    Else
        fra借入金データ.ForeColor = vbRed
    End If
    
End Sub

'------------------------------------------------
' 検索_Click
'------------------------------------------------
Private Sub 検索_Click()
    'Call 登録後初期セット
'
    GStr = wFname
    GStr_1 = ""
'
'    Unload Me
'    Me.Enabled = False
'
    frm_K借入金検索.Show
'
End Sub

'------------------------------------------------
' Copy_Click
'------------------------------------------------
Private Sub Copy_Click()

    Dim wRet As String

    If 借入番号 = "" Then
        GRet = MsgBox("既存の借入内容をセットしてください。")
            Exit Sub
    End If
    
    wRet = InputBox("新規借入番号を入力してください。", "借入金登録コピー")
    If wRet = "" Then
        Exit Sub
    End If
    
    '借入番号重複チェック
    GRet = Check_KARIIRENO(wRet)
    If GRet = False Then
        GRet = MsgBox("借入番号が重複しています。", vbOKOnly)
        Exit Sub
    End If

    If 借入番号 <> "" Then
        GRet = MsgBox("コピーして新規借入金登録データを作成します。よろしいですか？", vbYesNo)
        If GRet = vbNo Then
            Exit Sub
        End If
    End If

    'コピー
    GRet = Copy_KARIIRENO(借入番号.Text, wRet)
    If GRet = False Then
        GRet = MsgBox("コピーできませんでした。", vbOKOnly)
        Exit Sub
    End If

    '画面セット
    借入番号.Text = wRet
    Call 画面セット(False)

End Sub

'------------------------------------------------
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w借入番号 As String
    Dim w金融リストラ番号 As String
'
    w借入番号 = 借入番号
    w金融リストラ番号 = 金融リストラ番号
    
    借入番号 = ""
    金融リストラ番号 = ""
    Call 画面セット(False)
    新規変更.Caption = ""
    
    Call CEkey.SetFs(借入番号, True)
'
End Sub

'------------------------------------------------
' 借入番号_LostFocus
'------------------------------------------------
Private Sub 借入番号_LostFocus()
'
'    Call P8.FCControlLeft(借入番号, 30)
    
'    Select Case Screen.ActiveControl.Name
'        Case "閉じる", "検索", "借入番号", "銀行詳細", "金利変更", "明細入力", "内入入力", "明細書表示", _
'            "CSV出力", "CSV取込", "削除"
'            Exit Sub
'        Case "借入内容"
'            If FLG_Src = True Then
'                Exit Sub
'            End If
'    End Select
'
'    If 借入番号 = "" Then
''        MsgBox "コードを入力してください"
''        Call CEkey.SetFs(借入番号, True)
'        Exit Sub
'    End If
''
'    Select Case Screen.ActiveControl.Name
'        Case "登録"
''            Call CEkey.SetFs(借入内容, True)
'            MsgBox "該当データをセットします。登録処理は行いません。"
'            Call 画面セット(False)
'            Exit Sub
'    End Select
'
    If wTmp借入番号 = P8.FCStr(借入番号) Then
    Else
        Call 画面セット(False)
    End If
'
End Sub

'------------------------------------------------
' 借入番号_GotFocus
'------------------------------------------------
Private Sub 借入番号_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 金融リストラ番号_Change
'------------------------------------------------
Private Sub 金融リストラ番号_Change()
'
    Call P8.FCControlLeft(金融リストラ番号, 20)
    
    'If 借入計画番号 = "" _
    'And 金融リストラ番号 <> "" Then
    '    LblSM区分.Visible = True
    '    SM区分.Visible = True
    'Else
    '    LblSM区分.Visible = False
    '    SM区分.Visible = False
    'End If
'
End Sub

'------------------------------------------------
' 登録方法_Change
'------------------------------------------------
Private Sub 登録方法_Change()
'
    '明細入力画面
    L_登録方法.Visible = False
    
    If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
    '明細入力登録
'        明細入力.Enabled = True
'        内入入力.Enabled = False
'        金利変更.Enabled = False
            
        日割計算区分.Visible = True
    
        Select Case wi登録方法
        Case P8.FCDbl(XMXA020_区分("登録方法", "標準登録"))
            L_登録方法.Visible = False
        Case P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
            L_登録方法.Visible = False
        Case Else
            L_登録方法.Visible = True
        End Select
        
    Else
'        明細入力.Enabled = False
'        内入入力.Enabled = True
'        金利変更.Enabled = True
        
        日割計算区分.Visible = False
        日割計算区分.Text = CDbl(XMXA020_区分("日割計算区分", "自動計算"))
        
    End If
'
    Call 登録方法_画面セット
'
    If SM区分.Value = 0 Then
        fra借入金データ.ForeColor = vbBlack
    Else
        fra借入金データ.ForeColor = vbRed
    End If

End Sub

'------------------------------------------------
' 登録方法_画面セット
'------------------------------------------------
Private Sub 登録方法_画面セット()
'
    '金利種別.Text = 0
    '金利条件 = ""
    '利率 = 0
    '保証料率 = 0
    
    '金利初回年月 = ""
'    返済単位月数 = 1
    If P8.FCDbl(返済単位月数) < 1 Or P8.FCDbl(返済単位月数) > 12 Then
        返済単位月数 = 1
    End If
    解約年月日 = ""
    金融解約日 = ""
'    毎月返済額 = 0
'    初回返済額 = 0
'    最終返済額 = 0
    
    '金利種別.Visible = True
    '金利条件.Visible = True
    '利率.Visible = True
    '保証料率.Visible = True

    '金利初回年月.Visible = True
    C_金利初回年月.Visible = True
    返済単位月数.Visible = True
    解約年月日.Visible = True
    金融解約日.Visible = True
    毎月返済額.Visible = True
    初回返済額.Visible = True
    最終返済額.Visible = True
    
    L_金利初回年月.Caption = ""
    L_返済単位月数.Caption = ""
    L_解約年月日.Caption = ""
    L_金融解約日.Caption = ""
    L_毎月返済額.Caption = ""
    L_初回返済額.Caption = ""
    L_最終返済額.Caption = ""
    
    L_返済方法.Caption = "元金均等返済"

    If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        
        '2010/07/16 入力登録でも入力可
        '金利種別.Visible = False
        '金利条件.Visible = False
        '利率.Visible = False
        '保証料率.Visible = False
        L_返済方法.Caption = ""
        
        L_金利初回年月.Visible = True
        
        '金利初回年月.Visible = False
        C_金利初回年月.Visible = False
        返済単位月数.Visible = False
        解約年月日.Visible = False
        金融解約日.Visible = False
        
        毎月返済額.Visible = False
        初回返済額.Visible = False
        最終返済額.Visible = False
    
'        毎月返済額 = 0
'        初回返済額 = 0
'        最終返済額 = 0
    
        '金利初回年月 = ""
        If P8.FCDbl(返済単位月数) < 1 Or P8.FCDbl(返済単位月数) > 12 Then
            返済単位月数 = 1
        End If
        解約年月日 = ""
        金融解約日 = ""
    End If
'
End Sub

Private Sub 銀行_Change()
    If FLG_GSET = False Then
        Call 銀行_セット
    End If
End Sub

'------------------------------------------------
' 支払日_LostFocus
'------------------------------------------------
Private Sub 支払日_LostFocus()
    wi支払日 = P8.FCDbl(支払日.Text)
'
    If 支払日.Text = "" Then
        支払日.Text = 31
        wi支払日 = P8.FCDbl(支払日.Text)
    ElseIf 支払日.Text <> "" And 支払日.P8_Name = "" Then
        支払日.Text = 31
        wi支払日 = P8.FCDbl(支払日.Text)
    End If
'
End Sub

'------------------------------------------------
' LostFocus
'------------------------------------------------
Private Sub 借入計画番号_LostFocus()
    Call P8.FCControlLeft(借入計画番号, 20)
End Sub

Private Sub 借入内容_LostFocus()
    Call P8.FCControlLeft(借入内容, 50)
End Sub

Private Sub 登録方法_LostFocus()
    If 登録方法.Text = "" Then
         登録方法.Text = 0
    ElseIf 登録方法.Text <> "" And 登録方法.P8_Name = "" Then
         登録方法.Text = 0
    End If
End Sub

Private Sub 日割計算区分_LostFocus()
    If 日割計算区分.Text = "" Then
        日割計算区分.Text = 0
    ElseIf 日割計算区分.Text <> "" And 日割計算区分.P8_Name = "" Then
        日割計算区分.Text = 0
    End If
End Sub

Private Sub 実行日_LostFocus()
    実行日 = C年月日.FormatDate("年月日", 実行日)
End Sub

Private Sub 初回返済年月_LostFocus()
'
    Dim wvShokai As Variant, wvJikou As Variant, wv01 As Variant
'
    初回返済年月 = C年月日.FormatDate("年月", 初回返済年月)
    
    If wi単位 <> P8.FCDbl(返済単位月数) _
    Or wv実行日 <> 実行日.Text _
    Or wv初回返済年月 <> 初回返済年月.Text Then
        GVar1 = 金利初回年月_セット
        Call 金利年月_リスト作成(GVar1)
'
'        GVar1 = MXA030_金利初回年月(ws利息区分, wi利息支払, wi支払日, wi営業日, C年月日.平成To西暦("年月日", 実行日), C年月日.平成To西暦("年月日", 初回返済年月), P8.FCDbl(返済単位月数))
        Call 金利年月_リストセット(CStr(GVar1))
    Else
        If C_金利初回年月.ListCount = 0 Then
            GVar1 = 金利初回年月_セット
            Call 金利年月_リスト作成(GVar1)
            
            Call 金利年月_リストセット(CStr(GVar1))
        End If
    End If

    If P8.FCStr(初回返済年月) = "" Then
        初回返済実行日 = ""
        Exit Sub
    End If
    
    初回返済実行日 = 実行日計算(初回返済年月, 初回返済実行日)
'
End Sub

Private Sub 初回返済実行日_LostFocus()
'
    Dim wvJikou As Variant, wvShokai As Variant
'
    初回返済実行日 = C年月日.FormatDate("年月日", 初回返済実行日)
'
    '2010/04/30 変更　check
    wvJikou = C年月日.平成To西暦("年月日", P8.FCStr(実行日.Text))
    If wvJikou = 0 Then
        Exit Sub
    End If
    
    wvShokai = C年月日.平成To西暦("年月日", P8.FCStr(初回返済実行日.Text))
    If wvShokai = 0 Then
        Exit Sub
    End If
    
    If Format(CDate(wvJikou), "yyyy/mm/dd") = Format(CDate(wvShokai), "yyyy/mm/dd") Then
        Call 金利初回年月_実行日初回返済同一(wvJikou)
    End If
'
End Sub

Private Sub 金利初回年月_LostFocus()
    '金利初回年月 = C年月日.FormatDate("年月", 金利初回年月)
End Sub

Private Sub 最終返済年月_LostFocus()
    最終返済年月 = C年月日.FormatDate("年月", 最終返済年月)
    
    If P8.FCStr(最終返済年月) = "" Then
        最終返済実行日 = ""
        Exit Sub
    End If
    
    最終返済実行日 = 実行日計算(最終返済年月, 最終返済実行日)
End Sub

Private Sub 最終返済実行日_LostFocus()
    最終返済実行日 = C年月日.FormatDate("年月日", 最終返済実行日)
End Sub

Private Sub 解約年月日_LostFocus()
    解約年月日 = C年月日.FormatDate("年月日", 解約年月日)
End Sub

Private Sub 金融解約日_LostFocus()
    金融解約日 = C年月日.FormatDate("年月日", 金融解約日)
End Sub

Private Sub 返済単位月数_LostFocus()
    If P8.FCDbl(返済単位月数) < 1 Or P8.FCDbl(返済単位月数) > 12 Then
        返済単位月数 = 1
    End If
End Sub

Private Sub 融資金額_LostFocus()
    融資金額 = Right$(P8.FFormat(融資金額, "#,##0"), 15)
    
    Call 融資金額_セット
End Sub

Private Sub 毎月返済額_LostFocus()
    毎月返済額 = Right$(P8.FFormat(毎月返済額, "#,##0"), 15)
    wSmaituki = P8.FCDbl(毎月返済額)
End Sub

Private Sub 初回返済額_LostFocus()
    初回返済額 = Right$(P8.FFormat(初回返済額, "#,##0"), 15)
    wSshokai = P8.FCDbl(初回返済額)
End Sub

Private Sub 最終返済額_LostFocus()
    最終返済額 = Right$(P8.FFormat(最終返済額, "#,##0"), 15)
    wSsaishu = P8.FCDbl(最終返済額)
End Sub

Private Sub 金利条件_LostFocus()
    Call P8.FCControlLeft(金利条件, 50)
End Sub

Private Sub 利率_LostFocus()
    利率 = P8.FFormat(利率, "#,##0.00000")
End Sub

Private Sub 保証料率_LostFocus()
    保証料率 = P8.FFormat(保証料率, "#,##0.00000")
End Sub

'------------------------------------------------
' 実行日計算
'------------------------------------------------
Private Function 実行日計算(p年月 As String, p実行日 As String) As String
'
    On Error GoTo 実行日計算_ERR
'
    wi支払日 = P8.FCDbl(支払日.Text)
'
    GVar1 = C年月日.平成To西暦("年月", p年月)
    GVar2 = C年月日.平成To西暦("年月", p実行日)
    If Not IsNull(GVar1) Then
        GVar1 = MXA030_翌営業年月日計算(CDate(GVar1), wi支払日, wi営業日)
    Else
        実行日計算 = ""
        Exit Function
    End If
'
    If p実行日 = "" Then
        実行日計算 = Format(CDate(GVar1), Gfmt年月日)
    Else
        実行日計算 = p実行日
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
実行日計算_ERR:
    実行日計算 = ""
End Function

'------------------------------------------------
' 融資金額_セット
'------------------------------------------------
Private Sub 融資金額_セット()
'
    Dim wi01 As Integer, wi02 As Integer
    Dim w余 As Long
    Dim wSiharai As Integer, wHensai As Integer
    Dim wStrdate As Date, wEnddate As Date
    Dim wMaikin As Double, wEndkin As Double
    Dim wHyusi As Double, wHshokai As Double, wHsaishu As Double, wHmaituki As Double
'
    On Error GoTo 融資金額_セット_ERR
'
    wStrdate = C年月日.平成To西暦("年月", 初回返済年月)
    wEnddate = C年月日.平成To西暦("年月", 最終返済年月)
    wSiharai = DateDiff("m", wStrdate, wEnddate) + 1
'
    wi01 = P8.FCDbl(返済単位月数)
    If wi01 < 1 Or wi01 > 12 Then
        MsgBox "返済単位月数を確認してください": Call CEkey.SetFs(返済単位月数, True)
        Exit Sub
    End If
'
    If wi01 > 1 Then
        wi02 = Fix(wSiharai Mod wi01)
        'wi01=1になるように
        If wi02 <> 1 Then
            MsgBox "初回返済年月又は最終返済年月又は返済単位月数が違います"
            Call CEkey.SetFs(初回返済年月, True)
            Exit Sub
        End If
    End If
    
    
    w余 = wSiharai Mod wi01
    wHensai = Fix(wSiharai / wi01)
    If w余 <> 0 Then
        wHensai = wHensai + 1
    End If
    
    wHyusi = P8.FCDbl(融資金額)
    wHshokai = P8.FCDbl(初回返済額)
    wHsaishu = P8.FCDbl(最終返済額)
    wHmaituki = P8.FCDbl(毎月返済額)
    
    wMaikin = P8.FRound(P8.FCDiv(wHyusi, wHensai), 3)
    wEndkin = wHyusi - (wMaikin * (wHensai - 1))
    
    If wSyusi = wHyusi And wSshokai = wHshokai _
       And wSsaishu = wHsaishu And wSmaituki = wHmaituki _
       And wHyusi = wHshokai + (wHmaituki * (wHensai - 2)) + wHsaishu Then
        Exit Sub
    End If

    毎月返済額 = P8.FFormat(wMaikin, "#,##0")
    初回返済額 = P8.FFormat(wMaikin, "#,##0")
    最終返済額 = P8.FFormat(wEndkin, "#,##0")
    
    wSyusi = P8.FCDbl(融資金額)
    wSshokai = P8.FCDbl(初回返済額)
    wSsaishu = P8.FCDbl(最終返済額)
    wSmaituki = P8.FCDbl(毎月返済額)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
融資金額_セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 融資金額_セット() でエラー" + vbCrLf + vbCrLf + _
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
' 金利年月_リスト作成
'------------------------------------------------
Private Sub 金利年月_リスト作成(pDate As Variant)
'
    Dim wi01 As Integer, j As Integer
    Dim wvJikou As Variant, wvShokai As Variant, wv01 As Variant
'
    On Error GoTo 金利年月_リスト作成_ERR
'
    C_金利初回年月.Clear
'
    wi01 = P8.FCDbl(返済単位月数)
    If wi01 = 0 Then
        wi01 = 1
    End If
    
    If ws利息区分 = XMXA020_区分("利息区分", "利息先払") Then
        If CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
            wi01 = 1
        Else
            
        End If
    
    ElseIf ws利息区分 = XMXA020_区分("利息区分", "利息後払") Then
        If CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
            wi01 = 1
        ElseIf CStr(wi利息支払) = XMXA020_区分("利息支払", "一括") Then
        End If
    End If
    
'
    wvJikou = C年月日.平成To西暦("年月日", P8.FCStr(実行日.Text))
    If wvJikou = 0 Then
        wvJikou = Null
    End If
    
    wvShokai = C年月日.平成To西暦("年月日", P8.FCStr(初回返済年月.Text))
    If wvShokai = 0 Then
        wvShokai = Null
        Exit Sub
    End If
    
    wv01 = C年月日.平成To西暦("年月日", CStr(pDate))
    If wv01 = 0 Then
        wv01 = Null
        Exit Sub
    End If
    
    Do While Format(CDate(wv01), "yyyy/mm/01") <= Format(CDate(wvShokai), "yyyy/mm/01")
        C_金利初回年月.AddItem Format(wv01, Gfmt年月)
        
        wv01 = DateAdd("m", wi01, CDate(wv01))
    Loop
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利年月_リスト作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利年月_リスト作成() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
End Sub

'------------------------------------------------
' 金利年月_リストセット
'------------------------------------------------
Private Sub 金利年月_リストセット(p年月 As String)
'
    Dim j As Integer
    Dim wv01 As Variant, wv02 As Variant
'
    On Error GoTo 金利年月_リストセット_ERR
'
    wv01 = C年月日.平成To西暦("年月", p年月)
    If wv01 = 0 Then
        wv01 = Null
    End If
    If Not IsNull(wv01) And IsDate(Format(wv01, "yyyy/mm/dd")) Then
        For j = 0 To C_金利初回年月.ListCount
            wv02 = C年月日.平成To西暦("年月", C_金利初回年月.List(j))
            If Format(wv01, "yyyy/mm/dd") = Format(wv02, "yyyy/mm/dd") Then
                C_金利初回年月 = C_金利初回年月.List(j)
                Exit For
            End If
        Next j
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利年月_リストセット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利年月_リストセット() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
End Sub

'------------------------------------------------
' 金利初回年月_セット
'------------------------------------------------
Private Function 金利初回年月_セット() As Variant
'
    Dim wi01 As Integer, w支払日 As Integer
    Dim wvJikou As Variant, wvShokai As Variant, wvJikou2 As Variant
    Dim wv01 As Variant
    
'
    金利初回年月_セット = ""
    
    '金利初回年月.Text = ""
    'L_金利初回年月.Caption = ""
    
    '金利初回年月.Visible = True
    'L_金利初回年月.Visible = False
'
    wvJikou = C年月日.平成To西暦("年月日", P8.FCStr(実行日.Text))
    If wvJikou = 0 Then
        wvJikou = Null
    End If
    If Not IsDate(wvJikou) Then
        Exit Function
    End If
'
    wvShokai = C年月日.平成To西暦("年月日", P8.FCStr(初回返済年月.Text))
    If wvShokai = 0 Then
        wvShokai = Null
    End If
    If Not IsDate(wvShokai) Then
        Exit Function
    End If
'
    If FLG_New = True Then
        Call 銀行_セット
    End If
    
    wi支払日 = P8.FCDbl(支払日.Text)
    If wi支払日 < 1 Or wi支払日 > 31 Then
        Exit Function
    End If
'
    wvJikou2 = MBD010_実行日支払年月算出(wvJikou, wi営業日, wi支払日)
    
    wv01 = MXA030_金利初回年月(ws利息区分, wi利息支払, wi支払日, wi営業日, wvJikou2, wvShokai, P8.FCDbl(返済単位月数))
    If Not IsDate(wv01) Then
        Exit Function
    End If

    '金利初回年月.Text = Format(CStr(wv01), Gfmt年月)
'
    金利初回年月_セット = Format(CStr(wv01), Gfmt年月)
    'L_金利初回年月.Caption = 金利初回年月.Text
    
    'GRet = 金利初回年月_修正不可_区分支払(ws利息区分, wi利息支払)
    'If GRet = True Then
    '    L_金利初回年月.Caption = 金利初回年月.Text
        
        '金利初回年月.Visible = False
        'L_金利初回年月.Visible = True
    'End If
'
End Function

'------------------------------------------------
' 金利初回年月_修正不可_区分支払
'------------------------------------------------
Private Function 金利初回年月_修正不可_区分支払(p利息区分 As String, p利息支払 As Integer) As Boolean
'
    金利初回年月_修正不可_区分支払 = False
'
    If p利息区分 = XMXA020_区分("利息区分", "利息先払") Then
        金利初回年月_修正不可_区分支払 = True
    ElseIf p利息区分 = XMXA020_区分("利息区分", "利息後払") Then
        If CStr(p利息支払) = XMXA020_区分("利息支払", "一括") Then
            金利初回年月_修正不可_区分支払 = True
        End If
    End If
'
End Function

'------------------------------------------------
' 金利初回年月_実行日初回返済同一
'------------------------------------------------
Private Sub 金利初回年月_実行日初回返済同一(pJikou As Variant)
'
    Dim wvJikou As Variant
    Dim wv01 As Variant
'
    '2010/04/30 変更　check
    C_金利初回年月.Clear

    wvJikou = MBD010_実行日支払年月算出(pJikou, wi営業日, wi支払日)
    
    If CStr(wi利息支払) = XMXA020_区分("利息支払", "一括") Then
        wv01 = DateAdd("m", P8.FCDbl(返済単位月数) - 1, CDate(wvJikou))
        C_金利初回年月.AddItem Format(wv01, Gfmt年月)
    ElseIf CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
        wv01 = CDate(wvJikou)
        C_金利初回年月.AddItem Format(wv01, Gfmt年月)
    End If

    C_金利初回年月 = C_金利初回年月.List(0)
'
End Sub

'------------------------------------------------
' 銀行_セット
'------------------------------------------------
Private Function 銀行_セット()
'
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA040_銀行マスタ"
    wstr = wstr + " Where 銀行番号 = '" & P8.FCStr(銀行.Text) & "'"
    wstr = wstr + " And 取消フラグ=0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        
        If wi支払日 = 0 Then
            支払日.Text = P8.FCDbl(wRs("支払日"))
        End If

        wi営業日 = P8.FCDbl(wRs("営業日区分"))
        ws利息区分 = P8.FCStr(wRs("利息区分"))
        wi利息日数 = P8.FCDbl(wRs("利息計算日数区分"))
        wi利息支払 = P8.FCDbl(wRs("利息支払方法"))
        wi利息控除 = P8.FCDbl(wRs("利息控除区分"))
        wi金利計算 = P8.FCDbl(wRs("金利計算年間日数"))
        
    End If
    wRs.Close
    Set wRs = Nothing
'
End Function

'------------------------------------------------
' 登録_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim w借入金マスタ As MAA910_借入金
'    Dim w銀行マスタ As MAA030_銀行
    
    Dim j As Integer, wi01 As Integer
    Dim wFind As Boolean
    Dim ws01 As String
    Dim wdate As Date
    
    Dim wd01 As Date
    Dim w実行日 As Date, w初回返済年月 As Date, w最終返済年月 As Date
    Dim wc実行日 As Date, wc解約実行日 As Date, wc金融解約実行日 As Date
    Dim w初回返済実行日 As Date, w最終返済実行日 As Date
    Dim w初回返済1前 As Date, w初回返済1後 As Date
    Dim w最終返済1前 As Date
    Dim w実行支払年月 As Date, w金利初回年月 As Date
    
    Dim wv01 As Variant
    Dim w解約実行日 As Variant, w金融解約実行日 As Variant
    Dim w金利変更年月 As Variant, w金利変更年月前 As Variant
    
    Dim w支払日 As Integer
    Dim w支払回数 As Integer
    Dim w返済単位回数 As Integer
    Dim w融資額 As Double

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

    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    If P8.FCStr(借入番号) = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        Call CEkey.SetFs(借入番号, True)
        Exit Sub
    End If
'
    If 借入金種別.Text = "" Then
        MsgBox "借入金種別が不正です", vbExclamation
        Call CEkey.SetFs(借入金種別, True)
        Exit Sub
    End If

    If 借入金種別.Text <> "" And 借入金種別.P8_Name = "" Then
        MsgBox "借入金種別が不正です", vbExclamation
        Call CEkey.SetFs(借入金種別, True)
        Exit Sub
    End If
'
    If 部門.Text <> "" And 部門.P8_Name = "" Then
        MsgBox "部門が不正です", vbExclamation
        Call CEkey.SetFs(部門, True)
        Exit Sub
    End If
'
    If 銀行.Text = "" Then
        MsgBox "銀行が不正です", vbExclamation
        Call CEkey.SetFs(銀行, True)
        Exit Sub
    End If
    If 銀行.Text <> "" And 銀行.P8_Name = "" Then
        MsgBox "銀行が不正です", vbExclamation
        Call CEkey.SetFs(銀行, True)
        Exit Sub
    End If
'
    If 支払日.Text <> "" And 支払日.P8_Name = "" Then
        MsgBox "支払日が不正です", vbExclamation
        Call CEkey.SetFs(支払日, True)
        Exit Sub
    End If
'
    If 金利種別.ListIndex < 0 Then
        MsgBox "金利種別を選択してください", vbExclamation
        Call CEkey.SetFs(金利種別, True)
        Exit Sub
    End If
'
    If 担保区分.ListIndex < 0 Then
        MsgBox "担保区分を選択してください", vbExclamation
        Call CEkey.SetFs(担保区分, True)
        Exit Sub
    End If
'
    If 長短区分.ListIndex < 0 Then
        MsgBox "長短区分を選択してください", vbExclamation
        Call CEkey.SetFs(長短区分, True)
        Exit Sub
    End If
'
    If 担保区分.ListIndex < 0 Then
        MsgBox "担保区分を選択してください", vbExclamation
        Call CEkey.SetFs(担保区分, True)
        Exit Sub
    End If
    
    If 設備区分.ListIndex < 0 Then
        MsgBox "設備区分を選択してください", vbExclamation
        Call CEkey.SetFs(設備区分, True)
        Exit Sub
    End If
'
    If 金利グループ区分.Text <> "" And 金利グループ区分.P8_Name = "" Then
        MsgBox "金利グループが不正です", vbExclamation
        SSTab1.Tab = 1
        Call CEkey.SetFs(金利グループ区分, True)
        Exit Sub
    End If
'
    If 基準金利.Text <> "" And 基準金利.P8_Name = "" Then
        MsgBox "基準金利が不正です", vbExclamation
        Call CEkey.SetFs(基準金利, True)
        Exit Sub
    End If
'
    If GSys.Sit = True Then
'        If G基本情報.支店コード = G独算(0).支店コード And G基本情報.企業区分 <> "単独企業" Then
        If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結親会社" Then
            Exit Sub
        End If
        
        For j = 2 To UBound(G独算)
            If G基本情報.支店コード = G独算(j).支店コード Then
                If P8.FCStr(銀行.Text) <> "SS" Then
                    MsgBox "入力を確認してください": Call CEkey.SetFs(銀行, True)
                    
                    Exit Sub
                End If
                
                'If SM区分.Visible = True And P8.FCDbl(SM区分) = 0 Then
                '    MsgBox "入力を確認してください": Call CEkey.SetFs(SM区分, True)
                '
                '    Exit Sub
                'End If
                
                Exit For
            End If
        Next j
    End If
'
    If Not IsNumeric(融資金額) And 融資金額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(融資金額, True)
        Exit Sub
    End If
    If Not IsNumeric(毎月返済額) And 毎月返済額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(毎月返済額, True)
        Exit Sub
    End If
    If Not IsNumeric(初回返済額) And 初回返済額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(初回返済額, True)
        Exit Sub
    End If
    If Not IsNumeric(最終返済額) And 最終返済額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(最終返済額, True)
        Exit Sub
    End If
    If (Not IsNumeric(利率) And 利率 <> "") Or P8.FCDbl(利率) >= 100 Or P8.FCDbl(利率) < 0 Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(利率, True)
        Exit Sub
    End If
    If (Not IsNumeric(返済単位月数) And 返済単位月数 <> "") Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(返済単位月数, True)
        Exit Sub
    End If
    If P8.FCDbl(返済単位月数) < 1 Or P8.FCDbl(返済単位月数) > 12 Then
        MsgBox "返済単位月数を確認してください", vbExclamation: Call CEkey.SetFs(返済単位月数, True)
        Exit Sub
    End If
'
'    '保証会社
'    If (Not IsNumeric(保証料率) And 保証料率 <> "") Or P8.FCDbl(保証料率) >= 100 Or P8.FCDbl(保証料率) < 0 Then
'        MsgBox "入力を確認してください": Call CEkey.SetFs(保証料率, True)
'        Exit Sub
'    End If
''
'    If 保証会社区分.Text <> "" And 保証会社区分.P8_Name = "" Then
'        MsgBox "コードが違います"
'        Call CEkey.SetFs(保証会社区分, True)
'        Exit Sub
'    End If
''
'    If 融資区分.Text <> "" And 融資区分.P8_Name = "" Then
'        MsgBox "コードが違います"
'        Call CEkey.SetFs(融資区分, True)
'        Exit Sub
'    End If
''
'    wFind = False
'    For j = 0 To 借入計画番号.ListCount
'        If 借入計画番号 = 借入計画番号.List(j) Then
'            wFind = True
'            Exit For
'        End If
'    Next
'    If Not wFind Then
'        MsgBox "リストボックスより選択してください"
'        Call CEkey.SetFs(借入計画番号, True)
'        Exit Sub
'    End If
'
    If C年月日.平成To西暦("年月日", 実行日) = 0 Then
        MsgBox "実行日が不正です", vbExclamation
        Call CEkey.SetFs(実行日, True)
        Exit Sub
    End If
    
    If P8.FCDbl(登録方法.Text) = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        If C年月日.平成To西暦("年月", 初回返済年月) = 0 Then
            MsgBox "初回返済年月が不正です", vbExclamation
            Call CEkey.SetFs(初回返済年月, True)
            Exit Sub
        End If
        
        If C年月日.平成To西暦("年月", 初回返済実行日) = 0 Then
            MsgBox "初回返済年月日が不正です", vbExclamation
            Call CEkey.SetFs(初回返済実行日, True)
            Exit Sub
        End If
        
        If C年月日.平成To西暦("年月", C_金利初回年月.Text) = 0 Then
            MsgBox "金利初回年月が不正です", vbExclamation
            Call CEkey.SetFs(C_金利初回年月, True)
            Exit Sub
        End If
        
        If C年月日.平成To西暦("年月", 最終返済年月) = 0 Then
            MsgBox "最終返済年月が不正です", vbExclamation
            Call CEkey.SetFs(最終返済年月, True)
            Exit Sub
        End If
        
        If C年月日.平成To西暦("年月", 最終返済実行日) = 0 Then
            MsgBox "最終返済年月日が不正です", vbExclamation
            Call CEkey.SetFs(最終返済実行日, True)
            Exit Sub
        End If
        
        If C年月日.平成To西暦("年月日", 解約年月日, True) = 0 Then
            MsgBox "解約年月日が不正です", vbExclamation
            Call CEkey.SetFs(解約年月日, True)
            Exit Sub
        End If
        
        If C年月日.平成To西暦("年月日", 金融解約日, True) = 0 Then
            MsgBox "金融解約日が不正です", vbExclamation
            Call CEkey.SetFs(金融解約日, True)
            Exit Sub
        End If
    End If
'
    '手入力の場合
    '詳細TRにデータがある場合、初回返済年月と最終返済年月のCHECK
    w借入金マスタ.借入番号 = P8.FCStr(借入番号)
    w借入金マスタ.借入貸付 = wiTblNo
    Call MBD010_借入金入力明細Read(w借入金マスタ)
'
    w実行日 = C年月日.平成To西暦("年月日", 実行日)
    If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
    
        If UBound(G借入金入力) > 0 Then
            wdate = C年月日.平成To西暦("年月", P8.FCStr(初回返済年月))
            If Format(w実行日, "yyyymmdd") <> Format(G借入金入力(1).借入返済年月日, "yyyymmdd") _
            And Format(wdate, "yyyymm") < Format(G借入金入力(1).借入返済年月日, "yyyymm") Then
                MsgBox "年月が違います。入力登録の年月にセットします", vbExclamation
                初回返済年月 = Format(G借入金入力(1).借入返済年月日, Gfmt年月)
            End If
            
            wdate = C年月日.平成To西暦("年月", P8.FCStr(最終返済年月))
            If Format(wdate, "yyyymm") > Format(G借入金入力(UBound(G借入金入力)).借入返済年月日, "yyyymm") Then
                MsgBox "年月が違います。入力登録の年月にセットします", vbExclamation
                最終返済年月 = Format(G借入金入力(UBound(G借入金入力)).借入返済年月日, Gfmt年月)
            End If
        End If

    End If
'
    ' =========================================
    '                   前処理
    ' =========================================
    w実行日 = C年月日.平成To西暦("年月日", 実行日)
    w初回返済年月 = C年月日.平成To西暦("年月", 初回返済年月)
    w初回返済実行日 = C年月日.平成To西暦("年月", 初回返済実行日) 'V188
    w最終返済年月 = C年月日.平成To西暦("年月", 最終返済年月)
    w最終返済実行日 = C年月日.平成To西暦("年月", 最終返済実行日) 'V188
    w解約実行日 = C年月日.平成To西暦("年月日", 解約年月日, True)
    w金融解約実行日 = C年月日.平成To西暦("年月", 金融解約日, True)
    
    wd01 = DateAdd("m", -1, w初回返済年月)
    w初回返済1前 = MXA030_翌営業年月日計算(wd01, wi支払日, wi営業日)
    
    wd01 = DateAdd("m", 1, w初回返済年月)
    w初回返済1後 = MXA030_翌営業年月日計算(wd01, wi支払日, wi営業日)
    
    wd01 = DateAdd("m", -1, w最終返済年月)
    w最終返済1前 = MXA030_翌営業年月日計算(wd01, wi支払日, wi営業日)
    
    wi支払日 = P8.FCDbl(支払日.Text)
    w支払回数 = DateDiff("m", w初回返済年月, w最終返済年月) + 1
'
    ' =============================================================
    '                           実行日
    ' =============================================================
    If IsDate(w実行日) Then
        'Call C休日.計算(w実行日)
        'V180
        Call C休日.計算(w実行日, wi営業日)
        wc実行日 = C休日.次回稼働日
        
        If wc実行日 <> w実行日 Then
            
            If Format(w初回返済年月, "yyyy/mm") < Format(wc実行日, "yyyy/mm") Then ' 07/02/22 V180
                MsgBox "初回返済年月が不正です", vbExclamation
                Call CEkey.SetFs(初回返済年月, True)
                Exit Sub
            End If
            
            MsgBox "実行日を稼働日にセットします"
            w実行日 = wc実行日
            実行日 = Format(w実行日, Gfmt年月日)
        End If
    End If
'
    If P8.FCDbl(登録方法.Text) = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        If IsDate(w初回返済実行日) Then
            Call C休日.計算(w初回返済実行日, wi営業日)
            wd01 = C休日.次回稼働日
            
            If wd01 <> w初回返済実行日 Then
                MsgBox "初回返済年月日が稼働日ではありません", vbExclamation
                Call CEkey.SetFs(初回返済実行日, True)
                Exit Sub
            End If
        End If

        If IsDate(w最終返済実行日) Then
            Call C休日.計算(w最終返済実行日, wi営業日)
            wd01 = C休日.次回稼働日
            
            If wd01 <> w最終返済実行日 Then
                MsgBox "最終返済年月日が稼働日ではありません", vbExclamation
                Call CEkey.SetFs(最終返済実行日, True)
                Exit Sub
            End If
        End If
    End If
'
    ' =============================================================
    ' 初回返済年月　& 解約実行日 & 金融解約実行日　の　整合性check
    ' =============================================================
    If P8.FCDbl(登録方法.Text) = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        
        If Format(w初回返済実行日, "yyyy/mm/dd") < Format(w実行日, "yyyy/mm/dd") Then                     ' 07/02/22 V180
            MsgBox "初回返済年月が違います", vbExclamation
            Call CEkey.SetFs(初回返済年月, True)
            Exit Sub
        End If
            
        If Format(w初回返済年月, "yyyy/mm") < Format(w実行日, "yyyy/mm") Then
            MsgBox "初回返済年月が違います", vbExclamation
            Call CEkey.SetFs(初回返済年月, True)
            Exit Sub
        End If
        
        If Format(w最終返済実行日, "yyyy/mm/dd") < Format(w実行日, "yyyy/mm/dd") Then
            MsgBox "最終返済年月が違います", vbExclamation
            Call CEkey.SetFs(最終返済年月, True)
            Exit Sub
        End If
           
        If Format(w初回返済実行日, "yyyy/mm/dd") < Format(w実行日, "yyyy/mm/dd") Or Format(w初回返済実行日, "yyyy/mm/dd") > Format(w最終返済実行日, "yyyy/mm/dd") Then
            MsgBox "初回返済年月が違います", vbExclamation
            Call CEkey.SetFs(初回返済年月, True)
            Exit Sub
        End If
    
        '初回返済年月＝最終返済年月の場合、初回返済年月日＝最終返済年月日
        If Format(w初回返済年月, "yyyy/mm") > Format(w最終返済年月, "yyyy/mm") Then
            If Format(w初回返済実行日, "yyyy/mm/dd") <> Format(w最終返済実行日, "yyyy/mm/dd") Then
                MsgBox "年月が誤りです", vbExclamation
                Call CEkey.SetFs(初回返済年月, True)
                Exit Sub
            End If
        End If
    
    Else
    
        If Format(w初回返済年月, "yyyy/mm") < Format(w実行日, "yyyy/mm") Or w初回返済年月 > w最終返済年月 Then
            MsgBox "初回返済年月が違います", vbExclamation
            Call CEkey.SetFs(初回返済年月, True)
            Exit Sub
        End If
    
    End If
    
    If P8.FCDbl(登録方法.Text) = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
    
        '金利初回年月
'        GRet = 金利初回年月_修正不可_区分支払(ws利息区分, wi利息支払)
'        If GRet <> True Then
'
'            w金利初回年月 = C年月日.平成To西暦("年月日", C_金利初回年月.Text)
'
'            If ws利息区分 = XMXA020_区分("利息区分", "利息後払") _
'            And CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
'
'                'wv01 = MXA030_実行支払年月(w実行日, wi支払日, wi営業日, "=")
'                '据置回数
'                'wi01 = DateDiff("m", P8.FCDate(wv01), w金利初回年月)
'
'                '金利初回年月　据置回数1回目を通す
'
'                If Format(w実行日, "yyyy/mm") > Format(w金利初回年月, "yyyy/mm") Then
'                    MsgBox "金利初回年月が違います"
'                    Call CEkey.SetFs(C_金利初回年月, True)
'                    Exit Sub
'                End If
'            Else
'                If Format(w実行日, "yyyy/mm") > Format(w金利初回年月, "yyyy/mm") Then
'                    MsgBox "金利初回年月が違います"
'                    Call CEkey.SetFs(C_金利初回年月, True)
'                    Exit Sub
'                End If
'            End If
'
'        End If

        '2010/01/12 修正
        w金利初回年月 = C年月日.平成To西暦("年月日", C_金利初回年月.Text)
        If Format(w実行日, "yyyy/mm") > Format(w金利初回年月, "yyyy/mm") Then
            MsgBox "金利初回年月が違います", vbExclamation
            Call CEkey.SetFs(C_金利初回年月, True)
            Exit Sub
        End If
        
        '2010/04/30 変更　check
        If Format(w実行日, "yyyy/mm/dd") <> Format(w初回返済実行日, "yyyy/mm/dd") Then
            If Format(w初回返済年月, "yyyy/mm") < Format(w金利初回年月, "yyyy/mm") Then
                MsgBox "金利初回年月が違います", vbExclamation
                Call CEkey.SetFs(C_金利初回年月, True)
                Exit Sub
            End If
        End If

        If Not IsNull(w解約実行日) Then
            If w解約実行日 <= w実行日 Or w解約実行日 >= w最終返済実行日 Then
                MsgBox "解約年月日が違います", vbExclamation
                Call CEkey.SetFs(解約年月日, True)
                Exit Sub
            End If
        End If
        
        If Not IsNull(w金融解約実行日) Then
            If w金融解約実行日 <= w実行日 Or w金融解約実行日 >= w最終返済実行日 Then
                MsgBox "金融解約日が違います", vbExclamation
                Call CEkey.SetFs(金融解約日, True)
                Exit Sub
            End If
        End If
        
        ' =========================================
        '             融資額のcheck
        ' =========================================
         w返済単位回数 = Fix((w支払回数 + P8.FCDbl(返済単位月数) - 1) / P8.FCDbl(返済単位月数))
         If w返済単位回数 * P8.FCDbl(返済単位月数) <> (w支払回数 + P8.FCDbl(返済単位月数) - 1) Then
                 MsgBox "初回返済年月又は最終返済年月又は返済単位月数が違います"
                 Call CEkey.SetFs(初回返済年月, True)
                 Exit Sub
         End If
         
         If w返済単位回数 > 2 Then
             w融資額 = P8.FCDbl(初回返済額) + P8.FCDbl(最終返済額) + P8.FCDbl(毎月返済額) * (w返済単位回数 - 2)
         Else
             If w返済単位回数 = 2 Then
                 w融資額 = P8.FCDbl(初回返済額) + P8.FCDbl(最終返済額)
             Else
                 If w返済単位回数 = 1 Then
                    If P8.FCDbl(初回返済額) <> 0 Then
                        w融資額 = P8.FCDbl(初回返済額)  '10/01/01
                    Else
                        w融資額 = P8.FCDbl(最終返済額)  '10/01/01
                    End If                              '10/01/01
                 End If
             End If
         End If
        
         If w融資額 <> P8.FCDbl(融資金額) Then
             MsgBox "融資金額が違います", vbExclamation
                 Call CEkey.SetFs(融資金額, True)
                 Exit Sub
         End If
        
        ' =============================================================
        '     実行日 & 解約年月日 & 金融解約日　の　翌稼働日chck & set
        ' =============================================================
        If w初回返済実行日 < w実行日 Then                   ' 07/02/22 V180
                MsgBox "初回返済年月が違います", vbExclamation
                Call CEkey.SetFs(初回返済年月, True)
                Exit Sub
        End If
        
        If IsDate(w解約実行日) Then
            'Call C休日.計算(CDate(w解約実行日))
            'V180
            Call C休日.計算(CDate(w解約実行日), wi営業日)
            wc解約実行日 = C休日.次回稼働日
            If wc解約実行日 <> CDate(w解約実行日) Then
                MsgBox "解約年月日を稼働日にセットします", vbExclamation
                w解約実行日 = wc解約実行日
            End If
        End If
        
        If IsDate(w金融解約実行日) Then
            'Call C休日.計算(CDate(w金融解約実行日))
            'V180
            Call C休日.計算(CDate(w金融解約実行日), wi営業日)
            wc金融解約実行日 = C休日.次回稼働日
            If wc金融解約実行日 <> CDate(w金融解約実行日) Then
                MsgBox "金融解約年月日を稼働日にセットします"
                w金融解約実行日 = wc金融解約実行日
            End If
        End If
    
    Else
'        w借入金マスタ.返済単位月数 = 0
        w借入金マスタ.据置回数 = 0
        w借入金マスタ.支払回数 = 0
    
    End If
'
    ' =============================================================
    ' 初回返済実行日　& 最終返済実行日 & 金利変更の整合性check
    ' =============================================================
    'V188
    If P8.FCDbl(登録方法.Text) = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        
        '2010/04/30 変更　check
        '初回返済実行日＞実行日
        If Format(w初回返済実行日, "yyyy/mm/dd") < Format(w実行日, "yyyy/mm/dd") Then
            MsgBox "初回返済年月日が違います", vbExclamation
            Call CEkey.SetFs(初回返済実行日, True)
            Exit Sub
        End If
            
        '初回返済実行日＜2回目返済実行日
        wdate = DateAdd("m", 1, w初回返済年月)
        wdate = MXA030_翌営業年月日計算(wdate, wi支払日, wi営業日)
        If w初回返済実行日 >= wdate Then
            MsgBox "初回返済年月日が違います", vbExclamation
            Call CEkey.SetFs(初回返済実行日, True)
            Exit Sub
        End If
        
        '1ヶ月後＞初回返済実行日＞1ヶ月前はOK
        If w初回返済1後 <= w初回返済実行日 Or w初回返済実行日 <= w初回返済1前 Then
            MsgBox "初回返済年月日が違います", vbExclamation
            Call CEkey.SetFs(初回返済実行日, True)
            Exit Sub
        End If
    
        '2010/04/30 変更　check
        '1ヶ月後＞初回返済実行日＞実行日はOK
        'If w初回返済1後 <= w初回返済実行日 Or w初回返済実行日 <= w実行日 Then
        If w初回返済1後 <= w初回返済実行日 Or w初回返済実行日 < w実行日 Then
            MsgBox "初回返済年月日が違います", vbExclamation
            Call CEkey.SetFs(初回返済実行日, True)
            Exit Sub
        End If
    
        '最終返済実行日＞最終回-1返済実行日
        If w最終返済実行日 <= w最終返済1前 Then
            MsgBox "最終返済年月日が違います", vbExclamation
            Call CEkey.SetFs(最終返済実行日, True)
            Exit Sub
        End If
    
    End If
'
    '----------< Button制御 >-------------------------------------------------------
    登録.Enabled = False
'
    ' =========================================
    '   登録方法：標準登録→入力登録 ワークテーブル作成(変更時)
    ' =========================================
    If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        If wi登録方法_変更 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) _
        And 新規変更.Caption = "変更" Then
            wiRet = MsgBox("標準登録の明細データを移行しますか？", vbYesNo + vbQuestion)
            If wiRet = vbYes Then
                Call MXA040_借入明細移行_ワークテーブル作成(借入番号)
            End If
        End If
    End If
'
    ' =========================================
    '             テーブルにセット
    ' =========================================
    w借入金マスタ.借入番号 = 借入番号
    w借入金マスタ.借入内容 = 借入内容
    w借入金マスタ.借入金種別区分 = P8.FCStr(借入金種別.Text)
    w借入金マスタ.プロジェクト番号 = P8.FCStr(部門.Text)
    
    Select Case wFname  '06/02/01 V150
    Case "借入金登録"
        w借入金マスタ.借入貸付 = XMXA020_区分("借入貸付", "借入")
    Case "貸付登録"
        w借入金マスタ.借入貸付 = XMXA020_区分("借入貸付", "貸付")
    End Select
    
    'V180
    wi01 = P8.FCDbl(登録方法.Text)
    Select Case wi01
    Case P8.FCDbl(XMXA020_区分("登録方法", "標準登録"))
        w借入金マスタ.手入力区分 = wi01
    
    Case P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
                
        'If wi登録方法 <> wi01 Then
            GRet = 入力登録_残高CHECK
            Select Case GRet
            Case P8.FCDbl(XMXA020_区分("登録方法", "標準登録"))
            Case P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
            Case Else
                L_登録方法.Visible = True
            End Select
        'End If
        
        If L_登録方法.Visible <> True Then
            w借入金マスタ.手入力区分 = wi01
        Else
            w借入金マスタ.手入力区分 = 2
        End If
    End Select
'
    '
    w借入金マスタ.銀行番号 = P8.FCStr(銀行.Text)
    w借入金マスタ.返済単位月数 = P8.FCDbl(返済単位月数)
    
    w借入金マスタ.融資金額 = P8.FCDbl(融資金額)
    
    w借入金マスタ.金利種別 = P8.FCDbl(金利種別.ItemData(金利種別.ListIndex))
    w借入金マスタ.基準金利区分 = P8.FCStr(基準金利.Text)
    w借入金マスタ.利率 = P8.FCDbl(利率)
    w借入金マスタ.金利条件 = 金利条件
    
    w借入金マスタ.実行日 = w実行日
    w借入金マスタ.初回返済年月 = w初回返済年月
    'V188
    w借入金マスタ.初回返済実行日 = w初回返済実行日
    w借入金マスタ.最終返済年月 = w最終返済年月
    'V188
    w借入金マスタ.最終返済実行日 = w最終返済実行日
    w借入金マスタ.金利初回年月 = CDate(C年月日.平成To西暦("年月", C_金利初回年月.Text))
    
    'V180
    w借入金マスタ.解約年月 = MXA030_実行支払年月(w解約実行日, wi支払日, wi営業日, "")
    w借入金マスタ.解約実行日 = w解約実行日
    
    w借入金マスタ.借入計画番号 = P8.FCStr(借入計画番号)
    w借入金マスタ.金融リストラ番号 = P8.FCStr(金融リストラ番号)
    w借入金マスタ.金利グループ区分 = P8.FCStr(金利グループ区分.Text)
    
    'If 借入計画番号 <> "" Then
    '    w借入金マスタ.SM区分 = 1
    'Else
        If 金融リストラ番号 = "" Then
            w借入金マスタ.SM区分 = 0
        Else
            w借入金マスタ.SM区分 = SM区分
        End If
    'End If
    
    w借入金マスタ.金融解約年月 = C年月日.GetDate("月始", w金融解約実行日)
    w借入金マスタ.金融解約実行日 = w金融解約実行日
    
    w借入金マスタ.返済方法 = XMXA020_区分("返済方法", "元金均等返済")
    
    'w借入金マスタ.金融解約保証料戻
    w借入金マスタ.初回返済額 = P8.FCDbl(初回返済額)
    w借入金マスタ.毎月返済額 = P8.FCDbl(毎月返済額)
    w借入金マスタ.最終返済額 = P8.FCDbl(最終返済額)
    
    w借入金マスタ.長短区分 = P8.FCDbl(長短区分.ItemData(長短区分.ListIndex))
    w借入金マスタ.有担保フラグ = P8.FCDbl(担保区分.ItemData(担保区分.ListIndex))
    w借入金マスタ.担保名 = P8.FCStr(担保名)
    w借入金マスタ.資金用途 = P8.FCStr(資金用途)
    w借入金マスタ.設備フラグ = P8.FCDbl(担保区分.ItemData(設備区分.ListIndex))
    'w借入金マスタ.自己資金フラグ = 自己資金

    '********据置回数・支払回数********  2004/03/28
    'w銀行マスタ = MAA030_銀行マスタRead(w借入金マスタ.銀行番号)
    '
    'w支払日 = Day(C年月日.GetDate("月末", w借入金マスタ.実行日))
    'If w銀行マスタ.支払日 <> 31 Then
    '    w支払日 = w銀行マスタ.支払日
    'End If
    
    'V180
    w支払日 = Day(C年月日.GetDate("月末", w借入金マスタ.実行日))
    If wi支払日 <> 31 Then
        w支払日 = wi支払日
    End If
     
    'w実行支払年月 = MXA030_実行支払年月(w借入金マスタ.実行日, w支払日, "=")
    'V180
    w実行支払年月 = MXA030_実行支払年月(w借入金マスタ.実行日, w支払日, wi営業日, "*")   ' 07/02/22 V180
    
    w借入金マスタ.据置回数 = DateDiff("m", w実行支払年月, w借入金マスタ.初回返済年月)
 
    'w借入金マスタ.据置回数 = DateDiff("m", w実行日, w初回返済年月)
    w借入金マスタ.支払回数 = DateDiff("m", w初回返済年月, w最終返済年月) + 1
    'If Not IsNull(w借入金マスタ.解約年月) Then
    '    If w借入金マスタ.解約年月 < w初回返済年月 Then
    '        w借入金マスタ.据置回数 = DateDiff("m", w実行日, w借入金マスタ.解約年月) + 1
    '        w借入金マスタ.支払回数 = 0
    '
    '    ElseIf w借入金マスタ.解約年月 >= w初回返済年月 Then
    '        w借入金マスタ.支払回数 = DateDiff("m", w初回返済年月, w借入金マスタ.解約年月) + 1
    '    End If
    'End If
    '***  2004/03/28
    
    'w借入金マスタ.融資可能枠
    'w借入金マスタ.融資残高
    'w借入金マスタ.借入年度
    'w借入金マスタ.取消フラグ = 取消
    
    If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        'w借入金マスタ.金利種別 = 0
        'w借入金マスタ.金利条件 = ""
        'w借入金マスタ.利率 = 0
        w借入金マスタ.保証料率 = 0
        
'        w借入金マスタ.毎月返済額 = 0
    
        w借入金マスタ.金利初回年月 = w借入金マスタ.初回返済年月
        If P8.FCDbl(返済単位月数) < 1 Or P8.FCDbl(返済単位月数) > 12 Then
            w借入金マスタ.返済単位月数 = 1
        End If
        w借入金マスタ.解約実行日 = Null
        w借入金マスタ.解約年月 = Null
        w借入金マスタ.金融解約実行日 = Null
        w借入金マスタ.解約年月 = Null
    End If
'
    ' =========================================
    '           借入計画明細 作成処理
    ' =========================================
    'w借入金マスタ.解約保証料戻 = MBD010_借入金テーブル作成("", w借入金マスタ)
    
    'Call MBD010_借入明細作成処理
'
    ' =========================================
    '            借入金マスタ 更新処理
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & 借入番号 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            wRs.AddNew
            
            w借入金マスタ.借入番号 = LTrim(P8.FCStr(w借入金マスタ.借入番号))
            借入番号 = w借入金マスタ.借入番号
            wRs("借入番号") = w借入金マスタ.借入番号
        
            wslog = "追加"
        End If
     
            wRs("返済方法") = XMXA020_区分("返済方法", "元金均等返済")
            wRs("借入貸付") = w借入金マスタ.借入貸付
            wRs("借入金種別区分") = w借入金マスタ.借入金種別区分
            wRs("借入内容") = w借入金マスタ.借入内容
            wRs("プロジェクト番号") = w借入金マスタ.プロジェクト番号
            
            '登録方法
            wRs("手入力区分") = w借入金マスタ.手入力区分                'V180
            If w借入金マスタ.手入力区分 = CDbl(XMXA020_区分("登録方法", "標準登録")) Then
                wRs("日割計算区分") = CDbl(XMXA020_区分("日割計算区分", "自動計算"))
            Else
                wRs("日割計算区分") = CDbl(日割計算区分.Text)
            End If
            
            wRs("銀行番号") = w借入金マスタ.銀行番号
            wRs("支払日") = wi支払日
            wRs("返済単位月数") = w借入金マスタ.返済単位月数
            
            If FLG_New Then
                wRs("営業日区分") = wi営業日
                wRs("利息区分") = ws利息区分
                wRs("利息計算日数区分") = wi利息日数
                wRs("利息支払方法") = wi利息支払
                wRs("利息控除区分") = wi利息控除
                wRs("金利計算年間日数") = wi金利計算
            End If
            
            wRs("融資金額") = w借入金マスタ.融資金額
            
            wRs("金利種別") = w借入金マスタ.金利種別
            wRs("基準金利区分") = w借入金マスタ.基準金利区分
            wRs("利率") = w借入金マスタ.利率
            wRs("金利条件") = w借入金マスタ.金利条件
            wRs("金利グループ区分") = w借入金マスタ.金利グループ区分
            
            wRs("実行日") = w借入金マスタ.実行日
            wRs("初回返済年月") = w借入金マスタ.初回返済年月
            wRs("初回返済実行日") = w借入金マスタ.初回返済実行日
            wRs("金利初回年月") = w借入金マスタ.金利初回年月 'V180
            wRs("最終返済年月") = w借入金マスタ.最終返済年月
            wRs("最終返済実行日") = w借入金マスタ.最終返済実行日
            wRs("解約年月") = w借入金マスタ.解約年月
            wRs("解約実行日") = w借入金マスタ.解約実行日
            'w借入金マスタ.解約保証料戻
            
            wRs("借入計画番号") = P8.FCStr(w借入金マスタ.借入計画番号)
            wRs("Sm区分") = w借入金マスタ.SM区分
            wRs("金融リストラ番号") = P8.FCStr(w借入金マスタ.金融リストラ番号)
            wRs("金融解約年月") = w借入金マスタ.金融解約年月
            wRs("金融解約実行日") = w借入金マスタ.金融解約実行日
            wRs("金融解約保証料戻") = w借入金マスタ.金融解約保証料戻
                        
            wRs("初回返済額") = w借入金マスタ.初回返済額
            wRs("毎月返済額") = w借入金マスタ.毎月返済額
            wRs("最終返済額") = w借入金マスタ.最終返済額
            
            wRs("長短区分") = w借入金マスタ.長短区分
            wRs("有担保フラグ") = w借入金マスタ.有担保フラグ
            wRs("担保名") = w借入金マスタ.担保名
            wRs("資金用途") = w借入金マスタ.資金用途
            wRs("設備フラグ") = w借入金マスタ.設備フラグ
            'wRs("自己資金フラグ") = w借入金マスタ.自己資金フラグ
            wRs("据置回数") = w借入金マスタ.据置回数
            wRs("支払回数") = w借入金マスタ.支払回数
                    
            'w借入金マスタ.融資可能枠
            'w借入金マスタ.融資残高
            'w借入金マスタ.借入年度
 
             wRs("取消フラグ") = 0 'w借入金マスタ.取消フラグ
             
            '保証会社
            wRs("保証料率") = P8.FCDbl(保証料率)
            wRs("保証料分割フラグ") = P8.FCDbl(保証料分割)
            wRs("自己資金フラグ") = P8.FCDbl(自己資金)
            wRs("保証会社区分") = P8.FCStr(保証会社区分.Text)
            wRs("融資区分") = P8.FCStr(融資区分.Text)
            
            ' =========================================
            '   登録方法：標準登録→入力登録 明細データ作成
            ' =========================================
            If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
                If wi登録方法_変更 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) _
                And 新規変更.Caption = "変更" Then
                    If wiRet = vbYes Then
                        wRs("手入力区分") = P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
                        'wRs("日割計算区分") = CDbl(XMXA020_区分("日割計算区分", "自動計算"))
                    End If
                End If
            End If
                
        wRs.Update
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    If wslog <> "追加" Then
        wslog = "更新"
    End If
    
    GLogStr = "借入番号=" & w借入金マスタ.借入番号 & ","
    GLogStr = GLogStr & "借入内容=" & w借入金マスタ.借入内容 & ","
    GLogStr = GLogStr & "借入金種別区分=" & w借入金マスタ.借入金種別区分 & ","
    GLogStr = GLogStr & "部門=" & w借入金マスタ.プロジェクト番号 & ","
    GLogStr = GLogStr & "手入力区分=" & w借入金マスタ.手入力区分 & ","
    GLogStr = GLogStr & "日割計算区分=" & CDbl(XMXA020_区分("日割計算区分", "自動計算")) & ","
    GLogStr = GLogStr & "借入計画番号=" & w借入金マスタ.借入計画番号 & ","
    
    GLogStr = GLogStr & "銀行番号=" & w借入金マスタ.銀行番号 & ","
    GLogStr = GLogStr & "支払日=" & wi支払日 & ","
    GLogStr = GLogStr & "返済単位月数=" & w借入金マスタ.返済単位月数 & ","
    If wslog = "追加" Then
        GLogStr = GLogStr & "営業日区分=" & wi営業日 & ","
        GLogStr = GLogStr & "利息区分=" & ws利息区分 & ","
        GLogStr = GLogStr & "利息計算日数区分=" & wi利息日数 & ","
        GLogStr = GLogStr & "利息支払方法=" & wi利息支払 & ","
        GLogStr = GLogStr & "利息控除区分=" & wi利息控除 & ","
        GLogStr = GLogStr & "金利計算年間日数=" & wi金利計算 & ","
    End If
    GLogStr = GLogStr & "融資金額=" & w借入金マスタ.融資金額 & ","
    
    GLogStr = GLogStr & "金利種別=" & w借入金マスタ.金利種別 & ","
    GLogStr = GLogStr & "基準金利区分=" & w借入金マスタ.基準金利区分 & ","
    GLogStr = GLogStr & "利率=" & w借入金マスタ.利率 & ","
    GLogStr = GLogStr & "金利条件=" & w借入金マスタ.金利条件 & ","
    
    GLogStr = GLogStr & "実行日=" & w借入金マスタ.実行日 & ","
    GLogStr = GLogStr & "初回返済年月=" & w借入金マスタ.初回返済年月 & ","
    GLogStr = GLogStr & "初回返済年月日=" & w借入金マスタ.初回返済実行日 & ","
    GLogStr = GLogStr & "金利初回年月=" & w借入金マスタ.金利初回年月 & ","
    GLogStr = GLogStr & "最終返済年月=" & w借入金マスタ.最終返済年月 & ","
    GLogStr = GLogStr & "最終返済年月日=" & w借入金マスタ.最終返済実行日 & ","
    GLogStr = GLogStr & "解約年月日=" & w借入金マスタ.解約実行日 & ","
    
    GLogStr = GLogStr & "初回返済額=" & w借入金マスタ.初回返済額 & ","
    GLogStr = GLogStr & "毎月返済額=" & w借入金マスタ.毎月返済額 & ","
    GLogStr = GLogStr & "最終返済額=" & w借入金マスタ.最終返済額 & ","
    
    GLogStr = GLogStr & "長短区分=" & w借入金マスタ.長短区分 & ","
    GLogStr = GLogStr & "有担保フラグ=" & w借入金マスタ.有担保フラグ & ","
    GLogStr = GLogStr & "担保名=" & w借入金マスタ.担保名 & ","
    GLogStr = GLogStr & "資金用途=" & w借入金マスタ.資金用途 & ","
    GLogStr = GLogStr & "設備フラグ=" & w借入金マスタ.設備フラグ & ","
    
    GLogStr = GLogStr & "保証会社区分=" & P8.FCStr(保証会社区分.Text) & ","
    GLogStr = GLogStr & "融資区分=" & P8.FCStr(融資区分.Text)
    GLogStr = GLogStr & "保証料率=" & P8.FCDbl(保証料率) & ","
    GLogStr = GLogStr & "保証料分割フラグ=" & P8.FCDbl(保証料分割) & ","
    GLogStr = GLogStr & "自己資金フラグ=" & P8.FCDbl(自己資金) & ","
    
    GLogStr = GLogStr & "金融リストラ番号=" & w借入金マスタ.金融リストラ番号 & ","
    GLogStr = GLogStr & "Sm区分=" & w借入金マスタ.SM区分 & ","
    GLogStr = GLogStr & "金融解約年月日=" & w借入金マスタ.金融解約実行日 & ","
    GLogStr = GLogStr & "金利グループ区分=" & w借入金マスタ.金利グループ区分
    
    Call MXA030_LOG_WRITE("借入金登録", wslog, GLogStr)
'
    ''----------< TR削除フラグ >----------------------------------------------
    'If w借入金マスタ.手入力区分 <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
    '    Call 取消2_明細TR(w借入金マスタ.取消フラグ)
    'End If
'
    ' =========================================
    '   登録方法：標準登録→入力登録 明細データ作成
    ' =========================================
    If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        If wi登録方法_変更 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) _
        And 新規変更.Caption = "変更" Then
            If wiRet = vbYes Then
                Call MXA040_借入明細移行(借入番号)
            End If
        End If
    End If
'
    '----------------------------------------
    '            金融リストラ番号セット
    '----------------------------------------
    金融リストラ番号.Clear
    金融リストラ番号.AddItem ""
    
    wstr = ""
    wstr = wstr + "Select 金融リストラ番号"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 金融リストラ番号 <> '' "
    wstr = wstr + " Group By 金融リストラ番号"
    wstr = wstr + " Order By 金融リストラ番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.EOF
            金融リストラ番号.AddItem (P8.FCStr(wRs("金融リストラ番号")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    '----------------------------------------
    '            資金用途セット
    '----------------------------------------
    資金用途.Clear
    資金用途.AddItem ""
    
    wstr = ""
    wstr = wstr + "Select 資金用途"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 資金用途 <> ''"
    wstr = wstr + " Group By 資金用途"
    wstr = wstr + " Order By 資金用途"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.EOF
            資金用途.AddItem (P8.FCStr(wRs("資金用途")))
                         
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    'Call 登録後初期セット
    Call CEkey.SetFs(借入番号, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました。", vbInformation
'
    '----------< Button制御 >-------------------------------------------------------
    登録.Enabled = True
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
' 銀行詳細_Click
'------------------------------------------------
Private Sub 銀行詳細_Click()
'
    If P8.FCStr(借入番号.Text) = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        Exit Sub
    End If

    If FLG_New = True Then
        MsgBox "登録処理を行ってください"
        Exit Sub
    End If
'
    If P8.FCStr(借入番号.Text) = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        Exit Sub
    End If
'
    GStr = wFname
    GStr_2 = P8.FCStr(借入番号.Text)
    GStr_3 = ""
        
    DoEvents
'
'    Unload Me
    Me.Enabled = False
    frm_I借入金登録_銀行.Show
'
End Sub

'------------------------------------------------
' 金利変更_Click
'------------------------------------------------
Private Sub 金利変更_Click()
'
    If P8.FCStr(借入番号.Text) = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        Exit Sub
    End If
'
    If FLG_New = True Then
        MsgBox "登録処理を行ってください"
        Exit Sub
    End If
    
    GStr = wFname
    GStr_2 = P8.FCStr(借入番号.Text)
    GStr_3 = ""
    
    DoEvents
'
'    Unload Me
        
    frm_I借入金登録_金利変更.Show
    Me.Enabled = False
'
End Sub

'------------------------------------------------
' 明細入力_Click
'------------------------------------------------
Private Sub 明細入力_Click()
'
    If P8.FCStr(借入番号.Text) = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        Exit Sub
    End If
'
    If FLG_New = True Then
        MsgBox "登録処理を行ってください"
        Exit Sub
    End If
'
    GStr = wFname
    GStr_2 = P8.FCStr(借入番号.Text)
    GStr_3 = "明細入力"
    
    DoEvents
'
'    Unload Me
        
    Unload frm_F借入登録データ照会
    Unload frm_F借入金明細表
    
    frm_I借入金登録_明細.Show
    Me.Enabled = False
'
End Sub

'------------------------------------------------
' 内入入力_Click
'------------------------------------------------
Private Sub 内入入力_Click()
'
    If P8.FCStr(借入番号.Text) = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        Exit Sub
    End If
'
    If FLG_New = True Then
        MsgBox "登録処理を行ってください"
        Exit Sub
    End If
'
    GStr = wFname
    GStr_2 = P8.FCStr(借入番号.Text)
    GStr_3 = "内入入力"

    DoEvents
'
'    Unload Me

    frm_I借入金登録_内入.Show
    Me.Enabled = False
'
End Sub

'------------------------------------------------
' 取消2_明細TR
'------------------------------------------------
Private Sub 取消2_明細TR(p取消 As Integer)
'
    Dim wstr As String
'
    Select Case wFname
    Case "借入金登録"
        wstr = ""
        wstr = wstr & "Delete * "
        wstr = wstr & " From DBDA010_借入金明細TR"
        wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
        GDb.Execute wstr
        
        wstr = ""
        wstr = wstr & "Delete * "
        wstr = wstr & " From DBDA010_借入金明細TR2"
        wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
        GDb.Execute wstr
    Case "貸付登録"
        wstr = ""
        wstr = wstr & "Delete * "
        wstr = wstr & " From DBDA010_貸付金明細TR"
        wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
        GDb.Execute wstr
    End Select
    
    DoEvents
'
End Sub

'------------------------------------------------
' 削除_Click
'------------------------------------------------
Private Sub 削除_Click()
'
    Dim FLG_DEL As Boolean
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

    If P8.FCStr(借入番号.Text) = "" Then
        Exit Sub
    End If
'
    GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
    If GRet = vbNo Then
        Exit Sub
    End If
'
    FLG_DEL = False
    If P8.FCDbl(登録方法.Text) <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        FLG_DEL = True
    End If
'
    wstr = ""
    wstr = wstr & "Delete * From " & wsTbl
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
    
    DoEvents
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号=" & P8.FCStr(借入番号.Text)
    Call MXA030_LOG_WRITE("借入金登録", "削除", GLogStr)
'
    '----------< TR削除フラグ >----------------------------------------------
    If FLG_DEL = True Then
        Call 取消2_明細TR(1)
    End If
'
    '----------< 内入削除 >------------------------------------------------
    wstr = "Delete * from DBDA010_借入金内入1"
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入2"
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入3"
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入4"
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入5"
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入6"
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入7"
    wstr = wstr & " Where 借入番号='" & P8.FCStr(借入番号.Text) & "'"
    GDb.Execute wstr
'
    借入番号.Text = ""
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 登録後初期セット
    Call CEkey.SetFs(借入番号, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "削除しました。", vbInformation
'
End Sub

'------------------------------------------------
' 入力登録_残高CHECK
'------------------------------------------------
Private Function 入力登録_残高CHECK() As Integer
'
    Dim wd01 As Double, wiCnt As Integer
'
    wstr = ""
    wstr = wstr + "Select 融資残高"
    wstr = wstr + " From " & wsTbl2
    wstr = wstr + " Where 借入番号 = '" & P8.FCStr(借入番号.Text) & "'"
    wstr = wstr + " And 取消フラグ=0"
    wstr = wstr + " Order by 実際年月日 desc"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    wiCnt = wRs.RecordCount
    If Not wRs.EOF Then
        wd01 = P8.FCDbl(wRs("融資残高"))
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    If wiCnt = 0 Then
        入力登録_残高CHECK = 2
    
        Exit Function
    End If
'
    If wd01 <> 0 Then
        入力登録_残高CHECK = 2
    Else
        入力登録_残高CHECK = P8.FCDbl(XMXA020_区分("登録方法", "入力登録"))
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
入力登録_残高CHECK_ERR:
    pERR_MES = pPROGRAM_ID + "/ 入力登録_残高CHECK() でエラー" + vbCrLf + vbCrLf + _
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
' 明細書表示_Click
'------------------------------------------------
Private Sub 明細書表示_Click()
    
    G借入明細表照会.借入番号 = P8.FCStr(借入番号)
    G借入明細表照会.金融リストラ番号 = P8.FCStr(金融リストラ番号)
    G借入明細表照会.金融解約日 = P8.FCStr(金融解約日)
    G借入明細表照会.金利シミュレーションGP = P8.FCStr(金利グループ区分.Text)
    
    If G借入明細表照会.借入番号 = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        Exit Sub
    End If
    
    If wi登録方法 = 0 And P8.FCStr(金利種別) = "固定金利" Then
        G借入明細表照会.入力モード = 0
    ElseIf wi登録方法 = 0 And P8.FCStr(金利種別) = "変動金利" Then
        G借入明細表照会.入力モード = 1
    ElseIf wi登録方法 = 1 Then
        G借入明細表照会.入力モード = 0
    End If
    
    frm_F借入金明細表.Show
    Me.Enabled = False
End Sub

'------------------------------------------------
' CSV取込_Click
'------------------------------------------------
Private Sub CSV取込_Click()
'
    Dim w最終実績年月 As Date
    Dim w実績年月 As String
    Dim ws01 As String, wsRet As String
    Dim wi01 As Integer
'
    On Error GoTo 取込_Click_ERR
'
    GRet = MsgBox("CSVファイルをインポートします。" _
                    & vbCrLf & vbCrLf & "既存のデータは上書きされます。" & vbCrLf & _
                    "よろしいですか？", vbExclamation + vbYesNo, "取込")
    If GRet = vbNo Then
        Exit Sub
    End If
'
    wsRet = MXA040_COMDLG(CommonDialog1, "CSVファイル選択", "", _
                        "テキストファイル(*.csv)|*.csv", "借入明細表.csv")
    If wsRet = "" Then
        Exit Sub
    ElseIf wsRet = "キャンセル" Then
        Exit Sub
    End If
'
    GRet = MXA040_借入明細取込(wsRet, "")
    If GRet <> True Then
        MsgBox "CSVファイルをインポートできませんでした", vbInformation
        
        Exit Sub
    End If
'
    ' =========================================
    '               画面セット
    ' =========================================
    借入番号 = GStr_1
    Call 画面セット(False)
    'Call 登録後初期セット
    Call CEkey.SetFs(借入番号, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "CSVファイルをインポートしました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
取込_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 取込_Click() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Public Sub 画面セット呼出()

    Call 画面セット(False)

End Sub


'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    ReDim G借入金入力(0)
'
    Unload Me
End Sub

'------------------------------------------------
' Check_KARIIRENO
'------------------------------------------------
Private Function Check_KARIIRENO(pKririeNo As String) As Boolean
    Dim wRet As Boolean
'
    On Error GoTo Check_KARIIRENO_ERR
'
    wRet = False
    
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DBDA010_借入金"
    wstr = wstr + " Where 借入番号 = '" & pKririeNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.EOF Then
        wRet = True
    End If
    wRs.Close
    Set wRs = Nothing
    
    Check_KARIIRENO = wRet
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
Check_KARIIRENO_ERR:
    pERR_MES = pPROGRAM_ID + "/ Check_KARIIRENO() でエラー" + vbCrLf + vbCrLf + _
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
' Copy_KARIIRENO
'------------------------------------------------
Private Function Copy_KARIIRENO(pMotoNo As String, pNewNo As String) As Boolean
    
    Dim ws01 As String, ws02 As String
    Dim i As Integer
'
    On Error GoTo Copy_KARIIRENO_ERR
'
    Copy_KARIIRENO = False
    
    'DBDA010_借入金
    ws01 = ""
    ws01 = ws01 & ",プロジェクト番号,保証会社区分,融資区分,手入力区分,日割計算区分"
    ws01 = ws01 & ",借入内容,借入計画番号,sm区分,金融リストラ番号"
    ws01 = ws01 & ",銀行番号,支払日,営業日区分,利息区分,利息計算日数区分,利息支払方法"
    ws01 = ws01 & ",利息控除区分,金利計算年間日数"
    ws01 = ws01 & ",融資金額,利率,保証料率,保証料分割フラグ"
    ws01 = ws01 & ",実行日,初回返済年月,初回返済実行日,金利初回年月,最終返済年月,最終返済実行日"
    ws01 = ws01 & ",解約年月,解約実行日,解約保証料戻,金融解約年月,金融解約実行日,金融解約保証料戻"
    ws01 = ws01 & ",返済方法,借入貸付,借入金種別区分"
    ws01 = ws01 & ",初回返済額,毎月返済額,最終返済額,返済単位月数"
    ws01 = ws01 & ",有担保フラグ,担保名,設備フラグ,資金用途,自己資金フラグ,長短区分"
    ws01 = ws01 & ",支払回数,据置回数,金利種別,金利条件,基準金利区分,金利グループ区分"
    
    ws02 = ""
    For i = 2 To 100
        ws02 = ",金利変更" & CStr(i) & "回目年月,金利" & CStr(i) & "回目"
        ws01 = ws01 & ws02
    Next i
    
    ws01 = ws01 & ",融資可能枠,融資残高,借入年度,取消フラグ"

    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01 & ")"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & " From DBDA010_借入金"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents
    
    'DBDA010_借入金明細TR
    ws01 = ",返済回数,据置X回目,返済予定年月,実際年月日,利息計算年月日"
    ws01 = ws01 & ",返済金額,元金額,利息額,仮計上利息額"
    ws01 = ws01 & ",保証料,金融保証料,手数料,初期手数料,元金手数料,利息手数料"
    ws01 = ws01 & ",融資残高,日割日数,利息対象期間日数,利率"
    ws01 = ws01 & ",取消フラグ,取消フラグ２"
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金明細TR"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01 & ")"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & " From DBDA010_借入金明細TR"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents
    
    'DBDA010_借入金明細TR2
    ws01 = ",返済予定年月,実際年月日"
    ws01 = ws01 & ",保証料,初期手数料,元金手数料,利息手数料"
    ws01 = ws01 & ",取消フラグ,取消フラグ２"
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金明細TR2"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01 & ")"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & " From DBDA010_借入金明細TR2"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents
    
    'DBDA010_借入金内入
    ws01 = ""
    For i = 1 To 40
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,毎回支払額" & CStr(i) & "回目"
        ws01 = ws01 & ",最終支払額" & CStr(i) & "回目,最終支払年月" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents

    'DBDA010_借入金内入1
    ws01 = ""
    For i = 1 To 80
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入1"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入1"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents

    'DBDA010_借入金内入2
    ws01 = ""
    For i = 81 To 160
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入2"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入2"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents

    'DBDA010_借入金内入3
    ws01 = ""
    For i = 161 To 240
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入3"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入3"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents

    'DBDA010_借入金内入4
    ws01 = ""
    For i = 241 To 320
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入4"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入4"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents

    'DBDA010_借入金内入5
    ws01 = ""
    For i = 321 To 400
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入5"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入5"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents

    'DBDA010_借入金内入6
    ws01 = ""
    For i = 401 To 480
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入6"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入6"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents

    'DBDA010_借入金内入7
    ws01 = ""
    For i = 481 To 560
        ws01 = ws01 & ",内入" & CStr(i) & "回目年月日,内入金額" & CStr(i) & "回目,手数料" & CStr(i) & "回目"
    Next i
    
    wstr = ""
    wstr = "INSERT INTO DBDA010_借入金内入7"
    wstr = wstr & " (借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ)"
    wstr = wstr & " Select "
    wstr = wstr & " '" & pNewNo & "' As 借入番号"
    wstr = wstr & ws01
    wstr = wstr & ",取消フラグ"
    wstr = wstr & " From DBDA010_借入金内入7"
    wstr = wstr & " Where 借入番号='" & pMotoNo & "'"
    GDb.Execute wstr

    DoEvents
    
    Copy_KARIIRENO = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
Copy_KARIIRENO_ERR:
    pERR_MES = pPROGRAM_ID + "/ Copy_KARIIRENO() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

