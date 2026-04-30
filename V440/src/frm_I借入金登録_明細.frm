VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_I借入金登録_明細 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入登録金 明細入力"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   Icon            =   "frm_I借入金登録_明細.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "登録"
      Height          =   4215
      Left            =   120
      TabIndex        =   41
      Top             =   5300
      Width           =   12615
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
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox 調整日数 
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
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   4
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox 日割日数 
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
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   3
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox 調整利息額 
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
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   5520
         MaxLength       =   16
         TabIndex        =   7
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox 年月日2 
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
         IMEMode         =   2  'ｵﾌ
         Left            =   1560
         TabIndex        =   1
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox 元金額 
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
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   5520
         MaxLength       =   16
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox 利息額 
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
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   5520
         MaxLength       =   16
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox 年月日 
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
         IMEMode         =   2  'ｵﾌ
         Left            =   1560
         TabIndex        =   0
         Top             =   1080
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
         Left            =   10800
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton 登録 
         Caption         =   "登録"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   51
         Top             =   3600
         Width           =   1695
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
         Left            =   7320
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton 登録データ照会 
         Caption         =   "登録データ照会"
         Height          =   375
         Left            =   3840
         TabIndex        =   52
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton 利息額再計算 
         Caption         =   "利息額再計算"
         Height          =   375
         Left            =   6720
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton 利息額再計算ALL 
         Caption         =   "選択返済日以降の利息額再計算"
         Height          =   495
         Left            =   240
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3015
      End
      Begin VB.CommandButton CSV取込 
         Caption         =   "CSV取込"
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
         Left            =   5640
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
      End
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
         Left            =   3960
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Frame Frame_社債 
         Caption         =   "社債"
         Height          =   2775
         Left            =   8520
         TabIndex        =   42
         Top             =   360
         Width           =   3975
         Begin VB.TextBox 保証料 
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
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   14
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox 利息手数料 
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
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   12
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox 初期手数料 
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
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   10
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox 元金手数料 
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
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   11
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label L_支払計 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   "支払計"
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
            Left            =   240
            TabIndex        =   71
            Top             =   2160
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label 支払計 
            Alignment       =   1  '右揃え
            BackColor       =   &H80000000&
            BorderStyle     =   1  '実線
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
            Left            =   1560
            TabIndex        =   15
            Top             =   2160
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label L_手数料計 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   "手数料計"
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
            Left            =   240
            TabIndex        =   70
            Top             =   1440
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label 手数料計 
            Alignment       =   1  '右揃え
            BackColor       =   &H80000000&
            BorderStyle     =   1  '実線
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
            Left            =   1560
            TabIndex        =   13
            Top             =   1440
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label L_保証料 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   "保証料"
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
            Left            =   240
            TabIndex        =   46
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label L_利息手数料 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   "利息額手数料"
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
            Left            =   240
            TabIndex        =   45
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label L_初期手数料 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   "初期手数料"
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
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label L_元金手数料 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   "元金額手数料"
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
            Left            =   240
            TabIndex        =   43
            Top             =   720
            Width           =   1335
         End
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   240
         TabIndex        =   55
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
      Begin VB.Label L_利率 
         Alignment       =   1  '右揃え
         BorderStyle     =   1  '実線
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
         Left            =   1560
         TabIndex        =   69
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label L1_利率 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "利率"
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
         Left            =   240
         TabIndex        =   68
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label L1_調整日数 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "調整日数"
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
         Left            =   240
         TabIndex        =   67
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label L1_日割日数 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "日割日数"
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
         Left            =   240
         TabIndex        =   66
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label L1_調整利息額 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "調整利息額"
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
         Left            =   4200
         TabIndex        =   65
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label L1_年月日2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "利息計算日"
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
         Left            =   240
         TabIndex        =   64
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label L_元金額 
         Alignment       =   1  '右揃え
         BorderStyle     =   1  '実線
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
         Left            =   5520
         TabIndex        =   63
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label L_返済金額 
         Alignment       =   1  '右揃え
         BackColor       =   &H80000000&
         BorderStyle     =   1  '実線
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
         Left            =   5520
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label L_融資残高 
         Alignment       =   1  '右揃え
         BackColor       =   &H80000000&
         BorderStyle     =   1  '実線
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
         Left            =   5520
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label L1_1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "支払元金額"
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
         Left            =   4200
         TabIndex        =   62
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label L1_2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "支払利息額"
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
         Left            =   4200
         TabIndex        =   61
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label L1_年月日 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "返済日"
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
         Left            =   240
         TabIndex        =   60
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label L1_3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "返済金額"
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
         Left            =   4200
         TabIndex        =   59
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label L1_4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "融資残高"
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
         Left            =   4200
         TabIndex        =   58
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label L_番号 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
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
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label L_番号1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '実線
         Caption         =   "借入番号"
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
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "（単位：円）"
         Height          =   375
         Left            =   7200
         TabIndex        =   56
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   240
      Top             =   9000
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5150
      Left            =   120
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9075
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   13
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
         Size            =   8.25
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label L_合計調整利息額 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
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
      Left            =   9000
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_合計利息額 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
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
      Left            =   9000
      TabIndex        =   22
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_合計元金額 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
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
      Left            =   9000
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_合計融資残高 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
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
      Left            =   9000
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_融資金額 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   1920
      TabIndex        =   37
      Top             =   5760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_最終返済年月 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   1920
      TabIndex        =   33
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_初回返済年月 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   1920
      TabIndex        =   32
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_利息区分 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   1920
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label L_実行日 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   1920
      TabIndex        =   31
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "円"
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
      Left            =   4560
      TabIndex        =   39
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "融資金額"
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
      Left            =   0
      TabIndex        =   38
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label L_合計融資残高1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "融資残高"
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
      Left            =   6960
      TabIndex        =   35
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label L_合計融資残高2 
      Alignment       =   2  '中央揃え
      Caption         =   "円"
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
      Left            =   11640
      TabIndex        =   34
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "最終返済年月"
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
      Left            =   0
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label18 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "初回返済年月"
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
      Left            =   0
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label19 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "実行日"
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
      Left            =   0
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label L_合計元金額2 
      Alignment       =   2  '中央揃え
      Caption         =   "円"
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
      Left            =   11640
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label L_合計元金額1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "合計元金額"
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
      Left            =   6960
      TabIndex        =   26
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label L_合計利息額2 
      Alignment       =   2  '中央揃え
      Caption         =   "円"
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
      Left            =   11640
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label L_合計利息額1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "合計利息額"
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
      Left            =   6960
      TabIndex        =   23
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label L_合計調整利息額2 
      Alignment       =   2  '中央揃え
      Caption         =   "円"
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
      Left            =   11640
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label L_合計調整利息額1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "合計調整利息額"
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
      Left            =   6960
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "利息区分"
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
      Left            =   0
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frm_I借入金登録_明細"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_I借入金登録_明細"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim w借入データ As MAA910_借入金

Dim wFname As String, wFname2 As String
Dim wsTbl As String, wsTbl2 As String, wsTbl3 As String
Dim w初回 As Variant, w最終 As Variant
Dim wsBango As String

Dim wd前月残高 As Double, wd当月残高 As Double
Dim wv最新日 As Variant

Dim wd利率 As Double, wd融資残高 As Double
Dim wi日割日数 As Integer

Dim wi据置X回目 As Integer
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
'    Me.Caption = GFcap
    Me.Left = G_LEFT
    Me.Top = G_TOP
'
    wFname = GStr
    wFname2 = GStr_3
    'ZU050_Button1.Caption = wFname & wFname2 & Space(1) & "登録"
    
    wsBango = GStr_2
    
    L_番号.Caption = wsBango
    G借入明細表照会.借入番号 = wsBango
    
    登録.Caption = "金額設定" & vbCr & "登録"
        
    Select Case wFname
    Case "借入金登録"
        L_番号1.Caption = " 借入番号"
        
        wsTbl = "DBDA010_借入金明細TR"
        wsTbl2 = "DBDA010_借入金"
        
    Case "貸付登録"
        L_番号1.Caption = " 貸付番号"
    
        wsTbl = "DBDA010_貸付金明細TR"
        wsTbl2 = "DBDA010_貸付金"
    
    End Select
    
    GStr = "": GStr_1 = "": GStr_2 = ""
    GStr_3 = ""
'
    ' =========================================
    '                 初期設定
    ' =========================================
    w借入データ.借入番号 = wsBango
    
    w借入データ.実行日 = Null
    w借入データ.初回返済実行日 = Null
    w借入データ.最終返済実行日 = Null
    w借入データ.解約実行日 = Null

    w借入データ.金利種別 = 0
    w借入データ.利率 = 0
    
    L_利息区分.Caption = ""
    L_実行日.Caption = ""
    L_初回返済年月.Caption = ""
    L_最終返済年月.Caption = ""
    L_融資金額.Caption = ""
    '取消 = 0
    
    w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算"))

    L_利率.Caption = ""
    L_利率.Visible = False
    利率.Text = ""
    利率.Visible = True
'
    wstr = ""
    wstr = wstr + "Select * From " & wsTbl2
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
    
        w借入データ = MBD010_借入データセット(wRs)
        
        L_実行日.Caption = Format(P8.FCStr(wRs("実行日")), Gfmt年月日)
        L_初回返済年月.Caption = Format(P8.FCStr(wRs("初回返済年月")), Gfmt年月)
        L_最終返済年月.Caption = Format(P8.FCStr(wRs("最終返済年月")), Gfmt年月)
        L_融資金額.Caption = Format(w借入データ.融資金額, "#,##0")
                        
        If w借入データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
            L_利息区分.Caption = "利息先払"
        Else
            L_利息区分.Caption = "利息後払"
        End If
        
        If w借入データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
            wi据置X回目 = 3
        Else
            wi据置X回目 = 1
        End If
        
    End If
    wRs.Close
    Set wRs = Nothing
'
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA020_借入金明細"
    GDb.Execute wstr
'
    '** 明細ファイル 作成 **
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        Call MBD010_借入金入力明細_利息額自動計算(w借入データ, wsTbl)
    Else
        Call MBD010_借入金入力明細Read(w借入データ)
    End If
            
    If w借入データ.社債フラグ = 1 Then
        Call MDA020_借入金入力社債明細作成(w借入データ)
    End If
    
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) _
    Or w借入データ.社債フラグ = 1 Then
        Call MBD010_借入明細作成_入力登録(w借入データ)
    End If
'
'            Call MBD010_借入金入力明細Read(w借入データ)
'            If w借入データ.社債フラグ = 1 Then
'                Call MDA020_借入金入力社債明細作成(w借入データ)
'            End If
'            Call MBD010_借入明細作成_入力登録(w借入データ)
'
'
    '合計
    L_合計融資残高.Caption = ""
    
    wstr = ""
    wstr = wstr + "Select 融資残高"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And 取消フラグ=0"
    wstr = wstr + " Order by 実際年月日 desc"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        L_合計融資残高.Caption = Format(P8.FCDbl(wRs("融資残高")), "#,##0")
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    '合計集計
    Call 合計集計_セット
'
    L_元金額.Caption = ""
    L_返済金額.Caption = ""
    L_融資残高.Caption = ""
    
    年月日.Text = ""
    年月日2.Text = ""
    元金額.Text = ""
    利息額.Text = ""
    調整利息額.Text = ""
    日割日数.Text = ""
    調整日数.Text = ""
    利率.Text = ""
'
    L1_調整利息額.Visible = False
    L1_日割日数.Visible = False
    L1_調整日数.Visible = False
    調整利息額.Visible = False
    日割日数.Visible = False
    調整日数.Visible = False
    
    利息額再計算.Visible = True
    利息額再計算ALL.Visible = True
    
    '日割計算区分=自動計算の場合日数、調整額は表示しない。
    If w借入データ.日割計算区分 <> CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        L1_調整利息額.Visible = True
        L1_日割日数.Visible = True
        L1_調整日数.Visible = True
        調整利息額.Visible = True
        日割日数.Visible = True
        調整日数.Visible = True
    
        利息額再計算.Visible = False
        利息額再計算ALL.Visible = False
    End If
'
    '社債経費項目
    Frame_社債.Visible = False
    初期手数料.Visible = False
    元金手数料.Visible = False
    利息手数料.Visible = False
    手数料計.Visible = False
    保証料.Visible = False
    支払計.Visible = False
    L_初期手数料.Visible = False
    L_元金手数料.Visible = False
    L_利息手数料.Visible = False
    L_手数料計.Visible = False
    L_保証料.Visible = False
    L_支払計.Visible = False
    If w借入データ.社債フラグ = 1 Then
        Frame_社債.Visible = True
        初期手数料.Visible = True
        元金手数料.Visible = True
        利息手数料.Visible = True
        手数料計.Visible = True
        保証料.Visible = True
        支払計.Visible = True
        L_初期手数料.Visible = True
        L_元金手数料.Visible = True
        L_利息手数料.Visible = True
        L_手数料計.Visible = True
        L_保証料.Visible = True
        L_支払計.Visible = True
    End If
'
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
End Sub

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
'    Unload frm_I借入金登録
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
    
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
    GWhere = GWhere & " And M.借入番号='" & wsBango & "'"
    
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    
    'wstr = wstr + " IIF(取消フラグ２=0,1,2),"
    'wstr = wstr + " IIF(取消フラグ=0,1,2),"
    
    wstr = wstr + " M.借入番号,M.実際年月日,M.元金額,M.利息額,M.返済金額,M.融資残高,M.取消フラグ,"
    wstr = wstr + " M.返済回数 As Grd返済回数,"
    wstr = wstr + " Format(M.実際年月日,'" & Gfmt年月日 & "') As Grd年月日,"
    wstr = wstr + " Format(M.利息計算年月日,'" & Gfmt年月日 & "') As Grd利息計算日,"
    
    wstr = wstr + " Format(M.元金額,'#,##0') As Grd元金額,"
    wstr = wstr + " Format(M.利息額,'#,##0') As Grd利息額,"
    
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        'wstr = wstr + " Format(W.仮計上利息額,'#,##0') As Grd計算利息額,"
        wstr = wstr + " Format(M.利息額-W.仮計上利息額,'#,##0') As Grd利息差異,"
    End If
    
    wstr = wstr + " Format(M.返済金額,'#,##0') As Grd返済金額,"
    wstr = wstr + " Format(M.融資残高,'#,##0') As Grd融資残高,"
    wstr = wstr + " Format(M.利率,'#,##0.00000') As Grd利率,"
    wstr = wstr + " M.日割日数 As Grd日割日数,"
    wstr = wstr + " M.利息対象期間日数 As Grd調整日数"
    
    'wstr = wstr + " IIF(M.取消フラグ = 0,'','×') As Grd取消,"
    'wstr = wstr + " IIF(M.取消フラグ２ = 0,'','×') As Grd取消2"
    wstr = wstr + " From " & wsTbl & " As M"
    
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        wstr = wstr + " INNER JOIN DCDA020_借入金明細 As W"
        wstr = wstr + " ON (M.利息計算年月日 = W.利息計算年月日)"
        wstr = wstr + " AND (M.借入番号 = W.借入番号)"
    End If
    
    wstr = wstr + GWhere
    wstr = wstr + " Order By 1,2,M.実際年月日"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("返済回数", "回", 400, "C")
        Call XZMA010_DataGrid_Set("年月日", "返済日", 1300, "L")
        Call XZMA010_DataGrid_Set("利息計算日", "利息計算日", 1300, "L")
        Call XZMA010_DataGrid_Set("元金額", "返済元金", 1500, "R")
        Call XZMA010_DataGrid_Set("利息額", "支払利息", 1500, "R")
        
        If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
            'Call XZMA010_DataGrid_Set("計算利息額", "計算利息額", 1500, "R")
            Call XZMA010_DataGrid_Set("利息差異", "", 1000, "R")
        End If
        
        Call XZMA010_DataGrid_Set("返済金額", "支払合計", 1500, "R")
        Call XZMA010_DataGrid_Set("融資残高", "", 1500, "R")
        Call XZMA010_DataGrid_Set("利率", "", 900, "R")
        Call XZMA010_DataGrid_Set("日割日数", "日数", 600, "R")
        Call XZMA010_DataGrid_Set("調整日数", "調整日数", 1000, "R")
        'Call XZMA010_DataGrid_Set("取消", "", 550, "C")
        'Call XZMA010_DataGrid_Set("取消2", "", 550, "C")
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
' AdodcRefresh_社債
'------------------------------------------------
Private Sub AdodcRefresh_社債()
'
    On Error GoTo AdodcRefresh_社債_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
    Call MXA030_DataGridInit(DataGrid1)
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
    GWhere = GWhere & " And M.借入番号='" & wsBango & "'"
    
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    
    'wstr = wstr + " IIF(取消フラグ２=0,1,2),"
    'wstr = wstr + " IIF(取消フラグ=0,1,2),"
    
    wstr = wstr + " M.借入番号,M.実際年月日,M.元金額,M.利息額,M.返済金額,M.融資残高,"
    wstr = wstr + " M.返済回数 As Grd返済回数,"
    wstr = wstr + " Format(M.実際年月日,'" & Gfmt年月日 & "') As Grd年月日,"
    wstr = wstr + " Format(M.利息計算年月日,'" & Gfmt年月日 & "') As Grd利息計算日,"
    
    wstr = wstr + " Format(M.元金額,'#,##0') As Grd元金額,"
    wstr = wstr + " Format(M.利息額,'#,##0') As Grd利息額,"
    
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        'wstr = wstr + " Format(W.仮計上利息額,'#,##0') As Grd計算利息額,"
        wstr = wstr + " Format(M.利息額-M.仮計上利息額,'#,##0') As Grd利息差異,"
    End If
    
    wstr = wstr + " Format(M.返済金額,'#,##0') As Grd返済金額,"
    wstr = wstr + " Format(M.融資残高,'#,##0') As Grd融資残高,"
    wstr = wstr + " Format(M.利率,'#,##0.00000') As Grd利率,"
    wstr = wstr + " M.日割日数 As Grd日割日数,"
    wstr = wstr + " M.利息対象期間日数 As Grd調整日数,"
    wstr = wstr + " Format(M.初期手数料,'#,##0') As Grd初期手数料,"
    wstr = wstr + " Format(M.元金手数料,'#,##0') As Grd元金手数料,"
    wstr = wstr + " Format(M.利息手数料,'#,##0') As Grd利息手数料,"
    wstr = wstr + " Format(M.初期手数料+M.元金手数料+M.利息手数料,'#,##0') As Grd手数料計,"
    wstr = wstr + " Format(M.保証料,'#,##0') As Grd保証料,"
    wstr = wstr + " Format(M.返済金額+M.初期手数料+M.元金手数料+M.利息手数料+M.保証料,'#,##0') As Grd支払計"
    
    wstr = wstr + " From DCDA020_借入金明細 As M"
    wstr = wstr + GWhere
    wstr = wstr + " Order By 1,2,M.実際年月日"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("返済回数", "回", 400, "C")
        Call XZMA010_DataGrid_Set("年月日", "返済日", 1300, "L")
        Call XZMA010_DataGrid_Set("利息計算日", "利息計算日", 1300, "L")
        Call XZMA010_DataGrid_Set("元金額", "返済元金", 1500, "R")
        Call XZMA010_DataGrid_Set("利息額", "支払利息", 1500, "R")
        
        If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
            Call XZMA010_DataGrid_Set("利息差異", "", 1000, "R")
        End If
        
        Call XZMA010_DataGrid_Set("返済金額", "支払合計", 1500, "R")
        Call XZMA010_DataGrid_Set("融資残高", "", 1500, "R")
        Call XZMA010_DataGrid_Set("利率", "", 900, "R")
        Call XZMA010_DataGrid_Set("日割日数", "日数", 600, "R")
        Call XZMA010_DataGrid_Set("調整日数", "調整日数", 1000, "R")
        Call XZMA010_DataGrid_Set("初期手数料", "", 1500, "R")
        Call XZMA010_DataGrid_Set("元金手数料", "", 1500, "R")
        Call XZMA010_DataGrid_Set("利息手数料", "", 1500, "R")
        Call XZMA010_DataGrid_Set("手数料計", "", 1500, "R")
        Call XZMA010_DataGrid_Set("保証料", "", 1500, "R")
        Call XZMA010_DataGrid_Set("支払計", "", 1500, "R")
    Call XZMA010_DataGrid_Action(DataGrid1)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_社債_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh_社債() でエラー" + vbCrLf + vbCrLf + _
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
    Call CEkey.SetFs(年月日, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    Dim ws01 As String
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd年月日")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        年月日 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd年月日"))
    On Error GoTo 0
'
    Call 画面セット(True)
'
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

'    Call CEkey.SetFs(年月日, True)

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
'
    Call 画面セット(False)
'
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット(pGridClick As Boolean) As Boolean
'
    Dim j As Integer
    Dim wd01 As Double
    Dim ws01 As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
'
    ' =========================================
    '                画面クリア
    ' =========================================
    年月日2.Text = ""
    
    元金額 = ""
    利息額 = ""
    L_返済金額.Caption = ""
    L_融資残高.Caption = ""
    '取消 = 0
    調整利息額.Text = ""
    日割日数.Text = ""
    調整日数.Text = ""
    
    利率.Text = Format(w借入データ.利率, "#,##0.00000")
    L_利率.Caption = Format(w借入データ.利率, "#,##0.00000")
    L_利率.Visible = False
    If w借入データ.金利種別 = XMXA020_区分("金利種別", "変動金利") Then
        利率.Visible = True
    Else
        L_利率.Visible = True
        利率.Visible = False
    End If
    '
    初期手数料 = ""
    元金手数料 = ""
    利息手数料 = ""
    手数料計.Caption = ""
    保証料 = ""
    支払計.Caption = ""
    
    ' =========================================
    '            借入金マスタ セット
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 年月日.Text)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " 借入番号,実際年月日,利息計算年月日,"
    wstr = wstr & "元金額,利息額,仮計上利息額,返済金額,融資残高,"
    wstr = wstr & "日割日数,利息対象期間日数,利率,"
    wstr = wstr & "取消フラグ,取消フラグ２"
    wstr = wstr & " From " & wsTbl
    wstr = wstr & " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr & " And Format(実際年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            If 年月日 <> "" Then
'                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
'                If GRet = vbNo Then
'                    新規変更.Caption = ""
'                    wRs.Close
'                    Set wRs = Nothing
'
'                    Exit Function
'                End If
                
                新規変更.Caption = "新規登録"
                
                年月日2.Text = 年月日.Text
                
'                Call CEkey.SetFs(年月日2, True)
    
            End If
        Else
            画面セット = True
'            Call CEkey.SetFs(年月日2, True)
            新規変更.Caption = "変更"
            
            年月日2.Text = Format(wRs("利息計算年月日"), Gfmt年月日)
            
            元金額 = P8.FFormat(wRs("元金額"), "#,##0")
            利息額 = P8.FFormat(wRs("利息額"), "#,##0")
            L_返済金額.Caption = P8.FFormat(wRs("返済金額"), "#,##0")
            L_融資残高.Caption = P8.FFormat(wRs("融資残高"), "#,##0")
            
            調整利息額 = P8.FFormat(wRs("仮計上利息額"), "#,##0")
            利率 = P8.FFormat(wRs("利率"), "#,##0.00000")
            日割日数 = P8.FFormat(wRs("日割日数"), "#,##0")
            調整日数 = P8.FFormat(wRs("利息対象期間日数"), "#,##0")
            
            '取消 = wRs("取消フラグ")
        End If
    wRs.Close
    Set wRs = Nothing
'
    wd利率 = P8.FCDbl(利率)
'
    If w借入データ.社債フラグ = 1 Then
        wstr = ""
        wstr = wstr & "Select"
        wstr = wstr & " 借入番号,実際年月日,"
        wstr = wstr & "初期手数料,元金手数料,利息手数料,保証料,"
        wstr = wstr & "取消フラグ,取消フラグ２"
        wstr = wstr & " From DBDA010_借入金明細TR2"
        wstr = wstr & " Where 借入番号 = '" & wsBango & "'"
        wstr = wstr & " And Format(実際年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            If wRs.EOF Then
                If 年月日 <> "" Then
    '                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
    '                If GRet = vbNo Then
    '                    新規変更.Caption = ""
    '                    wRs.Close
    '                    Set wRs = Nothing
    '
    '                    Exit Function
    '                End If
                    
                    新規変更.Caption = "新規登録"
                    
                    年月日2.Text = 年月日.Text
                    
                End If
            Else
                画面セット = True
                
                新規変更.Caption = "変更"
                
                初期手数料 = P8.FFormat(wRs("初期手数料"), "#,##0")
                元金手数料 = P8.FFormat(wRs("元金手数料"), "#,##0")
                利息手数料 = P8.FFormat(wRs("利息手数料"), "#,##0")
                保証料 = P8.FFormat(wRs("保証料"), "#,##0")
            
            End If
        wRs.Close
        Set wRs = Nothing
        
        wd01 = P8.FCDbl(初期手数料) + P8.FCDbl(元金手数料) + P8.FCDbl(利息手数料)
        手数料計.Caption = P8.FFormat(wd01, "#,##0")
        支払計.Caption = P8.FFormat(P8.FCDbl(L_返済金額.Caption) + wd01 + P8.FCDbl(保証料), "#,##0")
    End If
'
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
                
        If w借入データ.社債フラグ = 0 Then
            Call AdodcRefresh
        Else
            Call AdodcRefresh_社債
        End If
    End If

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
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
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w年月日 As String
'
    年月日 = ""

    Call 画面セット(False)
    
    新規変更.Caption = ""
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
    Call CEkey.SetFs(元金額, True)
'
End Sub

Private Sub 登録データ照会_Click()
    
    frm_F借入登録データ照会.Show
    frm_F借入登録データ照会.SetFocus

End Sub

Private Sub 年月日_Change()
    
    Me.年月日2 = ""
    Me.利率 = ""
    Me.L_利率 = ""
    Me.日割日数 = ""
    Me.調整日数 = ""
    Me.元金額 = ""
    Me.利息額 = ""
    Me.調整利息額 = ""
    
End Sub

'------------------------------------------------
' 年月日_LostFocus
'------------------------------------------------
Private Sub 年月日_LostFocus()
'
    Dim ws01 As String
    Dim w年月日 As Date
'
'    Call P8.FCControlLeft(年月日, 30)
    
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1", "年月日"
            Exit Sub
    End Select
   
'    If 年月日 = "" Then
'        MsgBox "年月日を入力してください", vbExclamation
'        Call CEkey.SetFs(年月日, True)
'        Exit Sub
'    Else
'        If InStrRev(年月日, "年") Then
'            GVar1 = C年月日.平成To西暦("", 年月日)
'            If GVar1 = 0 Then
'                MsgBox "年月日が不正です", vbExclamation
'                年月日 = "": Call CEkey.SetFs(年月日, True)
'                Exit Sub
'            End If
'        Else
'            If Len(年月日) < 5 Then
'                MsgBox "年月日が不正です", vbExclamation
'                年月日 = "": Call CEkey.SetFs(年月日, True)
'                Exit Sub
'            End If
'
'            ws01 = Mid$(年月日 & "000000", 3, 2)
'            If ws01 < "01" Or ws01 > "12" Then
'                MsgBox "年月日が不正です", vbExclamation
'                年月日 = "": Call CEkey.SetFs(年月日, True)
'                Exit Sub
'            End If
'
'            ws01 = Right$("000000" & 年月日, 2)
'            If ws01 < "01" Or ws01 > "31" Then
'                MsgBox "年月日が不正です", vbExclamation
'                年月日 = "": Call CEkey.SetFs(年月日, True)
'                Exit Sub
'            End If
'
'        End If
'    End If
       
    年月日 = C年月日.FormatDate("年月日", 年月日)
'    If C年月日.平成To西暦("年月", 年月日) = 0 Then
'        MsgBox "年月日が不正です", vbExclamation
'        年月日 = "": Call CEkey.SetFs(年月日, True)
'        Exit Sub
'    End If
'
    Select Case Screen.ActiveControl.Name
        Case "登録", "削除"
'            Call CEkey.SetFs(元金額, True)
'            MsgBox "該当データをセットします。登録処理は行いません。"
            Exit Sub
        Case "利息額再計算", "利息額再計算ALL", "CSV取込", "CSV出力"
            Exit Sub
    End Select
'
    Call B_SET_Click

End Sub

'------------------------------------------------
' 年月日_GotFocus
'------------------------------------------------
Private Sub 年月日_GotFocus()
    Call CEkey.AllSelect
End Sub

Private Sub 年月日2_LostFocus()
'
    年月日2 = C年月日.FormatDate("年月日", 年月日2)
'
    '日割自動計算
    If w借入データ.日割計算区分 <> CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        If C年月日.平成To西暦("年月日", 年月日2, True) <> 0 _
        And 年月日2 <> "" Then
            If 新規変更.Caption = "新規登録" Then
                Call 日割日数セット(CDate(C年月日.平成To西暦("年月日", 年月日2, True)), CDate(C年月日.平成To西暦("年月日", 年月日, True)))
            End If
        End If
    End If
'
End Sub

Private Sub 元金額_LostFocus()
    元金額 = Right$(Format(元金額, "#,##0"), 15)
End Sub

Private Sub 利息額_LostFocus()
    利息額 = Right$(Format(利息額, "#,##0"), 15)
End Sub

Private Sub 調整利息額_LostFocus()
    調整利息額 = Right$(P8.FFormat(調整利息額, "#,##0"), 15)
End Sub

Private Sub 利率_LostFocus()
    利率 = P8.FFormat(利率, "#,##0.00000")
End Sub

Private Sub 初期手数料_LostFocus()
    初期手数料 = Right$(P8.FFormat(初期手数料, "#,##0"), 15)
    Call 手数料計_セット
End Sub

Private Sub 元金手数料_LostFocus()
    元金手数料 = Right$(P8.FFormat(元金手数料, "#,##0"), 15)
    Call 手数料計_セット
End Sub

Private Sub 利息手数料_LostFocus()
    利息手数料 = Right$(P8.FFormat(利息手数料, "#,##0"), 15)
    Call 手数料計_セット
End Sub

Private Sub 保証料_LostFocus()
    保証料 = Right$(P8.FFormat(保証料, "#,##0"), 15)
    Call 支払計_セット
End Sub

'------------------------------------------------
' 手数料計_セット
'------------------------------------------------
Private Sub 手数料計_セット()
'
    Dim wd01 As Double
'
    wd01 = P8.FCDbl(初期手数料) + P8.FCDbl(元金手数料) + P8.FCDbl(利息手数料)
    手数料計.Caption = P8.FFormat(wd01, "#,##0")
    
    Call 支払計_セット
'
End Sub

'------------------------------------------------
' 支払計_セット
'------------------------------------------------
Private Sub 支払計_セット()
'
    Dim wd01 As Double
'
    wd01 = P8.FCDbl(元金額) + P8.FCDbl(利息額) + P8.FCDbl(手数料計.Caption) + P8.FCDbl(保証料)
    支払計.Caption = P8.FFormat(wd01, "#,##0")
'
End Sub

'------------------------------------------------
' 日割日数セット
'------------------------------------------------
Private Sub 日割日数セット(pRDate As Date, pJDate As Date)
'
    Dim wi01 As Integer
    Dim wDate1 As Date, wDate2 As Date, wDate3 As Date
'
    wd融資残高 = 0
    wi日割日数 = 0

    If w借入データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
        wstr = "Select * From " & wsTbl
        wstr = wstr & " Where 借入番号 = '" & wsBango & "'"
        wstr = wstr & " And Format(利息計算年月日, 'yyyy/mm/dd') > '" & Format(pRDate, "yyyy/mm/dd") & "'"
        wstr = wstr & " order by 利息計算年月日"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
            Do Until wRs.EOF
                wDate3 = Format(wRs("実際年月日"), "yyyy/mm/dd")
                
                wDate1 = Format(wRs("利息計算年月日"), "yyyy/mm/dd")
                wi01 = DateDiff("d", pRDate, wDate1)
                
                If Format(pJDate, "yyyy/mm/dd") = Format(w借入データ.実行日, "yyyy/mm/dd") Then
                'wdate1が実行日の場合
                    
                    '実行日を含めた日数
                    wi01 = wi01 + 1
                    
                    '実行日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                
                ElseIf Format(wDate3, "yyyy/mm/dd") = Format(w借入データ.最終返済実行日, "yyyy/mm/dd") Then
                'wdate1が最終返済日の場合
                    
                    '最終返済日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                End If
                
                '融資残高
                wd融資残高 = P8.FCDbl(wRs("融資残高")) + P8.FCDbl(wRs("元金額"))
                
                Exit Do
                
            Loop
                    
        End If
        wRs.Close
        Set wRs = Nothing
        
    Else
    
        wstr = "Select * From " & wsTbl
        wstr = wstr & " Where 借入番号 = '" & wsBango & "'"
        wstr = wstr & " And Format(利息計算年月日, 'yyyy/mm/dd') < '" & Format(pRDate, "yyyy/mm/dd") & "'"
        wstr = wstr & " order by 利息計算年月日 desc"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
            Do Until wRs.EOF
                wDate3 = Format(wRs("実際年月日"), "yyyy/mm/dd")
                
                wDate1 = Format(wRs("利息計算年月日"), "yyyy/mm/dd")
                wi01 = DateDiff("d", wDate1, pRDate)
                
                If Format(wDate3, "yyyy/mm/dd") = Format(w借入データ.実行日, "yyyy/mm/dd") Then
                'wdate1が実行日の場合
                    
                    '実行日を含めた日数
                    wi01 = wi01 + 1
                    
                    '実行日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                
                ElseIf Format(pJDate, "yyyy/mm/dd") >= Format(w借入データ.最終返済実行日, "yyyy/mm/dd") Then
                'wdate1が最終返済日の場合
                    
                    '最終返済日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                End If
                
                '融資残高
                wd融資残高 = P8.FCDbl(wRs("融資残高"))
                
                Exit Do
            
            Loop
            
        Else
            wDate1 = Format(w借入データ.実行日, "yyyy/mm/dd")
            wi01 = DateDiff("d", wDate1, pRDate)
            
            '実行日を含めた日数
            wi01 = wi01 + 1
            
            '実行日控除
            If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日控除")) _
            Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                wi01 = wi01 - 1
            End If
            
            '融資残高
            wd融資残高 = w借入データ.融資金額
        
        End If
        
        wRs.Close
        Set wRs = Nothing
    End If
'
    If wi01 < 0 Then
        wi01 = 0
    End If
    日割日数 = wi01

    wi日割日数 = wi01
'
End Sub

'------------------------------------------------
' 利息額再計算_Click
'------------------------------------------------
Private Sub 利息額再計算_Click()
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
    If Not IsNumeric(元金額) And 元金額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(元金額, True)
        Exit Sub
    End If
    
    If Not IsNumeric(利息額) And 利息額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(利息額, True)
        Exit Sub
    End If
    
    If Not IsNumeric(初期手数料) And 初期手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(初期手数料, True)
        Exit Sub
    End If
    
    If Not IsNumeric(元金手数料) And 元金手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(元金手数料, True)
        Exit Sub
    End If

    If Not IsNumeric(利息手数料) And 利息手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(利息手数料, True)
        Exit Sub
    End If

    If Not IsNumeric(保証料) And 保証料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(保証料, True)
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "返済日が違います", vbExclamation
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    GVar1 = C年月日.平成To西暦("年月日", 年月日2.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "利息計算日が違います", vbExclamation
        Call CEkey.SetFs(年月日2, True)
        Exit Sub
    End If
'

'
    Call 日割日数セット(CDate(C年月日.平成To西暦("年月日", 年月日2, True)), CDate(C年月日.平成To西暦("年月日", 年月日, True)))
    
    If w借入データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
        利息額 = MBD010_利息計算小数点5桁(P8.FCDbl(利率), _
                    wd融資残高, wi日割日数, w借入データ.金利計算年間日数)
    Else
        利息額 = MBD010_利息計算小数点5桁(P8.FCDbl(利率), _
                    wd融資残高, wi日割日数, w借入データ.金利計算年間日数)
    End If
'
    利息額 = Format(利息額, "#,##0")
'
End Sub

'------------------------------------------------
' 返済残高_セット
'------------------------------------------------
Private Sub 返済残高_セット()
'
    Dim wi01 As Integer
    Dim w融資残高 As Double
    Dim FLG_Shokai As Boolean
'
    '融資残高
    w融資残高 = w借入データ.融資金額
    wi01 = 1
    
    w初回 = Null
    w最終 = Null
    
    FLG_Shokai = False
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " 実際年月日,元金額,融資残高,返済回数,取消フラグ"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And 取消フラグ=0"
    wstr = wstr + " Order by 実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.EOF
        wRs("融資残高") = w融資残高 - wRs("元金額")
        wRs("返済回数") = wi01
            
        If FLG_Shokai = False And P8.FCDbl(wRs("元金額")) > 0 Then
            w初回 = wRs("実際年月日")
            FLG_Shokai = True
        End If
        
        w融資残高 = wRs("融資残高")
        wi01 = wi01 + 1
        
        wRs.Update
    
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    '合計融資残高
    L_合計融資残高.Caption = ""
    
    wstr = ""
    wstr = wstr + "Select 実際年月日,融資残高"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And 取消フラグ=0"
    wstr = wstr + " Order by 実際年月日 desc"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        L_合計融資残高.Caption = Format(P8.FCDbl(wRs("融資残高")), "#,##0")
        w最終 = wRs("実際年月日")
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
返済残高_セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 返済残高_セット() でエラー" + vbCrLf + vbCrLf + _
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
' 合計集計_セット
'------------------------------------------------
Private Sub 合計集計_セット()
'
    L_合計元金額.Caption = ""
    L_合計利息額.Caption = ""
    L_合計調整利息額.Caption = ""
    
    wstr = ""
    wstr = wstr & "Select sum(元金額) As 合計元金額,"
    wstr = wstr & " sum(利息額) As 合計利息額,"
    wstr = wstr & " sum(仮計上利息額) As 合計調整利息額"
    wstr = wstr & " From " & wsTbl
    wstr = wstr & " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr & " And 取消フラグ=0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        L_合計元金額.Caption = Format(P8.FCDbl(wRs("合計元金額")), "#,##0")
        L_合計利息額.Caption = Format(P8.FCDbl(wRs("合計利息額")), "#,##0")
        L_合計調整利息額.Caption = Format(P8.FCDbl(wRs("合計調整利息額")), "#,##0")
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
合計集計_セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 合計集計_セット() でエラー" + vbCrLf + vbCrLf + _
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
    Dim wi01 As Integer
    Dim wd01 As Double, wd02 As Double
    Dim wDate1 As Date, wDate2 As Date
'
    On Error GoTo 削除_Click_ERR
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

    GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
    If GRet = vbNo Then
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
'
    ' =========================================
    '            明細TR
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
    
    wstr = ""
    wstr = wstr + "Delete * From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And Format(実際年月日,'yyyy/mm/dd') = '" & Format(GVar1, "yyyy/mm/dd") & "'"
    GDb.Execute wstr
    
    wstr = ""
    wstr = wstr + "Delete * From DBDA010_借入金明細TR2"
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And Format(実際年月日,'yyyy/mm/dd') = '" & Format(GVar1, "yyyy/mm/dd") & "'"
    GDb.Execute wstr
'
    Call 返済残高_セット
'
    w借入データ.初回返済実行日 = Format(P8.FCDate(w初回), "yyyy/mm/dd")
    w借入データ.最終返済実行日 = Format(P8.FCDate(w最終), "yyyy/mm/dd")
'
    ' =========================================
    '            借入金、貸付金
    ' =========================================
    wstr = ""
    wstr = wstr + "Select * From " & wsTbl2
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
    
        If P8.FCDbl(L_合計融資残高.Caption) = 0 Then
            wRs("手入力区分") = 1
        Else
            wRs("手入力区分") = 2
        End If
        
        If Not IsNull(P8.FCDate(w初回)) Then
            wRs("初回返済年月") = Format(P8.FCDate(w初回), "yyyy/mm/dd")
            wRs("初回返済実行日") = Format(P8.FCDate(w初回), "yyyy/mm/dd")
        End If
        
        If Not IsNull(P8.FCDate(w最終)) Then
            wRs("最終返済年月") = Format(P8.FCDate(w最終), "yyyy/mm/dd")
            wRs("最終返済実行日") = Format(P8.FCDate(w最終), "yyyy/mm/dd")
        End If
        
        wRs.Update
        
    End If
    wRs.Close
    Set wRs = Nothing
'
    If Not IsNull(P8.FCDate(w初回)) Then
        L_初回返済年月.Caption = Format(w初回, Gfmt年月)
    End If
    
    If Not IsNull(P8.FCDate(w最終)) Then
        L_最終返済年月.Caption = Format(w最終, Gfmt年月)
    End If
'
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        
        '日割自動計算
        Call MBD010_借入金入力明細作成_日割日数再計算(w借入データ, wsTbl)

        'GRID利息額自動計算値作成(DCDA020_借入金明細)
        wstr = ""
        wstr = wstr + "Select * From " & wsTbl2
        wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
        
            '** 明細ファイル 作成 **
            Call MBD010_借入金入力明細_利息額自動計算(w借入データ, wsTbl)
            
        End If
        wRs.Close
        Set wRs = Nothing
        
    End If
'
    If w借入データ.社債フラグ = 1 Then
        Call MDA020_借入金入力社債明細作成(w借入データ)
    End If
    
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) _
    Or w借入データ.社債フラグ = 1 Then
        Call MBD010_借入明細作成_入力登録(w借入データ)
    End If
'
    '----------< DataGrid Close >----------------------------------------------
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号=" & wsBango & ","
    GLogStr = "返済日=" & Format(GVar1, "yyyy/mm/dd")
    Call MXA030_LOG_WRITE("借入金入力登録", "削除", GLogStr)
'
    ' =========================================
    '               画面セット
    ' =========================================
    年月日.Text = ""
    
    Call 画面セット(False)
'    Call 登録後初期セット
    Call CEkey.SetFs(年月日, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "削除しました。", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
削除_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 削除_Click() でエラー" + vbCrLf + vbCrLf + _
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
' 登録_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim wi01 As Integer
    Dim wd01 As Double, wd02 As Double
    Dim wDate1 As Date, wDate2 As Date
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
    If Not IsNumeric(元金額) And 元金額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(元金額, True)
        Exit Sub
    End If
    
    If Not IsNumeric(利息額) And 利息額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(利息額, True)
        Exit Sub
    End If
    
    If Not IsNumeric(初期手数料) And 初期手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(初期手数料, True)
        Exit Sub
    End If
    
    If Not IsNumeric(元金手数料) And 元金手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(元金手数料, True)
        Exit Sub
    End If

    If Not IsNumeric(利息手数料) And 利息手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(利息手数料, True)
        Exit Sub
    End If

    If Not IsNumeric(保証料) And 保証料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(保証料, True)
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "返済日が違います", vbExclamation
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
    
    '実行日
    wDate1 = C年月日.平成To西暦("年月日", L_実行日.Caption)
    wDate2 = C年月日.平成To西暦("年月日", 年月日.Text)
    If wDate1 > wDate2 Then
        MsgBox "初回返済年月が違います", vbExclamation
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    GVar1 = C年月日.平成To西暦("年月日", 年月日2.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "利息計算日が違います", vbExclamation
        Call CEkey.SetFs(年月日2, True)
        Exit Sub
    End If
    
    '実行日
    wDate1 = C年月日.平成To西暦("年月日", L_実行日.Caption)
    wDate2 = C年月日.平成To西暦("年月日", 年月日2.Text)
    If wDate1 > wDate2 Then
        MsgBox "初回返済年月が違います", vbExclamation
        Call CEkey.SetFs(年月日2, True)
        Exit Sub
    End If
'
    ' =========================================
    '            明細TR
    ' =========================================
    wd01 = P8.FCDbl(元金額)
    wd02 = P8.FCDbl(利息額)
    L_融資残高.Caption = Format(wd01 + wd02, "#,##0")
'
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号="
    GLogStr = GLogStr & wsBango & ","
    GLogStr = GLogStr & "実際年月日="
    GLogStr = GLogStr & Format(GVar1, "yyyy/mm/dd") & ","
    
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " 借入番号,実際年月日,返済予定年月,利息計算年月日,"
    wstr = wstr & "元金額,利息額,仮計上利息額,返済金額,融資残高,"
    wstr = wstr & "日割日数,利息対象期間日数,利率,"
    wstr = wstr & "返済回数,取消フラグ"
    wstr = wstr & " From " & wsTbl
    wstr = wstr & " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr & " And Format(実際年月日,'yyyy/mm/dd') = '" & Format(GVar1, "yyyy/mm/dd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.EOF Then
        If wd01 + wd02 <> 0 Then
            wRs.AddNew
            
                wRs("借入番号") = wsBango
                wRs("実際年月日") = CDate(GVar1)
            
                GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
                wRs("返済予定年月") = CDate(Format(GVar1, "yyyy/mm") & "/01")
                
                GVar1 = C年月日.平成To西暦("年月日", 年月日2.Text)
                wRs("利息計算年月日") = CDate(Format(GVar1, "yyyy/mm/dd"))
                        
                wRs("元金額") = wd01
                wRs("利息額") = wd02
                wRs("返済金額") = wd01 + wd02
                
                wRs("取消フラグ") = 0 'P8.FCDbl(取消)
                
                If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) _
                Or wd01 + wd02 = 0 Then
                    wRs("仮計上利息額") = 0
                    wRs("日割日数") = 0
                    wRs("利息対象期間日数") = 0
                Else
                    wRs("仮計上利息額") = P8.FCDbl(調整利息額)
                    wRs("日割日数") = P8.FCDbl(日割日数)
                    wRs("利息対象期間日数") = P8.FCDbl(調整日数)
                End If
                
                wRs("利率") = P8.FCDbl(利率)
                
                If wRs("取消フラグ") = 1 Then
                    wRs("返済回数") = 0
                End If
                
                ' =========================================
                '               LOG_WRITE
                ' =========================================
                GLogStr = GLogStr & "利息計算年月日="
                GLogStr = GLogStr & wRs("利息計算年月日") & ","
                GLogStr = GLogStr & "元金額="
                GLogStr = GLogStr & wRs("元金額") & ","
                GLogStr = GLogStr & "利息額="
                GLogStr = GLogStr & wRs("利息額") & ","
                GLogStr = GLogStr & "返済金額="
                GLogStr = GLogStr & wRs("返済金額") & ","
                GLogStr = GLogStr & "調整利息額="
                GLogStr = GLogStr & wRs("仮計上利息額") & ","
                GLogStr = GLogStr & "日割日数="
                GLogStr = GLogStr & wRs("日割日数") & ","
                GLogStr = GLogStr & "調整日数="
                GLogStr = GLogStr & wRs("利息対象期間日数") & ","
                GLogStr = GLogStr & "利率="
                GLogStr = GLogStr & wRs("利率") & ","
                GLogStr = GLogStr & "返済回数="
                GLogStr = GLogStr & wRs("返済回数")
                    
            wRs.Update
        
        End If
    Else
            GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
            wRs("返済予定年月") = CDate(Format(GVar1, "yyyy/mm") & "/01")
            
            GVar1 = C年月日.平成To西暦("年月日", 年月日2.Text)
            wRs("利息計算年月日") = CDate(Format(GVar1, "yyyy/mm/dd"))
                    
            wRs("元金額") = wd01
            wRs("利息額") = wd02
            wRs("返済金額") = wd01 + wd02
            
            wRs("取消フラグ") = 0 'P8.FCDbl(取消)
            
            If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) _
            Or wd01 + wd02 = 0 Then
                wRs("仮計上利息額") = 0
                wRs("日割日数") = 0
                wRs("利息対象期間日数") = 0
            Else
                wRs("仮計上利息額") = P8.FCDbl(調整利息額)
                wRs("日割日数") = P8.FCDbl(日割日数)
                wRs("利息対象期間日数") = P8.FCDbl(調整日数)
            End If
            
            wRs("利率") = P8.FCDbl(利率)
            
            If wRs("取消フラグ") = 1 Then
                wRs("返済回数") = 0
            End If
            
            ' =========================================
            '               LOG_WRITE
            ' =========================================
            GLogStr = GLogStr & "利息計算年月日="
            GLogStr = GLogStr & wRs("利息計算年月日") & ","
            GLogStr = GLogStr & "元金額="
            GLogStr = GLogStr & wRs("元金額") & ","
            GLogStr = GLogStr & "利息額="
            GLogStr = GLogStr & wRs("利息額") & ","
            GLogStr = GLogStr & "返済金額="
            GLogStr = GLogStr & wRs("返済金額") & ","
            GLogStr = GLogStr & "調整利息額="
            GLogStr = GLogStr & wRs("仮計上利息額") & ","
            GLogStr = GLogStr & "日割日数="
            GLogStr = GLogStr & wRs("日割日数") & ","
            GLogStr = GLogStr & "調整日数="
            GLogStr = GLogStr & wRs("利息対象期間日数") & ","
            GLogStr = GLogStr & "利率="
            GLogStr = GLogStr & wRs("利率") & ","
            GLogStr = GLogStr & "返済回数="
            GLogStr = GLogStr & wRs("返済回数")
                
        wRs.Update
    End If
     
    wRs.Close
    Set wRs = Nothing
'
    '残高、w初回、w最終
    Call 返済残高_セット
'
    '社債 経費
    If w借入データ.社債フラグ = 1 Then
        wstr = ""
        wstr = wstr & "Select"
        wstr = wstr & " 借入番号,返済予定年月,実際年月日,"
        wstr = wstr & "初期手数料,元金手数料,利息手数料,保証料,"
        wstr = wstr & "取消フラグ"
        wstr = wstr & " From DBDA010_借入金明細TR2"
        wstr = wstr & " Where 借入番号 = '" & wsBango & "'"
        wstr = wstr & " And Format(実際年月日,'yyyy/mm/dd') = '" & Format(GVar1, "yyyy/mm/dd") & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            If wRs.EOF Then
                wRs.AddNew
                
                wRs("借入番号") = wsBango
                wRs("実際年月日") = CDate(GVar1)
            End If
                
                wRs("取消フラグ") = 0 'P8.FCDbl(取消)
                
                wRs("返済予定年月") = CDate(Format(GVar1, "yyyy/mm") & "/01")
                wRs("初期手数料") = P8.FCDbl(初期手数料)
                wRs("元金手数料") = P8.FCDbl(元金手数料)
                wRs("利息手数料") = P8.FCDbl(利息手数料)
                wRs("保証料") = P8.FCDbl(保証料)
            
            wRs.Update
        wRs.Close
        Set wRs = Nothing
        
    End If

    wstr = "Delete * From DBDA010_借入金明細TR2"
    wstr = wstr & " Where 保証料=0 And 初期手数料=0 And 元金手数料=0 And 利息手数料=0"
    GDb.Execute wstr
'
    w借入データ.初回返済実行日 = Format(P8.FCDate(w初回), "yyyy/mm/dd")
    w借入データ.最終返済実行日 = Format(P8.FCDate(w最終), "yyyy/mm/dd")
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = GLogStr & "初期手数料="
    GLogStr = GLogStr & P8.FCDbl(初期手数料) & ","
    GLogStr = GLogStr & "元金手数料="
    GLogStr = GLogStr & P8.FCDbl(元金手数料) & ","
    GLogStr = GLogStr & "利息手数料="
    GLogStr = GLogStr & P8.FCDbl(利息手数料) & ","
    GLogStr = GLogStr & "保証料="
    GLogStr = GLogStr & P8.FCDbl(保証料)
    Call MXA030_LOG_WRITE("借入金入力登録", "更新", GLogStr)
'
    ' =========================================
    '            借入金、貸付金
    ' =========================================
    wstr = ""
    wstr = wstr + "Select * From " & wsTbl2
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
    
        If P8.FCDbl(L_合計融資残高.Caption) = 0 Then
            wRs("手入力区分") = 1
        Else
            wRs("手入力区分") = 2
        End If
        
        If Not IsNull(P8.FCDate(w初回)) Then
            wRs("初回返済年月") = Format(P8.FCDate(w初回), "yyyy/mm/dd")
            wRs("初回返済実行日") = Format(P8.FCDate(w初回), "yyyy/mm/dd")
        End If
        
        If Not IsNull(P8.FCDate(w最終)) Then
            wRs("最終返済年月") = Format(P8.FCDate(w最終), "yyyy/mm/dd")
            wRs("最終返済実行日") = Format(P8.FCDate(w最終), "yyyy/mm/dd")
        End If
        
        wRs.Update
        
    End If
    wRs.Close
    Set wRs = Nothing
'
    If Not IsNull(P8.FCDate(w初回)) Then
        L_初回返済年月.Caption = Format(w初回, Gfmt年月)
    End If
    
    If Not IsNull(P8.FCDate(w最終)) Then
        L_最終返済年月.Caption = Format(w最終, Gfmt年月)
    End If
'
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        
        '日割自動計算
        Call MBD010_借入金入力明細作成_日割日数再計算(w借入データ, wsTbl)

        '利率変更時
        If wd利率 <> P8.FCDbl(利率) Then
            GRet = MsgBox("返済日:" & 年月日.Text & "以降の利率を変更しますか？", vbYesNo + vbQuestion)
            If GRet = vbYes Then
                GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
                Call MBD010_借入金入力明細作成_利率変更(w借入データ, wsTbl, CDate(GVar1), wd利率, P8.FCDbl(利率))
            End If
        End If
    
        'GRID利息額自動計算値作成(DCDA020_借入金明細)
        wstr = ""
        wstr = wstr + "Select * From " & wsTbl2
        wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
        
            '** 明細ファイル 作成 **
            Call MBD010_借入金入力明細_利息額自動計算(w借入データ, wsTbl)
            
        End If
        wRs.Close
        Set wRs = Nothing
        
    End If
'
    If w借入データ.社債フラグ = 1 Then
        If w借入データ.日割計算区分 <> CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
            '** 明細ファイル 削除 **
            wstr = ""
            wstr = wstr + "Delete * From DCDA020_借入金明細"
            GDb.Execute wstr
            '
            Call MBD010_借入金入力明細Read(w借入データ)
        End If
        
        Call MDA020_借入金入力社債明細作成(w借入データ)
    End If
    
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) _
    Or w借入データ.社債フラグ = 1 Then
        Call MBD010_借入明細作成_入力登録(w借入データ)
    End If
'
    '----------< DataGrid Close >----------------------------------------------
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
'    Call 登録後初期セット
    Call CEkey.SetFs(年月日, True)
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
' 利息額再計算ALL_Click
'------------------------------------------------
Private Sub 利息額再計算ALL_Click()
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
    If Not IsNumeric(元金額) And 元金額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(元金額, True)
        Exit Sub
    End If
    
    If Not IsNumeric(利息額) And 利息額 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(利息額, True)
        Exit Sub
    End If
    
    If Not IsNumeric(初期手数料) And 初期手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(初期手数料, True)
        Exit Sub
    End If
    
    If Not IsNumeric(元金手数料) And 元金手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(元金手数料, True)
        Exit Sub
    End If

    If Not IsNumeric(利息手数料) And 利息手数料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(利息手数料, True)
        Exit Sub
    End If

    If Not IsNumeric(保証料) And 保証料 <> "" Then
        MsgBox "入力を確認してください", vbExclamation: Call CEkey.SetFs(保証料, True)
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "返済日が違います", vbExclamation
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    GVar1 = C年月日.平成To西暦("年月日", 年月日2.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "利息計算日が違います", vbExclamation
        Call CEkey.SetFs(年月日2, True)
        Exit Sub
    End If
'

'
    If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
        
        GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
        If GVar1 = 0 Or GVar1 = Null Then
            Exit Sub
        End If
        
        GRet = MsgBox("返済日:" & 年月日.Text & "以降の利息額を再計算しますか？", vbYesNo + vbQuestion)
        If GRet = vbYes Then
            Call MBD010_借入金入力明細作成_利息額再計算(w借入データ, wsTbl, CDate(GVar1))
        Else
            Exit Sub
        End If
    
        'GRID利息額自動計算値作成(DCDA020_借入金明細)
        wstr = ""
        wstr = wstr + "Select * From " & wsTbl2
        wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
        
            '** 明細ファイル 作成 **
            Call MBD010_借入金入力明細_利息額自動計算(w借入データ, wsTbl)
            
        End If
        wRs.Close
        Set wRs = Nothing
        
        If w借入データ.社債フラグ = 1 Then
            Call MDA020_借入金入力社債明細作成(w借入データ)
        End If
        
        If w借入データ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) _
        Or w借入データ.社債フラグ = 1 Then
            Call MBD010_借入明細作成_入力登録(w借入データ)
        End If
    '
        '
            '----------< DataGrid Close >----------------------------------------------
            Set DataGrid1.DataSource = Nothing
            Adodc1.Recordset.Close
        '
            ' =========================================
            '               画面セット
            ' =========================================
            Call 画面セット(False)
        '    Call 登録後初期セット
            Call CEkey.SetFs(年月日, True)
        '
            ' =========================================
            '               メッセージ
            ' =========================================
            MsgBox "登録しました。", vbInformation
        '
    End If
'
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
    On Error GoTo CSV取込_Click_ERR
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

    GRet = MsgBox("CSVファイルをインポートします。" _
                    & vbCrLf & vbCrLf & "既存のデータは上書きされます。" & vbCrLf & _
                    "よろしいですか？", vbExclamation + vbYesNo, "取込")
    If GRet = vbNo Then
        Exit Sub
    End If
'
    If w借入データ.社債フラグ = 0 Then
        ws01 = "借入明細表.csv"
    ElseIf w借入データ.社債フラグ = 1 Then
        ws01 = "社債明細表.csv"
    End If
        wsRet = MXA040_COMDLG(CommonDialog1, "CSVファイル選択", "", _
                            "テキストファイル(*.csv)|*.csv", ws01)
    If wsRet = "" Then
        Exit Sub
    ElseIf wsRet = "キャンセル" Then
        Exit Sub
    End If
'
    '社債フラグ
    GInt1 = w借入データ.社債フラグ
'
    GRet = MXA040_借入明細取込(wsRet, w借入データ.借入番号)
    If GRet <> True Then
        MsgBox "CSVファイルをインポートできませんでした", vbInformation
        
        Exit Sub
    End If
'
    ' =========================================
    '               画面セット
    ' =========================================
    GStr = "借入金登録"
    GStr_2 = w借入データ.借入番号
    GStr_3 = "明細入力"
    Call Form_Load
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "CSVファイルをインポートしました", vbInformation
'
    ' =========================================
    '           Csv File Drive
    ' =========================================
    Call MX040_CsvPath
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
CSV取込_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ CSV取込_Click() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "金剛石を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' CSV出力_Click
'------------------------------------------------
Private Sub CSV出力_Click()
'
    Dim wsRet As String, wsFileName As String
'
    If L_番号.Caption = "" Then
        Exit Sub
    End If
'
    GRpt.帳票名 = "借入明細表"
    GRpt.コンボ_01 = w借入データ.借入番号
    GRpt.CSV = 1
'
    If w借入データ.社債フラグ = 0 Then
        wsFileName = w借入データ.借入番号 & "借入明細表.csv"
    ElseIf w借入データ.社債フラグ = 1 Then
        wsFileName = w借入データ.借入番号 & "社債明細表.csv"
    End If
    wsRet = MXA040_COMDLG(CommonDialog1, "CSVファイル出力", "", _
                        "テキストファイル(*.csv)|*.csv", wsFileName)
    If wsRet = "" Then
        Exit Sub
    ElseIf wsRet = "キャンセル" Then
        Exit Sub
    End If
'
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA020_借入金明細"
    GDb.Execute wstr
    
    '** 明細ファイル 作成 **
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From " & wsTbl2
    wstr = wstr + " Where 借入番号 = '" & GRpt.コンボ_01 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
      Do Until wRs.EOF
      
          w借入データ = MBD010_借入データセット(wRs)
          If P8.FCDbl(wRs("手入力区分")) = "0" Then
          '標準
              Call MBD010_借入金テーブル作成("", w借入データ)
              Call MBD010_借入明細作成("", w借入データ)        ' 07/02/21 V180
          Else
          '入力登録
              Call MBD010_借入金入力明細Read(w借入データ)
              If w借入データ.社債フラグ = 1 Then
                Call MDA020_借入金入力社債明細作成(w借入データ)
              End If
              Call MBD010_借入明細作成_入力登録(w借入データ)
          End If

          wRs.MoveNext
      Loop
    Else
        wRs.Close
        Set wRs = Nothing
    
        Exit Sub
    End If
    wRs.Close
    Set wRs = Nothing
'
    Call MX040_CsvPath_CDL(wsRet)
    wsFileName = Mid(wsRet, InStrRev(wsRet, "\") + 1)
    If w借入データ.社債フラグ = 0 Then
        Call MX040_借入明細表(wsFileName)
    ElseIf w借入データ.社債フラグ = 1 Then
        Call MX040_社債明細表(wsFileName)
    End If
'
    ' =========================================
    '           Csv File Drive
    ' =========================================
    Call MX040_CsvPath
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
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
'    Unload frm_I借入金登録
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
    
'
End Sub
