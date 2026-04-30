VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_M銀行マスタ 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "銀行マスタ"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   Icon            =   "frm_M銀行マスタ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox 削除データを表示 
      Caption         =   "削除データを表示"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton 登録 
      Caption         =   "登録(F11)"
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
      Left            =   9000
      TabIndex        =   22
      Top             =   9360
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
      Height          =   3975
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   12615
      Begin VB.CheckBox 変更 
         Caption         =   "銀行番号変更"
         Height          =   255
         Left            =   7560
         TabIndex        =   54
         Top             =   360
         Width           =   1695
      End
      Begin VB.Frame Frame_Henko 
         Height          =   1815
         Left            =   240
         TabIndex        =   47
         Top             =   2040
         Width           =   12255
         Begin VB.TextBox 変更金融機関名 
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
            Left            =   7320
            MaxLength       =   20
            TabIndex        =   18
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox 変更支店名 
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
            Left            =   7320
            MaxLength       =   20
            TabIndex        =   19
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox 変更支店番号 
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
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   16
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox 変更金融機関番号 
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
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   15
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label L_変更銀行名 
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
            Left            =   7320
            TabIndex        =   20
            Top             =   960
            Width           =   4815
         End
         Begin VB.Label L_変更銀行番号 
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
            Left            =   1920
            TabIndex        =   17
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label L_H6 
            Alignment       =   1  '右揃え
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  '実線
            Caption         =   " 変更銀行名"
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
            Left            =   5520
            TabIndex        =   53
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label L_H4 
            Alignment       =   1  '右揃え
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  '実線
            Caption         =   " 変更金融機関名"
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
            Left            =   5520
            TabIndex        =   52
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label L_H5 
            Alignment       =   1  '右揃え
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  '実線
            Caption         =   " 変更支店名"
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
            Left            =   5520
            TabIndex        =   51
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label L_H2 
            Alignment       =   1  '右揃え
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  '実線
            Caption         =   " 変更支店番号"
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
            TabIndex        =   50
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label L_H1 
            Alignment       =   1  '右揃え
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  '実線
            Caption         =   "変更金融機関番号"
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
            TabIndex        =   49
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label L_H3 
            Alignment       =   1  '右揃え
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  '実線
            Caption         =   " 変更銀行番号"
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
            TabIndex        =   48
            Top             =   960
            Width           =   1815
         End
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   120
         TabIndex        =   37
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
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   5295
         Begin VB.TextBox 金融機関番号 
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
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   0
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox 支店番号 
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
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   1
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label L_銀行番号 
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
            Left            =   1920
            TabIndex        =   2
            Top             =   960
            Width           =   3255
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
            TabIndex        =   46
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label8 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   " 金融機関番号"
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
            TabIndex        =   45
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label12 
            Alignment       =   1  '右揃え
            BackColor       =   &H00D6DBBD&
            BorderStyle     =   1  '実線
            Caption         =   " 支店番号"
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
            TabIndex        =   44
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.TextBox 預金種別 
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
         Left            =   7560
         MaxLength       =   1
         TabIndex        =   13
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox 口座番号 
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
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   14
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox 支店名 
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
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox 金融機関名 
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
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox 削除 
         Caption         =   "削除"
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin 借換たろう.ZU020_ComboBox 支払日 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
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
      Begin 借換たろう.ZU020_ComboBox 営業日 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
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
      Begin 借換たろう.ZU020_ComboBox 利息区分 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   2760
         Width           =   3255
         _ExtentX        =   5741
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
      Begin 借換たろう.ZU020_ComboBox 利息日数 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   3120
         Width           =   3255
         _ExtentX        =   5741
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
      Begin 借換たろう.ZU020_ComboBox 利息支払 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   3480
         Width           =   3255
         _ExtentX        =   5741
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
      Begin 借換たろう.ZU020_ComboBox 利息控除 
         Height          =   315
         Left            =   7560
         TabIndex        =   11
         Top             =   2040
         Width           =   4455
         _ExtentX        =   7858
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
      Begin 借換たろう.ZU020_ComboBox 金利計算 
         Height          =   315
         Left            =   7560
         TabIndex        =   12
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.Label L_銀行名 
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
         Left            =   7560
         TabIndex        =   5
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label L_8 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "預金種別"
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
         Left            =   5760
         TabIndex        =   42
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label L_9 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 口座番号"
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
         Left            =   5760
         TabIndex        =   41
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 支店名"
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
         Left            =   5760
         TabIndex        =   40
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 金融機関名"
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
         Left            =   5760
         TabIndex        =   39
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 銀行名"
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
         Left            =   5760
         TabIndex        =   34
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label L_2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 営業日"
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
         TabIndex        =   33
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label L_1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 支払日"
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
         TabIndex        =   32
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label L_3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 利息区分"
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
         TabIndex        =   31
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label L_4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 利息計算日数"
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
         TabIndex        =   30
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label L_5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 利息支払方法"
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
         TabIndex        =   29
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label L_6 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 利息控除区分"
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
         Left            =   5760
         TabIndex        =   28
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label L_7 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 金利計算日数"
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
         Left            =   5760
         TabIndex        =   27
         Top             =   2400
         Width           =   1815
      End
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
      Left            =   1440
      TabIndex        =   26
      Top             =   9480
      Visible         =   0   'False
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
      Left            =   10920
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   9360
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   240
      Top             =   10320
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
      Height          =   4005
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7064
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
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "銀行マスタ"
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
      Left            =   240
      TabIndex        =   35
      Top             =   9840
      Visible         =   0   'False
      Width           =   15015
   End
End
Attribute VB_Name = "frm_M銀行マスタ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "銀行マスタ"
'
' =========================================
'             修正履歴
' =========================================
' @001 2018/05/16 銀行番号変更対応
'
' =========================================
'
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
    ' =========================================
    '                 初期設定
    ' =========================================
'    Me.Caption = GFcap
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    '銀行名.MaxLength = 50
'
    ' =========================================
    '                 初期設定
    ' =========================================
    Call 登録後初期セット
    メッセージ = ""
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
    ' =========================================
    '             コンボボックス
    ' =========================================
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
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("営業日", "翌営業日"), "翌営業日")
        Call .AddItem(XMXA020_区分("営業日", "前営業日"), "前営業日")
    End With
    営業日.CreateCombo
'
    With 利息区分
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("利息区分", "利息先払"), "利息先払")
        Call .AddItem(XMXA020_区分("利息区分", "利息後払"), "利息後払")
    End With
    利息区分.CreateCombo
'
    With 利息日数
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("利息日数", "営業日数"), "営業日数")
        Call .AddItem(XMXA020_区分("利息日数", "固定日数"), "固定日数")
    End With
    利息日数.CreateCombo
'
    With 利息支払
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("利息支払", "毎月"), "毎月")
        Call .AddItem(XMXA020_区分("利息支払", "一括"), "一括")
    End With
    利息支払.CreateCombo
'
    With 利息控除
        .P8_Clear
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
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(XMXA020_区分("金利計算", "365日"), "365日")
        Call .AddItem(XMXA020_区分("金利計算", "360日"), "360日")
    End With
    金利計算.CreateCombo
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
    メッセージ = ""
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
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " IIF(取消フラグ=0,'','*') As Grd削除,"
    wstr = wstr + " 銀行番号,銀行名,支払日,営業日区分,利息区分,利息計算日数区分,利息控除区分,"
    wstr = wstr + " 金利計算年間日数,取消フラグ,"
    wstr = wstr + " 銀行番号 As Grd銀行番号,"
    wstr = wstr + " 銀行名 As Grd銀行名,"
    wstr = wstr + " 支払日 As Grd支払日,"
    wstr = wstr + " IIF(営業日区分=0,'翌営業日','前営業日') As Grd営業日,"
    wstr = wstr + " IIF(利息区分='1','先払','後払') As Grd利息区分,"
    wstr = wstr + " IIF(利息計算日数区分=0,'営業日数','固定日数') As Grd利息日数,"
    wstr = wstr + " IIF(利息支払方法=0,'毎月','一括') As Grd利息支払,"
    wstr = wstr + " IIF(利息控除区分=0,'控除無し',IIF(利息控除区分=1,'実行日控除',"
    wstr = wstr + " IIF(利息控除区分=2,'最終返済日控除',IIF(利息控除区分=3,'実行日及び最終返済日控除','')))) As Grd利息控除,"
    wstr = wstr + " IIF(金利計算年間日数=0,'365日','360日') As Grd金利計算"
    'wstr = wstr + " IIF(取消フラグ=0,'','×') As Grd取消"
    wstr = wstr + " From  DAAA040_銀行マスタ"
    wstr = wstr + GWhere
    If Me.削除データを表示.Value = 0 Then
        wstr = wstr & " AND 取消フラグ = 0"
    End If
    wstr = wstr + " Order By 銀行番号"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        If 削除データを表示.Value = 1 Then
            Call XZMA010_DataGrid_Set("削除", "削", 300, "C")
        End If
        Call XZMA010_DataGrid_Set("銀行番号", "", 1050, "L")
        Call XZMA010_DataGrid_Set("銀行名", "", 3400, "L")
        Call XZMA010_DataGrid_Set("支払日", "", 800, "L")
        Call XZMA010_DataGrid_Set("営業日", "", 1050, "L")
        Call XZMA010_DataGrid_Set("利息区分", "", 1050, "L")
        Call XZMA010_DataGrid_Set("利息日数", "", 1050, "L")
        Call XZMA010_DataGrid_Set("利息支払", "", 1050, "L")
        Call XZMA010_DataGrid_Set("利息控除", "", 1050, "L")
        Call XZMA010_DataGrid_Set("金利計算", "", 1050, "L")
        'Call XZMA010_DataGrid_Set("取消", "", 550, "C")
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
    メッセージ = ""
    Call CEkey.SetFs(支店番号, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("銀行番号")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        L_銀行番号.Caption = P8.FCStr(Adodc1.Recordset.Fields.Item("銀行番号"))
    On Error GoTo 0
    
    Call 画面セット(True)
   
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(金融機関名, True)

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
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    L_銀行名.Caption = ""
    金融機関名 = ""
    支店名 = ""
    利息区分.Text = ""
    支払日.Text = ""
    営業日.Text = ""
    利息区分.Text = ""
    利息日数.Text = ""
    利息支払.Text = ""
    利息控除.Text = ""
    金利計算.Text = ""
    
    預金種別.Text = ""
    口座番号.Text = ""
    
    '取消 = 0
    
    '@001 ADD STR
    変更金融機関番号.Text = ""
    変更支店番号.Text = ""
    L_変更銀行番号.Caption = ""
    変更金融機関名 = ""
    変更支店名.Text = ""
    L_変更銀行名.Caption = ""
    
    変更 = 0
    変更.Enabled = False
    
    支払日.Enabled = True
    営業日.Enabled = True
    利息区分.Enabled = True
    利息日数.Enabled = True
    利息支払.Enabled = True
    利息控除.Enabled = True
    金利計算.Enabled = True
    預金種別.Enabled = True
    口座番号.Enabled = True
    
    Frame_Henko.Visible = False
    '@001 ADD END
    
    ' =========================================
    '            銀行マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA040_銀行マスタ"
    wstr = wstr + " Where 銀行番号 = '" & L_銀行番号.Caption & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            If L_銀行番号.Caption <> "" Then
'                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
'                If GRet = vbNo Then
'                    新規変更.Caption = ""
'                    wRs.Close
'                    Set wRs = Nothing
'
'                    Exit Function
'                End If
                
                新規変更.Caption = "新規登録"
                
                If L_銀行番号.Caption = "SS" Then
                    L_銀行名.Caption = "支店貸付銀行"
                ElseIf L_銀行番号.Caption = "ZZ" Then
                    L_銀行名.Caption = "全社借入銀行"
                End If
            
                Call 新規金融機関セット
                                
                Call CEkey.SetFs(金融機関名, True)
    
            End If
        Else
            画面セット = True
            
            Call CEkey.SetFs(支店名, True)
            新規変更.Caption = "変更"
            
            L_銀行名.Caption = P8.FCStr(wRs("銀行名"))
            金融機関番号 = P8.FCStr(wRs("金融機関番号"))
            金融機関名 = P8.FCStr(wRs("金融機関名"))
            支店番号 = P8.FCStr(wRs("支店番号"))
            支店名 = P8.FCStr(wRs("支店名"))
            
            支払日.Text = P8.FCStr(wRs("支払日"))
            営業日.Text = P8.FCStr(wRs("営業日区分"))
            利息区分.Text = P8.FCStr(wRs("利息区分"))
            利息日数.Text = P8.FCStr(wRs("利息計算日数区分"))
            利息支払.Text = P8.FCStr(wRs("利息支払方法"))
            利息控除.Text = P8.FCStr(wRs("利息控除区分"))
            金利計算.Text = P8.FCStr(wRs("金利計算年間日数"))
            
            預金種別.Text = P8.FCStr(wRs("預金種別"))
            口座番号.Text = P8.FCStr(wRs("口座番号"))
            
            削除 = wRs("取消フラグ")

            '@001
            変更.Enabled = True
        
        End If
    wRs.Close
    Set wRs = Nothing
    
    'If 銀行番号 = "SS" Or 銀行番号 = "ZZ" Then
    '    銀行名.Visible = False
    '
    '    L_銀行名.Caption = 銀行名
    '    L_銀行名.Visible = True
    'End If
    
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "銀行番号 = '" + L_銀行番号.Caption + "'")
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
' 新規金融機関セット
'------------------------------------------------
Private Sub 新規金融機関セット()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
'
    On Error GoTo 新規金融機関セット_ERR
'
    ' =========================================
    '            銀行マスタ セット
    ' =========================================
    wstr1 = ""
    wstr1 = wstr1 + "Select *"
    wstr1 = wstr1 + " From DAAA040_銀行マスタ"
    wstr1 = wstr1 + " Where 金融機関番号 = '" & P8.FCStr(金融機関番号) & "'"
    wstr1 = wstr1 + " Order by 銀行番号"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
    If Not wRs1.eof Then
            
        金融機関名 = P8.FCStr(wRs1("金融機関名"))
        
        支払日.Text = P8.FCStr(wRs1("支払日"))
        営業日.Text = P8.FCStr(wRs1("営業日区分"))
        利息区分.Text = P8.FCStr(wRs1("利息区分"))
        利息日数.Text = P8.FCStr(wRs1("利息計算日数区分"))
        利息支払.Text = P8.FCStr(wRs1("利息支払方法"))
        利息控除.Text = P8.FCStr(wRs1("利息控除区分"))
        金利計算.Text = P8.FCStr(wRs1("金利計算年間日数"))
        
    End If
    wRs1.Close
    Set wRs1 = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
新規金融機関セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 新規金融機関セット() でエラー" + vbCrLf + vbCrLf + _
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
' 検索_Click
'------------------------------------------------
Private Sub 検索_Click()
    Call 登録後初期セット
End Sub

Private Sub 銀行番号_Change()
    
    L_銀行名.Caption = ""
    金融機関番号 = ""
    金融機関名 = ""
    支店番号 = ""
    支店名 = ""
    
    利息区分.Text = ""
    支払日.Text = ""
    営業日.Text = ""
    利息区分.Text = ""
    利息日数.Text = ""
    利息支払.Text = ""
    利息控除.Text = ""
    金利計算.Text = ""
    
    預金種別.Text = ""
    口座番号.Text = ""

End Sub

'------------------------------------------------
' 金融機関番号番号_GotFocus
'------------------------------------------------
Private Sub 金融機関番号_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 金融機関番号_LostFocus
'------------------------------------------------
Private Sub 金融機関番号_LostFocus()
'
    Call P8.FCControlLeft(金融機関番号, 10)
    
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1", "金融機関番号", "支店番号"
            Exit Sub
    End Select

    Call CEkey.SetFs(支店番号, True)

End Sub

'------------------------------------------------
' 支店番号_LostFocus
'------------------------------------------------
Private Sub 支店番号_LostFocus()
'
    On Error GoTo 支店番号_LostFocus_ERR
'
    Call P8.FCControlLeft(支店番号, 8)
    
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1"
            Exit Sub
        Case "金融機関番号", "支店番号"
            Exit Sub
'        Case Else
'            Exit Sub
    End Select
   
'    If 銀行番号 = "" Then
'        MsgBox "コードを入力してください"
'        Call CEkey.SetFs(銀行番号, True)
'        Exit Sub
'    End If
''
'    Select Case Screen.ActiveControl.Name
'        Case "保存"
'            Call CEkey.SetFs(銀行名, True)
'            MsgBox "該当データをセットします。保存処理は行いません。"
'            Exit Sub
'    End Select
'
    If P8.FCStr(支店番号) <> "" Then
        L_銀行番号.Caption = 金融機関番号 & "-" & 支店番号
    Else
        L_銀行番号.Caption = 金融機関番号
    End If
    
    Call 画面セット(False)
    Call CEkey.AllSelect

    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
支店番号_LostFocus_ERR:
    pERR_MES = pPROGRAM_ID + "/ 支店番号_LostFocus() でエラー" + vbCrLf + vbCrLf + _
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
    Dim w銀行番号 As String
'
    w銀行番号 = L_銀行番号.Caption
    
    L_銀行番号.Caption = ""
    
    '@001
    金融機関番号.Text = ""
    支店番号.Text = ""
    L_銀行番号.Caption = ""
    金融機関名.Text = ""
    支店名.Text = ""
    L_銀行名.Caption = ""
    
    変更金融機関番号.Text = ""
    変更支店番号.Text = ""
    L_変更銀行番号.Caption = ""
    変更金融機関名.Text = ""
    変更支店名.Text = ""
    L_変更銀行名.Caption = ""
    
    Frame_Henko.Visible = False

    変更 = 0
    変更.Enabled = False

    Call 画面セット(False)
    新規変更.Caption = ""
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "銀行番号 = '" + w銀行番号 + "'")
    Call CEkey.SetFs(支店番号, True)
'
End Sub

'------------------------------------------------
' LostFocus
'------------------------------------------------
Private Sub 金融機関名_LostFocus()
    Call P8.FCControlLeft(金融機関名, 20)
End Sub
Private Sub 支店名_LostFocus()
    Call P8.FCControlLeft(支店名, 20)
    If P8.FCStr(支店名) <> "" Then
        L_銀行名.Caption = 金融機関名 & Space(1) & 支店名
    Else
        L_銀行名.Caption = 金融機関名
    End If
End Sub

'Private Sub 銀行名_LostFocus()
'    Call P8.FCControlLeft(銀行名, 20)
'End Sub
Private Sub 預金種別_LostFocus()
    Call P8.FCControlLeft(預金種別, 1)
End Sub

Private Sub 口座番号_LostFocus()
    Call P8.FCControlLeft(口座番号, 10)
End Sub

'@001 ADD STR
Private Sub 変更金融機関番号_LostFocus()
    Call P8.FCControlLeft(変更金融機関番号, 10)
End Sub

Private Sub 変更支店番号_LostFocus()
    Call P8.FCControlLeft(変更支店番号, 8)
    If P8.FCStr(変更支店番号) <> "" Then
        L_変更銀行番号.Caption = 変更金融機関番号 & "-" & 変更支店番号
    Else
        L_変更銀行番号.Caption = 変更金融機関番号
    End If
End Sub

Private Sub 変更金融機関名_Change()
    Call P8.FCControlLeft(変更金融機関名, 20)
End Sub
'@001 ADD END

Private Sub 変更支店名_LostFocus()
    Call P8.FCControlLeft(変更支店名, 20)
    If P8.FCStr(変更支店名) <> "" Then
        L_変更銀行名.Caption = 変更金融機関名 & Space(1) & 変更支店名
    Else
        L_変更銀行名.Caption = 変更金融機関名
    End If
End Sub

Private Sub 削除データを表示_Click()
    
    Call AdodcRefresh
    
End Sub

Private Sub 変更_Click()
    If 変更 = 0 Then
        支払日.Enabled = True
        営業日.Enabled = True
        利息区分.Enabled = True
        利息日数.Enabled = True
        利息支払.Enabled = True
        利息控除.Enabled = True
        金利計算.Enabled = True
        預金種別.Enabled = True
        口座番号.Enabled = True
        
        Frame_Henko.Visible = False
        
    Else
        支払日.Enabled = False
        営業日.Enabled = False
        利息区分.Enabled = False
        利息日数.Enabled = False
        利息支払.Enabled = False
        利息控除.Enabled = False
        金利計算.Enabled = False
        預金種別.Enabled = False
        口座番号.Enabled = False
        
        Frame_Henko.Visible = True
    End If
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim wRet As Boolean
    Dim wslog As String
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
'
    '@001
    If 変更 = 1 Then
        Call 銀行番号変更
        Call 登録後初期セット
        
        Exit Sub
    End If
'
    If P8.FCStr(金融機関番号) = "" Then
        MsgBox "金融機関番号が未入力です。", vbExclamation
        Call CEkey.SetFs(金融機関番号, True)
        Exit Sub
    End If

    If 金融機関名 = "" Then
        MsgBox "金融機関名が未入力です。", vbExclamation
        Call CEkey.SetFs(金融機関名, True)
        Exit Sub
    End If
    
    If P8.FCStr(支店番号) = "" Then
        GRet = MsgBox("支店番号が未入力です。登録しますか？", vbYesNo + vbExclamation)
        If GRet = vbNo Then
            Call CEkey.SetFs(支店番号, True)
            Exit Sub
        End If
    End If

'    If 支店名 = "" Then
'        MsgBox "支店名が未入力です。", vbExclamation
'        Call CEkey.SetFs(支店名, True)
'        Exit Sub
'    End If
'
    '2018/02/02
'    GRet = 金融機関名CHECK
'    If GRet <> True Then
'        MsgBox "金融機関名を確認してください。", vbExclamation
'        Call CEkey.SetFs(金融機関名, True)
'        Exit Sub
'    End If
'
'    If P8.FCStr(銀行番号) = "" Then
'        MsgBox "銀行番号が未入力です。", vbExclamation
'        Call CEkey.SetFs(銀行番号, True)
'        Exit Sub
'    End If
'
'    If 銀行名 = "" Then
'        MsgBox "銀行名が未入力です。", vbExclamation
'        Call CEkey.SetFs(銀行名, True)
'        Exit Sub
'    End If
'
    If Not IsNumeric(支払日.Text) Or 支払日.Text = "" Then
        MsgBox "支払日を選択してください。", vbExclamation
        Call CEkey.SetFs(支払日, True)
        Exit Sub
    End If
'
    If 営業日.Text = "" Or (営業日.Text < "0" Or 営業日.Text > "1") Then
        MsgBox "営業日を選択してください。", vbExclamation
        Call CEkey.SetFs(営業日, True)
        Exit Sub
    End If
'
    If 利息区分.Text = "" Or (利息区分.Text < "1" Or 利息区分.Text > "2") Then
        MsgBox "利息区分を選択してください。", vbExclamation
        Call CEkey.SetFs(利息区分, True)
        Exit Sub
    End If
'
    If 利息日数.Text = "" Or (利息日数.Text < "0" Or 利息日数.Text > "1") Then
        MsgBox "利息計算日数を選択してください。", vbExclamation
        Call CEkey.SetFs(利息日数, True)
        Exit Sub
    End If
'
    If 利息支払.Text = "" Or (利息支払.Text < "0" Or 利息支払.Text > "1") Then
        MsgBox "利息支払方法を選択してください。", vbExclamation
        Call CEkey.SetFs(利息支払, True)
        Exit Sub
    End If
'
'    If 利息控除.Text = "" Or (利息控除.Text < "0" Or 利息控除.Text > "3") Then
    If 利息控除.Text = "" Or (利息控除.Text < "0" Or 利息控除.Text > "4") Then
        MsgBox "利息控除区分を選択してください。", vbExclamation
        Call CEkey.SetFs(利息控除, True)
        Exit Sub
    End If
'
    If 金利計算.Text = "" Or (金利計算.Text < "0" Or 金利計算.Text > "1") Then
        MsgBox "金利計算年間日数を選択してください。", vbExclamation
        Call CEkey.SetFs(金利計算, True)
        Exit Sub
    End If
'
    '-----------------------------------------
    '               支払日check
    '-----------------------------------------
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAB020_支払区分マスタ"
    wstr = wstr + " Where 支払日 =" & 支払日.Text
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            MsgBox "支払日を選択してください。", vbExclamation
            Call CEkey.SetFs(支払日, True)
            Exit Sub
        End If
    wRs.Close
    Set wRs = Nothing
    
'
    '2018/02/02
    ' =========================================
    '            銀行マスタ 更新処理
    ' =========================================
'    wRet = False
'    GRet = 金融機関名CHECK
'    If GRet <> True Then
'        GRet = MsgBox("金融機関名と銀行名を変更します。よろしいですか？", vbYesNo + vbExclamation)
'        If GRet = vbNo Then
'            Call CEkey.SetFs(金融機関名, True)
'            Exit Sub
'        End If
'        wRet = True
'    End If
        
    ' =========================================
    '            銀行マスタ 更新処理
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA040_銀行マスタ"
    wstr = wstr + " Where 銀行番号 = '" & L_銀行番号.Caption & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            wRs.AddNew
            
            金融機関番号 = LTrim(P8.FCStr(金融機関番号))
            金融機関名 = LTrim(P8.FCStr(金融機関名))
            支店番号 = LTrim(P8.FCStr(支店番号))
            支店名 = LTrim(P8.FCStr(支店名))
            
            If 支店番号 <> "" Then
                L_銀行番号.Caption = 金融機関番号 & "-" & 支店番号
            Else
                L_銀行番号.Caption = 金融機関番号
            End If
            
            wRs("銀行番号") = L_銀行番号.Caption
            wRs("銀行名") = 金融機関名 & Space(1) & 支店名
            
            wslog = "追加"
        End If
     
        金融機関名 = LTrim(P8.FCStr(金融機関名))
        支店名 = LTrim(P8.FCStr(支店名))
    
        wRs("銀行番号") = L_銀行番号.Caption
        wRs("銀行名") = 金融機関名 & Space(1) & 支店名
        
        wRs("金融機関番号") = 金融機関番号
        wRs("金融機関名") = 金融機関名
        wRs("支店番号") = 支店番号
        wRs("支店名") = 支店名
        
        wRs("支払日") = P8.FCStr(支払日.Text)
        wRs("営業日区分") = P8.FCStr(営業日.Text)
        wRs("利息区分") = P8.FCStr(利息区分.Text)
        wRs("利息計算日数区分") = P8.FCStr(利息日数.Text)
        wRs("利息支払方法") = P8.FCStr(利息支払.Text)
        wRs("利息控除区分") = P8.FCStr(利息控除.Text)
        wRs("金利計算年間日数") = P8.FCStr(金利計算.Text)

        wRs("預金種別") = P8.FCStr(預金種別.Text)
        wRs("口座番号") = P8.FCStr(口座番号.Text)
        
        wRs("取消フラグ") = P8.FCStr(削除.Value)
 
        wRs.Update
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '           金融機関名更新処理
    ' =========================================
'    If wRet = True Then
'        wstr = ""
'        wstr = wstr + "Select * "
'        wstr = wstr + " From DAAA040_銀行マスタ"
'        wstr = wstr + " Where 金融機関番号 = '" & P8.FCStr(金融機関番号) & "'"
'        wstr = wstr + " Order by 銀行番号"
'        Call AdoRecordsetOpen(GDb, wRs, wstr)
'        Do Until wRs.eof
'
'            wRs("金融機関名") = LTrim(P8.FCStr(金融機関名))
'            wRs("銀行名") = LTrim(P8.FCStr(金融機関名)) & Space(1) & wRs("支店名")
'
'            wRs.Update
'
'            wRs.MoveNext
'        Loop
'
'        wRs.Close
'        Set wRs = Nothing
'    End If
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
    GLogStr = "銀行番号=" & P8.FCStr(L_銀行番号.Caption) & ","
    GLogStr = GLogStr & "銀行名=" & P8.FCStr(L_銀行名.Caption) & ","
    GLogStr = GLogStr & "金融機関番号=" & P8.FCStr(金融機関番号.Text) & ","
    GLogStr = GLogStr & "金融機関名=" & P8.FCStr(金融機関名.Text) & ","
    GLogStr = GLogStr & "支店番号=" & P8.FCStr(支店番号.Text) & ","
    GLogStr = GLogStr & "支店名=" & P8.FCStr(支店名.Text) & ","
    GLogStr = GLogStr & "支払日=" & P8.FCStr(支払日.Text) & ","
    GLogStr = GLogStr & "営業日区分=" & P8.FCStr(営業日.Text) & ","
    GLogStr = GLogStr & "利息区分=" & P8.FCStr(利息区分.Text) & ","
    GLogStr = GLogStr & "利息計算日数区分=" & P8.FCStr(利息日数.Text) & ","
    GLogStr = GLogStr & "利息支払方法=" & P8.FCStr(利息支払.Text) & ","
    GLogStr = GLogStr & "利息控除区分=" & P8.FCStr(利息控除.Text) & ","
    GLogStr = GLogStr & "金利計算年間日数=" & P8.FCStr(金利計算.Text) & ","
    GLogStr = GLogStr & "預金種別=" & P8.FCStr(預金種別.Text) & ","
    GLogStr = GLogStr & "口座番号=" & P8.FCStr(口座番号.Text) & ","
    GLogStr = GLogStr & "削除=" & P8.FCStr(削除.Value)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
'
    ' =========================================
    '                テーブル変更
    ' =========================================
    Call MAA030_銀行マスタ設定
'
    Adodc1.Refresh
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(金融機関番号, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました", vbInformation
'
    Call UNLOAD_借入金FRM
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

'------------------------------------------------
' 金融機関名CHECK
'------------------------------------------------
Private Function 金融機関名CHECK()
'
    Dim FLG_ERR As Boolean
'
    On Error GoTo 金融機関名CHECK_ERR
'
    金融機関名CHECK = False
'
    FLG_ERR = False
    
    wstr = ""
    wstr = wstr + "Select 銀行番号,金融機関番号,金融機関名"
    wstr = wstr + " From DAAA040_銀行マスタ"
    wstr = wstr + " Where 金融機関番号 = '" & P8.FCStr(金融機関番号) & "'"
    wstr = wstr + " And 銀行番号 <> '" & P8.FCStr(L_銀行番号.Caption) & "'"
    wstr = wstr + " Order by 銀行番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
    
        If wRs("金融機関名") <> P8.FCStr(金融機関名) Then
            FLG_ERR = True
            Exit Do
        End If
    
        wRs.MoveNext
    Loop
    
    wRs.Close
    Set wRs = Nothing
'
    If FLG_ERR = True Then
        Exit Function
    End If
'
    金融機関名CHECK = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
金融機関名CHECK_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金融機関名CHECK() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

''------------------------------------------------
'' 削除_Click
''------------------------------------------------
'Private Sub 削除_Click()
''
'    If P8.FCStr(L_銀行番号.Caption) = "" Then
'        Exit Sub
'    End If
''
'    GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
'    If GRet = vbNo Then
'        Exit Sub
'    End If
''
'    wstr = ""
'    wstr = wstr & "Delete * From DAAA040_銀行マスタ"
'    wstr = wstr & " Where 銀行番号='" & P8.FCStr(L_銀行番号.Caption) & "'"
'    GDb.Execute wstr
'
'    DoEvents
''
'    ' =========================================
'    '               LOG_WRITE
'    ' =========================================
'    GLogStr = "銀行番号=" & P8.FCStr(L_銀行番号.Caption)
'    'Call MXA030_LOG_WRITE("銀行登録", "削除", GLogStr)
''
'    L_銀行番号.Caption = ""
''
'    Adodc1.Refresh
''
'    ' =========================================
'    '               画面セット
'    ' =========================================
'    Call 登録後初期セット
'    Call CEkey.SetFs(銀行番号, True)
''
'    ' =========================================
'    '               メッセージ
'    ' =========================================
'    メッセージ = "削除処理は終了しました"
''
'End Sub
'
'@001 ADD
'------------------------------------------------
' 銀行番号変更
'------------------------------------------------
Private Sub 銀行番号変更()
'
    Dim wRet As Boolean
    Dim wslog As String
'
    On Error GoTo 銀行番号変更_ERR
'
    If P8.FCStr(変更金融機関番号) = "" Then
        MsgBox "変更金融機関番号が未入力です。", vbExclamation
        Call CEkey.SetFs(変更金融機関番号, True)
        Exit Sub
    End If

    If 変更金融機関名 = "" Then
        MsgBox "変更金融機関名が未入力です。", vbExclamation
        Call CEkey.SetFs(変更金融機関名, True)
        Exit Sub
    End If
    
    If P8.FCStr(変更支店番号) = "" Then
        GRet = MsgBox("変更支店番号が未入力です。登録しますか？", vbYesNo + vbExclamation)
        If GRet = vbNo Then
            Call CEkey.SetFs(変更支店番号, True)
            Exit Sub
        End If
    End If

'    If 変更支店名 = "" Then
'        MsgBox "変更支店名が未入力です。", vbExclamation
'        Call CEkey.SetFs(変更支店名, True)
'        Exit Sub
'    End If
'
    GRet = MsgBox("銀行マスタ、借入金データの銀行番号を変更します。" + vbCrLf + "よろしいですか？", vbYesNo + vbQuestion)
    If GRet = vbNo Then
        Call 登録後初期セット
        Exit Sub
    End If
'
    '重複チェック
    wstr = ""
    wstr = wstr + "Select Count(*) As カウント From DAAA040_銀行マスタ"
    wstr = wstr + " Where 銀行番号 = '" & P8.FCStr(L_変更銀行番号.Caption) & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs("カウント") > 0 Then
        GRet = MsgBox("銀行番号が重複しています。", vbOKOnly + vbExclamation)
        
            wRs.Close
            Set wRs = Nothing
                
        Exit Sub
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    'DAAA040_銀行マスタ
    wstr = ""
    wstr = wstr + "Select * From DAAA040_銀行マスタ"
    wstr = wstr + " Where 銀行番号 = '" & P8.FCStr(L_銀行番号.Caption) & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
   
         wRs("金融機関番号") = P8.FCStr(変更金融機関番号.Text)
         wRs("支店番号") = P8.FCStr(変更支店番号.Text)
         wRs("銀行番号") = P8.FCStr(L_変更銀行番号.Caption)
         wRs("金融機関名") = P8.FCStr(変更金融機関名)
         wRs("支店名") = P8.FCStr(変更支店名.Text)
         wRs("銀行名") = P8.FCStr(L_変更銀行名.Caption)
    
        wRs.Update
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    'DBDA010_借入金
    wstr = "UPDATE DBDA010_借入金"
    wstr = wstr & " SET"
    wstr = wstr & " 銀行番号='" + P8.FCStr(L_変更銀行番号.Caption) + "'"
    wstr = wstr + " Where 銀行番号='" + P8.FCStr(L_銀行番号.Caption) + "'"
    
    GDb.Execute wstr
'
    GRet = MsgBox("銀行番号を変更しました。", vbOKOnly + vbInformation)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
銀行番号変更_ERR:
    pERR_MES = pPROGRAM_ID + "/ 銀行番号変更() でエラー" + vbCrLf + vbCrLf + _
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
    '----------< DataGrid Close >----------------------------------------------
    If Not DataGrid1.DataSource Is Nothing Then
        Set DataGrid1.DataSource = Nothing
    End If
    
    Adodc1.Recordset.Close
'
    Unload Me
End Sub


