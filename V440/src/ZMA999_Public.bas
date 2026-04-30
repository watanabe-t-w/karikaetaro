Attribute VB_Name = "ZMA999_PUBLIC"

'######################################################
'#
'#           アプリケーション　Ｐｕｂｌｉｃ変数
'#
'######################################################

'帳票構造体---------------------------------2011.11.14 By m.mino

Public G借入明細表照会 As MAA910_借入明細表照会


'---------------------------------------------------------------

Type Type_Print
    帳票名 As String
    
    テキスト_01 As String
    テキスト_02 As String
    
    コンボ_01 As String
    コンボ_02 As String
    コンボ_03 As String
    コンボ_04 As String
    コンボ_05 As String
'    コンボ_06 As String
'    コンボ_07 As String
'    コンボ_08 As String
'    コンボ_09 As String
'    コンボ_10 As String
'    コンボ_11 As String
'    コンボ_12 As String
        
    推移 As String
    選択 As String
    作業 As String
    実績 As String
    集計 As String
    指定 As String
    
    連結売上  As String
    連結売上2  As String
    売上  As String
    売上2 As String
    借入  As String
    借入2 As String
    設備  As String
    設備2 As String
    金融  As String
    金融2 As String
    リス  As String
    リス2 As String
    
    借入金管理区分 As Integer
    詳細表示 As Integer
    CSV As Integer
    千円単位 As Integer
    金利SM As Integer
      
    設備R  As String
    設備R2 As String

    チェック_01 As Integer
    チェック_02 As Integer
    チェック_03 As Integer
    チェック_04 As Integer

    C_種別 As String
    C_部門 As String
    C_金融 As String
    C_銀行 As String

    S_種別 As String
    S_部門 As String
    S_金融 As String
    S_銀行 As String
    S_金利 As String
    S_利息 As String
    
    NewPage1 As Integer
    NewPage2 As Integer
    NewPage3 As Integer
    NewPage4 As Integer
    
End Type
'
Public wレコード() As MGG010_Typeレコード
Type MGG010_Typeレコード
    xName()  As String
    xValue() As Variant
End Type
'
Public GRpt As Type_Print
Public GCsvName As String

Public GSys As MAA910_システム
Public G基本情報 As MAA010_基本情報ファイル
Public Gコントロール As MAA020_コントロールファイル
Public G償却率マスタ() As MAA040_償却率
Public G保証率マスタ As MAA050_保証率
Public G税率マスタ() As MAA060_税率
Public G科目マスタ() As MAA910_科目マスタ
Public G独算() As MAA910_企業情報
Public G連結() As MAA910_企業情報
'wk
Public G年度計画 As MAA910_年度計画
Public G売上計画テーブル() As MAA910_売上計画テーブル
Public G設備計画テーブル() As MAA910_設備計画テーブル
Public G借入金テーブル() As MAA910_借入金テーブル
Public G利息未払前払テーブル() As MAA910_利息未払前払テーブル
Public G借入金入力() As MAA910_借入金入力
Public G社債入力() As MAA910_借入金入力
Public Gリーステーブル() As MAA910_リーステーブル   '08/04/08 V182
Public Gリース入力() As MAA910_リース入力           '08/04/08 V182
Public Gリース消費税総額 As Double                  '08/04/08 V182
Public G会議計画() As MAA910_会議計画

Public G基準金利() As MAA200_基準金利レート                   '2015/06/01 借入金時価評価
Public G適用金利() As MAA910_適用金利
Public G時価明細() As MAA910_時価評価明細

'2017/12/01 祝日マスタ ADD
Public G祝日マスタ() As MAA090_祝日
Public GCal() As MAA090_KakuninMsg  '2018/03 追加

Public G借換 As MAA910_借入金借換                 '09/09/01
Public G借現 As MAA910_借入金借換現状

'######################################################
'#
'#                クラスモジュール
'#
'######################################################
Public P8 As New CZC010_CommandX
Public CEkey As New CZC020_EnterKeyP8
Public C休日 As New CAA010_休日
Public C年月日 As New CAA020_年月日

'######################################################
'#
'#               共通　Ｐｕｂｌｉｃ変数
'#
'######################################################
Public GDefaultPrinter  As String  'Default Printer の名前
Public GSystemDir       As String  'SYSYTEM の Drive 名
Public GCurDir          As String  'App.Path 名（このプログラムの存在場所)
Public GSerDir          As String  '
Public GMyComputerName  As String
Public GSerComputerName As String
Public GCsvPath          As String  '2018/05/30

Public GPwd As String               'gdb Open 時のパスワード
Public GP2  As String               'Serial data P2
Public G実績共有 As String          '実績データ共有区分 0:単独 1:共有
Public GTbl_売上実績 As String
Public Gcsv_Jiseki As String
Public Gcsv_Setubi As String

Public GVerNo As String             'バージョン
Public GFcap As String              'Form.caption
Public GKeyName As String           '企業名Key
Public GCoName As String            '企業名
Public GProduct As String           '製品名

Public G金額1 As Double           '04/07/23 V120
Public G金額2 As Double           '04/07/23 V120
Public G金額3 As Double           '04/07/23 V120
Public G期末在庫 As Double        '5/5/16 V127
Public G仮払消費税22F As Integer    ' 07/03/02 V180

Public G償却最終年月 As Date        ' 07/02/16 V180 設備減価償却最終年月
Public G償却額整合性F As Integer    ' 07/02/16 V180 設備購入額>=減価償却＋調整償却＋特別償却　のCHECK
                                    ' 0(正常)  1(異常)
'Public G支店 As Integer             '5/7/28 V128 支店の時　過去の資金自動調達　有効
'Public G本部振替 As Boolean
'Public G基幹調整 As Boolean

Public G実績調整 As Integer
Public G販売売仕 As Integer

Public G会議 As String             '06/02/01 V150
Public G会議資金調達 As String

Public G利益調整(10) As Double    '5/12/31 V131

Public G管理年月 As Variant         '5/9/8 V129 借入金管理
Public G実績年月 As Variant         '5/9/8 V129 借入金管理

Public G金利SM As Boolean           '2010/11/01
Public GSstrt帳票Msg As String

Public G決算日(1) As Date           '2015/06/01 v430時価評価

Public GSqlOlegdb   As New ADODB.Connection
Public GServerName As String
Public GSQLgdbName  As String
Public GSQLUid     As String
Public GSQLPwd     As String

Public GDb     As New ADODB.Connection
Public GDb2    As New ADODB.Connection
Public GDb3    As New ADODB.Connection
Public GDbName As String

Public GErrSwich As Boolean
Public GMyFuncName As String
Public GMyFuncNames As New Collection
Public GMyControlNames As New Collection

Public Ws As Workspace   ' ワーク

Public GStr   As String  'ワーク
Public GStr_1 As String
Public GStr_2 As String
Public GStr_3 As String

Public GInt1 As Integer
Public GInt2 As Integer
Public GInt3 As Integer

Public GLong1 As Long
Public GLong2 As Long

Public GDbl1 As Double
Public GDbl2 As Double

Public GDate1 As Date
Public GDate2 As Date

Public GDate1利息対象年月日 As Variant
Public GDate2利息対象年月日 As Variant
Public GDate利息対象年月日 As Variant

Public GVar1 As Variant
Public GVar2 As Variant
Public GVar3 As Variant

Public GRet As Long

Public GWhere  As String            'ワーク(ＳＱＬで使用）
Public GSelect As String            'ワーク(ＳＱＬで使用）
Public GFrom   As String            'ワーク(ＳＱＬで使用）
Public GJoin   As String            'ワーク(ＳＱＬで使用）
Public GOrder  As String            'ワーク(ＳＱＬで使用）

Public GFind As Integer             'ワーク（While 文 で Loop時 使用）

Public GUserID As String            '2010/01/01
Public GUserKen As String           '2010/01/01
Public GLogStr As String            '2010/01/01

Public GInputDateKbn As String      '2011/11/17 追加 0:和暦入出力、1:西暦入出力
Public Gfmt年 As String
Public Gfmt年月 As String
Public Gfmt年月日 As String

Public Gfmtcsv年 As String          '2012/10/23 追加
Public Gfmtcsv年月 As String
Public Gfmtcsv年月日 As String

Public GPriCnt As Integer
Public GForm As Object

Public GstrDenNo As String          '神姫バス
Public GstrDenNo2 As String         '神姫バス

'######################################################
'#
'#               共通　Public Const
'#
'######################################################
Public Const GMain = "List.mdb"     'ListDB名
Public Const GTemp = "K000.mdb"     'ﾃﾝﾌﾟﾚｰﾄMDB名

Public Const G_TOP = 0              'フォームの出現位置Y
Public Const G_LEFT = 2325          'フォームの出現位置X

'######################################################
'#
'#               Color
'#
'######################################################
Public Const C_White = &HFFFFFF     '白
Public Const C_Gray = &HE0E0E0      '灰色
Public Const C_DGray = &H404040     '灰色
Public Const C_Black = &H0&         '黒
Public Const C_Blue = &HFF0000      '青
Public Const C_Red = &HC0&          '赤
Public Const C_Pink = &HC0C0FF      'ピンク
Public Const C_Orange = &H80FF&     'オレンジ
Public Const C_Yellow = &HC0FFFF    '黄色
Public Const C_Green = &H8000&      '緑
Public Const C_LGreen = &HD6DBBD    '薄緑

Public Const C_PGreen = &HC0C000    '薄緑
Public Const C_PSky = &HE2B685      '薄青

'######################################################
'#
'#               CSV FileName
'#
'######################################################
'金剛石 or 借換たろう！
'Public Const Gcsv_DirName = "金剛石CSV"
Public Const Gcsv_DirName = "借換たろうCSV"
Public Const Gcsv_Shu1 = "Kcsv_科目集計1.csv"
Public Const Gcsv_Shu2 = "Kcsv_科目集計2.csv"
Public Const Gcsv_Set1 = "Kcsv_設備計画1.csv"
Public Const Gcsv_Set2 = "Kcsv_設備計画2.csv"
Public Const Gcsv_Ktl1 = "Kcsv_科目テーブル1.csv"
Public Const Gcsv_Ktl2 = "Kcsv_科目テーブル2.csv"

Public Const GJiseki_DirName = "金剛石実績変換"
Public Const GJDbName = "金剛石実績変換.mdb"
'
'----------< PUBLIC CONSTANT >--------------------------------------------
Public Const pSERIAL_NAME = "SERIAL.FIL"                                  ' Serial File Name
Public Const pSYSLOG_NAME = "FSYSLOG.FIL"                                 ' System Log
Public Const pERRLOG_NAME = "ERRLOG.FIL"                                  ' Error Log
Public Const pMSGBOX_TYTLE = "ｱﾌﾟﾘｹｰｼｮﾝｴﾗｰ"                               ' Message
Public Const pMSGBOX_OPERAT = "ｵﾍﾟﾚｰｼｮﾝｴﾗｰ"                               ' Message
Public Const pMSGBOX_INFO = "ｵﾍﾟﾚｰｼｮﾝｰ"                                   ' Message
Public Const pREC_LOGCAP = 1000                                           ' System Log
Public Const pMAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
'
'----------< PUBLIC ARGUMENT >--------------------------------------------
Public pERR_MES As String                                                 ' Message
Public pERR_RET As Long                                                   ' Message
Public pREC_SYSLOG As TYPE_SYSLOG                                         ' System Log Buffer
Public DriveN As String
Public kf As String
Public pVER As String
'

'######################################################
'#
'#           Ｄｅｃｌａｒｅ　Ｆｕｎｃｔｉｏｎ , S u b
'#
'######################################################
' =========================================
'       Sleep  (単位 ミリ秒)
' =========================================
Public Declare Sub Sleep Lib "KERNEL32" (ByVal DWMILLISECOND As Long)

' =========================================
'       ShellExecute
' =========================================
Public Const SW_SHOWNORMAL = 1 'ｳｲﾝﾄﾞｳをｱｸﾃｨﾌﾞ化し、表示する。ｳｲﾝﾄﾞｳが最小化または最大化されているときには、元のｻｲｽﾞと位置に復元する

Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameteGRs As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' =========================================
'       GetComputerName
' =========================================
Public Const Max_ComputerName_Length  As Long = 15

Public Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" _
    (ByVal lpgstrfer As String, nSize As Long) As Long

'
'---------< API DECLARE >-------------------------------------------------------------
'Public Declare Function GetLogicalDrives Lib "KERNEL32" () As Long
'
Public Declare Function GetVolumeInformation Lib "KERNEL32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, _
     ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
     lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
     ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'
Public Declare Function DeleteFile Lib "KERNEL32" Alias "DeleteFileA" _
    (ByVal FileName As String) As Long

Public Declare Function FindFirstFile Lib "KERNEL32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile Lib "KERNEL32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * pMAX_PATH
    cAlternate As String * 14
End Type
'

' フォルダの参照
Public Const CSIDL_DESKTOP = &H0
Public Const BIF_RETURNONLYFSDIRS = &H1

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
    (lpBROWSEINFO As BROWSEINFO) As Long
    
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long
    
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Type BROWSEINFO
      hWndOwner As Long
      pidlRoot As Long
      pszDisplayName As String
      lpszTitle As String
      ulFlags As Long
      lpfn As Long
      lParam As Long
      iImage As Long
End Type
'

' ファイルのバージョン製品名取得
Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long

Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" _
    (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
    (dest As Any, ByVal Source As Long, ByVal length As Long)

Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Type CODEPAGE
    lngLOW As Integer
    lngHIGH As Integer
End Type
'

'csvfile
'クラス名・キャプションタイトル名を与えてウインドウのハンドルを取得
Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

