' ------------------------------
' SUNCO年次更新処理スクリプト
' ------------------------------
Option Explicit
'
Dim connSrc
Dim connTgt
Dim sql
Dim rs
Dim fso
Dim basePath
Dim fieldNames
Dim fieldCount
Dim value
Dim values
Dim i
Dim connList
Dim rs2
Dim listDbPath
Dim listDBNAME
Dim TBDB
Dim fieldName
Dim val
Dim adSchemaTables
Dim rsSchema
Dim backupPath
Dim timestamp
Dim backupFile

' ----- DB・テーブル定義 -----
Dim srcDbName
Dim tgtDbName
Dim TBK
Dim TBK_TEMP
Dim TBKTR
Dim TBKTR_TEMP
Dim TBKTR2
Dim TBKTR2_TEMP
Dim TBKU1
Dim TBKU1_TEMP
Dim TBKU2
Dim TBKU2_TEMP
Dim MG
Dim keyFieldName
Dim DbPassword
Dim baseDate
Dim yearOffset
Dim insertCountTBK
Dim insertCountTBKTR
Dim insertCountTBKTR2
Dim insertCountTBKU1
Dim insertCountTBKU2
Dim recordsAffected
Dim logFile
' ----- パス -----
Dim srcDbPath
Dim tgtDbPath

' -------------------------
' 設定値
' -------------------------
listDBNAME = "List"
srcDbName= "サンコーインダストリー"
tgtDbName= "全データ_サンコーインダストリー"
TBDB = "DAAA070_企業名マスタ"
TBK = "DBDA010_借入金"
TBK_TEMP = "DBDA010_借入金_TEMP"
TBKTR = "DBDA010_借入金明細TR"
TBKTR_TEMP = "DBDA010_借入金明細TR_TEMP"
TBKTR2 = "DBDA010_借入金明細TR2"
TBKTR2_TEMP = "DBDA010_借入金明細TR2_TEMP"
TBKU1 = "DBDA010_借入金内入1"
TBKU1_TEMP = "DBDA010_借入金内入1_TEMP"
TBKU2 = "DBDA010_借入金内入2"
TBKU2_TEMP = "DBDA010_借入金内入2_TEMP"
MG = "DAAA040_銀行マスタ"
keyFieldName = "借入番号"
DbPassword = "inkinhkheshh2IHPDKPI"
adSchemaTables = 20

' -------------------------
' 同一フォルダのパス取得
' -------------------------
Set fso = CreateObject("Scripting.FileSystemObject")
basePath = fso.GetParentFolderName(WScript.ScriptFullName)

Set logFile = fso.OpenTextFile(basePath & "\年次更新_log.txt", 8, True)

srcDbPath = basePath & "\" & srcDbName & ".mdb"
tgtDbPath = basePath & "\" & tgtDbName & ".mdb"
listDbPath = basePath & "\" & listDBNAME & ".mdb"

If not fso.FileExists(srcDbPath) then
  WScript.Echo "元DBが見つかりません: " & srcDbPath
  WScript.Quit
end if

If Month(Date) >= 4 Then
    yearOffset = 4
Else
    yearOffset = 5
End If
baseDate = "#" & (Year(Date) - yearOffset) & "/03/01#"
If MsgBox("基準日：" & baseDate & " より前のデータを移行" & vbCrLf & "完済データは " & tgtDbPath & " に移動します。", vbokcancel, "年次更新") = vbCancel Then
    WScript.Quit
End If
logFile.WriteLine "基準日：" & baseDate & " より前のデータを移行 Start:" & Now()

' -------------------------
' バックアップ作成
' -------------------------
backupPath = basePath & "\backup"
If Not fso.FolderExists(backupPath) Then
    fso.CreateFolder backupPath
End If
timestamp = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & Right("0" & Hour(Now()), 2) & Right("0" & Minute(Now()), 2) & Right("0" & Second(Now()), 2)
backupFile = backupPath & "\" & srcDbName & "_" & timestamp & ".mdb"
If fso.FileExists(srcDbPath) then
  fso.CopyFile srcDbPath, backupFile
end if
backupFile = backupPath & "\" & tgtDbName & "_" & timestamp & ".mdb"
If fso.FileExists(tgtDbPath) then
  fso.CopyFile tgtDbPath, backupFile
end if
backupFile = backupPath & "\" & listDBNAME & "_" & timestamp & ".mdb"
If fso.FileExists(listDbPath) then
  fso.CopyFile listDbPath, backupFile
end if

' -------------------------
' DBコピー（tgtDbが存在しない場合）
' -------------------------
If not fso.FileExists(tgtDbPath) then
  fso.CopyFile srcDbPath, tgtDbPath
  Set connTgt = CreateObject("ADODB.Connection")
  connTgt.Open _
    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=" & tgtDbPath & ";" & _
    "Jet OLEDB:Database Password=" & DbPassword & ";"
      sql = ""
      sql = sql & "UPDATE [" & TBDB & "] "
      sql = sql & "SET 企業名Key = '" & tgtDbName & "', "
      sql = sql & "企業名 = '" & tgtDbName & "', "
      sql = sql & "DB名 = '" & tgtDbName & ".mdb', "
      sql = sql & "最新処理日 = #" & FormatDateTime(Now(), vbGeneralDate) & "#, "
      sql = sql & "作成日 = #" & FormatDateTime(Now(), vbGeneralDate) & "#, "
      sql = sql & "復元日 = NULL, "
      sql = sql & "削除日 = NULL, "
      sql = sql & "端末コンピュータ名 = ''"
      connTgt.Execute sql
  connTgt.Close
  logFile.WriteLine "DBをコピーしました: " & tgtDbPath
end if

' -------------------------
' list.mdb操作
' -------------------------
Set connList = CreateObject("ADODB.Connection")
connList.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & listDbPath & ";"

Set rs = connList.Execute("SELECT * FROM [" & TBDB & "] WHERE [企業名Key] = '" & srcDbName & "'")
If not rs.EOF then
  rs.MoveFirst
  Set rs2 = connList.Execute("SELECT * FROM [" & TBDB & "] WHERE [企業名Key] = '" & tgtDbName & "'")
  If rs2.EOF then
    ' INSERT
    fieldNames = ""
    values = ""
    for i = 0 to rs.Fields.Count - 1
      fieldName = rs.Fields(i).Name
      val = rs.Fields(i).Value
      if fieldName = "企業名Key" then val = tgtDbName
      if fieldName = "企業名" then val = tgtDbName
      if fieldName = "DB名" then val = tgtDbName & ".mdb"
      if fieldName = "最新処理日" then val = FormatDateTime(Now(), vbGeneralDate)
      if fieldName = "作成日" then val = FormatDateTime(Now(), vbGeneralDate)
      if fieldName = "保存日" then val = Null
      if fieldName = "復元日" then val = Null
      if fieldName = "削除日" then val = Null
      if fieldName = "端末コンピュータ名" then val = ""
      fieldNames = fieldNames & "[" & fieldName & "]"
      if IsNull(val) then
        values = values & "NULL"
      ElseIf rs.Fields(i).Type = 202 or rs.Fields(i).Type = 203 or rs.Fields(i).Type = 129 then
        values = values & "'" & Replace(val, "'", "''") & "'"
      ElseIf rs.Fields(i).Type = 7 or rs.Fields(i).Type = 133 or rs.Fields(i).Type = 134 or rs.Fields(i).Type = 135 then
        values = values & "#" & val & "#"
      else
        values = values & val
      end if
      if i < rs.Fields.Count - 1 then
        fieldNames = fieldNames & ", "
        values = values & ", "
      end if
    next
    sql = "INSERT INTO [" & TBDB & "] (" & fieldNames & ") VALUES (" & values & ")"
    connList.Execute sql
    logFile.WriteLine "list.mdbに新規追加しました: " & tgtDbName
  else
    ' UPDATE
    sql = "UPDATE [" & TBDB & "] SET "
    for i = 0 to rs.Fields.Count - 1
      fieldName = rs.Fields(i).Name
      val = rs.Fields(i).Value
      if fieldName = "企業名Key" then val = tgtDbName
      if fieldName = "企業名" then val = tgtDbName
      if fieldName = "DB名" then val = tgtDbName & ".mdb"
      if fieldName = "最新処理日" then val = FormatDateTime(Date(), vbShortDate)
      if fieldName = "作成日" then val = FormatDateTime(Date(), vbShortDate)
      if fieldName = "保存日" then val = Null
      if fieldName = "復元日" then val = Null
      if fieldName = "削除日" then val = Null
      if fieldName = "端末コンピュータ名" then val = ""
      sql = sql & "[" & fieldName & "] = "
      if IsNull(val) then
        sql = sql & "NULL"
      ElseIf rs.Fields(i).Type = 202 or rs.Fields(i).Type = 203 or rs.Fields(i).Type = 129 then
        sql = sql & "'" & Replace(val, "'", "''") & "'"
      ElseIf rs.Fields(i).Type = 7 or rs.Fields(i).Type = 133 or rs.Fields(i).Type = 134 or rs.Fields(i).Type = 135 then
        sql = sql & "#" & val & "#"
      else
        sql = sql & val
      end if
      if i < rs.Fields.Count - 1 then sql = sql & ", "
    next
    sql = sql & " WHERE [企業名Key] = '" & tgtDbName & "'"
    connList.Execute sql
    logFile.WriteLine "List.mdbを更新しました: " & tgtDbName
  end if
  rs2.Close
  Set rs2 = Nothing
end if
rs.Close
Set rs = Nothing
connList.Close
Set connList = Nothing

' -------------------------
' DB接続
' -------------------------
Set connSrc = CreateObject("ADODB.Connection")
connSrc.Open _
 "Provider=Microsoft.ACE.OLEDB.12.0;" & _
 "Data Source=" & srcDbPath & ";" & _
 "Jet OLEDB:Database Password=" & DbPassword & ";"

Set connTgt = CreateObject("ADODB.Connection")
connTgt.Open _
 "Provider=Microsoft.ACE.OLEDB.12.0;" & _
 "Data Source=" & tgtDbPath & ";" & _
 "Jet OLEDB:Database Password=" & DbPassword & ";"

' -------------------------
' TEMPテーブル作成（存在しない場合）
' -------------------------
Set rsSchema = connTgt.OpenSchema(adSchemaTables)
rsSchema.MoveFirst

' DBDA010_借入金_TEMP
' 空の TEMP テーブルを作成する WHERE 1=0
rsSchema.Find "TABLE_NAME = '" & TBK_TEMP & "'"
If rsSchema.EOF then
  connTgt.Execute "SELECT * INTO [" & TBK_TEMP & "] FROM [" & TBK & "] WHERE 1=0"
  logFile.WriteLine "TEMPテーブル [" & TBK_TEMP & "] を作成しました。"
end if
rsSchema.MoveFirst

' DBDA010_借入金明細TR_TEMP
rsSchema.Find "TABLE_NAME = '" & TBKTR_TEMP & "'"
If rsSchema.EOF then
  connTgt.Execute "SELECT * INTO [" & TBKTR_TEMP & "] FROM [" & TBKTR & "] WHERE 1=0"
  logFile.WriteLine "TEMPテーブル [" & TBKTR_TEMP & "] を作成しました。"
end if
rsSchema.MoveFirst

' DBDA010_借入金明細TR2_TEMP
rsSchema.Find "TABLE_NAME = '" & TBKTR2_TEMP & "'"
If rsSchema.EOF then
  connTgt.Execute "SELECT * INTO [" & TBKTR2_TEMP & "] FROM [" & TBKTR2 & "] WHERE 1=0"
  logFile.WriteLine "TEMPテーブル [" & TBKTR2_TEMP & "] を作成しました。"
end if
rsSchema.MoveFirst

' DBDA010_借入金内入1_TEMP
rsSchema.Find "TABLE_NAME = '" & TBKU1_TEMP & "'"
If rsSchema.EOF then
  connTgt.Execute "SELECT * INTO [" & TBKU1_TEMP & "] FROM [" & TBKU1 & "] WHERE 1=0"
  logFile.WriteLine "TEMPテーブル [" & TBKU1_TEMP & "] を作成しました。"
end if
rsSchema.MoveFirst

' DBDA010_借入金内入2_TEMP
rsSchema.Find "TABLE_NAME = '" & TBKU2_TEMP & "'"
If rsSchema.EOF then
  connTgt.Execute "SELECT * INTO [" & TBKU2_TEMP & "] FROM [" & TBKU2 & "] WHERE 1=0"
  logFile.WriteLine "TEMPテーブル [" & TBKU2_TEMP & "] を作成しました。"
end if

rsSchema.Close
Set rsSchema = Nothing

' -------------------------
' DELETE(全データDB)
' -------------------------
' DBDA010_借入金_TEMP
sql = ""
sql = sql & "DELETE * FROM [" & TBK_TEMP & "] "
connTgt.Execute sql

' DBDA010_借入金明細TR_TEMP
sql = ""
sql = sql & "DELETE * FROM [" & TBKTR_TEMP & "] "
connTgt.Execute sql

' DBDA010_借入金明細TR2_TEMP
sql = ""
sql = sql & "DELETE * FROM [" & TBKTR2_TEMP & "] "
connTgt.Execute sql

' DBDA010_借入金内入1_TEMP
sql = ""
sql = sql & "DELETE * FROM [" & TBKU1_TEMP & "] "
connTgt.Execute sql

' DBDA010_借入金内入2_TEMP
sql = ""
sql = sql & "DELETE * FROM [" & TBKU2_TEMP & "] "
connTgt.Execute sql

' DBDA010_銀行マスタ
sql = ""
sql = sql & "DELETE * FROM [" & MG & "] "
connTgt.Execute sql

' -------------------------
' INSERT（差分追加を全データ.TEMPテーブルに追加）
' -------------------------
' DBDA010_借入金
Call InsertByRecordset(connSrc, connTgt, TBK, TBK_TEMP, logFile)
' DBDA010_借入金明細TR
Call InsertByRecordset(connSrc, connTgt, TBKTR, TBKTR_TEMP, logFile)
' DBDA010_借入金明細TR2
Call InsertByRecordset(connSrc, connTgt, TBKTR2, TBKTR2_TEMP, logFile)
' DBDA010_借入金内入1
Call InsertByRecordset(connSrc, connTgt, TBKU1, TBKU1_TEMP, logFile)
' DBDA010_借入金内入2
Call InsertByRecordset(connSrc, connTgt, TBKU2, TBKU2_TEMP, logFile)

' -------------------------
' INSERT（TEMPテーブルから追加）
' -------------------------
' DBDA010_借入金
Call InsertFromTemp(TBK_TEMP,  TBK,  connTgt, keyFieldName, logFile)
' DBDA010_借入金明細TR
Call InsertFromTemp(TBKTR_TEMP, TBKTR, connTgt, keyFieldName, logFile)
' DBDA010_借入金明細TR2
Call InsertFromTemp(TBKTR2_TEMP, TBKTR2, connTgt, keyFieldName, logFile)
' DBDA010_借入金内入2
Call InsertFromTemp(TBKU1_TEMP, TBKU1, connTgt, keyFieldName, logFile)
' DBDA010_借入金内入2
Call InsertFromTemp(TBKU2_TEMP, TBKU2, connTgt, keyFieldName, logFile)

' -------------------------
' INSERT（全データに追加）
' -------------------------
sql = ""
Set rs = connSrc.Execute("SELECT * FROM [" & MG & "] WHERE 1=0")
fieldNames = ""
fieldCount = rs.Fields.Count
for i = 0 to fieldCount - 1
  fieldNames = fieldNames & "[" & rs.Fields(i).Name & "]"
  if i < fieldCount - 1 then fieldNames = fieldNames & ", "
next
rs.Close
Set rs = connSrc.Execute("SELECT * FROM [" & MG & "]")
while not rs.EOF
  values = ""
  for i = 0 to fieldCount - 1
    value = rs.Fields(i).Value
    if IsNull(value) then
      values = values & "NULL"
    ElseIf rs.Fields(i).Type = 202 or rs.Fields(i).Type = 203 or rs.Fields(i).Type = 129 then
      values = values & "'" & Replace(value, "'", "''") & "'"
    ElseIf rs.Fields(i).Type = 7 Then
      values = values & "#" & Format(value, "yyyy/mm/dd hh:nn:ss") & "#"
    else
      values = values & value
    end if
    if i < fieldCount - 1 then values = values & ", "
  next
  sql = "INSERT INTO [" & MG & "] (" & fieldNames & ") VALUES (" & values & ")"
  connTgt.Execute sql
  rs.MoveNext
wend
rs.Close
Set rs = Nothing

' -------------------------
' DELETE（完済データ削除、条件年度内は残す）
' -------------------------
' DBDA010_借入金
sql = ""
sql = sql & "DELETE K.* "
sql = sql & "FROM [" & TBK & "] AS K "
sql = sql & "WHERE NOT ("
sql = sql & " (K.解約実行日 IS NULL "
sql = sql & " AND K.最終返済実行日 >= " & baseDate & ")"
sql = sql & " OR "
sql = sql & " (K.解約実行日 IS NOT NULL "
sql = sql & " AND K.解約実行日 >= " & baseDate & ")"
sql = sql & ");"
connSrc.Execute sql

' DBDA010_借入金明細TR
sql = ""
sql = sql & "DELETE TR.* "
sql = sql & "FROM [" & TBKTR & "] AS TR "
sql = sql & "LEFT JOIN [" & TBK & "] AS K ON TR.借入番号 = K.借入番号 "
sql = sql & "WHERE"
sql = sql & " K.借入番号 IS NULL;"
connSrc.Execute sql

' DBDA010_借入金明細TR2
sql = ""
sql = sql & "DELETE TR2.* "
sql = sql & "FROM [" & TBKTR2 & "] AS TR2 "
sql = sql & "LEFT JOIN [" & TBK & "] AS K ON TR2.借入番号 = K.借入番号 "
sql = sql & "WHERE"
sql = sql & " K.借入番号 IS NULL;"
connSrc.Execute sql

' DBDA010_借入金内入1
sql = ""
sql = sql & "DELETE U1.* "
sql = sql & "FROM [" & TBKU1 & "] AS U1 "
sql = sql & "LEFT JOIN [" & TBK & "] AS K ON U1.借入番号 = K.借入番号 "
sql = sql & "WHERE"
sql = sql & " K.借入番号 IS NULL;"
connSrc.Execute sql

' DBDA010_借入金内入2
sql = ""
sql = sql & "DELETE U2.* "
sql = sql & "FROM [" & TBKU2 & "] AS U2 "
sql = sql & "LEFT JOIN [" & TBK & "] AS K ON U2.借入番号 = K.借入番号 "
sql = sql & "WHERE"
sql = sql & " K.借入番号 IS NULL;"
connSrc.Execute sql

' -------------------------
' 終了処理
' -------------------------
logFile.WriteLine "処理が完了しました。 End:" & Now()

connSrc.Close
connTgt.Close
Set connSrc = Nothing
Set connTgt = Nothing
logFile.Close
Set logFile = Nothing
Set fso = Nothing

MsgBox "処理が完了しました。" & vbCrLf & "（基準日：" & baseDate & " より前のデータを移行しました。）" , vbOKOnly, "年次更新完了"

' -------------------------
' INSERT（TEMPテーブルから追加）
' -------------------------
Sub InsertByRecordset( _
    srcConn, _
    tgtConn, _
    srcTableName, _
    tgtTableName, _
    logFile _
)

  Dim rs
  Dim fieldNames
  Dim fieldCount
  Dim values
  Dim value
  Dim sql
  Dim i
  Dim recordsAffected
  Dim insertCount

  '-------------------------
  ' フィールド定義取得（src）
  '-------------------------
  Set rs = srcConn.Execute("SELECT * FROM [" & srcTableName & "] WHERE 1=0")

  fieldCount = rs.Fields.Count
  fieldNames = ""

  For i = 0 To fieldCount - 1
    fieldNames = fieldNames & "[" & rs.Fields(i).Name & "]"
    If i < fieldCount - 1 Then fieldNames = fieldNames & ", "
  Next

  rs.Close
  Set rs = Nothing

  '-------------------------
  ' データ取得（src）
  '-------------------------
  Set rs = srcConn.Execute("SELECT * FROM [" & srcTableName & "]")
  insertCount = 0

  Do While Not rs.EOF
    values = ""

    For i = 0 To fieldCount - 1
      value = rs.Fields(i).Value

      If IsNull(value) Then
        values = values & "NULL"

      Else
        Select Case rs.Fields(i).Type
          Case 202, 203, 129
            values = values & "'" & Replace(value, "'", "''") & "'"

          Case 7
            values = values & "#" & _
                     FormatDateTime(value, vbGeneralDate) & "#"

          Case Else
            values = values & value
        End Select
      End If

      If i < fieldCount - 1 Then values = values & ", "
    Next

    sql = "INSERT INTO [" & tgtTableName & "] (" & _
          fieldNames & ") VALUES (" & values & ")"

    tgtConn.Execute sql, recordsAffected
    insertCount = insertCount + recordsAffected

    rs.MoveNext
  Loop

  rs.Close
  Set rs = Nothing

  logFile.WriteLine _
    "テーブル [" & tgtTableName & "] に " & insertCount & _
    " 件追加しました。（元: " & srcTableName & "）"
End Sub

' -------------------------
' INSERT（TEMPテーブルから追加）
' -------------------------
Sub InsertFromTemp( _
    tempTableName, _
    mainTableName, _
    conn, _
    keyFieldName, _
    logFile _
)

  Dim sql
  Dim recordsAffected

  sql = ""
  sql = sql & "INSERT INTO [" & mainTableName & "] "
  sql = sql & "SELECT T.* "
  sql = sql & "FROM [" & tempTableName & "] AS T "
  sql = sql & "LEFT JOIN [" & mainTableName & "] AS M "
  sql = sql & "ON T.[" & keyFieldName & "] = M.[" & keyFieldName & "] "
  sql = sql & "WHERE M.[" & keyFieldName & "] IS NULL;"

  conn.Execute sql, recordsAffected

  logFile.WriteLine _
    "テーブル [" & mainTableName & "] に " & _
    recordsAffected & " 件追加しました。（元: " & tempTableName & "）"

End Sub
