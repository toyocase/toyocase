'20231216 作成　小原征史　v1.0.0 変数、自作関数は大文字、その他はVBSの言語？

Set objFSO = CreateObject("Scripting.FileSystemObject")

INPUT_FILE_PATH = "\\192.168.2.250\zaiko\SEND\F04.CSV"
'INPUT_FILE_PATH = "\\192.168.2.2\share\IS200\SYUKKA"
'WScript.Echo INPUT_FILE_PATH
'現在の西暦年月日時間を取得
'WScript.Echo Now
CURRENT_DATE_TIME = Replace(Replace(Replace(Now, "/", ""), ":", ""), " ", "")

'WScript.Echo CURRENT_DATE_TIME
' 新しいファイル名を生成
NEW_FILE_NAME = "JISSEKI_" & mid(CURRENT_DATE_TIME,1,8)&"_"&Right(CURRENT_DATE_TIME, 6) & ".csv"
OUTPUT_FILE_PATH = "\\192.168.2.2\share\IS200\SYUKKA\" & NEW_FILE_NAME

' ファイルからテキストデータを読み込む
Set INPUT_FILE = objFSO.OpenTextFile(INPUT_FILE_PATH, 1)
CONTENT = INPUT_FILE.ReadAll
INPUT_FILE.Close

' 各行を処理
LINES = Split(CONTENT, vbCrLf)
ReDim MODIFIED_LINES(UBound(LINES) - 1)

For i = 0 To UBound(LINES) - 1
    MODIFIED_LINES(i) = MODIFY_ROW(LINES(i))
Next

Function MODIFY_ROW(ROW)
    ' 行をカンマで分割
    COLUMNS = Split(ROW, ",")

    ' 列の入れ替え (1, 2, 3 列はそのまま、9 列目を 4 列目に、20 列目を 5 列目に)
    COLUMNS(4) = COLUMNS(20) ' 20列目を 5 列目に
    COLUMNS(3) = COLUMNS(8)  ' 9列目を 4 列目に

    ' 各列をダブルクオートで括る
    For j = LBound(COLUMNS) To UBound(COLUMNS)
        COLUMNS(j) = """" & COLUMNS(j) & """"
    Next
    
    ' 新しい行を構築 (5列目まで)
    If UBound(COLUMNS) >= 4 Then
        ReDim Preserve COLUMNS(4)
    End If
    MODIFIED_ROW = Join(COLUMNS, ",")

    ' 5列目までのデータを含む新しい行を返す
    MODIFY_ROW = MODIFIED_ROW
End Function

Set OUTPUT_FILE = objFSO.CreateTextFile(OUTPUT_FILE_PATH, True)
OUTPUT_FILE.Write Join(MODIFIED_LINES, vbCrLf)
OUTPUT_FILE.Close
