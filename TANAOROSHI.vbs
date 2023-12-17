'20231216 作成　小原征史　v1.0.0 変数、自作関数は大文字、その他はVBSの言語？

Set objFSO = CreateObject("Scripting.FileSystemObject")

INPUT_FILE_PATH = "\\192.168.2.250\zaiko\SEND\F05.CSV"

' 現在の西暦年月日時間を取得
' WScript.Echo Now
CURRENT_DATE_TIME = Replace(Replace(Replace(Now, "/", ""), ":", ""), " ", "")

' 新しいファイル名を生成
NEW_FILE_NAME = "TANAOROSHI_" & mid(CURRENT_DATE_TIME,1,8)&"_"&Right(CURRENT_DATE_TIME, 6) & ".csv"
OUTPUT_FILE_PATH = "\\192.168.2.2\share\IS200\TANAOROSHI\" & NEW_FILE_NAME

' ファイルからテキストデータを読み込む
Set INPUT_FILE = objFSO.OpenTextFile(INPUT_FILE_PATH, 1)
CONTENT = INPUT_FILE.ReadAll
INPUT_FILE.Close

' 各行を処理
LINES = Split(CONTENT, vbCrLf)
ReDim MODIFIED_LINES(UBound(LINES) - 1)

' MODIFY_ROW関数: 行の変更処理を行う
Function MODIFY_ROW(ROW)
    ' 行をカンマで分割
    COLUMNS = Split(ROW, ",")
    formattedDate = GetFormattedDate()

    ' 列の入れ替え 
    COLUMNS(9) = COLUMNS(3)
    COLUMNS(3) = COLUMNS(0)
    COLUMNS(0) = formattedDate
    COLUMNS(1) = "00000001"
    COLUMNS(2) = "000241"
    COLUMNS(4) = ""
    COLUMNS(5) = ""
    COLUMNS(6) = ""
    COLUMNS(7) = ""
    COLUMNS(8) = ""

    ' 各列をダブルクオートで括る
    For j = LBound(COLUMNS) To UBound(COLUMNS)
        COLUMNS(j) = """" & COLUMNS(j) & """"
    Next

    ' 新しい行を構築 (9列目まで)
    If UBound(COLUMNS) >= 9 Then
        ReDim Preserve COLUMNS(9)
    End If
    ' Join関数: 配列の要素を指定した区切り文字で結合する
    MODIFIED_ROW = Join(COLUMNS, ",")

    ' MODIFY_ROW関数の戻り値
    MODIFY_ROW = MODIFIED_ROW
End Function

' GetFormattedDate関数: 今日の日付を整形して返す
Function GetFormattedDate()
    'today = Date

    ' 現在の年と月を取得
currentYear = Year(Date)
currentMonth = Month(Date)
lastDay = GetLastDayOfMonth(currentYear, currentMonth)

    'currentYear = Year(today)
    'currentMonth = Right("0" & Month(today), 2)
    currentDay = Right("0" & lastDay, 2)
    formattedDate = currentYear & currentMonth & currentDay
    ' GetFormattedDate関数の戻り値
    GetFormattedDate = formattedDate
End Function

Function GetLastDayOfMonth(year, month)
    ' 月の最後の日を取得する関数
    Dim firstDayOfNextMonth
    firstDayOfNextMonth = DateSerial(year, month + 1, 1)
    GetLastDayOfMonth = DateAdd("d", -1, firstDayOfNextMonth)
End Function


' 各行の変更処理を実施
For i = 0 To UBound(LINES) - 1
    MODIFIED_LINES(i) = MODIFY_ROW(LINES(i))
Next

' ファイル出力処理
Set OUTPUT_FILE = objFSO.CreateTextFile(OUTPUT_FILE_PATH, True)
OUTPUT_FILE.Write Join(MODIFIED_LINES, vbCrLf)
OUTPUT_FILE.Close
