Sub TANAOROSHITORIKOMI()
'v1.0.0 20231216 作成
'v1.0.1 20240202 行をクリアではなく削除に変更
    Dim WS As Worksheet

    ' ワークシートをループして、オートフィルタが設定されているか確認する
    For Each WS In ThisWorkbook.Worksheets
        If WS.AutoFilterMode = True Then
            WS.AutoFilterMode = False    ' オートフィルタを解除する
        End If
    Next WS

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("棚卸DATA")
    Dim qt As QueryTable
    Dim FileToOpen As Variant

    Dim WS1_LR_AC As Long
    WS1_LR_AC = WS1.Cells(Rows.Count, 1).End(xlUp).Row

    '既存データ削除
    If WS1_LR_AC > 1 Then

        WS1.Range(Cells(2, 1), Cells(WS1_LR_AC, 10)).Delete

    End If
    'WS1.Columns("B:C").NumberFormatLocal = "@"

    ' CSVファイルを開く
    FileToOpen = Application.GetOpenFilename("CSVファイル (*.csv),*.csv", , "CSVファイルを選択してください")
    If Not IsEmpty(FileToOpen) Then
        ' クエリテーブルを作成
        Set qt = WS1.QueryTables.Add(Connection:="TEXT;" & FileToOpen, Destination:=WS1.Range("$A$2"))
        With qt
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 932
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 3, 3, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With

        ' クエリテーブルを削除
        For Each qt In WS1.QueryTables
            qt.Delete
        Next qt

        ' 接続を削除
        Dim cn As WorkbookConnection
        For Each cn In ThisWorkbook.Connections
            cn.Delete
        Next cn
    End If

    WS1_LR_AC = WS1.Cells(Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To WS1_LR_AC
        WS1.Cells(i, 2) = Format(WS1.Cells(i, 2), "'00000000")
        WS1.Cells(i, 3) = Format(WS1.Cells(i, 3), "'000000")
        WS1.Cells(i, 9).Value = 0
        If i > 2 Then
            WS1.Cells(i, 1).Formula = "=A" & i - 1
        End If
    Next i

    WS1.Columns.AutoFit

    '新シートを作って閉じる
    Dim newWorkbook As Workbook
    Dim currentSheet As Worksheet
    Dim fileName As String

    ' コピー元のシートを指定（例: Sheet1）
    Set currentSheet = ThisWorkbook.Sheets("棚卸DATA")

    ' 新しいブックを作成
    Set newWorkbook = Workbooks.Add

    ' シートをコピー
    currentSheet.Copy Before:=newWorkbook.Sheets(1)

    ' 日付を取得してファイル名に追加
    fileName = Format(Now, "yyyymmdd") & "_棚卸データ.xlsx"

    ' 新しいブックを保存
    newWorkbook.SaveAs fileName:=fileName

    ' ブックを閉じる
    newWorkbook.Close

End Sub
