Attribute VB_Name = "GenbaReport"
Option Explicit
'
' 点検表から、日ごとの点検報告書を作る（Excel のシートと PDF）。
' 点検表を読むところは GenbaNorm の ReadTable を呼ぶ（Python 版と全225セル一致を確かめた部分）。
' 報告書のセルから読み戻した内容を CSV にも出し、compare.py --report で Python 版の正解と突き合わせる。
'
' 使い方
'   1. 新しい空白のブックで VBA エディタ（Alt+F11）→ ファイル → ファイルのインポート
'      → GenbaNorm.bas と GenbaReport.bas の2つ
'   2. ReportSelfTest を実行（Alt+F8）。「ReportSelfTest OK」と出ることを確かめる
'   3. MakeReports を実行 → 1つ目に点検表の xlsx、2つ目にひな形（点検報告書_ひな形.xlsx）を選ぶ。
'      逆に選ぶと、何も作らずに止まる
'   4. 点検表と同じフォルダに、次ができる
'        <元の名前>_点検報告書 フォルダ … 日ごとの PDF と、全日のシートをまとめた 点検報告書.xlsx
'        <元の名前>_報告書_VBA.csv   … 報告書のセルから読み戻した内容（突き合わせ用）
'
' 元の点検表とひな形は読み取り専用で開く。書き換えない。
'
' ============================================================
' ひな形の決まり
' ============================================================
'
' ・次の {名前} を書いたセルに値が入る。場所・書式・見た目は、ひな形の側で自由に変えてよい
'     {日付} {部署} {担当} {点検回数}
'     {異常} {条件付き許容} {未記入} {判別不能} {対象外} {記入あり}（その日の件数）
' ・{点検結果の見出し} は、点検結果の表の見出しを横1行に書く（列の名前は点検表の見出しから取る）
' ・{確認が必要な項目} と {点検結果} は表の左上のセル。件数に合わせて下に行を足し、
'   足した行には、そのセルの行の書式（罫線も）を写してから書く。表の行には結合セルを使わない
' ・知らない {名前} があると止まる。報告書に {名前} の文字のまま残さないため
'
' ============================================================
' 書くときに避けた罠
' ============================================================
'
' ・日付は文字ではなく日付の値で書く。表示の形（2026年9月1日（火））は、ひな形のセルの書式で決まる
' ・文字は先に文字列書式にしてから書く（表も1マスも）。しないと "14.0" が数値の 14 になり、
'   報告書に「14」と出る。数字や日付に見える担当・部署の名前も、同じように化ける
' ・行を足すのは、ひな形を上から1行ずつ読みながら。上の表で行を足すと下の表の位置がずれるが、
'   下の {点検結果} はその時点でまだ読んでいないので、ずれた後の位置で見つかる
' ・行を足すだけでは罫線が写らなかった（PDF で、表の2行目から下の罫線が消えていた）。
'   足した行には、ひな形の行を書式ごと写してから値を書く
' ・表を書いた行の右側を読むときは、表の列を飛ばす。表の中身が {名前} の形でも差し込み欄と見間違えない
' ・ファイルを2つ選ばせると、逆に選ばれる（実際に起きた。ひな形を点検表として読み、
'   日付不明の報告書が1枚だけ黙ってできた）。点検欄の列が無い点検表と、差し込み欄の無いひな形は、
'   何も作らないうちに止める
' ・SelfTest という名前は GenbaNorm と重なるので、こちらは ReportSelfTest にした
'
' ============================================================
' 既知の限界
' ============================================================
'
' ・報告書は日付が出てきた順に並ぶ（並べ替えない）。日付が読めない行は「日付不明」の1枚にまとめる
' ・1日を1枚に収めるよう、ひな形の印刷設定を「横1ページ・縦1ページ」にしてある。項目が多い日は文字が小さくなる
' ・PDF の見た目（はみ出し・読みやすさ）は数字では確かめていない。目で見る

' --- 入口 -------------------------------------------------------------------

Public Sub MakeReports()
    Dim src As Variant, tpl As Variant
    src = Application.GetOpenFilename("Excel ファイル (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , _
                                      "1/2 点検表（点検記録の Excel）を選んでください")
    If VarType(src) = vbBoolean Then Exit Sub
    tpl = Application.GetOpenFilename("Excel ファイル (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , _
                                      "2/2 報告書のひな形（点検報告書_ひな形.xlsx）を選んでください")
    If VarType(tpl) = vbBoolean Then Exit Sub
    MakeReportsFile CStr(src), CStr(tpl)
End Sub

Public Sub MakeReportsFile(ByVal src As String, ByVal tpl As String)
    Dim res As Variant, n As Long, nCheck As Long
    Dim days() As String, nDay As Long, d As Long, r As Long
    Dim wbTpl As Workbook, wbOut As Workbook, ws As Worksheet
    Dim stem As String, folder As String, rec() As Variant, nRec As Long
    Dim alerts As Boolean, updating As Boolean, msg As String

    alerts = Application.DisplayAlerts
    updating = Application.ScreenUpdating
    On Error GoTo Fail

    ' 点検表を読む。点検欄の列が無いファイル（ひな形など）は、ReadTable がここで止める
    res = ReadTable(src, n, nCheck)
    If n = 0 Then Err.Raise vbObjectError + 513, "GenbaReport", "点検の行が1行もありません: " & LeafName(src)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ひな形を開く。差し込み欄が1つも無いファイル（点検表など）は、フォルダも作らないうちに止める
    Set wbTpl = Workbooks.Open(Filename:=tpl, ReadOnly:=True)
    If Not HasKey(wbTpl.Worksheets(1)) Then
        Err.Raise vbObjectError + 515, "GenbaReport", _
                  "ひな形に {日付} などの差し込み欄が見つかりません。" & _
                  "2つ目にひな形を選んだか確かめてください: " & LeafName(tpl)
    End If

    ' 日付の一覧（出てきた順）
    ReDim days(1 To n)
    For r = 1 To n
        If Not Seen(days, nDay, CStr(res(r, 2))) Then
            nDay = nDay + 1
            days(nDay) = res(r, 2)
        End If
    Next

    stem = Left$(src, InStrRev(src, ".") - 1)
    folder = stem & "_点検報告書"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder

    For d = 1 To nDay
        ' ひな形のシートを写し、1日1枚にする
        If wbOut Is Nothing Then
            wbTpl.Worksheets(1).Copy
            Set wbOut = ActiveWorkbook
        Else
            wbTpl.Worksheets(1).Copy After:=wbOut.Worksheets(wbOut.Worksheets.Count)
        End If
        Set ws = wbOut.Worksheets(wbOut.Worksheets.Count)
        ws.Name = SheetName(days(d))

        FillReport ws, res, n, nCheck, days(d), rec, nRec

        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & "\" & ws.Name & ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=False, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
    Next

    wbTpl.Close SaveChanges:=False
    Set wbTpl = Nothing
    wbOut.Activate                    ' シートより先にブックを前に出す。別のブックが前だとシートを出せない
    wbOut.Worksheets(1).Activate      ' 開いたときに最初の日が出るように
    wbOut.SaveAs Filename:=folder & "\点検報告書.xlsx", FileFormat:=51   ' 51 = xlsx
    wbOut.Close SaveChanges:=False
    Set wbOut = Nothing

    WriteRec stem & "_報告書_VBA.csv", rec, nRec

    Application.DisplayAlerts = alerts
    Application.ScreenUpdating = updating
    MsgBox "報告書を作りました。" & vbCrLf & vbCrLf & _
           "点検表: " & LeafName(src) & vbCrLf & _
           "ひな形: " & LeafName(tpl) & vbCrLf & _
           "点検の行: " & n & vbCrLf & _
           "報告書: " & nDay & " 日分（PDF " & nDay & " 件）" & vbCrLf & vbCrLf & _
           "出力: " & folder, vbInformation, "GenbaReport"
    Exit Sub

Fail:
    msg = Err.Description & vbCrLf & "（エラー番号 " & Err.Number & "）"
    On Error Resume Next
    If Not wbTpl Is Nothing Then wbTpl.Close SaveChanges:=False
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = alerts
    Application.ScreenUpdating = updating
    MsgBox "止まりました。" & vbCrLf & vbCrLf & msg, vbExclamation, "GenbaReport"
End Sub

' シートに {名前} のセルが1つでもあるか
Private Function HasKey(ByVal ws As Worksheet) As Boolean
    Dim cell As Range
    For Each cell In ws.UsedRange.Cells
        If Len(KeyOf(cell.Value2)) > 0 Then HasKey = True: Exit Function
    Next
End Function

' "H:\...\点検記録_2026-09.xlsx" → "点検記録_2026-09.xlsx"
Private Function LeafName(ByVal path As String) As String
    LeafName = Mid$(path, InStrRev(path, "\") + 1)
End Function

' --- 1日分を書く ---------------------------------------------------------------

Private Sub FillReport(ByVal ws As Worksheet, ByVal res As Variant, ByVal n As Long, _
                       ByVal nCheck As Long, ByVal dayKey As String, _
                       ByRef rec() As Variant, ByRef nRec As Long)
    Dim keys As Variant, vals(0 To 9) As Variant, cnt(0 To 5) As Long
    Dim items() As Variant, nItem As Long, head() As Variant, detail() As Variant, nRow As Long
    Dim people As String, depts As String, pri As Variant
    Dim r As Long, c As Long, k As Long, i As Long, j As Long
    Dim lastRow As Long, lastCol As Long, used As Long, key As String, skipTo As Long

    ' その日の担当・部署・件数
    For r = 1 To n
        If res(r, 2) = dayKey Then
            nRow = nRow + 1
            people = AddUnique(people, CStr(res(r, 3)))
            depts = AddUnique(depts, CStr(res(r, 4)))
            For k = 1 To nCheck
                cnt(KindNo(res(r, 5 + k))) = cnt(KindNo(res(r, 5 + k))) + 1
            Next
        End If
    Next

    ' 確認が必要な項目。区分ごとに、異常 → 条件付き許容 → 未記入 → 判別不能 の順
    ReDim items(1 To nRow * nCheck + 1, 1 To 5)   ' +1 は点検欄が0列でも ReDim を通すため
    For Each pri In Array("異常", "条件付き許容", "未記入", "判別不能")
        For r = 1 To n
            If res(r, 2) = dayKey Then
                For k = 1 To nCheck
                    If res(r, 5 + k) = pri Then
                        nItem = nItem + 1
                        items(nItem, 1) = pri
                        items(nItem, 2) = res(0, 5 + k)
                        items(nItem, 3) = res(r, 5 + nCheck + k)
                        items(nItem, 4) = res(r, 1)
                        items(nItem, 5) = res(r, 3)
                    End If
                Next
            End If
        Next
    Next

    ' 点検結果。見出しの1行と点検の行を分けて持つ。ひな形で、見出しと中身の書式を別にできるように
    ReDim head(1 To 1, 1 To 4 + nCheck)
    head(1, 1) = "原票の行": head(1, 2) = "担当": head(1, 3) = res(0, 5)
    For k = 1 To nCheck
        head(1, 3 + k) = res(0, 5 + k)
    Next
    head(1, 4 + nCheck) = "備考"
    ReDim detail(1 To nRow, 1 To 4 + nCheck)
    i = 0
    For r = 1 To n
        If res(r, 2) = dayKey Then
            i = i + 1
            detail(i, 1) = res(r, 1): detail(i, 2) = res(r, 3): detail(i, 3) = res(r, 5)
            For k = 1 To nCheck
                detail(i, 3 + k) = res(r, 5 + k)
            Next
            detail(i, 4 + nCheck) = res(r, 6 + 2 * nCheck)
        End If
    Next

    ' 1マスに入る値。件数の並びは GenbaNorm と同じ（記入あり・異常・条件付き許容・対象外・未記入・判別不能）
    keys = Array("日付", "部署", "担当", "点検回数", _
                 "記入あり", "異常", "条件付き許容", "対象外", "未記入", "判別不能")
    vals(0) = DateValueOf(dayKey)
    vals(1) = depts
    vals(2) = people
    vals(3) = nRow
    For i = 0 To 5
        vals(4 + i) = cnt(i)
    Next

    ' ひな形を上から1行ずつ読み、{名前} のセルに書く。表を書いたら、その行数ぶん先へ進む
    r = 1
    Do
        With ws.UsedRange
            lastRow = .Row + .Rows.Count - 1
            lastCol = .Column + .Columns.Count - 1
        End With
        If r > lastRow Then Exit Do
        used = 1
        skipTo = 0
        For c = 1 To lastCol
            ' 表を書いた列は読み飛ばす。表の中身が {名前} の形でも、差し込み欄と見間違えないため
            If c > skipTo Then key = KeyOf(ws.Cells(r, c).Value2) Else key = ""
            If Len(key) > 0 Then
                Select Case key
                Case "確認が必要な項目"
                    i = PutTable(ws, r, c, items, nItem, 5, dayKey & "/" & key, rec, nRec, 1)
                    skipTo = c + 4
                Case "点検結果の見出し"
                    i = PutTable(ws, r, c, head, 1, 4 + nCheck, dayKey & "/点検結果", rec, nRec, 0)
                    skipTo = c + 3 + nCheck
                Case "点検結果"
                    i = PutTable(ws, r, c, detail, nRow, 4 + nCheck, dayKey & "/" & key, rec, nRec, 1)
                    skipTo = c + 3 + nCheck
                Case Else
                    j = IndexOf(keys, key)
                    If j < 0 Then
                        Err.Raise vbObjectError + 514, "GenbaReport", _
                                  "ひな形に知らない差し込み欄があります: {" & key & "}"
                    End If
                    If VarType(vals(j)) = vbString Then ws.Cells(r, c).NumberFormat = "@"
                    ws.Cells(r, c).Value = vals(j)
                    AddRec rec, nRec, Array(dayKey & "/" & key, CellText(ws.Cells(r, c).Value))
                    i = 1
                End Select
                If i > used Then used = i
            End If
        Next
        r = r + used
    Loop
End Sub

' 表を書く。行が足りなければ、左上のセルの行の下に足し、その行の書式（罫線も）を写す。
' 書いたあとのセルから読み戻して、突き合わせ用に残す。使った行数を返す
Private Function PutTable(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, _
                          ByVal data As Variant, ByVal nr As Long, ByVal nc As Long, _
                          ByVal prefix As String, ByRef rec() As Variant, ByRef nRec As Long, _
                          ByVal firstNo As Long) As Long
    Dim i As Long, j As Long, fields() As Variant

    If nr = 0 Then
        ws.Cells(r, c).Value = "なし"
        AddRec rec, nRec, Array(prefix & "/" & firstNo, CellText(ws.Cells(r, c).Value))
        PutTable = 1
        Exit Function
    End If

    If nr > 1 Then
        ws.Rows(r + 1).Resize(nr - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ' 行を足しただけでは罫線が写らなかった。値を書く前に、左上のセルの行を書式ごと写す
        ws.Cells(r, c).Resize(1, nc).Copy Destination:=ws.Cells(r + 1, c).Resize(nr - 1, nc)
    End If
    ' 配列が範囲より大きいときは、範囲に収まる左上の部分だけが書かれる
    With ws.Cells(r, c).Resize(nr, nc)
        .NumberFormat = "@"
        .Value = data
    End With

    For i = 0 To nr - 1
        ReDim fields(0 To nc)
        fields(0) = prefix & "/" & (firstNo + i)
        For j = 1 To nc
            fields(j) = CellText(ws.Cells(r + i, c + j - 1).Value)
        Next
        AddRec rec, nRec, fields
    Next
    PutTable = nr
End Function

' --- 突き合わせ用の CSV ---------------------------------------------------------
' 1行 = 欄の名前（例: 2026-09-01/異常）と、そのセルの値。表は1行ずつ（例: 2026-09-01/点検結果/1）

Private Sub AddRec(ByRef rec() As Variant, ByRef nRec As Long, ByVal fields As Variant)
    nRec = nRec + 1
    ReDim Preserve rec(1 To nRec)
    rec(nRec) = fields
End Sub

Private Sub WriteRec(ByVal path As String, ByRef rec() As Variant, ByVal nRec As Long)
    Dim wb As Workbook, ws As Worksheet, i As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Cells.NumberFormat = "@"
    ws.Range("A1:B1").Value = Array("欄", "値")
    For i = 1 To nRec
        ws.Cells(i + 1, 1).Resize(1, UBound(rec(i)) + 1).Value = rec(i)
    Next
    wb.SaveAs Filename:=path, FileFormat:=62   ' 62 = CSV UTF-8
    wb.Close SaveChanges:=False
End Sub

' --- 小さな部品 -----------------------------------------------------------------

' "{日付}" → "日付"。{名前} の形でなければ ""
Private Function KeyOf(ByVal v As Variant) As String
    If VarType(v) <> vbString Then Exit Function
    If Len(v) > 2 And Left$(v, 1) = "{" And Right$(v, 1) = "}" Then KeyOf = Mid$(v, 2, Len(v) - 2)
End Function

' 「山田太郎、佐藤花子」のように、出てきた順に重ねずにつなぐ
Private Function AddUnique(ByVal joined As String, ByVal s As String) As String
    AddUnique = joined
    If Len(s) = 0 Then Exit Function
    If InStr("、" & joined & "、", "、" & s & "、") > 0 Then Exit Function
    If Len(joined) = 0 Then AddUnique = s Else AddUnique = joined & "、" & s
End Function

' "2026-09-01" → 日付の値。形が違えば（「読めない」など）文字のまま返す
Private Function DateValueOf(ByVal s As String) As Variant
    DateValueOf = s
    If Len(s) <> 10 Or Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    DateValueOf = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2)))
End Function

' セルの値を、CSV に書く文字にする。日付は 2026-09-01 の形
Private Function CellText(ByVal v As Variant) As String
    If VarType(v) = vbDate Then
        CellText = Format$(v, "yyyy-mm-dd")
    ElseIf IsError(v) Or IsEmpty(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

Private Function SheetName(ByVal dayKey As String) As String
    If dayKey = "読めない" Then SheetName = "日付不明" Else SheetName = dayKey
End Function

Private Function KindNo(ByVal kind As String) As Long
    Select Case kind
    Case "記入あり": KindNo = 0
    Case "異常": KindNo = 1
    Case "条件付き許容": KindNo = 2
    Case "対象外": KindNo = 3
    Case "未記入": KindNo = 4
    Case Else: KindNo = 5
    End Select
End Function

Private Function IndexOf(ByVal list As Variant, ByVal s As String) As Long
    Dim i As Long
    For i = LBound(list) To UBound(list)
        If list(i) = s Then IndexOf = i: Exit Function
    Next
    IndexOf = -1
End Function

Private Function Seen(ByRef list() As String, ByVal n As Long, ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To n
        If list(i) = s Then Seen = True: Exit Function
    Next
End Function

' --- 自己確認 -----------------------------------------------------------------

Public Sub ReportSelfTest()
    Dim fails As String

    Check fails, "差し込み欄", KeyOf("{日付}"), "日付"
    Check fails, "中身の無い波かっこ", KeyOf("{}"), ""
    Check fails, "ふつうの文字", KeyOf("日付"), ""
    Check fails, "数値のセル", KeyOf(12), ""

    Check fails, "担当を重ねない", AddUnique(AddUnique(AddUnique("", "山田太郎"), "佐藤花子"), "山田太郎"), "山田太郎、佐藤花子"
    Check fails, "空の担当は足さない", AddUnique("山田太郎", ""), "山田太郎"
    Check fails, "名前の一部を同じと見ない", AddUnique("山田太郎", "山田"), "山田太郎、山田"

    Check fails, "日付の値", CellText(DateValueOf("2026-09-01")), "2026-09-01"
    Check fails, "日付の型", CStr(VarType(DateValueOf("2026-09-01")) = vbDate), "True"
    Check fails, "読めない日付", CellText(DateValueOf("読めない")), "読めない"
    Check fails, "空のセル", CellText(Empty), ""
    Check fails, "件数のセル", CellText(3#), "3"

    Check fails, "日付不明のシート名", SheetName("読めない"), "日付不明"
    Check fails, "シート名", SheetName("2026-09-01"), "2026-09-01"
    Check fails, "区分の番号", CStr(KindNo("未記入")), "4"
    Check fails, "欄の番号", CStr(IndexOf(Array("日付", "部署"), "部署")), "1"
    Check fails, "知らない欄", CStr(IndexOf(Array("日付", "部署"), "日時")), "-1"

    If Len(fails) = 0 Then
        MsgBox "ReportSelfTest OK", vbInformation, "GenbaReport"
    Else
        MsgBox "ReportSelfTest NG" & vbCrLf & fails, vbExclamation, "GenbaReport"
    End If
End Sub

Private Sub Check(ByRef fails As String, ByVal label As String, ByVal got As String, ByVal want As String)
    If got <> want Then fails = fails & label & ": 期待 [" & want & "] / 実際 [" & got & "]" & vbCrLf
End Sub
