$ErrorActionPreference = 'Stop'
$root = 'D:\project\ProjectionManager'
$outDir = Join-Path $root 'docs'
$docx = Join-Path $outDir 'Glory_榮耀_投影控制台_操作手冊_V2.3.docx'
$pdf = Join-Path $outDir 'Glory_榮耀_投影控制台_操作手冊_V2.3.pdf'
$logo = Join-Path $root 'Assets\glory_logo.png'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()
$sel = $word.Selection

# 版面：A4、四邊 2.5 cm
$doc.PageSetup.PageWidth = $word.CentimetersToPoints(21)
$doc.PageSetup.PageHeight = $word.CentimetersToPoints(29.7)
$doc.PageSetup.TopMargin = $word.CentimetersToPoints(2.5)
$doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2.5)
$doc.PageSetup.LeftMargin = $word.CentimetersToPoints(2.5)
$doc.PageSetup.RightMargin = $word.CentimetersToPoints(2.5)

$normal = $doc.Styles.Item('內文')
$normal.Font.Name = 'Times New Roman'
$normal.Font.NameFarEast = '宋体'
$normal.Font.Size = 12
$normal.Font.Color = 0
$normal.ParagraphFormat.Alignment = 3
$normal.ParagraphFormat.LineSpacingRule = 1
$normal.ParagraphFormat.FirstLineIndent = 24
$normal.ParagraphFormat.SpaceBefore = 0
$normal.ParagraphFormat.SpaceAfter = 0

$title = $doc.Styles.Item('標題')
$title.Font.Name = 'Times New Roman'
$title.Font.NameFarEast = '黑体'
$title.Font.Size = 18
$title.Font.Bold = $true
$title.Font.Color = 0
$title.ParagraphFormat.Alignment = 1
$title.ParagraphFormat.SpaceAfter = 18

$subtitle = $doc.Styles.Item('副標題')
$subtitle.Font.Name = 'Times New Roman'
$subtitle.Font.NameFarEast = '楷体'
$subtitle.Font.Size = 12
$subtitle.Font.Color = 0
$subtitle.ParagraphFormat.Alignment = 1
$subtitle.ParagraphFormat.SpaceAfter = 12

foreach ($name in @('標題 1','標題 2','標題 3')) {
    $s = $doc.Styles.Item($name)
    $s.Font.Name = 'Times New Roman'
    $s.Font.NameFarEast = '黑体'
    $s.Font.Bold = $true
    $s.Font.Color = 0
    $s.Font.Italic = $false
    $s.ParagraphFormat.Alignment = 0
    $s.ParagraphFormat.FirstLineIndent = 0
    $s.ParagraphFormat.KeepWithNext = $true
    $s.ParagraphFormat.SpaceBefore = if ($name -eq '標題 1') { 14 } elseif ($name -eq '標題 2') { 11 } else { 9 }
    $s.ParagraphFormat.SpaceAfter = if ($name -eq '標題 1') { 6 } elseif ($name -eq '標題 2') { 5 } else { 4 }
    $s.Font.Size = if ($name -eq '標題 1') { 16 } elseif ($name -eq '標題 2') { 14 } else { 12 }
}

function H1([string]$text) {
    $sel.Style = $doc.Styles.Item('標題 1')
    $sel.TypeText($text)
    $sel.TypeParagraph()
}
function H2([string]$text) {
    $sel.Style = $doc.Styles.Item('標題 2')
    $sel.TypeText($text)
    $sel.TypeParagraph()
}
function H3([string]$text) {
    $sel.Style = $doc.Styles.Item('標題 3')
    $sel.TypeText($text)
    $sel.TypeParagraph()
}
function P([string]$text) {
    $sel.Style = $doc.Styles.Item('內文')
    $sel.TypeText($text)
    $sel.TypeParagraph()
}
function Note([string]$text) {
    $sel.Style = $doc.Styles.Item('內文')
    $sel.Font.Size = 10.5
    $sel.Font.Color = 4210752
    $sel.ParagraphFormat.FirstLineIndent = 0
    $sel.TypeText($text)
    $sel.TypeParagraph()
    $sel.Font.Size = 12
    $sel.Font.Color = 0
    $sel.ParagraphFormat.FirstLineIndent = 24
}
function Bullet([string]$text) {
    $sel.Style = $doc.Styles.Item('內文')
    $sel.ParagraphFormat.FirstLineIndent = 0
    $sel.ParagraphFormat.LeftIndent = 24
    $sel.Range.ListFormat.ApplyBulletDefault()
    $sel.TypeText($text)
    $sel.TypeParagraph()
    $sel.Range.ListFormat.RemoveNumbers()
    $sel.Style = $doc.Styles.Item('內文')
    $sel.ParagraphFormat.LeftIndent = 0
    $sel.ParagraphFormat.FirstLineIndent = 24
}

function Table([string[]]$headers, [object[]]$rows) {
    $cols = $headers.Count
    $r = $sel.Range
    $table = $doc.Tables.Add($r, $rows.Count + 1, $cols)
    $table.Borders.Enable = $true
    $table.Borders.OutsideLineStyle = 1
    $table.Borders.InsideLineStyle = 1
    $table.Borders.OutsideColor = 8421504
    $table.Borders.InsideColor = 12632256
    $table.Rows.Item(1).HeadingFormat = -1
    $table.Rows.Item(1).Shading.BackgroundPatternColor = 15921906
    for ($i = 1; $i -le $table.Rows.Count; $i++) { $table.Rows.Item($i).AllowBreakAcrossPages = 0 }
    for ($c = 1; $c -le $cols; $c++) {
        $cell = $table.Cell(1, $c)
        $cell.Range.Text = $headers[$c-1]
        $cell.Range.Font.Name = 'Times New Roman'
        $cell.Range.Font.NameFarEast = '宋体'
        $cell.Range.Font.Size = 10.5
        $cell.Range.Font.Bold = $true
        $cell.Range.ParagraphFormat.Alignment = 1
        $cell.Range.ParagraphFormat.FirstLineIndent = 0
        $cell.Range.ParagraphFormat.LeftIndent = 0
        $cell.Range.ParagraphFormat.LineSpacingRule = 0
        $cell.Range.ParagraphFormat.SpaceBefore = 0
        $cell.Range.ParagraphFormat.SpaceAfter = 0
    }
    for ($i = 0; $i -lt $rows.Count; $i++) {
        for ($c = 1; $c -le $cols; $c++) {
            $cell = $table.Cell($i + 2, $c)
            $cell.Range.Text = [string]$rows[$i][$c-1]
            $cell.Range.Font.Name = 'Times New Roman'
            $cell.Range.Font.NameFarEast = '宋体'
            $cell.Range.Font.Size = 10.5
            $cell.Range.Font.Bold = $false
            $cell.Range.ParagraphFormat.Alignment = if ($c -eq 1) { 1 } else { 0 }
            $cell.Range.ParagraphFormat.FirstLineIndent = 0
            $cell.Range.ParagraphFormat.LeftIndent = 0
            $cell.Range.ParagraphFormat.LineSpacingRule = 0
            $cell.Range.ParagraphFormat.SpaceBefore = 0
            $cell.Range.ParagraphFormat.SpaceAfter = 0
        }
    }
    for ($i = 1; $i -le $table.Rows.Count; $i++) {
        for ($c = 1; $c -le $cols; $c++) {
            $cellPara = $table.Cell($i, $c).Range.Paragraphs.Item(1)
            $cellPara.Format.KeepWithNext = if ($i -lt $table.Rows.Count) { $true } else { $false }
        }
    }
    $sel.EndKey(6) | Out-Null
    $sel.TypeParagraph()
}

# ================= 封面 =================
$sel.ParagraphFormat.Alignment = 1
$sel.Style = $doc.Styles.Item('標題')
if (Test-Path -LiteralPath $logo) {
    $pic = $sel.InlineShapes.AddPicture($logo, $false, $true)
    $pic.LockAspectRatio = -1
    $pic.Width = 150
    $sel.TypeParagraph()
}
$sel.Font.Size = 22
$sel.TypeText('Glory 榮耀')
$sel.TypeParagraph()
$sel.Font.Size = 30
$sel.TypeText('投影控制台')
$sel.TypeParagraph()
$sel.Font.Size = 24
$sel.TypeText('操作手冊')
$sel.TypeParagraph()
$sel.Style = $doc.Styles.Item('副標題')
$sel.TypeText('V2.3 交付版｜2026 年 11 月 14 日')
$sel.TypeParagraph()
$sel.Font.Size = 12
$sel.TypeText('適用於聚會現場操作、實機測試與故障排查')
$sel.TypeParagraph()
$sel.TypeText('軟件全名：Glory 榮耀 · 投影控制台')
$sel.TypeParagraph()
$sel.TypeText('交付內容：自包含執行環境、安裝包、操作手冊、持久化操作日誌')
$sel.TypeParagraph()
$sel.InsertBreak(7)

# ================= 目錄 =================
$sel.Style = $doc.Styles.Item('內文')
$sel.Font.NameFarEast = '黑体'
$sel.Font.Size = 16
$sel.Font.Bold = $true
$sel.ParagraphFormat.Alignment = 0
$sel.ParagraphFormat.FirstLineIndent = 0
$sel.TypeText('目錄')
$sel.TypeParagraph()
$sel.Font.Size = 12
$sel.Font.Bold = $false
$sel.Font.NameFarEast = '宋体'
$tocRange = $sel.Range
$doc.TablesOfContents.Add($tocRange, $true, 1, 2) | Out-Null
$sel.EndKey(6) | Out-Null
$sel.InsertBreak(7)

# ================= 正文 =================
H1('一、文件說明')
H2('（一）目的')
P('本手冊說明 Glory 榮耀投影控制台的安裝、投影操作、素材推送、手機控制、主題設定與日誌排查方法。操作員可按第一章至第九章完成現場投影；遇到問題時，直接查閱第十章。')
H2('（二）版本與適用範圍')
P('本手冊對應 Glory 榮耀投影控制台 V2.3 交付版。此版本包含黑金、深灰青綠、暖白三種介面主題；主 PPT 與素材分開控制；素材預設只在控制台預覽，必須按「推送投影」才會送到第二螢幕。')
H2('（三）名詞說明')
Bullet('主 PPT：由 Microsoft PowerPoint 開啟並放映的簡報，通常是聚會的主要投影片。')
Bullet('素材：另行加入的簡報、圖片或資料，透過控制台推送或更新到投影螢幕。')
Bullet('操作員：持有控制台、負責選擇投影內容的人。')
Bullet('講員：使用翻頁筆推進頁面的人；控制權可切換，但兩種角色按鈕仍可使用。')
Bullet('投影螢幕：第二台顯示器或投影機，Windows 顯示設定需選「延伸」。')

H1('二、系統需求與安裝')
H2('（一）硬體與系統')
Table @('項目','建議需求') @(
    @('作業系統','Windows 10 或 Windows 11，64 位元'),
    @('處理器與記憶體','四核心以上；建議 8 GB 以上記憶體'),
    @('顯示器','建議兩台：控制台與投影螢幕；單螢幕僅能在控制台預覽'),
    @('網路','手機控制需與電腦在同一個 Wi-Fi 或同一個區網'),
    @('連接埠','手機控制預設使用 TCP 5180；防火牆需允許')
)
H2('（二）必要軟件')
P('安裝包已包含 .NET 8 桌面執行環境，目標電腦不需要另外安裝 .NET。安裝包亦附 WebView2 安裝程式；若目標電腦已有 WebView2，安裝程式會跳過。')
P('主 PPT 放映依賴本機的 Microsoft PowerPoint。PowerPoint 不隨安裝包提供，目標電腦需已安裝 PowerPoint 2016 或更新版本。素材圖片、經文與網頁投影不依賴 PowerPoint。')
H2('（三）安裝步驟')
P('步驟 1：把安裝包複製到目標電腦，雙擊執行。若 Windows 顯示 SmartScreen 警告，選擇「更多資訊」，確認發行者为 Glory，再按「仍要執行」。')
P('步驟 2：依安裝精靈選擇安裝位置。預設為系統程式目錄下的 Glory 資料夾。')
P('步驟 3：安裝完成後，桌面與開始功能表會出現「Glory 榮耀 · 投影控制台」捷徑。')
P('步驟 4：第一次啟動前，先在 Windows「顯示設定」把投影機設為「延伸」，並確認第二螢幕能顯示桌面。')
P('步驟 5：雙擊捷徑啟動。第一次啟動會建立設定與日誌目錄。')
H2('（四）資料儲存位置')
Table @('資料','位置') @(
    @('設定檔','%AppData%\ProjectionManager\settings.json'),
    @('素材索引','%AppData%\ProjectionManager\materials.json'),
    @('素材圖檔','%AppData%\ProjectionManager\materials'),
    @('預覽暫存','%Temp%\ProjectionManager\preview'),
    @('操作日誌','%LocalAppData%\Glory\logs')
)
Note('安裝新版本不會清除上述資料；如需備份或搬到另一台電腦，備份 AppData 下的 ProjectionManager 與 Glory 資料夾即可。')
H2('（五）移除程式')
P('在 Windows「設定 → 應用程式 → 已安裝的應用程式」找到 Glory 榮耀投影控制台，按「解除安裝」。移除程式不會自動刪除 AppData 內的設定、素材與日誌；如要完全清除，需手動刪除上述資料夾。')

H1('三、介面導覽')
P('主視窗分為左、中、右三欄，底部是狀態列。第一次開啟時建議先確認投影目標、主題與控制者。')
Table @('區域','主要內容','用途') @(
    @('投影源','主 PPT、未分組素材、經文、網頁','選擇要操作的投影內容'),
    @('預覽','左欄目前播出／目前選擇頁；右欄下一頁','確認目前畫面與下一頁'),
    @('放映設定','投影目標螢幕、視窗化放映、控制台不搶焦點','選擇投影螢幕與放映方式'),
    @('翻頁筆','講員翻頁筆綁定','綁定後講員可用翻頁筆推進'),
    @('控制者','操作員／講員、交給講員／收回控制、手機控制','標示控制權並管理手機'),
    @('底部狀態列','頁碼、控制者、投影源','快速確認目前狀態'),
    @('日誌','右下角「日誌」按鈕','展開本次操作記錄')
)
H2('（一）預覽語意')
P('預覽左欄是「目前播出」；主 PPT 放映中會顯示播出頁，素材尚未推送時顯示「已選擇（未投影）」。右欄是「下一頁」，主 PPT 會顯示真正的下一頁縮圖；到了最後一頁會顯示「已到最後一頁」。')
P('結束放映後，左欄狀態改為「已停止放映」，不會再顯示「播出中」。素材未推送時，卡片顯示「已選擇 · 待推送」，避免誤以為已經投出。')
H2('（二）控制者顏色')
P('黑金主題下，操作員以金色顯示，講員以天空藍顯示；深灰青綠與暖白主題各有獨立色值。控制者名稱、狀態 chip 與底部 chip 會一起更新，切換主題後不需要重啟。')

H1('四、主 PPT 放映')
H2('（一）加入主 PPT')
P('步驟 1：按左上角「加入主 PPT」。')
P('步驟 2：在檔案對話框選擇 .pptx 或 .ppt 檔案。')
P('步驟 3：加入後，投影源列表會出現檔名與頁數；右欄「放映設定」可選擇投影目標。')
P('步驟 4：若尚未加入，列表顯示「尚未加入」，此時不能開始主 PPT 放映。')
H2('（二）開始放映')
P('步驟 1：確認「放映設定」的投影目標是正確的第二螢幕，例如 \\.\DISPLAY2。')
P('步驟 2：按「開始放映」。程式會透過 PowerPoint 開啟簡報並在指定螢幕放映。')
P('步驟 3：若只有一台螢幕，程式會提示僅能在控制台預覽；請先接上投影機並設為「延伸」。')
P('步驟 4：放映開始後，左欄顯示「播出中」，頁碼徽章顯示目前頁次。')
H2('（三）放映控制')
Table @('按鈕','作用','注意事項') @(
    @('黑屏','暫時遮黑投影畫面','不停止放映；按「恢復畫面」恢復'),
    @('恢復畫面','取消黑屏','回到原本頁面'),
    @('重新定位','把放映視窗移回指定螢幕','更換投影機或解析度後使用'),
    @('結束放映','結束 PowerPoint 放映','結束後左欄改為「已停止放映」'),
    @('上一頁／下一頁','切換主 PPT 頁面','按鈕與翻頁筆皆可操作'),
    @('跳到／跳至','輸入頁碼後跳到指定頁','超出總頁數會拒絕並記錄')
)
H2('（四）下一頁預覽')
P('主 PPT 放映時，右欄會從預覽暫存載入下一頁縮圖。縮圖还在匯出時，右欄可能暫時空白；程式完成後會自動補上。若一直空白，先按「下一頁」再返回，或按「重新定位」讓程式重繪。')
H2('（五）結束放映')
P('按「結束放映」後，程式會結束 PowerPoint 放映並把左欄狀態改為「已停止放映」。此時「開始放映」可再次使用。若 PowerPoint 詢問是否儲存，這是 PowerPoint 自身的檔案狀態；本程式只控制放映，不會修改簡報內容。')

H1('五、素材庫與推送投影')
H2('（一）加入素材')
P('步驟 1：按左下角「+ 素材」，選擇簡報、圖片或資料夾。')
P('步驟 2：素材會加入左側列表，顯示名稱與頁數。')
P('步驟 3：素材檔案會快取到 AppData，原始檔搬移或改名不影響已加入的素材。')
H2('（二）選擇素材')
P('在投影源列表點選素材。左欄會顯示第 1 頁，狀態為「已選擇（未投影）」。此時投影螢幕不會變動，適合先確認內容。')
P('素材預設不投影，這是為避免誤把尚未確認的內容送上投影幕。確認後必須按「推送投影」。')
H2('（三）推送投影')
P('步驟 1：確認投影目標與素材頁碼。')
P('步驟 2：按導覽列上的「推送投影」。')
P('步驟 3：程式會設定素材為放映中、取消黑屏，並把目前頁送到第二螢幕。')
P('步驟 4：推送後按鈕文字改為「更新投影」；後續翻頁或改頁時，若已在推送狀態，投影畫面會同步更新。')
P('步驟 5：只有一台螢幕時，程式只記錄「僅控制台預覽」，不會假裝已投出；接上第二螢幕並設為「延伸」後再推送。')
H2('（四）素材翻頁與更新')
P('推送後可用「上一頁／下一頁」或「跳到」切換頁碼。若只改了控制台頁碼而沒有按「更新投影」，投影畫面不會變；此時按「更新投影」同步。')
H2('（五）移除與重新匯出')
P('右欄「選中的素材」提供「重新匯出」與「移除」。重新匯出用於簡報更新後重建縮圖；移除只移除素材索引，不會刪除原始檔。')

H1('六、經文投影')
H2('（一）加入與選擇')
P('按左下角「+ 經文」開啟選擇經文視窗。可選書卷、章、節，按「投影」送到投影螢幕。經文列表會顯示出處與字級。')
H2('（二）搜尋')
P('在搜尋框輸入關鍵字，例如「葡萄樹」或「約翰福音 15:1」，按「搜尋」。結果會列出符合節數；點選結果可跳到該節。')
H2('（三）字級與樣式')
P('在經文視窗可選字級與樣式。投影後仍可在右欄調整字級；調整會套用到目前經文。')
H2('（四）單螢幕與雙螢幕')
P('只有一台螢幕時，經文只在控制台預覽；接上第二螢幕後按「投影」才會開投影視窗。投影視窗固定深色，不跟隨控制台主題，避免投影內容變色。')

H1('七、網頁投影')
H2('（一）加入網頁')
P('按左下角「+ 網頁」，輸入名稱與网址，按確定。網頁會加入投影源列表。')
H2('（二）投影與滾動')
P('選擇網頁後按「投影」。程式會在第二螢幕開網頁視窗。可用「上一頁」「下一頁」或滾動按鈕控制網頁。')
P('網頁投影需要 WebView2 Runtime。Windows 10/11 通常已內建；若網頁視窗空白，先查看第十章日誌與 WebView2 排查。')
H2('（三）投影視窗固定配色')
P('經文、素材、網頁的投影輸出視窗固定使用深色背景，刻意不跟隨控制台主題。這樣切換黑金、暖白時，投影內容不會突然變色。')

H1('八、手機控制')
H2('（一）啟用')
P('按右欄「手機控制...」開啟手機控制視窗。程式會在區網啟動 HTTP 服務，預設連接埠 5180。視窗會顯示網址與配對碼。')
H2('（二）配對')
P('步驟 1：確認手機與電腦連同一個 Wi-Fi。')
P('步驟 2：在手機瀏覽器輸入視窗顯示的網址，例如 http://192.168.0.170:5180。')
P('步驟 3：輸入配對碼。配對成功後，手機可看到目前投影內容與控制按鈕。')
P('步驟 4：已配對的手機下次可自動連線；不需要每次重新輸入配對碼。')
H2('（三）操作員與講員')
P('手機裝置列表可把角色切換為「操作員」或「講員」。操作員可切換投影源與控制；講員主要用翻頁筆或手機上一頁／下一頁推進。控制權切換只影響標示與按鈕提示，不會鎖住另一個角色。')
H2('（四）撤銷裝置')
P('按裝置右側「撤銷」，該手機需要重新配對。若手機遺失或更換，請先撤銷舊裝置。')
H2('（五）網絡與防火牆')
P('若手機打不開網址，先確認兩台裝置在同一個 Wi-Fi、沒有開啟 AP 隔離，且 Windows 防火牆允許 Glory 投影控制台使用私人網路。連接埠 5180 被其他程式佔用時，手機控制會啟動失败，日誌會記錄原因。')

H1('九、設定與主題')
H2('（一）三種主題')
Table @('主題','外觀','適用場景') @(
    @('黑金（預設）','純黑背景、金色強調、奶油金文字','昏暗聚會現場，高對比'),
    @('深灰青綠','深灰底、青綠重點色','已熟悉舊版的操作員'),
    @('暖白背景','暖白底、深墨文字、深金重點','明亮室內或需要列印截圖')
)
P('按右上角「設定」開啟設定視窗。選擇主題後立即套用，不需要重啟。設定會寫入 settings.json，下次啟動沿用。')
H2('（二）選檔起始資料夾')
P('設定頁可指定主 PPT 與素材的起始資料夾。这只影响下次選檔對話框的起始位置，不會搬移已加入的素材。')
H2('（三）日誌與診斷')
P('設定頁的「儲存、快取與日誌」區塊可看到操作日誌目錄。按「開啟」直接打開資料夾；按「匯出診斷包」會把最近日誌、設定檔、素材索引與系統資訊打包到桌面。')

H1('十、操作日誌與故障排查')
H2('（一）日誌位置與格式')
P('日誌寫入 %LocalAppData%\Glory\logs，檔名為 glory-YYYYMMDD.log。每行包含日期、時間、等級、分類與訊息；錯誤會附上完整 Exception 堆疊。')
P('單一日誌檔超過 8 MB 會自動輪替；程式啟動時清除 30 天前的舊日誌。日誌包含啟動版本、作業系統、執行檔路徑、命令列、主题切換、投影操作與錯誤。')
H2('（二）匯出診斷包')
P('步驟 1：開啟「設定」。')
P('步驟 2：在「儲存、快取與日誌」按「匯出診斷包」。')
P('步驟 3：桌面會出現 Glory_診斷包_日期_時間.zip。把這個 ZIP 交給技術支援，即可重現問題。')
H2('（三）常見問題')
Table @('現象','可能原因','處理方式') @(
    @('無法開始放映','未加入主 PPT、未安裝 PowerPoint、COM 呼叫失败','先加入主 PPT；確認 PowerPoint 可正常開啟檔案；查看日誌 [ERROR] 堆疊'),
    @('第二螢幕沒畫面','Windows 顯示設定不是「延伸」、投影目標選錯','把投影機設為延伸；在放映設定重選目標；按「重新定位」'),
    @('素材按推送沒反應','只選了素材但沒按推送、只有一台螢幕','按「推送投影」；接上第二螢幕並設為延伸'),
    @('預覽右欄空白','下一頁縮圖仍在匯出、簡報只有一頁','等缩圖完成；到最後一頁會顯示「已到最後一頁」'),
    @('預覽比例被壓扁','預覽卡比例與來源比例不一致','本版預覽採 4:3 容器，16:9 內容會上下留黑；若仍變形，查看日誌中的預覽尺寸'),
    @('彈窗文字看不清','主題切換後樣式未更新、對比不足','切換主題後關閉再開彈窗；確認使用 V2.3 版本'),
    @('手機打不開網址','不同 Wi-Fi、AP 隔離、防火牆阻擋','連同一個網路；允許私人網路；確認 5180 沒有被佔用'),
    @('手機控制啟動失败','連接埠 5180 已被其他程式佔用','關閉佔用程式，或在日誌確認後改用其他連接埠'),
    @('網頁投影空白','WebView2 Runtime 缺失或載入失败','安裝 WebView2 Runtime；查看日誌 [ERROR] [WEB]'),
    @('關閉 PowerPoint 提示儲存','PowerPoint 自身檔案狀態或增益集提示','本程式不修改簡報；如未手動改動，可直接不儲存關閉，並查看日誌確認有無 COM 錯誤')
)
H2('（四）回報問題時請提供')
Bullet('發生時間與操作步驟。')
Bullet('診斷包 ZIP（含日誌、設定、素材索引與系統資訊）。')
Bullet('投影螢幕解析度、Windows 版本、PowerPoint 版本。')
Bullet('若為畫面問題，附上控制台與投影幕照片。')

H1('十一、實機測試檢查清單')
P('以下項目建議在交付現場逐項確認；全部通過後再進行正式聚會。')
Table @('編號','檢查項目','通過標準') @(
    @('1','安裝與啟動','雙擊安裝包完成安裝；桌面捷徑可啟動；目標機未另裝 .NET'),
    @('2','投影目標','第二螢幕設為延伸；放映設定可選到正確螢幕'),
    @('3','主 PPT 放映','加入主 PPT、開始放映；投影幕出現簡報；頁碼正確'),
    @('4','主 PPT 預覽','左欄播出中；右欄顯示下一頁；最後一頁顯示「已到最後一頁」'),
    @('5','結束放映','按結束放映；左欄改為「已停止放映」；可再次開始'),
    @('6','素材預設','選擇素材後只在控制台預覽；投影幕不變'),
    @('7','推送投影','按推送後投影幕出現素材；按鈕變「更新投影」；翻頁同步'),
    @('8','單螢幕保護','只有一台螢幕時按推送，投影幕不會誤開；日誌記錄僅預覽'),
    @('9','經文','開啟經文視窗、搜尋、投影；主題切換後仍可讀'),
    @('10','網頁','加入網址、投影；WebView2 正常載入'),
    @('11','手機配對','手機同 Wi-Fi 開啟網址、輸入配對碼；可控制上一頁／下一頁'),
    @('12','角色切換','操作員／講員顏色明顯不同；按鈕文字正確'),
    @('13','主題切換','黑金、深灰青綠、暖白都可切換；彈窗跟隨主題；投影視窗保持深色'),
    @('14','日誌','操作後日誌檔有新增內容；錯誤有堆疊；可匯出診斷包'),
    @('15','資料保存','重啟後主題、素材、設定仍在')
)

H1('十二、版本與已知限制')
P('本版為 V2.3 交付版，包含 V2.2 以來的四項修正：控制者顏色區分、彈窗主題化、預覽左當前右下一頁、素材需按推送投影。')
Bullet('主 PPT 放映需要目標電腦已安裝 Microsoft PowerPoint，安裝包不包含 Office。')
Bullet('網頁投影需要 WebView2 Runtime；安裝包附安裝程式，Windows 10/11 通常已內建。')
Bullet('手機控制需要電腦與手機在同一個區網，且防火牆允許 5180。')
Bullet('單螢幕僅能在控制台預覽，无法真正投出到第二螢幕。')
Bullet('經文與網頁仍使用原本的「投影」流程；「推送投影」按鈕目前只套用到素材庫頁面。')
Bullet('投影輸出視窗固定深色，不跟隨控制台主題，這是刻意設計。')

H1('十三、快速操作卡')
P('開場前：接投影機 → Windows 設「延伸」 → 啟動 Glory → 選投影目標 → 加入主 PPT → 開始放映。')
P('素材：選素材 → 確認左欄第 1 頁 → 按「推送投影」 → 翻頁或按「更新投影」。')
P('經文：按「+ 經文」 → 選書卷章節 → 按「投影」。')
P('手機：按「手機控制...」 → 手機連同一 Wi-Fi → 輸入網址與配對碼 → 選角色。')
P('出錯：按右下「日誌」查看 → 設定 → 匯出診斷包 → 把 ZIP 交給技術支援。')
P('收場：按「結束放映」 → 關閉投影視窗 → 關閉 Glory。')

# 頁尾頁碼（封面不顯示）
$doc.Sections.Item(1).PageSetup.DifferentFirstPageHeaderFooter = -1
$footer = $doc.Sections.Item(1).Footers.Item(1)
$footer.Range.ParagraphFormat.Alignment = 1
$footer.Range.Fields.Add($footer.Range, 33) | Out-Null
$footer.Range.Font.Name = 'Times New Roman'
$footer.Range.Font.Size = 9

# 更新目錄
if ($doc.TablesOfContents.Count -gt 0) {
    $doc.TablesOfContents.Item(1).Update()
}

$doc.SaveAs2($docx, 16)
$doc.ExportAsFixedFormat($pdf, 17)
$doc.Close(0)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect()
Write-Output "DOCX=$docx"
Write-Output "PDF=$pdf"
