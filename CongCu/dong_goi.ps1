# =====================================================================
#  ĐÓNG GÓI CÔNG CỤ ĐỂ GỬI CHO NGƯỜI KHÁC
#  Tạo file zip trong thư mục "GoiGui" (cạnh thư mục CongCu), gồm:
#    - công cụ + cửa sổ + hướng dẫn
#    - khung mẫu ĐÃ XOÁ tên trường, lớp, tên người ký
#    - cấu hình đã bỏ ngày nghỉ riêng
#    - các thư mục môn học trống để người nhận chép file vào
#  KHÔNG gồm: giáo án, file môn học, kế hoạch khối, bản mốc kiểm tra.
# =====================================================================
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$goc  = Split-Path -Parent $here
$utf8Bom = New-Object Text.UTF8Encoding $true
$utf8    = New-Object Text.UTF8Encoding $false
$ngay = Get-Date
$ban  = $ngay.ToString('yyyy.MM.dd')
try { $h = (& git -C $goc rev-parse --short HEAD 2>$null); if ($LASTEXITCODE -eq 0 -and $h) { $ban += "-$h" } } catch {}

$dirGui = Join-Path $goc 'GoiGui'
$tam    = Join-Path $dirGui 'tam'
$ten    = 'SoanGiaoAn'
$dich   = Join-Path $tam $ten
if (Test-Path -LiteralPath $tam) { Remove-Item -LiteralPath $tam -Recurse -Force }
[void](New-Item -ItemType Directory -Force -Path (Join-Path $dich 'CongCu'))

Write-Host "Đóng gói bản $ban ..." -ForegroundColor Cyan

# ---------- 1. Các file công cụ ----------
foreach ($f in 'tao_giao_an.ps1', 'giao_dien.ps1', 'MO CUA SO.bat', 'TAO GIAO AN.bat', 'HUONG_DAN.md') {
    Copy-Item -LiteralPath (Join-Path $here $f) -Destination (Join-Path $dich "CongCu\$f")
}
[IO.File]::WriteAllText((Join-Path $dich 'CongCu\PHIEN_BAN.txt'), "$ban`r`n", $utf8)

# ---------- 2. Cấu hình: bỏ ngày nghỉ riêng, chữ "Lớp 2C" in đậm -> "Lớp" ----------
$cfgText = [IO.File]::ReadAllText((Join-Path $here 'cauhinh.json'), [Text.Encoding]::UTF8)
$cfg = $cfgText | ConvertFrom-Json
$cfgText = [regex]::Replace($cfgText, '"ngayNghi"\s*:\s*\[[^\]]*\]', '"ngayNghi": []')
$cfgText = [regex]::Replace($cfgText, '"tuanNghi"\s*:\s*\[[^\]]*\]', '"tuanNghi": []')
$cfgText = [regex]::Replace($cfgText, '("chuInDamDauTrang"\s*:\s*\[[^\]]*?)"Lớp [^"]*"', '$1"Lớp"')
$cfgText = [regex]::Replace($cfgText, '"thuMucGoc"\s*:\s*"[^"]*"', '"thuMucGoc": ""')
[void]($cfgText | ConvertFrom-Json)
[IO.File]::WriteAllText((Join-Path $dich 'CongCu\cauhinh.json'), $cfgText, $utf8)

# ---------- 3. Khung mẫu: xoá tên trường, lớp, tên người ký, thông tin tác giả ----------
$khungNguon = Join-Path $goc ($cfg.fileMauKhung -replace '/', '\')
$khungDich  = Join-Path $dich ($cfg.fileMauKhung -replace '/', '\')
[void](New-Item -ItemType Directory -Force -Path (Split-Path -Parent $khungDich))
Copy-Item -LiteralPath $khungNguon -Destination $khungDich
$daXoa = New-Object System.Collections.ArrayList
$wordTruoc = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
$wordCuaToi = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $wordTruoc -notcontains $_.Id } | ForEach-Object { $_.Id })
try {
    $d = $w.Documents.Open($khungDich)
    # Hai dòng đầu: phần trước dấu tab là tên trường / tên lớp
    $thay = @('Trường Tiểu học ...........', 'Lớp ......')
    for ($i = 1; $i -le 2; $i++) {
        $r = $d.Paragraphs.Item($i).Range
        $raw = $r.Text; $ix = $raw.IndexOf("`t")
        if ($ix -gt 0) {
            [void]$daXoa.Add($raw.Substring(0, $ix).Trim())
            $d.Range($r.Start, $r.Start + $ix).Text = $thay[$i - 1] + ' '
        }
    }
    # Bảng thời khoá biểu: xoá tên bài, PPCT, ghi chú của tuần cũ (công cụ sẽ điền lại)
    $t1 = $d.Tables.Item(1)
    foreach ($c in $t1.Range.Cells) { if ($c.RowIndex -ge 2 -and $c.ColumnIndex -in 4, 6, 7) { $rr = $c.Range; [void]$rr.MoveEnd(1, -1); if ($rr.Text) { $rr.Text = '' } } }
    # Bảng chữ ký (bảng cuối): giữ dòng chức danh đầu ô, các dòng có chữ phía dưới (họ tên) thay bằng dấu chấm
    if ($d.Tables.Count -ge 2) {
        foreach ($c in $d.Tables.Item($d.Tables.Count).Range.Cells) {
            $ps = @(); foreach ($p in $c.Range.Paragraphs) { $ps += $p }
            for ($k = $ps.Count - 1; $k -ge 1; $k--) {
                $rr = $ps[$k].Range; $txt = ($rr.Text -replace "[\r\a]", '').Trim()
                if ($txt) { [void]$daXoa.Add($txt); [void]$rr.MoveEnd(1, -1); $rr.Text = '....................' }
            }
        }
    }
    try { $d.RemoveDocumentInformation(8) } catch {}   # thuộc tính tài liệu (tác giả...)
    try { $d.RemoveDocumentInformation(4) } catch {}   # thông tin cá nhân
    $d.Save(); $d.Close(0)
} finally {
    try { $w.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($w) } catch {}
    $w = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($wordCuaToi.Count -eq 1) { $wp = Get-Process -Id $wordCuaToi[0] -ErrorAction SilentlyContinue; if ($wp -and -not $wp.WaitForExit(10000)) { $wp.Refresh(); if (-not $wp.MainWindowTitle) { try { $wp.Kill() } catch {} } } }
}

# ---------- 4. Thư mục môn học trống + file hướng dẫn bắt đầu ----------
foreach ($p in $cfg.monHoc.PSObject.Properties) {
    $m = $p.Value
    if (-not $m.thuMuc -or $m.cheDo -eq 'tieuDe') { continue }
    $dm = Join-Path $dich ($m.thuMuc -replace '/', '\')
    [void](New-Item -ItemType Directory -Force -Path $dm)
    $vd = if ($m.mauTenFile) { $m.mauTenFile -f 11 } else { '' }
    [IO.File]::WriteAllText((Join-Path $dm 'CHEP FILE VAO DAY.txt'), "Chép các file giáo án môn $($m.tenTKB) vào thư mục này.`r`nTên file phải theo mẫu, ví dụ tuần 11:  $vd`r`n", $utf8Bom)
}
$batDau = @"
SOẠN GIÁO ÁN TUẦN TỰ ĐỘNG  -  bản $ban
==========================================================

CẦN CÓ
  - Máy tính Windows có cài Microsoft Word.
  - Bộ giáo án từng môn theo tuần (file Word) và file kế hoạch dạy học của khối.

CÀI ĐẶT (làm một lần)
  1. Giải nén cả thư mục "$ten" ra ổ D hoặc Desktop. Không chạy thẳng trong file zip.
  2. Chép file kế hoạch dạy học khối vào thư mục "$ten", đặt tên đúng là:
         $($cfg.fileKeHoachKhoi)
  3. Chép giáo án từng môn vào các thư mục con tương ứng (mỗi thư mục có sẵn một
     file "CHEP FILE VAO DAY.txt" ghi rõ cách đặt tên file).
  4. Bấm đúp "SOAN GIAO AN.bat".
     Nếu Windows hiện bảng xanh "Windows protected your PC": bấm "More info" rồi "Run anyway".
  5. Trong cửa sổ, bấm "Thiết lập lớp":
       - chọn ngày thứ Hai của tuần 1;
       - chọn thời khoá biểu của lớp mình (bấm từng ô để chọn môn);
       - bấm "Sửa khung mẫu" để ghi tên trường, tên lớp và tên người ký; lưu lại trong Word;
       - bấm "Lưu".

DÙNG HẰNG TUẦN
  Bấm đúp "SOAN GIAO AN.bat" -> chọn tuần -> (thêm ngày nghỉ nếu có) -> "TẠO GIÁO ÁN"
  -> đọc khung "Lưu ý cần kiểm tra" -> "Mở giáo án vừa tạo" để xem và sửa.

CẦN BIẾT
  - Công cụ ghép và sửa trình bày theo một kiểu giáo án nhất định (mỗi bài bắt đầu bằng
    dòng "KẾ HOẠCH..." hoặc "BÀI n.", có các mục I, II, III, IV). Giáo án nguồn trình bày
    khác đi thì công cụ sẽ báo trong khung lưu ý và cần sửa tay nhiều hơn.
  - Cấu hình sẵn cho khối 2 (số tiết mỗi tuần: Toán 5, Tiếng Việt 10, HĐTN 3, TNXH 2,
    Đạo đức 1). Khối khác phải sửa mục "monHoc" trong CongCu\cauhinh.json.
  - Công cụ không gửi dữ liệu đi đâu; mọi thứ nằm trên máy.
  - Hướng dẫn chi tiết: CongCu\HUONG_DAN.md (mở bằng Notepad).
"@
[IO.File]::WriteAllText((Join-Path $dich 'BAT DAU - DOC TRUOC.txt'), ($batDau -replace "`r?`n", "`r`n"), $utf8Bom)
$bat = "@echo off`r`nrem Mo cua so soan giao an.`r`nstart `"`" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"%~dp0CongCu\giao_dien.ps1`"`r`n"
[IO.File]::WriteAllText((Join-Path $dich 'SOAN GIAO AN.bat'), $bat, $utf8)

# ---------- 5. Kiểm tra không lọt thông tin riêng, rồi nén ----------
$lot = @()
foreach ($f in Get-ChildItem -LiteralPath $dich -Recurse -File | Where-Object { $_.Extension -in '.json', '.txt', '.bat', '.md', '.ps1' }) {
    $noiDung = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
    foreach ($x in $daXoa) { if ($x.Length -ge 4 -and $noiDung.Contains($x)) { $lot += "$($f.Name): còn chữ '$x'" } }
}
$docx = @(Get-ChildItem -LiteralPath $dich -Recurse -File -Filter '*.docx')
Add-Type -AssemblyName System.IO.Compression.FileSystem
foreach ($f in $docx) {
    $z = [IO.Compression.ZipFile]::OpenRead($f.FullName)
    try {
        foreach ($e in $z.Entries) {
            if ($e.FullName -notmatch '\.xml$|\.rels$') { continue }
            $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); $x = $sr.ReadToEnd(); $sr.Close()
            $x = $x -replace '<[^>]+>', ''
            foreach ($s in $daXoa) { if ($s.Length -ge 4 -and $x.Contains($s)) { $lot += "$($f.Name) ($($e.FullName)): còn chữ '$s'" } }
        }
    } finally { $z.Dispose() }
}
$zip = Join-Path $dirGui "$ten-$($ngay.ToString('yyyyMMdd')).zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
[IO.Compression.ZipFile]::CreateFromDirectory($tam, $zip, [IO.Compression.CompressionLevel]::Optimal, $false)
$soFile = @(Get-ChildItem -LiteralPath $dich -Recurse -File).Count
Remove-Item -LiteralPath $tam -Recurse -Force

Write-Host ''
Write-Host "Đã xoá khỏi khung mẫu: $(($daXoa | ForEach-Object { "'$_'" }) -join ', ')"
if ($lot.Count -gt 0) {
    Write-Host 'CẢNH BÁO: trong gói vẫn còn thông tin riêng, CHƯA NÊN GỬI:' -ForegroundColor Red
    $lot | Select-Object -Unique | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
} else {
    Write-Host 'Đã kiểm tra: trong gói không còn các thông tin riêng ở trên.' -ForegroundColor Green
}
Write-Host "XONG: $zip  ($soFile file, $([Math]::Round((Get-Item -LiteralPath $zip).Length / 1KB)) KB)" -ForegroundColor Green
Write-Host 'Gửi file zip này cho người nhận (Zalo, email, USB).'
