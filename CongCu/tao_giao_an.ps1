# =====================================================================
#  CÔNG CỤ SOẠN GIÁO ÁN TUẦN TỰ ĐỘNG - LỚP 2C
#  Cách dùng:  tao_giao_an.ps1 -Tuan 11
#              tao_giao_an.ps1 -Tuan 11 -NgayThuHai 16/11/2026   (nếu lịch lệch)
#  Cần: Microsoft Word đã cài trên máy.
# =====================================================================
param(
    [int]$Tuan = 0,
    [string]$NgayThuHai = "",
    [switch]$HienWord,
    [string]$CauHinh = "",        # file cấu hình khác (VD cauhinh-2A.json); để trống = cauhinh.json
    [string]$ThuMucKetQua = "",   # ghi kết quả vào thư mục khác (dùng khi kiểm tra, không đè file thật)
    [switch]$KhongMoLuuY          # không tự mở file lưu ý bằng Notepad
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$cfgPath = if ($CauHinh) { if ([IO.Path]::IsPathRooted($CauHinh)) { $CauHinh } else { Join-Path $here $CauHinh } } else { Join-Path $here 'cauhinh.json' }
function Dung([string[]]$dong) {
    # Dừng công cụ với thông báo dễ hiểu (không hiện lỗi PowerShell)
    Write-Host ''
    Write-Host 'KHÔNG CHẠY ĐƯỢC:' -ForegroundColor Red
    foreach ($d in $dong) { Write-Host "  $d" -ForegroundColor Red }
    exit 1
}
$tenCfg = [IO.Path]::GetFileName($cfgPath)
if (-not (Test-Path -LiteralPath $cfgPath)) { Dung "Không thấy file cấu hình: $cfgPath" }
$cfgText = Get-Content -Raw -Encoding UTF8 -LiteralPath $cfgPath
try { $cfg = $cfgText | ConvertFrom-Json }
catch {
    $msg = $_.Exception.Message
    $goiY = @("File $tenCfg bị sai cú pháp nên không đọc được.")
    $mPos = [regex]::Match($msg, '\((\d+)\)')
    if ($mPos.Success -and [int]$mPos.Groups[1].Value -le $cfgText.Length) {
        $dongSo = ($cfgText.Substring(0, [int]$mPos.Groups[1].Value) -split "`n").Count
        $goiY += "Chỗ sai ở khoảng dòng $dongSo (hoặc dòng ngay phía trên)."
    }
    $goiY += 'Thường là do thiếu dấu phẩy cuối dòng, thừa dấu phẩy trước dấu ] hoặc }, hoặc thiếu dấu ngoặc kép ".'
    $goiY += 'Cách nhanh nhất: lấy lại bản cũ của file này trên GitHub rồi sửa lại từ từ.'
    Dung $goiY
}
# Thư mục gốc = thư mục chứa "CongCu" (chép cả thư mục sang ổ khác vẫn chạy đúng).
# Chỉ dùng "thuMucGoc" trong cauhinh.json khi được điền rõ ràng.
$thuMucCha = Split-Path -Parent $here
if (-not $cfg.thuMucGoc) {
    $cfg | Add-Member -NotePropertyName thuMucGoc -NotePropertyValue $thuMucCha -Force
} elseif (-not (Test-Path -LiteralPath $cfg.thuMucGoc)) {
    Write-Host "  ! Không thấy thư mục '$($cfg.thuMucGoc)', dùng '$thuMucCha'" -ForegroundColor Yellow
    $cfg.thuMucGoc = $thuMucCha
}
Write-Host "Thư mục giáo án: $($cfg.thuMucGoc)" -ForegroundColor DarkGray

if ($Tuan -le 0) {
    $nhap = Read-Host 'Nhập số tuần cần soạn'
    if (-not [int]::TryParse("$nhap".Trim(), [ref]$Tuan) -or $Tuan -le 0) { Dung "Số tuần phải là một số (ví dụ 11), cô đã nhập: '$nhap'" }
}

# ---------- Kiểm tra cấu hình trước khi mở Word ----------
$loiCfg = New-Object System.Collections.ArrayList
foreach ($k in 'fileKeHoachKhoi', 'fileMauKhung', 'tenFileKetQua', 'ngayThuHaiTuan1', 'thoiKhoaBieu', 'monHoc') {
    if (-not $cfg.$k) { [void]$loiCfg.Add("Thiếu mục `"$k`" trong $tenCfg.") }
}
function KiemNgay([string]$ten, $giaTri) {
    foreach ($v in @($giaTri)) {
        if (-not $v) { continue }
        $tmp = [datetime]::MinValue
        if (-not [datetime]::TryParseExact("$v", 'yyyy-MM-dd', $null, 'None', [ref]$tmp)) {
            [void]$loiCfg.Add("Mục `"$ten`": ngày '$v' phải viết dạng năm-tháng-ngày, ví dụ 2026-09-07.")
        }
    }
}
KiemNgay 'ngayThuHaiTuan1' $cfg.ngayThuHaiTuan1
KiemNgay 'tuanNghi' $cfg.tuanNghi
KiemNgay 'ngayNghi' $cfg.ngayNghi
if ($NgayThuHai) {
    $tmp = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($NgayThuHai, 'dd/MM/yyyy', $null, 'None', [ref]$tmp)) { [void]$loiCfg.Add("-NgayThuHai '$NgayThuHai' phải viết dạng ngày/tháng/năm, ví dụ 23/11/2026.") }
}
if ($cfg.monHoc -and $cfg.thoiKhoaBieu) {
    $maMon = @($cfg.monHoc.PSObject.Properties.Name)
    $thuHopLe = 'Hai', 'Ba', 'Tư', 'Năm', 'Sáu', 'Bảy'
    foreach ($ses in $cfg.thoiKhoaBieu) {
        if ($thuHopLe -notcontains "$($ses.thu)".Normalize([Text.NormalizationForm]::FormC)) { [void]$loiCfg.Add("thoiKhoaBieu: thứ '$($ses.thu)' không hợp lệ (dùng: $($thuHopLe -join ', ')).") }
        foreach ($m in @($ses.mon)) {
            if ($maMon -notcontains $m) { [void]$loiCfg.Add("thoiKhoaBieu (thứ $($ses.thu), $($ses.buoi)): mã môn '$m' chưa có trong monHoc. Các mã đang có: $($maMon -join ', ').") }
        }
    }
}
foreach ($k in 'fileKeHoachKhoi', 'fileMauKhung') {
    if ($cfg.$k) {
        $p = Join-Path $cfg.thuMucGoc ($cfg.$k -replace '/', '\')
        if (-not (Test-Path -LiteralPath $p)) { [void]$loiCfg.Add("Không thấy file: $p  (mục `"$k`").") }
    }
}
if (-not [Type]::GetTypeFromProgID('Word.Application')) { [void]$loiCfg.Add('Máy này chưa cài Microsoft Word. Công cụ cần Word để ghép giáo án.') }
if ($loiCfg.Count -gt 0) { Dung @($loiCfg | Select-Object -Unique) }

$warn = New-Object System.Collections.ArrayList
function Warn([string]$m) { [void]$warn.Add($m); Write-Host "  ! $m" -ForegroundColor Yellow }
function Info([string]$m) { Write-Host $m -ForegroundColor Cyan }

# ---------- Tiện ích chuỗi ----------
function N([string]$s) { if ($null -eq $s) { return '' }; return $s.Normalize([Text.NormalizationForm]::FormC) }
function CleanText([string]$s) {
    $s = N $s
    $s = $s -replace "[\r\n\a\v\f]", " " -replace "\s+", " "
    return $s.Trim()
}
function CleanTenBai([string]$s) {
    $s = CleanText $s
    $s = $s -replace '^Bài\s*\d+\s*[:.]\s*', ''
    $s = $s -replace '\s*[-–]\s*Tiết\s*(\d+)\s*(:.*)?$', ' (T$1)'
    return $s.Trim()
}
function ApDungQuyTac([string]$s, $rules) {
    # Quy tắc sửa tên bài trong cauhinh.json: [["mẫu tìm (regex)", "thay bằng"], ...]
    foreach ($rule in @($rules)) { if ($rule) { $s = [regex]::Replace($s, (N $rule[0]), (N $rule[1])) } }
    return $s.Trim()
}
function TieuDeTuTen([string]$s) { return $s.Trim().TrimEnd('.').Trim().ToUpper() }
function ApDungGhiChu([string]$s) {
    $s = CleanText $s
    foreach ($pair in $cfg.thayTheGhiChu) { $s = $s.Replace((N $pair[0]), (N $pair[1])) }
    return $s
}
function Path-Of([string]$rel) { return Join-Path $cfg.thuMucGoc ($rel -replace '/', '\') }

# ---------- Ngày tháng ----------
if ($NgayThuHai) {
    $monday = [datetime]::ParseExact($NgayThuHai, 'dd/MM/yyyy', $null)
} else {
    $monday = [datetime]::ParseExact($cfg.ngayThuHaiTuan1, 'yyyy-MM-dd', $null).AddDays(7 * ($Tuan - 1))
    # Tuần nghỉ hẳn (Tết...) không tính là tuần học: mỗi tuần nghỉ nằm trước/trùng thì lùi thêm 1 tuần
    $dsTuanNghi = @($cfg.tuanNghi | Where-Object { $_ } | ForEach-Object {
        $d = [datetime]::ParseExact($_, 'yyyy-MM-dd', $null); $d.AddDays(-((([int]$d.DayOfWeek) + 6) % 7)) } | Sort-Object -Unique)
    foreach ($d in $dsTuanNghi) { if ($d -le $monday) { $monday = $monday.AddDays(7) } }
}
# Ngày nghỉ lẻ trong tuần (lễ...): các buổi của ngày đó ghi "Nghỉ", không lấy tiết từ kế hoạch khối
$dsNgayNghi = @($cfg.ngayNghi | Where-Object { $_ } | ForEach-Object { [datetime]::ParseExact($_, 'yyyy-MM-dd', $null).Date })
$chuNghi = if ($cfg.chuNghi) { $cfg.chuNghi } else { 'Nghỉ' }
$dayOffset = @{ 'Hai' = 0; 'Ba' = 1; 'Tư' = 2; 'Năm' = 3; 'Sáu' = 4; 'Bảy' = 5 }
function NgayCua([string]$thu) { return $monday.AddDays($dayOffset[(N $thu)]) }

# ---------- Bảng tra môn ----------
$monCodes = @($cfg.monHoc.PSObject.Properties.Name)
$aliasMap = @{}
foreach ($code in $monCodes) {
    foreach ($a in $cfg.monHoc.$code.tenTrongKeHoachKhoi) { $aliasMap[(CleanText $a).ToLower()] = $code }
}
function MonCfg([string]$code) { return $cfg.monHoc.$code }

Info "=== SOẠN GIÁO ÁN TUẦN $Tuan (Thứ Hai: $($monday.ToString('dd/MM/yyyy'))) ==="

$wordTruoc = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$word = New-Object -ComObject Word.Application
# Ghi nhớ tiến trình Word do công cụ mở ra, để cuối cùng chắc chắn đóng đúng nó (không đụng Word cô đang dùng)
$wordCuaToi = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $wordTruoc -notcontains $_.Id } | ForEach-Object { $_.Id })
$word.Visible = [bool]$HienWord
$word.DisplayAlerts = 0
$openDocs = New-Object System.Collections.ArrayList

try {
    # =================================================================
    # 1. ĐỌC KẾ HOẠCH DẠY HỌC KHỐI -> danh sách tiết theo môn
    # =================================================================
    Info "1. Đọc kế hoạch dạy học khối..."
    $khPath = Path-Of $cfg.fileKeHoachKhoi
    $kh = $word.Documents.Open($khPath, $false, $true)
    [void]$openDocs.Add($kh)

    $tblTuan = $null
    foreach ($tbl in $kh.Tables) {
        $s = [Math]::Max(0, $tbl.Range.Start - 400)
        $before = N ($kh.Range($s, $tbl.Range.Start).Text)
        $m = [regex]::Matches($before, 'KẾ HOẠCH DẠY HỌC TUẦN\s*(\d+)')
        if ($m.Count -gt 0 -and [int]$m[$m.Count - 1].Groups[1].Value -eq $Tuan) { $tblTuan = $tbl; break }
    }
    if (-not $tblTuan) { throw "Không tìm thấy bảng 'KẾ HOẠCH DẠY HỌC TUẦN $Tuan' trong file $($cfg.fileKeHoachKhoi)" }

    # Đọc ô theo (dòng, cột thật) để không lệch cột ở các ô bị gộp
    $grid = [ordered]@{}
    $maxCol = 0
    foreach ($c in $tblTuan.Range.Cells) {
        $ri = "r$($c.RowIndex)"
        if (-not $grid.Contains($ri)) { $grid[$ri] = @{} }
        $grid[$ri][[int]$c.ColumnIndex] = CleanText $c.Range.Text
        if ($c.ColumnIndex -gt $maxCol) { $maxCol = [int]$c.ColumnIndex }
    }
    # 4 cột cuối của bảng: PPCT | Môn | Tên bài dạy | Điều chỉnh
    $colP = $maxCol - 3; $colM = $maxCol - 2; $colT = $maxCol - 1; $colD = $maxCol
    $rows = [ordered]@{}
    foreach ($ri in $grid.Keys) {
        $g = $grid[$ri]
        $rows[$ri] = @('', '', '', '', $g[$colP], $g[$colM], $g[$colT], $g[$colD])
    }
    $hang = @{}   # code -> list of lessons
    foreach ($code in $monCodes) { $hang[$code] = New-Object System.Collections.ArrayList }
    foreach ($ri in $rows.Keys) {
        $cells = $rows[$ri]
        if ($cells.Count -lt 5) { continue }
        $n = $cells.Count
        $ppctTxt = $cells[$n - 4]; $monTxt = $cells[$n - 3]; $ten = $cells[$n - 2]; $dc = $cells[$n - 1]
        $ppct = 0
        if (-not [int]::TryParse($ppctTxt, [ref]$ppct)) { continue }
        $code = $aliasMap[$monTxt.ToLower()]
        if (-not $code) { Warn "Môn '$monTxt' (PPCT $ppct) trong kế hoạch khối chưa có trong cauhinh.json - bỏ qua"; continue }
        [void]$hang[$code].Add([pscustomobject]@{ PPCT = $ppct; Ten = $ten; GhiChu = $dc; TenTKB = '' })
    }
    foreach ($code in $monCodes) {
        $sorted = @($hang[$code] | Sort-Object PPCT)
        $hang[$code] = New-Object System.Collections.ArrayList
        $prev = $null
        foreach ($x in $sorted) {
            # Tên bài hiển thị: bỏ "Bài n:", đổi "- Tiết n" thành "(Tn)", áp quy tắc riêng của môn
            $t = ApDungQuyTac (CleanTenBai $x.Ten) $cfg.monHoc.$code.quyTacTenBai
            # Kế hoạch khối ghi trùng số tiết (VD: 2 lần "Luyện tập chung (T1)") -> tự tăng thành (T2)
            if ($prev -and $t -match '^(.*)\(T(\d+)\)$') {
                $goc = $Matches[1]
                if ($prev -match '^(.*)\(T(\d+)\)$' -and $Matches[1] -eq $goc) {
                    $k = [int]$Matches[2]
                    $mine = [int]([regex]::Match($t, '\(T(\d+)\)$').Groups[1].Value)
                    if ($mine -le $k) {
                        $t = "$goc(T$($k + 1))"
                        Warn "$($cfg.monHoc.$code.tenTKB) PPCT $($x.PPCT): kế hoạch khối ghi trùng số tiết, đã sửa thành '$t'"
                    }
                }
            }
            $x.TenTKB = $t; $prev = $t
            [void]$hang[$code].Add($x)
        }
    }

    # =================================================================
    # 2. XẾP TIẾT VÀO THỜI KHOÁ BIỂU CỐ ĐỊNH CỦA LỚP
    # =================================================================
    Info "2. Xếp tiết vào thời khoá biểu lớp..."
    $slots = New-Object System.Collections.ArrayList
    $ptr = @{}; foreach ($code in $monCodes) { $ptr[$code] = 0 }
    foreach ($ses in $cfg.thoiKhoaBieu) {
        foreach ($code in $ses.mon) {
            $mc = MonCfg $code
            if (-not $mc) { throw "Mã môn '$code' trong thoiKhoaBieu không có trong monHoc" }
            $lesson = $null
            $nghi = $dsNgayNghi -contains (NgayCua $ses.thu).Date
            if ($nghi) { }
            elseif ($ptr[$code] -lt $hang[$code].Count) { $lesson = $hang[$code][$ptr[$code]]; $ptr[$code]++ }
            else { Warn "Thứ $($ses.thu) ($($ses.buoi)): môn $($mc.tenTKB) không còn tiết nào trong kế hoạch khối tuần $Tuan" }
            [void]$slots.Add([pscustomobject]@{
                Thu = (N $ses.thu); Buoi = (N $ses.buoi); Code = $code; Lesson = $lesson; Nghi = $nghi
                Mode = ''; Local = 0; Block = $null; PPCTs = @(); Skip = $false; Lessons = @($lesson)
            })
        }
    }
    foreach ($code in $monCodes) {
        for ($i = $ptr[$code]; $i -lt $hang[$code].Count; $i++) {
            Warn "Môn $((MonCfg $code).tenTKB) còn tiết PPCT $($hang[$code][$i].PPCT) chưa xếp được (TKB thiếu chỗ)"
        }
    }

    # =================================================================
    # 3. TÌM NỘI DUNG TỪNG TIẾT TRONG FILE MÔN HỌC
    # =================================================================
    Info "3. Tìm nội dung bài dạy trong các file môn..."
    $srcCache = @{}

    function Get-Blocks($sd) {
        # Đọc toàn bộ đoạn văn một lượt (không dùng Find để tránh Word treo vì hộp thoại ẩn)
        $paras = New-Object System.Collections.ArrayList
        foreach ($p in $sd.Paragraphs) {
            $r = $p.Range
            [void]$paras.Add([pscustomobject]@{ S = $r.Start; E = $r.End; T = (N $r.Text) })
        }
        $hdr = New-Object System.Collections.ArrayList
        for ($j = 0; $j -lt $paras.Count; $j++) {
            if ($paras[$j].T.TrimStart() -cmatch '^KẾ HOẠCH') { [void]$hdr.Add($j) }
        }
        if ($hdr.Count -eq 0) {
            # File không có dòng "KẾ HOẠCH..." -> mỗi bài bắt đầu bằng dòng "BÀI n." (ngoài bảng)
            for ($j = 0; $j -lt $paras.Count; $j++) {
                if ($paras[$j].T.TrimStart() -cmatch '^BÀI\s*\d+') {
                    $pp = $sd.Range($paras[$j].S, $paras[$j].S)
                    if ($pp.Tables.Count -eq 0) { [void]$hdr.Add($j) }
                }
            }
        }
        if ($hdr.Count -eq 0) { [void]$hdr.Add(0) }
        $blocks = New-Object System.Collections.ArrayList
        $from = 1
        for ($k = 0; $k -lt $hdr.Count; $k++) {
            $i0 = $hdr[$k]
            $i1 = if ($k + 1 -lt $hdr.Count) { $hdr[$k + 1] } else { $paras.Count }
            # Phần thân bắt đầu ở đoạn "I. YÊU CẦU CẦN ĐẠT" (hoặc "MỤC TIÊU")
            $b = $i0
            for ($j = $i0; $j -lt $i1; $j++) { if ($paras[$j].T -match 'YÊU CẦU CẦN ĐẠT|MỤC TIÊU') { $b = $j; break } }
            # Các dòng tiêu đề nằm giữa dòng "KẾ HOẠCH..." và phần thân
            $titleLines = New-Object System.Collections.ArrayList
            for ($j = $i0; $j -lt $b; $j++) {
                foreach ($piece in ($paras[$j].T -split "[\r\v\a\f\n]")) {
                    $t = CleanText $piece
                    if ($t) { [void]$titleLines.Add($t) }
                }
            }
            # Bỏ các đoạn trống / ngắt trang ở cuối bài
            $e = $i1 - 1
            while ($e -gt $b -and $paras[$e].T -match '^[\r\f\x0c\a ]*$') { $e-- }
            # Số tiết của bài (VD: "TIẾT 1 - 2" = 2 tiết)
            $span = 1
            foreach ($t in $titleLines) {
                $mm = [regex]::Match($t, 'TIẾT\s*(\d+)\s*[-–]\s*(\d+)', 'IgnoreCase')
                if ($mm.Success) { $span = [int]$mm.Groups[2].Value - [int]$mm.Groups[1].Value + 1; break }
            }
            [void]$blocks.Add([pscustomobject]@{ Start = $paras[$b].S; End = $paras[$e].E; Titles = $titleLines; From = $from; To = $from + $span - 1 })
            $from += $span
        }
        return ,$blocks
    }

    function Get-Source([string]$code, [int]$wk) {
        $mc = MonCfg $code
        $path = Join-Path (Path-Of $mc.thuMuc) ($mc.mauTenFile -f $wk)
        if ($srcCache.ContainsKey($path)) { return $srcCache[$path] }
        $entry = $null
        if (Test-Path -LiteralPath $path) {
            Write-Host "   - mở $([IO.Path]::GetFileName($path))"
            $sd = $word.Documents.Open($path, $false, $true)
            [void]$openDocs.Add($sd)
            $entry = [pscustomobject]@{ Path = $path; Doc = $sd; Blocks = (Get-Blocks $sd) }
            # File môn có số tiết khác với số tiết/tuần -> có thể thiếu bài hoặc dòng mở đầu bài viết khác quy ước
            $tong = 0; if ($entry.Blocks.Count -gt 0) { $tong = $entry.Blocks[$entry.Blocks.Count - 1].To }
            if ($mc.soTietTuan -and $tong -ne [int]$mc.soTietTuan) {
                Warn "File $([IO.Path]::GetFileName($path)): đọc được $($entry.Blocks.Count) bài = $tong tiết, nhưng môn $($mc.tenTKB) có $($mc.soTietTuan) tiết/tuần - kiểm tra dòng mở đầu từng bài ('KẾ HOẠCH...' hoặc 'BÀI n.') và ghi chú 'TIẾT 1 - 2'"
            }
        } else {
            Warn "Không thấy file: $path"
        }
        $srcCache[$path] = $entry
        return $entry
    }

    # Trả về các dòng tiêu đề bài (mỗi phần tử là một dòng in hoa)
    function Get-Title([string]$code, $titles, $lessons) {
        $mc = MonCfg $code
        $lessons = @($lessons | Where-Object { $_ })
        $L = @($titles | Where-Object {
            $_ -notmatch '^KẾ HOẠCH' -and $_ -notmatch '^(CHỦ ĐỀ|Thời gian)' -and
            $_ -notmatch '^TUẦN\s*\d+\s*(:\s*TIẾT\s*\d+)?\s*$'
        })
        # Tên bài lấy từ kế hoạch khối (đã áp quy tắc), gộp các tiết: "A; B"
        $khoi = (@($lessons | ForEach-Object { TieuDeTuTen $_.TenTKB } | Where-Object { $_ } | Select-Object -Unique)) -join '; '
        $res = @()
        switch ($mc.kieuTieuDe) {
            'tiet' {
                $t = $L | Where-Object { $_ -match '^TIẾT' } | Select-Object -First 1
                $src = ''
                if ($t) {
                    $src = $t -replace '^TIẾT\s*\d+(\s*[-–]\s*\d+)?\s*[:.\-–]?\s*', ''
                    $src = [regex]::Replace($src, '\s+–\s+', ': ', 1)   # "ĐỌC – CHỮ A" -> "ĐỌC: CHỮ A"
                }
                $pre = { param($x) ($x -split '[:;]')[0].Trim() }
                if (-not $src) { $src = $khoi }
                elseif ($khoi -and (& $pre $khoi) -eq (& $pre $src) -and $khoi.Length -gt $src.Length) { $src = $khoi }
                if ($src) { $res = @($src) }
            }
            'bai' {
                $bai  = $L | Where-Object { $_ -cmatch '^BÀI' } | Select-Object -First 1
                $tiet = $L | Where-Object { $_ -match '^TIẾT\s*\d+' } | Select-Object -First 1
                if ($bai) {
                    $r1 = $bai -creplace '^BÀI\s*\d+\s*[.:\-–]\s*', ''
                    $k = $null
                    if ($tiet) { $k = [regex]::Match($tiet, '\d+').Value }
                    elseif ($lessons.Count -gt 0) {
                        $mk = [regex]::Match($lessons[0].TenTKB, '(?:\(T|Tiết\s*)(\d+)')
                        if ($mk.Success) { $k = $mk.Groups[1].Value }
                    }
                    if ($k -and $r1 -notmatch 'TIẾT\s*\d') { $r1 = "$r1 (TIẾT $k)" }
                    $res = @($r1)
                }
            }
            'giuNguyen' {
                # Giữ nguyên dòng "CHỦ ĐỀ ..." và "BÀI ..." như file môn (TNXH)
                $res = @($titles | Where-Object { $_ -cmatch '^(CHỦ ĐỀ|BÀI)' })
            }
        }
        if ($res.Count -eq 0 -and $L.Count -gt 0) { $res = @(($L -join ' - ') -creplace '^BÀI\s*\d+\s*[.:]\s*', '') }
        if ($res.Count -eq 0 -and $khoi) { $res = @($khoi) }
        # Tiêu đề dài có " – " thì xuống dòng tại đó
        $out = @()
        foreach ($x in $res) {
            $x = $x.ToUpper().Trim()
            if ($x.Length -gt 60 -and $x -match ' – ') { $out += ($x -split ' – ', 2 | ForEach-Object { $_.Trim() }) }
            elseif ($x -match '; ' -and ($lessons.Count -gt 1 -or $x.Length -gt 75)) {
                # Tiêu đề gộp 2 tiết hoặc quá dài: xuống dòng sau dấu ";"
                $ix = $x.IndexOf('; '); $out += $x.Substring(0, $ix + 1); $out += $x.Substring($ix + 2).Trim()
            }
            else { $out += $x }
        }
        return ,$out
    }

    foreach ($sl in $slots) {
        $mc = MonCfg $sl.Code
        $sl.Mode = $mc.cheDo
        if ($sl.Nghi) { $sl.Mode = 'nghi'; continue }
        if (-not $sl.Lesson) { $sl.Mode = 'trong'; continue }
        $sl.PPCTs = @($sl.Lesson.PPCT)
        if ($mc.soTietTuan) {
            $wk = [int][Math]::Ceiling($sl.Lesson.PPCT / [double]$mc.soTietTuan)
            $sl.Local = $sl.Lesson.PPCT - ($wk - 1) * $mc.soTietTuan
            if ($mc.tietGVBMTrongTuan -and (@($mc.tietGVBMTrongTuan) -contains $sl.Local)) { $sl.Mode = 'gvbm_tieude' }
            if ($sl.Mode -in @('soan', 'gvbm_tieude')) {
                $src = Get-Source $sl.Code $wk
                if ($src) {
                    $blk = $src.Blocks | Where-Object { $_.From -le $sl.Local -and $_.To -ge $sl.Local } | Select-Object -First 1
                    if ($blk) { $sl.Block = $blk }
                    elseif ($sl.Mode -eq 'soan') { Warn "$($mc.tenTKB) PPCT $($sl.Lesson.PPCT): file tuần $wk không có tiết thứ $($sl.Local) (file có $($src.Blocks.Count) bài)" }
                }
            }
        }
    }
    # Gộp các tiết liền nhau cùng một bài (VD: Tiếng Việt Tiết 1-2 Đọc)
    for ($i = 1; $i -lt $slots.Count; $i++) {
        $a = $slots[$i - 1]; $b = $slots[$i]
        if ($b.Mode -eq 'soan' -and $a.Mode -eq 'soan' -and $a.Block -and ($a.Block -eq $b.Block) -and $a.Thu -eq $b.Thu -and $a.Buoi -eq $b.Buoi) {
            $first = $a; $j = $i - 1
            while ($first.Skip -and $j -gt 0) { $j--; $first = $slots[$j] }
            $first.PPCTs = @($first.PPCTs) + $b.Lesson.PPCT
            $first.Lessons = @($first.Lessons) + $b.Lesson
            $b.Skip = $true
        }
    }

    # =================================================================
    # 4. TẠO FILE GIÁO ÁN TUẦN TỪ MẪU KHUNG
    # =================================================================
    Info "4. Tạo file giáo án..."
    $outDir = $cfg.thuMucGoc
    if ($ThuMucKetQua) {
        if (-not (Test-Path -LiteralPath $ThuMucKetQua)) { [void](New-Item -ItemType Directory -Force -Path $ThuMucKetQua) }
        $outDir = (Resolve-Path -LiteralPath $ThuMucKetQua).Path
    }
    $outPath = Join-Path $outDir ($cfg.tenFileKetQua -f $Tuan)
    if (Test-Path -LiteralPath $outPath) {
        # File cũ đang mở trong Word -> lưu sang tên mới, không báo lỗi
        $locked = $false
        try { $fs = [IO.File]::Open($outPath, 'Open', 'ReadWrite', 'None'); $fs.Close() } catch { $locked = $true }
        if ($locked) {
            $outPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($outPath),
                [IO.Path]::GetFileNameWithoutExtension($outPath) + " (bản " + (Get-Date -Format 'HH''h''mm') + ").docx")
            Warn "File cũ đang mở trong Word nên lưu sang tên mới: $([IO.Path]::GetFileName($outPath))"
        }
    }
    Copy-Item -LiteralPath (Path-Of $cfg.fileMauKhung) -Destination $outPath -Force
    $doc = $word.Documents.Open($outPath)
    [void]$openDocs.Add($doc)

    $lastThu = ($slots | Select-Object -Last 1).Thu
    $friday = NgayCua $lastThu
    function Set-RangeText($rng, [string]$text) { [void]$rng.MoveEnd(1, -1); $rng.Text = $text }

    for ($i = 1; $i -le [Math]::Min(15, $doc.Paragraphs.Count); $i++) {
        $p = $doc.Paragraphs.Item($i); $t = CleanText $p.Range.Text
        if ($t -match '^KẾ HOẠCH DẠY HỌC TUẦN') { Set-RangeText $p.Range "KẾ HOẠCH DẠY HỌC TUẦN $Tuan" }
        elseif ($t -match '^\(Từ ngày') { Set-RangeText $p.Range "(Từ ngày $($monday.ToString('dd/MM/yyyy')) đến ngày $($friday.ToString('dd/MM/yyyy')))" }
    }

    # Phần đầu trang: in đậm các chữ trong "chuInDamDauTrang" (tính từ chữ đó đến dấu tab hoặc hết dòng)
    foreach ($cum in @($cfg.chuInDamDauTrang)) {
        if (-not $cum) { continue }
        $cum = N $cum
        for ($i = 1; $i -le [Math]::Min(6, $doc.Paragraphs.Count); $i++) {
            $pr = $doc.Paragraphs.Item($i).Range
            $raw = N $pr.Text
            $ix = $raw.IndexOf($cum)
            if ($ix -lt 0) { continue }
            $stop = $raw.IndexOfAny([char[]]"`t`r`a`v", $ix)
            if ($stop -lt 0) { $stop = $raw.Length }
            $doc.Range($pr.Start + $ix, $pr.Start + $stop).Font.Bold = -1
            break
        }
    }

    # Bảng thời khoá biểu
    $t1 = $doc.Tables.Item(1)
    if ($t1.Rows.Count - 1 -ne $slots.Count) {
        throw "Mẫu khung có $($t1.Rows.Count - 1) dòng tiết nhưng thoiKhoaBieu có $($slots.Count) tiết. Hãy sửa mẫu MAU_KHUNG.docx cho khớp."
    }
    $prevThu = $null
    for ($i = 0; $i -lt $slots.Count; $i++) {
        $sl = $slots[$i]; $mc = MonCfg $sl.Code; $r = $i + 2
        $ppct = ''; $ten = ''; $gc = ''
        if ($sl.Nghi) { $ten = N $chuNghi }
        if ($sl.Lesson) {
            $ppct = "$($sl.Lesson.PPCT)"
            if ($sl.Mode -eq 'gvbm_tieude' -or ($sl.Mode -eq 'gvbm' -and -not $mc.hienTenBaiTrongTKB)) { $ten = N $cfg.chuGVBM }
            else {
                $ten = $sl.Lesson.TenTKB; $gc = ApDungGhiChu $sl.Lesson.GhiChu
                if (-not $ten -and $sl.Block) {
                    # Kế hoạch khối để trống tên bài -> lấy tên từ file môn học
                    $tt = ((Get-Title $sl.Code $sl.Block.Titles $null) -join ' – ').ToLower()
                    if ($tt) { $ten = ApDungQuyTac ($tt.Substring(0, 1).ToUpper() + $tt.Substring(1)) $mc.quyTacTenBai }
                    Warn "$($mc.tenTKB) PPCT ${ppct}: kế hoạch khối không có tên bài, đã lấy từ file môn: '$ten'"
                }
            }
        }
        Set-RangeText $t1.Cell($r, 4).Range $ppct
        Set-RangeText $t1.Cell($r, 5).Range (N $mc.tenTKB)
        Set-RangeText $t1.Cell($r, 6).Range $ten
        Set-RangeText $t1.Cell($r, 7).Range $gc
        if ($sl.Thu -ne $prevThu) {
            # Ô "Thứ/ngày": chỉ thay phần ngày dd/MM, giữ nguyên chữ "Hai", "Ba"...
            $rng = $t1.Cell($r, 1).Range
            [void]$rng.MoveEnd(1, -1)
            $old = $rng.Text
            $newTxt = [regex]::Replace($old, '\d{1,2}/\d{1,2}', (NgayCua $sl.Thu).ToString('dd/MM'))
            if ($newTxt -ne $old) { $rng.Text = $newTxt }
            $prevThu = $sl.Thu
        }
    }

    # ---------- Phần nội dung bài dạy ----------
    function Add-Para([string]$text, [int]$size = 13, [bool]$bold = $false, [int]$align = 1, [int]$ulLen = 0, [int]$color = -1, [switch]$Keep) {
        $end = $doc.Content.End - 1
        $r = $doc.Range($end, $end)
        $r.InsertAfter($text)
        $r.InsertParagraphAfter()
        $r.Style = $doc.Styles.Item(-1)
        $r.Font.Reset()
        $r.Font.Name = 'Times New Roman'; $r.Font.Size = $size; $r.Font.Bold = [int]$bold
        $r.Font.Italic = 0; $r.Font.Underline = 0
        if ($color -ge 0) { $r.Font.Color = $color }
        $pf = $r.ParagraphFormat
        $pf.Alignment = $align; $pf.LeftIndent = 0; $pf.FirstLineIndent = 0
        $pf.SpaceBefore = 0; $pf.SpaceAfter = 0; $pf.LineSpacingRule = 0
        # Giữ dòng tiêu đề đi cùng đoạn sau, không bị lẻ ở cuối trang
        $pf.KeepWithNext = $(if ($Keep) { -1 } else { 0 })
        if ($ulLen -gt 0) { $doc.Range($r.Start, $r.Start + $ulLen).Font.Underline = 1 }
        $script:lastPara = $r
    }
    function Add-Heading([string]$tenDayDu, $ppcts) {
        $ten = N $tenDayDu
        Add-Para "${ten}: TIẾT $((@($ppcts)) -join ', ')" 13 $true 1 $ten.Length -Keep
    }
    # Dòng ngăn cách giữa các bài: "(((" được thay bằng ký hiệu hoa văn + quyển sách
    function Add-Separator {
        Add-Para $sep 13 $false 1
        $syms = @($cfg.kyHieuNganCach)
        if ($syms.Count -eq 0 -or -not $syms[0]) { return }
        $r = $script:lastPara
        $ix = $r.Text.IndexOf('(((')
        if ($ix -lt 0) { return }
        $pos = $r.Start + $ix
        $doc.Range($pos, $pos + 3).Delete() | Out-Null
        foreach ($s in $syms) {
            # mỗi phần tử: "mã hex|phông", VD "F026|Wingdings" (quyển sách), "2619|Segoe UI Symbol" (hoa văn)
            $code, $font = ($s -split '\|', 2)
            $n = [Convert]::ToInt32($code, 16); if ($n -gt 32767) { $n -= 65536 }
            $ins = $doc.Range($pos, $pos)
            $ins.InsertSymbol($n, $font, $true)
            $doc.Range($pos, $pos + 1).Font.Size = 13
            $pos++
        }
    }
    function Add-Title($lines, [bool]$bold = $true) {
        foreach ($ln in @($lines)) { if ($ln) { Add-Para $ln 13 $bold 1 -Keep } }
    }
    # Chiều rộng vùng chữ (để đặt tab chấm chấm "IV. ĐIỀU CHỈNH")
    $ps0 = $doc.PageSetup
    $textWidth = $ps0.PageWidth - $ps0.LeftMargin - $ps0.RightMargin
    function Fix-DieuChinh([int]$from) {
        # Sau dòng "IV. ĐIỀU CHỈNH..." thay các dòng chấm chấm bằng 2 dòng tab chấm như mẫu
        $pas = @(); foreach ($p in $doc.Range($from, $doc.Content.End).Paragraphs) { $pas += $p }
        $iv = -1
        for ($j = $pas.Count - 1; $j -ge 0; $j--) {
            if ((CleanText $pas[$j].Range.Text) -match '^IV\.') { $iv = $j; break }
        }
        if ($iv -lt 0 -or $pas[$iv].Range.Tables.Count -gt 0) { return }
        for ($j = $iv + 1; $j -lt $pas.Count; $j++) {
            if ($pas[$j].Range.Text -match '[\p{L}\d]') { return }   # có chữ -> không đụng vào
        }
        $doc.Range($pas[$iv].Range.End, $doc.Content.End - 1).Delete() | Out-Null
        for ($k = 0; $k -lt [int]$cfg.soDongDieuChinh; $k++) {
            Add-Para "`t" 13 $false 3
            $r = $script:lastPara
            $r.ParagraphFormat.LineSpacingRule = 5; $r.ParagraphFormat.LineSpacing = 13.8
            $r.ParagraphFormat.TabStops.ClearAll()
            [void]$r.ParagraphFormat.TabStops.Add($textWidth - 5, 2, 1)
        }
    }
    # ================= CHUẨN HOÁ NỘI DUNG BÀI (theo cách cô sửa tay) =================
    function Para-List($rng) {
        $L = New-Object System.Collections.ArrayList
        foreach ($p in $rng.Paragraphs) {
            [void]$L.Add([pscustomobject]@{ P = $p; T = (CleanText $p.Range.Text); InT = ($p.Range.Tables.Count -gt 0) })
        }
        return ,$L
    }
    function Set-ParaText($p, [string]$text) { $r = $p.Range; [void]$r.MoveEnd(1, -1); $r.Text = $text }
    function Is-UpperWord([string]$t) {
        $w = (($t -replace '^\d+\.\s*', '') -split '\s+')[0]
        return ($w.Length -ge 2 -and $w -cmatch '\p{Lu}' -and $w -ceq $w.ToUpper())
    }
    function Normalize-Content([int]$from) {
        $docEnd = { $doc.Content.End - 1 }
        $all = $doc.Range($from, (& $docEnd))
        $headEnd = if ($all.Tables.Count -gt 0) { $all.Tables.Item(1).Range.Start } else { & $docEnd }

        # --- 0. Phần đầu bài: dòng ngắt (Shift+Enter) -> đoạn riêng, để xử lý từng dòng an toàn ---
        foreach ($p in @($doc.Range($from, $headEnd).Paragraphs)) {
            if ($p.Range.Tables.Count -gt 0) { continue }
            $s0 = $p.Range.Start; $raw = $p.Range.Text
            for ($q = $raw.Length - 1; $q -ge 0; $q--) {
                if ([int]$raw[$q] -eq 11) { $doc.Range($s0 + $q, $s0 + $q + 1).Text = "`r" }
            }
        }
        $headEnd = if ($doc.Range($from, (& $docEnd)).Tables.Count -gt 0) { $doc.Range($from, (& $docEnd)).Tables.Item(1).Range.Start } else { & $docEnd }

        # --- 1. Mục I: "1. Kiến thức / 2. Phẩm chất / 3. Năng lực", bỏ "4. Tích hợp" ---
        # Chỉ thay phần chữ tiêu đề ở đầu đoạn, không đụng nội dung phía sau
        $repl = {
            param($x, [string]$re, [string]$new)
            $raw = $x.P.Range.Text
            $mm = [regex]::Match($raw, $re)
            if ($mm.Success -and $mm.Index -eq 0) { $s = $x.P.Range.Start; $doc.Range($s, $s + $mm.Length).Text = $new }
        }
        $H = Para-List $doc.Range($from, $headEnd)
        $hasNLDT = $false
        foreach ($x in $H) {
            if ($x.InT) { continue }
            if ($x.T -match '^1\.\s*Năng lực đặc thù') { & $repl $x '^1\.\s*Năng lực đặc thù[ \t]*:?' '1. Kiến thức:'; $hasNLDT = $true }
            elseif ($x.T -match '^2\.\s*Năng lực chung') { & $repl $x '^2\.\s*Năng lực chung[ \t]*:?' '3. Năng lực:' }
            elseif ($x.T -match '^3\.\s*Phẩm chất') { & $repl $x '^3\.\s*Phẩm chất' '2. Phẩm chất' }
            elseif ($x.T -match '^II\.\s*CHUẨN BỊ') { & $repl $x '^II\.\s*CHUẨN BỊ' 'II. ĐỒ DÙNG DẠY HỌC' }
        }
        $headEnd = if ($doc.Range($from, (& $docEnd)).Tables.Count -gt 0) { $doc.Range($from, (& $docEnd)).Tables.Item(1).Range.Start } else { & $docEnd }
        $H = Para-List $doc.Range($from, $headEnd)
        $iD = -1; $iII = -1
        for ($j = 0; $j -lt $H.Count; $j++) {
            if ($iD -lt 0 -and $H[$j].T -match '^4\.\s*Tích hợp') { $iD = $j }
            if ($iII -lt 0 -and $H[$j].T -match '^II\.\s') { $iII = $j }
        }
        if ($iD -ge 0 -and $iII -gt $iD) { $doc.Range($H[$iD].P.Range.Start, $H[$iII].P.Range.Start).Delete() | Out-Null }
        if ($hasNLDT) {
            $headEnd = if ($doc.Range($from, (& $docEnd)).Tables.Count -gt 0) { $doc.Range($from, (& $docEnd)).Tables.Item(1).Range.Start } else { & $docEnd }
            $H = Para-List $doc.Range($from, $headEnd)
            $iB = -1; $iC = -1; $iII = -1
            for ($j = 0; $j -lt $H.Count; $j++) {
                if ($iB -lt 0 -and $H[$j].T -match '^3\.\s*Năng lực') { $iB = $j }
                elseif ($iC -lt 0 -and $H[$j].T -match '^2\.\s*Phẩm chất') { $iC = $j }
                elseif ($iII -lt 0 -and $H[$j].T -match '^II\.\s') { $iII = $j }
            }
            if ($iB -ge 0 -and $iC -gt $iB -and $iII -gt $iC) {
                # Chuyển phần "2. Phẩm chất" lên trước "3. Năng lực"
                $cS = $H[$iC].P.Range.Start; $cE = $H[$iII].P.Range.Start; $len = $cE - $cS
                $bS = $H[$iB].P.Range.Start
                $doc.Range($bS, $bS).FormattedText = $doc.Range($cS, $cE).FormattedText
                $doc.Range($cS + $len, $cE + $len).Delete() | Out-Null
            }
        }
        # Đồ dùng GV: bỏ máy tính, tivi, máy chiếu
        $headEnd = if ($doc.Range($from, (& $docEnd)).Tables.Count -gt 0) { $doc.Range($from, (& $docEnd)).Tables.Item(1).Range.Start } else { & $docEnd }
        foreach ($x in (Para-List $doc.Range($from, $headEnd))) {
            $m = [regex]::Match($x.T, '^(-\s*GV\s*:|1\.\s*Giáo viên\s*:)\s*')
            if (-not $m.Success) { continue }
            $rest = $x.T.Substring($m.Length).TrimEnd('.')
            $parts = [regex]::Split($rest, '\s*([,;])\s*')
            $keep = New-Object System.Collections.ArrayList; $sepNext = ''
            for ($k = 0; $k -lt $parts.Count; $k += 2) {
                $item = $parts[$k].Trim()
                $sepAfter = if ($k + 1 -lt $parts.Count) { $parts[$k + 1] } else { '' }
                if ($item -match '^(máy tính|tivi(\s*/\s*máy chiếu)?|máy chiếu(\s*/\s*tivi)?)$') { continue }
                if ($keep.Count -gt 0) { [void]$keep.Add($sepNext + ' ') }
                [void]$keep.Add($item); $sepNext = $sepAfter
            }
            $newRest = (-join $keep).Trim()
            if ($newRest -and $newRest -ne $rest) {
                $newRest = $newRest.Substring(0, 1).ToUpper() + $newRest.Substring(1) + '.'
                $r = $x.P.Range; [void]$r.MoveEnd(1, -1)
                $doc.Range($r.Start + $m.Length, $r.End).Text = $newRest
            }
        }

        # --- 2. Sinh hoạt lớp: bỏ mục "HOẠT ĐỘNG TỔNG KẾT TUẦN", đánh số lại ---
        if ($cfg.boTongKetTuan -ne $false) {
            $A = Para-List $doc.Range($from, (& $docEnd))
            for ($j = 0; $j -lt $A.Count; $j++) {
                if (-not ($A[$j].InT -and $A[$j].T -match '^(\d+)\.\s*HOẠT ĐỘNG TỔNG KẾT TUẦN')) { continue }
                $num = [int]$Matches[1]; $nx = -1
                for ($k = $j + 1; $k -lt $A.Count; $k++) {
                    if ($A[$k].InT -and $A[$k].T -match '^(\d+)\.\s' -and $A[$k].T -notmatch '^\d+\.\d' -and (Is-UpperWord $A[$k].T)) { $nx = $k; break }
                }
                try {
                    $rs = $A[$j].P.Range.Rows.Item(1).Range.Start
                    $re = if ($nx -ge 0) { $A[$nx].P.Range.Rows.Item(1).Range.Start } else { $A[$j].P.Range.Tables.Item(1).Range.End }
                    $later = @(); if ($nx -ge 0) { for ($k = $nx; $k -lt $A.Count; $k++) { if ($A[$k].InT -and $A[$k].T -match '^(\d+)\.\s' -and $A[$k].T -notmatch '^\d+\.\d' -and (Is-UpperWord $A[$k].T)) { $later += $A[$k] } } }
                    $doc.Range($rs, $re).Rows.Delete()
                    foreach ($x in $later) {
                        $n2 = [int]([regex]::Match($x.T, '^\d+').Value)
                        if ($n2 -gt $num) { $r = $x.P.Range; $doc.Range($r.Start, $r.Start + "$n2".Length).Text = "$($n2 - 1)" }
                    }
                } catch { Warn "Không xoá được mục 'Tổng kết tuần' trong Sinh hoạt lớp - cô xoá tay giúp" }
                break
            }
        }

        # --- 3. Từng dòng: bỏ "Mục tiêu", "Cách tiến hành", nhãn tích hợp, tách tiêu đề hoạt động ---
        $A = Para-List $doc.Range($from, (& $docEnd))
        $del = New-Object 'System.Collections.Generic.HashSet[int]'
        for ($j = 0; $j -lt $A.Count; $j++) {
            $t = $A[$j].T
            if ($t -match '^(-\s*)?Mục tiêu\b') {
                [void]$del.Add($j)
                if ($t -match ':\s*$') {   # "- Mục tiêu:" rồi các gạch đầu dòng bên dưới
                    for ($k = $j + 1; $k -lt $A.Count -and $A[$k].T -match '^-\s' -and $A[$k].T -notmatch 'Cách tiến hành'; $k++) { [void]$del.Add($k) }
                }
            } elseif ($t -match '^\*\s*Cách tiến hành') { [void]$del.Add($j) }
            elseif ($A[$j].InT -and $t -match '^\d\.\s+(Khởi động|Khám phá|Thực hành|Luyện tập|Vận dụng)') {
                # Ô tiêu đề hoạt động (gộp cả hàng): bỏ các dòng mô tả bên dưới tiêu đề
                try {
                    $pr = $A[$j].P.Range
                    $cellR = $pr.Cells.Item(1).Range
                    if ($pr.Start -eq $cellR.Start -and $pr.Rows.Item(1).Cells.Count -eq 1) {
                        for ($k = $j + 1; $k -lt $A.Count -and $A[$k].P.Range.Start -lt $cellR.End; $k++) {
                            if ($A[$k].T -and $A[$k].T -notmatch '^(-\s*)?Hoạt động\s*\d' -and $A[$k].T -notmatch '^\d\.') { [void]$del.Add($k) }
                        }
                    }
                } catch {}
            }
        }
        for ($j = $A.Count - 1; $j -ge 0; $j--) {
            $x = $A[$j]; $p = $x.P; $t = $x.T
            if ($del.Contains($j)) {
                $r = $p.Range
                if ($x.InT) {
                    $cell = $r.Cells.Item(1).Range
                    if ($cell.Paragraphs.Count -le 1) { Set-ParaText $p '' }
                    elseif ($r.End -ge $cell.End) { $doc.Range($r.Start - 1, $r.End - 1).Delete() | Out-Null }
                    else { $r.Delete() | Out-Null }
                } else { $r.Delete() | Out-Null }
                continue
            }
            $s = $p.Range.Start
            # Nhãn tích hợp đầu dòng -> "- "
            $m = [regex]::Match($t, '^(-\s*)?((Câu hỏi dẫn|Diễn giải)\s+)?tích hợp[^:]{0,100}:\s*', 'IgnoreCase')
            if ($m.Success -and $t -notmatch '^\d') {
                $doc.Range($s, $s + $m.Length).Text = '- '
                continue
            }
            if ($x.InT -and $t -cmatch '^HS (trả lời|lắng nghe|chia sẻ|nêu)') { $doc.Range($s, $s).Text = '- '; continue }
            # Tiêu đề hoạt động: "2. Khám phá kiến thức - Hoạt động 1: ... (27 phút) – mô tả"
            if ($x.InT -and $t -match '^\d\.\s+\S') {
                $raw = $p.Range.Text
                $mt = [regex]::Match($raw, '\(\d+(\s*-\s*\d+)?\s*phút\)(\s+–\s+[^\r\a]*)')
                if ($mt.Success) { $g = $mt.Groups[2]; $doc.Range($s + $g.Index, $s + $g.Index + $g.Length).Delete() | Out-Null }
                $raw = $p.Range.Text
                $ms = [regex]::Match($raw, '^\d\.\s+[^\r\a]+?(\s+-\s+)Hoạt động\s+\d+\s*:')
                if ($ms.Success) { $g = $ms.Groups[1]; $doc.Range($s + $g.Index, $s + $g.Index + $g.Length).Text = "`r" }
            }
        }
    }

    # Sau khi chuẩn hoá: báo nếu bài thiếu mục I-IV hoặc còn sót cụm lẽ ra đã được sửa/xoá
    function Check-Content([int]$from, [string]$nhan) {
        $lines = @((N $doc.Range($from, $doc.Content.End - 1).Text) -split "[\r\n\a\v\f]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $thieu = @()
        foreach ($muc in 'I', 'II', 'III', 'IV') {
            if (-not ($lines | Where-Object { $_ -cmatch "^$muc\.\s" } | Select-Object -First 1)) { $thieu += $muc }
        }
        if ($thieu.Count -gt 0) { Warn "${nhan}: bài không có mục $($thieu -join ', ') - kiểm tra lại nội dung" }
        foreach ($re in @($cfg.tuKhoaCanhBao)) {
            if (-not $re) { continue }
            $hit = @($lines | Where-Object { $_ -match (N $re) })
            if ($hit.Count -gt 0) {
                $vd = $hit[0]; if ($vd.Length -gt 60) { $vd = $vd.Substring(0, 60) + '…' }
                Warn "${nhan}: còn $($hit.Count) dòng chưa được sửa theo khung, VD: '$vd'"
            }
        }
    }

    function Paste-Block($src, $blk, [string]$nhan = '') {
        $end = $doc.Content.End - 1
        $dst = $doc.Range($end, $end)
        $dst.FormattedText = $src.Doc.Range($blk.Start, $blk.End).FormattedText
        if ($cfg.chuanHoaNoiDung -ne $false) {
            try { Normalize-Content $end } catch { Warn "Chuẩn hoá nội dung gặp lỗi ở một bài: $($_.Exception.Message)" }
            try { Check-Content $end $nhan } catch {}
        }
        $pasted = $doc.Range($end, $doc.Content.End - 1)
        # Thu nhỏ hình rộng hơn "anhRongToiDaCm" (giữ tỉ lệ); 0 = không đụng tới hình
        $anhMax = 0.0; if ($cfg.anhRongToiDaCm) { $anhMax = [double]$cfg.anhRongToiDaCm * 28.3465 }
        if ($anhMax -gt 0) {
            foreach ($sh in $pasted.InlineShapes) {
                try {
                    if ($sh.Width -gt $anhMax) { $h = $sh.Height * $anhMax / $sh.Width; $sh.LockAspectRatio = 0; $sh.Width = [single]$anhMax; $sh.Height = [single]$h }
                } catch {}
            }
        }
        # Chữ đen hết (bỏ màu đỏ/xanh của file môn) và cỡ chữ thống nhất
        if ($cfg.chuDen -ne $false) { $pasted.Font.Color = -16777216 }
        if ($cfg.coChu -gt 0) { $pasted.Font.Size = [single]$cfg.coChu }
        foreach ($p in $pasted.Paragraphs) {
            # Đoạn nào đang căn trái -> căn đều hai bên (giữ nguyên đoạn căn giữa / căn phải)
            if ($cfg.canDeuHaiBen -ne $false -and $p.Alignment -eq 0) { $p.Alignment = 3 }
            # Tiêu đề mục I, II, III, IV và "1. Kiến thức / 2. Phẩm chất / 3. Năng lực": đậm, không nghiêng
            $tt = $p.Range.Text
            # Mục IV: thống nhất một tên cho mọi tiết (không có "(nếu có)")
            if ($tt -match '^\s*IV\.\s') {
                $mIV = [regex]::Match($tt, '^\s*IV\.[^\r\a\v]*')
                $newIV = if ($cfg.tieuDeMucIV) { N $cfg.tieuDeMucIV } else { ($mIV.Value -replace '\s*\(\s*nếu\s+có\s*\)', '').Trim() }
                if ($mIV.Value -ne $newIV) {
                    $s0 = $p.Range.Start
                    $doc.Range($s0 + $mIV.Index, $s0 + $mIV.Index + $mIV.Length).Text = $newIV
                    $tt = $p.Range.Text
                }
            }
            if ($tt -match '^\s*(I|II|III|IV)\.\s' -or $tt -match '^\s*\d\.\s*(Kiến thức|Phẩm chất|Năng lực)\s*:?\s*[\r\a]') {
                if ($p.Range.Tables.Count -eq 0) { $p.Range.Font.Italic = 0; $p.Range.Font.Bold = -1 }
            }
        }
        if ($cfg.bangRongToanTrang -ne $false) {
            # Bảng hoạt động giãn vừa khổ giấy (giữ tỉ lệ các cột)
            foreach ($tb in $doc.Range($end, $doc.Content.End - 1).Tables) {
                try { $tb.Rows.LeftIndent = 0; $tb.AutoFitBehavior(2) } catch {}
            }
        }
        if ($cfg.soDongDieuChinh -gt 0) { Fix-DieuChinh $end }
    }
    $RED = 255
    $sep = N $cfg.dongNganCach
    $gvbm = N $cfg.chuGVBM

    $curThu = $null; $curBuoi = $null; $hasPM = $false
    Add-Para '' 13
    foreach ($sl in $slots) {
        if ($sl.Skip) { continue }
        $mc = MonCfg $sl.Code
        if ($sl.Mode -eq 'nghi') {
            # Ngày nghỉ: chỉ ghi một lần "Thứ ... ngày ..." + "Nghỉ"
            if ($sl.Thu -ne $curThu) {
                $d = NgayCua $sl.Thu
                Add-Para "Thứ $($sl.Thu) ngày $($d.ToString('dd')) tháng $($d.ToString('MM')) năm $($d.ToString('yyyy'))" 13 $false 1 -Keep
                if ($null -eq $curThu -and $cfg.trangMoiChoBaiDau -ne $false) { $script:lastPara.ParagraphFormat.PageBreakBefore = -1 }
                Add-Para (N $chuNghi).ToUpper() 13 $true 1
                Add-Separator
                $curThu = $sl.Thu; $curBuoi = $null
            }
            continue
        }
        if ($sl.Thu -ne $curThu) {
            $d = NgayCua $sl.Thu
            Add-Para "Thứ $($sl.Thu) ngày $($d.ToString('dd')) tháng $($d.ToString('MM')) năm $($d.ToString('yyyy'))" 13 $false 1 -Keep
            # Ngày đầu tiên (tiết 1 thứ Hai) luôn bắt đầu ở trang mới, tách khỏi trang kế hoạch tuần
            if ($null -eq $curThu -and $cfg.trangMoiChoBaiDau -ne $false) { $script:lastPara.ParagraphFormat.PageBreakBefore = -1 }
            $curThu = $sl.Thu; $curBuoi = $null
            $thuNay = $sl.Thu
            $hasPM = @($slots | Where-Object { $_.Thu -eq $thuNay } | Select-Object -ExpandProperty Buoi -Unique).Count -gt 1
        }
        if ($hasPM -and $sl.Buoi -ne $curBuoi) { Add-Para "BUỔI $($sl.Buoi.ToUpper()):" 13 $false 3 -Keep }
        $curBuoi = $sl.Buoi

        switch ($sl.Mode) {
            'trong' {
                Add-Heading $mc.tenDayDu @('...')
                Add-Para "[Không có tiết $($mc.tenTKB) trong kế hoạch khối tuần này - kiểm tra lại]" 13 $false 1 0 $RED
            }
            'tieuDe' {
                Add-Heading $mc.tenDayDu $sl.PPCTs
                Add-Title @(TieuDeTuTen $sl.Lesson.TenTKB) $false
            }
            'gvbm' {
                Add-Heading $mc.tenDayDu $sl.PPCTs
                if ($mc.hienTieuDe) { Add-Title @(TieuDeTuTen $sl.Lesson.TenTKB) }
                Add-Para $gvbm 13 $false 1
            }
            'gvbm_tieude' {
                Add-Heading $mc.tenDayDu $sl.PPCTs
                $tt = if ($sl.Block) { Get-Title $sl.Code $sl.Block.Titles $sl.Lessons } else { @(TieuDeTuTen $sl.Lesson.TenTKB) }
                Add-Title $tt
                Add-Para $gvbm 13 $false 1
            }
            default {
                Add-Heading $mc.tenDayDu $sl.PPCTs
                if ($sl.Block) {
                    $wk = [int][Math]::Ceiling($sl.Lesson.PPCT / [double]$mc.soTietTuan)
                    $src = Get-Source $sl.Code $wk
                    Add-Title (Get-Title $sl.Code $sl.Block.Titles $sl.Lessons)
                    Add-Para '' 13 $true 1 -Keep
                    Paste-Block $src $sl.Block "Thứ $($sl.Thu), $($mc.tenTKB) PPCT $($sl.PPCTs -join ',')"
                    Write-Host ("   + {0,-4} {1,-12} PPCT {2}" -f $sl.Thu, $mc.tenTKB, ($sl.PPCTs -join ','))
                } else {
                    Add-Title @(TieuDeTuTen $sl.Lesson.TenTKB)
                    Add-Para "[CHƯA TÌM THẤY NỘI DUNG BÀI DẠY - cần dán thủ công]" 13 $true 1 0 $RED
                }
            }
        }
        Add-Separator
    }

    # Lặp lại bảng chữ ký ở cuối giáo án (như mẫu tuần 11)
    if ($cfg.chuKyCuoiGiaoAn -ne $false -and $doc.Tables.Count -ge 2) {
        Add-Para '' 13
        $sig = $doc.Tables.Item(2).Range.FormattedText
        $end = $doc.Content.End - 1
        $doc.Range($end, $end).FormattedText = $sig
    }

    $doc.Save()
    Info "`nXONG! File đã tạo: $outPath"
    # Ghi các lưu ý ra file cạnh giáo án để đọc lại sau khi đóng cửa sổ đen
    $luuYPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($outPath), "TUẦN $Tuan - LƯU Ý.txt")
    if ($warn.Count -gt 0) {
        Write-Host "`nCó $($warn.Count) lưu ý cần kiểm tra:" -ForegroundColor Yellow
        $warn | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
        $noiDung = @("LƯU Ý KHI TẠO GIÁO ÁN TUẦN $Tuan  (tạo lúc $(Get-Date -Format 'HH:mm dd/MM/yyyy'))", '') + @($warn | ForEach-Object { "- $_" })
        [IO.File]::WriteAllLines($luuYPath, [string[]]$noiDung, (New-Object Text.UTF8Encoding $true))
        Write-Host "Đã ghi các lưu ý vào: $luuYPath" -ForegroundColor Yellow
        if (-not $KhongMoLuuY) { try { Start-Process notepad.exe -ArgumentList "`"$luuYPath`"" } catch {} }
    } elseif (Test-Path -LiteralPath $luuYPath) {
        Remove-Item -LiteralPath $luuYPath -Force
    }
}
catch {
    $coLoi = $true
    Write-Host ''
    Write-Host 'CÓ LỖI, CHƯA TẠO XONG GIÁO ÁN:' -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  (dòng $($_.InvocationInfo.ScriptLineNumber) trong tao_giao_an.ps1)" -ForegroundColor DarkGray
    Write-Host '  Thử: đóng hết các file Word đang mở rồi chạy lại. Nếu vẫn lỗi, chụp màn hình này để nhờ sửa.' -ForegroundColor Yellow
}
finally {
    # Đóng mọi file, thoát Word và giải phóng hẳn để không sót tiến trình Word chạy ngầm
    foreach ($d in $openDocs) { try { $d.Close(0) } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($d) } catch {} }
    try { $word.Quit() } catch { Write-Host "  (Word chưa nhận lệnh thoát: $($_.Exception.Message))" -ForegroundColor DarkGray }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch {}
    $word = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    # Chờ Word thoát; nếu sau 10 giây vẫn còn chạy ngầm (không có cửa sổ) thì tắt hẳn
    if ($wordCuaToi.Count -eq 1) {
        $wp = Get-Process -Id $wordCuaToi[0] -ErrorAction SilentlyContinue
        if ($wp -and -not $wp.WaitForExit(10000)) {
            $wp.Refresh()
            if (-not $wp.MainWindowTitle) { try { $wp.Kill(); Write-Host '  (đã tắt tiến trình Word chạy ngầm còn sót)' -ForegroundColor DarkGray } catch {} }
        }
    }
}
if ($coLoi) { exit 1 }
