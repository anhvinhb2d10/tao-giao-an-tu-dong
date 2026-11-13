# =====================================================================
#  CỬA SỔ SOẠN GIÁO ÁN TUẦN - LỚP 2C
#  Chỉ là lớp vỏ: thu thập lựa chọn rồi gọi tao_giao_an.ps1.
#  Mọi cách ghép / chuẩn hoá giáo án vẫn nằm trong tao_giao_an.ps1.
#  (Các tham số bên dưới chỉ dùng khi tự kiểm tra, cô không cần quan tâm.)
# =====================================================================
param(
    [switch]$ChayThu,
    [int]$TuanThu = 11,
    [string]$CauHinh = "",
    [string]$ThuMucKetQua = "",
    [string]$AnhChup = "",
    [string]$FileKetQuaThu = "",
    [string[]]$NghiThu = @(),
    [int]$XoaThu = -1,
    [string]$TkbThu = ""
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$goc     = Split-Path -Parent $here
$tool    = Join-Path $here 'tao_giao_an.ps1'
$cfgPath = if ($CauHinh) { $CauHinh } else { Join-Path $here 'cauhinh.json' }
$utf8    = New-Object Text.UTF8Encoding $false

$script:cfg = $null
$script:ds = @{ ngayNghi = New-Object System.Collections.ArrayList; tuanNghi = New-Object System.Collections.ArrayList }
$script:mucNghi = New-Object System.Collections.ArrayList   # mỗi dòng trong danh sách: @{ Loai; Ngay }
$script:proc = $null; $script:docTask = $null; $script:fileKetQua = ''; $script:wordTruoc = @(); $script:daDung = $false

# ---------- Đọc cấu hình ----------
function Doc-CauHinh {
    try {
        $script:cfg = [IO.File]::ReadAllText($cfgPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        foreach ($k in 'ngayNghi', 'tuanNghi') {
            $script:ds[$k].Clear()
            foreach ($v in @($script:cfg.$k)) { if ($v) { [void]$script:ds[$k].Add("$v") } }
        }
        return $true
    } catch { $script:cfg = $null; return $false }
}
function Ngay([string]$s) { return [datetime]::ParseExact($s, 'yyyy-MM-dd', $null) }
function ThuHai-CuaNgay([datetime]$d) { return $d.Date.AddDays(-((([int]$d.DayOfWeek) + 6) % 7)) }
# Cùng cách tính với tao_giao_an.ps1: tuần nghỉ không tính là tuần học
function ThuHai-CuaTuan([int]$n) {
    $m = (Ngay $script:cfg.ngayThuHaiTuan1).AddDays(7 * ($n - 1))
    foreach ($d in @($script:ds.tuanNghi | ForEach-Object { ThuHai-CuaNgay (Ngay $_) } | Sort-Object -Unique)) { if ($d -le $m) { $m = $m.AddDays(7) } }
    return $m
}
$tenThu = @('Chủ nhật', 'thứ Hai', 'thứ Ba', 'thứ Tư', 'thứ Năm', 'thứ Sáu', 'thứ Bảy')

# ---------- Cửa sổ ----------
$fontThuong = New-Object Drawing.Font('Segoe UI', 10.5)
$fontDam    = New-Object Drawing.Font('Segoe UI', 10.5, [Drawing.FontStyle]::Bold)
$fontTo     = New-Object Drawing.Font('Segoe UI', 13, [Drawing.FontStyle]::Bold)

$form = New-Object Windows.Forms.Form
$phienBan = ''
try { $pb = Join-Path $here 'PHIEN_BAN.txt'; if (Test-Path -LiteralPath $pb) { $phienBan = '  (bản ' + ([IO.File]::ReadAllLines($pb)[0]).Trim() + ')' } } catch {}
$form.Text = 'Soạn giáo án tuần tự động' + $phienBan
$form.Font = $fontThuong
$form.ClientSize = New-Object Drawing.Size(760, 742)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'; $form.MaximizeBox = $false

function Them($ctl, $cha, [int]$x, [int]$y, [int]$w, [int]$h) {
    $ctl.Location = New-Object Drawing.Point($x, $y); $ctl.Size = New-Object Drawing.Size($w, $h)
    [void]$cha.Controls.Add($ctl); return $ctl
}
function Nhan([string]$chu, $cha, [int]$x, [int]$y, [int]$w, [int]$h = 26) {
    $l = New-Object Windows.Forms.Label; $l.Text = $chu; $l.TextAlign = 'MiddleLeft'
    return (Them $l $cha $x $y $w $h)
}
function Nut([string]$chu, $cha, [int]$x, [int]$y, [int]$w, [int]$h = 34) {
    $b = New-Object Windows.Forms.Button; $b.Text = $chu
    return (Them $b $cha $x $y $w $h)
}

# --- 1. Chọn tuần ---
$gbTuan = New-Object Windows.Forms.GroupBox; $gbTuan.Text = '1. Chọn tuần cần soạn'; $gbTuan.Font = $fontDam
[void](Them $gbTuan $form 14 10 732 74)
$lb = Nhan 'Tuần:' $gbTuan 16 30 54; $lb.Font = $fontThuong
$numTuan = New-Object Windows.Forms.NumericUpDown
$numTuan.Minimum = 1; $numTuan.Maximum = 45; $numTuan.Font = $fontTo; $numTuan.TextAlign = 'Center'
[void](Them $numTuan $gbTuan 72 26 76 32)
$lblNgay = Nhan '' $gbTuan 164 30 556; $lblNgay.Font = $fontThuong

# --- 2. Ngày nghỉ ---
$gbNghi = New-Object Windows.Forms.GroupBox; $gbNghi.Text = '2. Ngày nghỉ (bỏ qua nếu tuần này học bình thường)'; $gbNghi.Font = $fontDam
[void](Them $gbNghi $form 14 92 732 150)
$lstNghi = New-Object Windows.Forms.ListBox; $lstNghi.Font = $fontThuong; $lstNghi.IntegralHeight = $false
[void](Them $lstNghi $gbNghi 16 28 400 108)
$lb = Nhan 'Chọn ngày:' $gbNghi 430 28 90; $lb.Font = $fontThuong
$dtNghi = New-Object Windows.Forms.DateTimePicker; $dtNghi.Format = 'Custom'; $dtNghi.CustomFormat = 'dd/MM/yyyy'; $dtNghi.Font = $fontThuong
[void](Them $dtNghi $gbNghi 522 28 194 28)
$btnThemNgay = Nut 'Nghỉ ngày này' $gbNghi 430 64 140 32;  $btnThemNgay.Font = $fontThuong
$btnThemTuan = Nut 'Nghỉ cả tuần này' $gbNghi 576 64 140 32; $btnThemTuan.Font = $fontThuong
$btnXoaNghi  = Nut 'Xoá dòng đang chọn' $gbNghi 430 104 286 32; $btnXoaNghi.Font = $fontThuong

# --- 3. Tạo ---
$btnTao = Nut 'TẠO GIÁO ÁN' $form 14 252 590 50; $btnTao.Font = $fontTo
$btnTao.BackColor = [Drawing.Color]::FromArgb(22, 110, 60); $btnTao.ForeColor = [Drawing.Color]::White; $btnTao.FlatStyle = 'Flat'
$btnDung = Nut 'Dừng' $form 612 252 134 50; $btnDung.Enabled = $false
$bar = New-Object Windows.Forms.ProgressBar; $bar.Style = 'Blocks'; $bar.Minimum = 0; $bar.Maximum = 100
[void](Them $bar $form 14 310 732 14)
$lblTrangThai = Nhan 'Sẵn sàng.' $form 14 328 732 28; $lblTrangThai.Font = $fontDam

# --- Lưu ý ---
$gbLuuY = New-Object Windows.Forms.GroupBox; $gbLuuY.Text = 'Lưu ý cần kiểm tra trong giáo án'; $gbLuuY.Font = $fontDam
[void](Them $gbLuuY $form 14 362 732 176)
$txtLuuY = New-Object Windows.Forms.TextBox
$txtLuuY.Multiline = $true; $txtLuuY.ReadOnly = $true; $txtLuuY.ScrollBars = 'Vertical'; $txtLuuY.Font = $fontThuong
$txtLuuY.BackColor = [Drawing.Color]::White
[void](Them $txtLuuY $gbLuuY 12 26 708 138)

$btnMoFile   = Nut 'Mở giáo án vừa tạo' $form 14 548 240 38;  $btnMoFile.Enabled = $false; $btnMoFile.Font = $fontDam
$btnMoThuMuc = Nut 'Mở thư mục giáo án' $form 262 548 180 38
$btnCauHinh  = Nut 'Thiết lập lớp' $form 450 548 144 38
$btnHuongDan = Nut 'Hướng dẫn' $form 602 548 144 38

# --- Quá trình chạy ---
$lb = Nhan 'Quá trình chạy:' $form 14 592 300 24
$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Multiline = $true; $txtLog.ReadOnly = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.WordWrap = $false
$txtLog.Font = New-Object Drawing.Font('Consolas', 9); $txtLog.BackColor = [Drawing.Color]::FromArgb(246, 246, 246)
[void](Them $txtLog $form 14 618 732 112)

$timer = New-Object Windows.Forms.Timer; $timer.Interval = 150

# ---------- Hiển thị ----------
$mauThuong = [Drawing.Color]::Black
$mauTot  = [Drawing.Color]::FromArgb(22, 110, 60)
$mauLoi  = [Drawing.Color]::FromArgb(190, 30, 30)
$mauVang = [Drawing.Color]::FromArgb(160, 100, 0)
function TrangThai([string]$chu, $mau = $mauThuong) { $lblTrangThai.Text = $chu; $lblTrangThai.ForeColor = $mau }

function CapNhat-Ngay {
    if (-not $script:cfg) { $lblNgay.Text = ''; return }
    try {
        $m = ThuHai-CuaTuan ([int]$numTuan.Value)
        $nghi = @($script:ds.ngayNghi | ForEach-Object { Ngay $_ } | Where-Object { $_ -ge $m -and $_ -lt $m.AddDays(7) } | Sort-Object)
        $s = "Từ thứ Hai $($m.ToString('dd/MM/yyyy')) đến thứ Sáu $($m.AddDays(4).ToString('dd/MM/yyyy'))"
        if ($nghi.Count -gt 0) { $s += "   (nghỉ: " + (($nghi | ForEach-Object { $_.ToString('dd/MM') }) -join ', ') + ")" }
        $lblNgay.Text = $s
    } catch { $lblNgay.Text = 'Không tính được ngày - kiểm tra "ngayThuHaiTuan1" trong cấu hình.' }
}
function CapNhat-DanhSachNghi {
    $lstNghi.Items.Clear(); $script:mucNghi.Clear()
    $tam = @()
    foreach ($v in $script:ds.tuanNghi) { try { $d = ThuHai-CuaNgay (Ngay $v); $tam += [pscustomobject]@{ Loai = 'tuanNghi'; Goc = $v; D = $d; Chu = "Nghỉ cả tuần  $($d.ToString('dd/MM/yyyy')) - $($d.AddDays(4).ToString('dd/MM/yyyy'))" } } catch {} }
    foreach ($v in $script:ds.ngayNghi) { try { $d = Ngay $v; $tam += [pscustomobject]@{ Loai = 'ngayNghi'; Goc = $v; D = $d; Chu = "Nghỉ $($tenThu[[int]$d.DayOfWeek]) ngày $($d.ToString('dd/MM/yyyy'))" } } catch {} }
    foreach ($x in @($tam | Sort-Object D)) { [void]$script:mucNghi.Add($x); [void]$lstNghi.Items.Add($x.Chu) }
}
function Nap-CauHinh {
    if (Doc-CauHinh) {
        CapNhat-DanhSachNghi; CapNhat-Ngay
        $btnTao.Enabled = $true
        return $true
    }
    $lstNghi.Items.Clear(); $lblNgay.Text = ''
    TrangThai 'File cấu hình đang bị sai cú pháp. Bấm "Tạo giáo án" để xem sai ở dòng nào.' $mauLoi
    return $false
}

# ---------- Ghi ngày nghỉ vào cấu hình (chỉ thay đúng 2 dòng, giữ nguyên phần còn lại) ----------
function Luu-NgayNghi {
    $t = [IO.File]::ReadAllText($cfgPath, [Text.Encoding]::UTF8)
    foreach ($k in 'ngayNghi', 'tuanNghi') {
        $arr = @($script:ds[$k] | Sort-Object -Unique)
        $val = if ($arr.Count -gt 0) { '["' + ($arr -join '", "') + '"]' } else { '[]' }
        $re = New-Object Text.RegularExpressions.Regex ('"' + $k + '"\s*:\s*\[[^\]]*\]')
        if (-not $re.IsMatch($t)) {
            [void][Windows.Forms.MessageBox]::Show("Không tìm thấy mục `"$k`" trong file cấu hình nên chưa lưu được ngày nghỉ.`nHãy thêm dòng  `"$k`": [],  vào cauhinh.json.", 'Chưa lưu được', 'OK', 'Warning')
            return $false
        }
        $t = $re.Replace($t, ('"' + $k + '": ' + $val), 1)
    }
    [IO.File]::WriteAllText($cfgPath, $t, $utf8)
    return $true
}
function Them-Nghi([string]$loai, [datetime]$d) {
    if ($loai -eq 'tuanNghi') { $d = ThuHai-CuaNgay $d }
    $s = $d.ToString('yyyy-MM-dd')
    if ($script:ds[$loai] -contains $s) { return }
    [void]$script:ds[$loai].Add($s)
    if (-not (Luu-NgayNghi)) { $script:ds[$loai].Remove($s) }
    [void](Nap-CauHinh)
}
function Xoa-Nghi([int]$i) {
    if ($i -lt 0 -or $i -ge $script:mucNghi.Count) { return }
    $x = $script:mucNghi[$i]
    $script:ds[$x.Loai].Remove($x.Goc)
    [void](Luu-NgayNghi); [void](Nap-CauHinh)
}

# ---------- Chạy công cụ ----------
function Dang-Chay([bool]$chay) {
    $btnTao.Enabled = -not $chay; $btnDung.Enabled = $chay
    $numTuan.Enabled = -not $chay; $gbNghi.Enabled = -not $chay; $btnCauHinh.Enabled = -not $chay
    if ($chay) { $bar.Style = 'Marquee'; $bar.MarqueeAnimationSpeed = 30 } else { $bar.Style = 'Blocks'; $bar.MarqueeAnimationSpeed = 0 }
}
function Ghi-Log([string]$dong) { $txtLog.AppendText($dong + "`r`n") }

function Bat-Dau {
    if ($script:proc) { return }
    $n = [int]$numTuan.Value
    [void](Doc-CauHinh)
    # Hỏi trước khi ghi đè file đã có (có thể cô đã sửa tay trong đó)
    if (-not $ThuMucKetQua -and $script:cfg -and $script:cfg.tenFileKetQua) {
        $cu = Join-Path $goc ($script:cfg.tenFileKetQua -f $n)
        if (Test-Path -LiteralPath $cu) {
            $tl = [Windows.Forms.MessageBox]::Show("File `"$([IO.Path]::GetFileName($cu))`" đã có.`n`nTạo lại sẽ GHI ĐÈ lên file này. Những chỗ cô đã sửa tay trong file đó sẽ mất.`n`nCô có muốn tạo lại không?", 'File đã có', 'YesNo', 'Warning', 'Button2')
            if ($tl -ne 'Yes') { return }
        }
    }
    $txtLog.Clear(); $txtLuuY.Clear(); $txtLuuY.ForeColor = $mauThuong
    $script:fileKetQua = ''; $script:daDung = $false; $btnMoFile.Enabled = $false
    $script:wordTruoc = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

    $q = { param($s) "'" + ($s -replace "'", "''") + "'" }
    $lenh = "[Console]::OutputEncoding = [Text.Encoding]::UTF8; & $(& $q $tool) -Tuan $n -KhongMoLuuY"
    if ($CauHinh)      { $lenh += " -CauHinh $(& $q $CauHinh)" }
    if ($ThuMucKetQua) { $lenh += " -ThuMucKetQua $(& $q $ThuMucKetQua)" }
    $lenh += '; exit $LASTEXITCODE'
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($lenh))
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    # Cửa sổ tự giữ cả 3 luồng vào/ra của công cụ, để công cụ không bị kẹt vì luồng thừa hưởng từ nơi khác
    $psi.RedirectStandardInput = $true; $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true; $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
    try { $script:proc = [Diagnostics.Process]::Start($psi) }
    catch { TrangThai "Không khởi động được công cụ: $($_.Exception.Message)" $mauLoi; $script:proc = $null; return }
    try { $script:proc.StandardInput.Close() } catch {}
    $script:loiTask = $script:proc.StandardError.ReadToEndAsync()
    $script:docTask = $script:proc.StandardOutput.ReadLineAsync()
    Dang-Chay $true
    TrangThai "Đang tạo giáo án tuần $n... (khoảng 1 phút, cô đừng mở file Word trong lúc này)"
    $timer.Start()
}

function Xu-Ly-Dong([string]$dong) {
    $dong = $dong.TrimEnd()
    if (-not $dong) { return }
    Ghi-Log $dong
    if ($dong -match '^\d\.\s+(.+?)\.*$') { TrangThai "Đang làm: $($Matches[1])..." }
    elseif ($dong -match '^\s+\+\s+(\S+)\s+(.+?)\s+PPCT\s+(.+)$') { TrangThai "Đang dán bài: thứ $($Matches[1]), $($Matches[2].Trim()), tiết $($Matches[3])" }
    elseif ($dong -match 'File đã tạo:\s*(.+)$') { $script:fileKetQua = $Matches[1].Trim() }
}

function Ket-Thuc([int]$ma) {
    $timer.Stop()
    $n = [int]$numTuan.Value
    try { $script:proc.Dispose() } catch {}
    $script:proc = $null; $script:docTask = $null
    Dang-Chay $false
    $ketQua = 'loi'
    if ($script:daDung) {
        TrangThai 'Đã dừng. Chưa tạo giáo án.' $mauVang; $ketQua = 'dung'
    } elseif ($ma -eq 0 -and $script:fileKetQua -and (Test-Path -LiteralPath $script:fileKetQua)) {
        $btnMoFile.Enabled = $true
        $fileLuuY = [IO.Path]::Combine([IO.Path]::GetDirectoryName($script:fileKetQua), "TUẦN $n - LƯU Ý.txt")
        $luuY = @()
        if (Test-Path -LiteralPath $fileLuuY) { $luuY = @([IO.File]::ReadAllLines($fileLuuY, [Text.Encoding]::UTF8) | Where-Object { $_ -like '- *' }) }
        if ($luuY.Count -gt 0) {
            $txtLuuY.Text = ($luuY -join "`r`n`r`n")
            TrangThai "Đã tạo xong tuần $n. Có $($luuY.Count) lưu ý cần cô kiểm tra ở khung bên dưới." $mauVang
            $ketQua = "xong, $($luuY.Count) luu y"
        } else {
            $txtLuuY.ForeColor = $mauTot; $txtLuuY.Text = 'Không có lưu ý nào. Cô mở giáo án để xem và sửa những phần lặt vặt như thường lệ.'
            TrangThai "Đã tạo xong giáo án tuần $n." $mauTot
            $ketQua = 'xong, 0 luu y'
        }
    } else {
        # Lỗi: đưa phần giải thích trong quá trình chạy lên khung lưu ý cho dễ đọc
        $dongLoi = @($txtLog.Lines | Where-Object { $_.Trim() })
        $tu = -1
        for ($i = 0; $i -lt $dongLoi.Count; $i++) { if ($dongLoi[$i] -match 'KHÔNG CHẠY ĐƯỢC|CÓ LỖI') { $tu = $i; break } }
        $txtLuuY.ForeColor = $mauLoi
        $txtLuuY.Text = if ($tu -ge 0) { (@($dongLoi[$tu..($dongLoi.Count - 1)] | ForEach-Object { $_.Trim() }) -join "`r`n") } else { 'Công cụ dừng giữa chừng mà không báo lý do. Xem khung "Quá trình chạy" bên dưới.' }
        TrangThai 'Chưa tạo được giáo án. Xem lý do ở khung bên dưới.' $mauLoi
    }
    if ($ChayThu) {
        if ($AnhChup) { try { $form.Refresh(); $bmp = New-Object Drawing.Bitmap($form.Width, $form.Height); $form.DrawToBitmap($bmp, (New-Object Drawing.Rectangle(0, 0, $form.Width, $form.Height))); $bmp.Save($AnhChup); $bmp.Dispose() } catch {} }
        if ($FileKetQuaThu) { [IO.File]::WriteAllLines($FileKetQuaThu, [string[]]@("ketqua=$ketQua", "ma=$ma", "file=$($script:fileKetQua)", "trangthai=$($lblTrangThai.Text)", "ngay=$($lblNgay.Text)", "nghi=$(@($lstNghi.Items) -join ' | ')", "luuy=$($txtLuuY.Text -replace '\s+', ' ')"), (New-Object Text.UTF8Encoding $true)) }
        $form.Close()
    }
}

function Dung-Lai {
    if (-not $script:proc) { return }
    $script:daDung = $true
    try { if (-not $script:proc.HasExited) { $script:proc.Kill() } } catch {}
    # Tắt Word chạy ngầm do lần chạy này mở ra (chỉ tiến trình mới, không có cửa sổ)
    Start-Sleep -Milliseconds 400
    Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $script:wordTruoc -notcontains $_.Id -and -not $_.MainWindowTitle } | ForEach-Object { try { $_.Kill() } catch {} }
}

$timer.Add_Tick({
    if (-not $script:proc) { $timer.Stop(); return }
    $dem = 0
    while ($script:docTask -and $script:docTask.IsCompleted -and $dem -lt 200) {
        $dong = $null
        try { $dong = $script:docTask.Result } catch {}
        if ($null -eq $dong) { $script:docTask = $null; break }
        Xu-Ly-Dong $dong
        $script:docTask = $script:proc.StandardOutput.ReadLineAsync(); $dem++
    }
    if (-not $script:docTask -and $script:proc.HasExited) { Ket-Thuc $script:proc.ExitCode }
})

# ---------- Thiết lập lớp: ngày bắt đầu năm học + thời khoá biểu ----------
$dsThu = 'Hai', 'Ba', 'Tư', 'Năm', 'Sáu', 'Bảy'
$dsBuoi = 'Sáng', 'Chiều'
function Chuan([string]$s) { return "$s".Normalize([Text.NormalizationForm]::FormC) }

# Thay đúng khối "thoiKhoaBieu": [ ... ] và dòng "ngayThuHaiTuan1" trong file cấu hình, giữ nguyên phần còn lại
function Luu-ThietLap([string[]]$dongTkb, [datetime]$ngayDau) {
    $t = [IO.File]::ReadAllText($cfgPath, [Text.Encoding]::UTF8)
    $nl = if ($t.Contains("`r`n")) { "`r`n" } else { "`n" }
    $m = [regex]::Match($t, '"thoiKhoaBieu"\s*:\s*\[')
    if (-not $m.Success) { return 'Không tìm thấy mục "thoiKhoaBieu" trong file cấu hình.' }
    $mo = $m.Index + $m.Length - 1; $sau = 0; $dong = -1
    for ($i = $mo; $i -lt $t.Length; $i++) {
        if ($t[$i] -eq '[') { $sau++ } elseif ($t[$i] -eq ']') { $sau--; if ($sau -eq 0) { $dong = $i; break } }
    }
    if ($dong -lt 0) { return 'Mục "thoiKhoaBieu" trong file cấu hình thiếu dấu ].' }
    $khoi = '"thoiKhoaBieu": [' + $nl + ($dongTkb -join (',' + $nl)) + $nl + '  ]'
    $t = $t.Substring(0, $m.Index) + $khoi + $t.Substring($dong + 1)
    $re = New-Object Text.RegularExpressions.Regex '"ngayThuHaiTuan1"\s*:\s*"[^"]*"'
    if (-not $re.IsMatch($t)) { return 'Không tìm thấy mục "ngayThuHaiTuan1" trong file cấu hình.' }
    $t = $re.Replace($t, ('"ngayThuHaiTuan1": "' + $ngayDau.ToString('yyyy-MM-dd') + '"'), 1)
    try { [void]($t | ConvertFrom-Json) } catch { return 'Nội dung sau khi sửa không hợp lệ nên chưa lưu.' }
    [IO.File]::WriteAllText($cfgPath, $t, $utf8)
    return ''
}

function Mo-ThietLap([string]$tuThu = '') {
    if (-not (Doc-CauHinh)) {
        [void][Windows.Forms.MessageBox]::Show('File cấu hình đang bị sai cú pháp nên chưa mở được phần thiết lập. Bấm "Tạo giáo án" để xem sai ở dòng nào.', 'Cấu hình sai', 'OK', 'Warning')
        return
    }
    $dlg = New-Object Windows.Forms.Form
    $dlg.Text = 'Thiết lập lớp'; $dlg.Font = $fontThuong; $dlg.ClientSize = New-Object Drawing.Size(740, 612)
    $dlg.StartPosition = 'CenterParent'; $dlg.FormBorderStyle = 'FixedDialog'; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false

    [void](Nhan 'Thứ Hai của tuần 1 (tuần đầu năm học):' $dlg 14 14 290)
    $dt1 = New-Object Windows.Forms.DateTimePicker; $dt1.Format = 'Custom'; $dt1.CustomFormat = 'dd/MM/yyyy'
    try { $dt1.Value = Ngay $script:cfg.ngayThuHaiTuan1 } catch {}
    [void](Them $dt1 $dlg 308 13 150 28)
    $l = Nhan 'Thời khoá biểu của lớp: bấm vào ô rồi chọn môn. Ô nào không có tiết thì để trống.' $dlg 14 50 712; $l.Font = $fontDam

    $maTheoTen = [ordered]@{}; $tenTheoMa = @{}
    foreach ($p in $script:cfg.monHoc.PSObject.Properties) { $ten = Chuan $p.Value.tenTKB; $maTheoTen[$ten] = $p.Name; $tenTheoMa[$p.Name] = $ten }
    $cacTen = @('') + @($maTheoTen.Keys)

    $grid = New-Object Windows.Forms.DataGridView
    $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.AllowUserToResizeRows = $false
    $grid.RowHeadersVisible = $false; $grid.SelectionMode = 'CellSelect'; $grid.MultiSelect = $false
    $grid.EditMode = 'EditOnEnter'; $grid.AutoSizeColumnsMode = 'Fill'; $grid.BackgroundColor = [Drawing.Color]::White
    $grid.ColumnHeadersHeightSizeMode = 'DisableResizing'; $grid.ColumnHeadersHeight = 30; $grid.RowTemplate.Height = 30
    [void](Them $grid $dlg 14 80 712 394)
    foreach ($h in 'Thứ', 'Buổi') {
        $c = New-Object Windows.Forms.DataGridViewTextBoxColumn; $c.HeaderText = $h; $c.ReadOnly = $true; $c.FillWeight = 55; $c.SortMode = 'NotSortable'
        $c.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(240, 240, 240); $c.DefaultCellStyle.Font = $fontDam
        [void]$grid.Columns.Add($c)
    }
    for ($k = 1; $k -le 5; $k++) {
        $c = New-Object Windows.Forms.DataGridViewComboBoxColumn; $c.HeaderText = "Tiết $k"; $c.FlatStyle = 'Flat'; $c.SortMode = 'NotSortable'
        $c.DisplayStyle = 'ComboBox'; [void]$c.Items.AddRange([object[]]$cacTen)
        [void]$grid.Columns.Add($c)
    }
    foreach ($thu in $dsThu) {
        foreach ($buoi in $dsBuoi) {
            $o = @($thu, $buoi, '', '', '', '', '')
            $ses = @($script:cfg.thoiKhoaBieu | Where-Object { (Chuan $_.thu) -eq $thu -and (Chuan $_.buoi) -eq $buoi }) | Select-Object -First 1
            if ($ses) { $k = 2; foreach ($ma in @($ses.mon)) { if ($k -le 6 -and $tenTheoMa[$ma]) { $o[$k] = $tenTheoMa[$ma]; $k++ } } }
            [void]$grid.Rows.Add([object[]]$o)
        }
    }
    $grid.Add_DataError({ param($s, $e) $e.ThrowException = $false })

    $lblDem = Nhan '' $dlg 14 480 712 66
    $lblDem.TextAlign = 'TopLeft'
    $dem = {
        $n = @{}; $tong = 0
        foreach ($r in $grid.Rows) { for ($c = 2; $c -le 6; $c++) { $v = "$($r.Cells[$c].Value)"; if ($v) { $n[$v] = 1 + [int]$n[$v]; $tong++ } } }
        $phan = @(); $lech = $false
        foreach ($ten in $maTheoTen.Keys) {
            $can = $script:cfg.monHoc.($maTheoTen[$ten]).soTietTuan; $co = [int]$n[$ten]
            if ($can) { $phan += "$ten $co/$can"; if ($co -ne [int]$can) { $lech = $true } } elseif ($co -gt 0) { $phan += "$ten $co" }
        }
        $lblDem.Text = "Tổng $tong tiết/tuần:  " + ($phan -join ',  ') + $(if ($lech) { "`r`nCó môn chưa đủ hoặc thừa số tiết so với chương trình (số sau dấu / là số tiết mỗi tuần)." } else { '' })
        $lblDem.ForeColor = if ($lech) { $mauVang } else { $mauTot }
    }
    $grid.Add_CellValueChanged({ & $dem })
    $grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { [void]$grid.CommitEdit('Commit') } })
    & $dem

    $bKhung = Nut 'Sửa khung mẫu (tên trường, lớp, chữ ký)' $dlg 14 552 330 38
    $bFile  = Nut 'Mở file cấu hình' $dlg 352 552 150 38
    $bLuu   = Nut 'Lưu' $dlg 510 552 104 38; $bLuu.Font = $fontDam
    $bHuy   = Nut 'Đóng' $dlg 622 552 104 38
    $bKhung.Add_Click({
        $k = Join-Path $goc ("$($script:cfg.fileMauKhung)" -replace '/', '\')
        if (Test-Path -LiteralPath $k) {
            [void][Windows.Forms.MessageBox]::Show("Word sẽ mở file khung mẫu.`n`nCô sửa tên trường, tên lớp ở đầu trang và tên người ký ở bảng chữ ký cuối trang, rồi bấm Lưu (Ctrl+S) và đóng Word.`n`nKhông cần sửa bảng thời khoá biểu trong đó: công cụ tự dựng lại theo lịch cô chọn.", 'Sửa khung mẫu', 'OK', 'Information')
            Start-Process -FilePath $k
        } else { [void][Windows.Forms.MessageBox]::Show("Không thấy file khung mẫu: $k", 'Thiếu file', 'OK', 'Warning') }
    })
    $bFile.Add_Click({ Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$cfgPath`"" })
    $bHuy.Add_Click({ $dlg.Close() })
    $luu = {
        [void]$grid.EndEdit()
        $dong = @()
        foreach ($r in $grid.Rows) {
            $ma = @(); for ($c = 2; $c -le 6; $c++) { $v = "$($r.Cells[$c].Value)"; if ($v) { $ma += '"' + $maTheoTen[$v] + '"' } }
            if ($ma.Count -gt 0) { $dong += '    { "thu": "' + $r.Cells[0].Value + '", "buoi": "' + $r.Cells[1].Value + '", "mon": [' + ($ma -join ', ') + '] }' }
        }
        if ($dong.Count -eq 0) { [void][Windows.Forms.MessageBox]::Show('Thời khoá biểu đang trống. Hãy chọn môn cho ít nhất một tiết.', 'Chưa lưu', 'OK', 'Warning'); return }
        $loi = Luu-ThietLap $dong $dt1.Value.Date
        if ($loi) { [void][Windows.Forms.MessageBox]::Show($loi, 'Chưa lưu được', 'OK', 'Warning'); return }
        $dlg.Close()
    }
    $bLuu.Add_Click($luu)
    if ($tuThu) {
        # Tự kiểm tra: "Tư|Chiều|TCTV,GDTC;Ba|Chiều|TOAN,TCTV" -> đặt lại các dòng đó rồi bấm Lưu
        $dlg.Add_Shown({
            foreach ($muc in ($tuThu -split ';')) {
                $p = $muc -split '\|'
                foreach ($r in $grid.Rows) {
                    if ($r.Cells[0].Value -eq $p[0] -and $r.Cells[1].Value -eq $p[1]) {
                        $ds = @($p[2] -split ',' | Where-Object { $_ })
                        for ($c = 2; $c -le 6; $c++) { $r.Cells[$c].Value = $(if ($c - 2 -lt $ds.Count) { $tenTheoMa[$ds[$c - 2]] } else { '' }) }
                    }
                }
            }
            if ($AnhChup) { try { $dlg.Refresh(); $bmp = New-Object Drawing.Bitmap($dlg.Width, $dlg.Height); $dlg.DrawToBitmap($bmp, (New-Object Drawing.Rectangle(0, 0, $dlg.Width, $dlg.Height))); $bmp.Save(($AnhChup -replace '\.png$', '-thietlap.png')); $bmp.Dispose() } catch {} }
            & $luu
        })
    }
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
    [void](Nap-CauHinh)
}

# ---------- Sự kiện ----------
$numTuan.Add_ValueChanged({ CapNhat-Ngay })
$btnThemNgay.Add_Click({ if ($script:cfg) { Them-Nghi 'ngayNghi' $dtNghi.Value.Date } })
$btnThemTuan.Add_Click({ if ($script:cfg) { Them-Nghi 'tuanNghi' $dtNghi.Value.Date } })
$btnXoaNghi.Add_Click({ Xoa-Nghi $lstNghi.SelectedIndex })
$btnTao.Add_Click({ Bat-Dau })
$btnDung.Add_Click({ Dung-Lai })
$btnMoFile.Add_Click({ if ($script:fileKetQua -and (Test-Path -LiteralPath $script:fileKetQua)) { Start-Process -FilePath $script:fileKetQua } })
$btnMoThuMuc.Add_Click({ Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$goc`"" })
$btnCauHinh.Add_Click({ Mo-ThietLap })
$btnHuongDan.Add_Click({ Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$(Join-Path $here 'HUONG_DAN.md')`"" })
$form.Add_Activated({ if (-not $script:proc) { [void](Doc-CauHinh); if ($script:cfg) { CapNhat-DanhSachNghi; CapNhat-Ngay } } })
$form.Add_FormClosing({
    param($s, $e)
    if ($script:proc -and -not $ChayThu) {
        $tl = [Windows.Forms.MessageBox]::Show('Công cụ đang tạo giáo án. Dừng lại và đóng cửa sổ?', 'Đang chạy', 'YesNo', 'Question', 'Button2')
        if ($tl -ne 'Yes') { $e.Cancel = $true; return }
        Dung-Lai
    }
    $timer.Stop()
})

# ---------- Khởi động ----------
if (Nap-CauHinh) {
    # Gợi ý tuần: tuần học chứa ngày "hôm nay + 3" (từ thứ Sáu trở đi gợi ý tuần sau)
    $goiY = 1
    try { for ($k = 1; $k -le 45; $k++) { if ((ThuHai-CuaTuan $k) -le (Get-Date).Date.AddDays(3)) { $goiY = $k } else { break } } } catch {}
    $numTuan.Value = [Math]::Max(1, [Math]::Min(45, $goiY))
    CapNhat-Ngay
}
if ($ChayThu) {
    $form.Add_Shown({
        foreach ($x in @(($NghiThu -join ',') -split ',' | Where-Object { $_ })) { $p = $x -split ':'; Them-Nghi $p[0] (Ngay $p[1]) }
        if ($XoaThu -ge 0) { Xoa-Nghi $XoaThu }
        if ($TkbThu) { Mo-ThietLap $TkbThu }
        $numTuan.Value = $TuanThu; [Windows.Forms.Application]::DoEvents(); Bat-Dau
    })
}
[void]$form.ShowDialog()
