# =====================================================================
#  KIỂM TRA CÔNG CỤ SAU KHI SỬA
#  Chạy công cụ cho các tuần mẫu, lấy toàn bộ chữ của file kết quả rồi
#  so với "bản mốc" đã lưu trong thư mục KiemTra (cạnh thư mục CongCu).
#    kiem_tra.ps1                 -> so với bản mốc, liệt kê chỗ khác
#    kiem_tra.ps1 -LuuMoc         -> lưu kết quả hiện tại làm bản mốc mới
#    kiem_tra.ps1 -Tuan 10,11,12  -> chọn tuần mẫu
# =====================================================================
param(
    [int[]]$Tuan = @(10, 11),
    [switch]$LuuMoc,
    [string]$CauHinh = ""
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$goc  = Split-Path -Parent $here
$dir  = Join-Path $goc 'KiemTra'
$tam  = Join-Path $dir 'tam'
[void](New-Item -ItemType Directory -Force -Path $tam)
$utf8 = New-Object Text.UTF8Encoding $true

function Lay-Chu([string]$docx) {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false; $w.DisplayAlerts = 0
    try {
        $d = $w.Documents.Open($docx, $false, $true)
        $t = $d.Content.Text
        $d.Close(0)
    } finally {
        try { $w.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
    }
    # Mỗi đoạn / ô bảng một dòng, bỏ khoảng trắng thừa
    $lines = $t.Normalize([Text.NormalizationForm]::FormC) -split "[\r\n\a\v\f]+" |
        ForEach-Object { ($_ -replace '\s+', ' ').Trim() } | Where-Object { $_ }
    return $lines
}

$tongKhac = 0
foreach ($n in $Tuan) {
    Write-Host "`n===== TUẦN $n =====" -ForegroundColor Cyan
    Get-ChildItem -LiteralPath $tam -File | Remove-Item -Force
    $thamSo = @{ Tuan = $n; ThuMucKetQua = $tam; KhongMoLuuY = $true }
    if ($CauHinh) { $thamSo.CauHinh = $CauHinh }
    try { & (Join-Path $here 'tao_giao_an.ps1') @thamSo | Out-Null }
    catch { Write-Host "  LỖI khi tạo giáo án tuần ${n}: $($_.Exception.Message)" -ForegroundColor Red; $tongKhac++; continue }

    $docx = Get-ChildItem -LiteralPath $tam -Filter '*.docx' | Select-Object -First 1
    if (-not $docx) { Write-Host "  LỖI: không thấy file kết quả" -ForegroundColor Red; $tongKhac++; continue }
    $moi = @(Lay-Chu $docx.FullName)
    $luuY = Get-ChildItem -LiteralPath $tam -Filter '*.txt' | Select-Object -First 1
    if ($luuY) { $moi += '##### LƯU Ý #####'; $moi += @(Get-Content -LiteralPath $luuY.FullName -Encoding UTF8 | Where-Object { $_ -like '- *' }) }

    $fileMoc = Join-Path $dir "TUAN $n - moc.txt"
    $fileMoi = Join-Path $dir "TUAN $n - moi.txt"
    if ($LuuMoc -or -not (Test-Path -LiteralPath $fileMoc)) {
        [IO.File]::WriteAllLines($fileMoc, [string[]]$moi, $utf8)
        Write-Host "  Đã lưu bản mốc ($($moi.Count) dòng): $fileMoc" -ForegroundColor Green
        continue
    }
    [IO.File]::WriteAllLines($fileMoi, [string[]]$moi, $utf8)
    $cu = @(Get-Content -LiteralPath $fileMoc -Encoding UTF8)
    $khac = @(Compare-Object -ReferenceObject $cu -DifferenceObject $moi -CaseSensitive)
    if ($khac.Count -eq 0) {
        Write-Host "  GIỐNG bản mốc ($($moi.Count) dòng)." -ForegroundColor Green
        Remove-Item -LiteralPath $fileMoi -Force
    } else {
        $tongKhac += $khac.Count
        Write-Host "  KHÁC bản mốc $($khac.Count) dòng  (- = chỉ có ở bản mốc, + = chỉ có ở bản mới):" -ForegroundColor Yellow
        $khac | Select-Object -First 60 | ForEach-Object {
            $dau = if ($_.SideIndicator -eq '<=') { '-' } else { '+' }
            $s = "$($_.InputObject)"; if ($s.Length -gt 150) { $s = $s.Substring(0, 150) + '…' }
            Write-Host "    $dau $s"
        }
        if ($khac.Count -gt 60) { Write-Host "    … còn $($khac.Count - 60) dòng nữa, xem file: $fileMoi" }
    }
}
Remove-Item -LiteralPath $tam -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ''
if ($LuuMoc) { Write-Host 'Đã lưu bản mốc.' -ForegroundColor Green }
elseif ($tongKhac -eq 0) { Write-Host 'KẾT QUẢ: không có gì thay đổi so với bản mốc.' -ForegroundColor Green }
else { Write-Host "KẾT QUẢ: có $tongKhac chỗ khác. Nếu đúng là thay đổi cô muốn, chạy lại với -LuuMoc để lấy làm bản mốc mới." -ForegroundColor Yellow }
exit $(if ($tongKhac -eq 0) { 0 } else { 1 })
