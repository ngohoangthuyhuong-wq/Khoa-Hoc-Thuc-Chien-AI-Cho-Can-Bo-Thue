# BO TRO GIUP - APP TONG QUAN MAU (giao dien cho BOT_TONG_QUAN.vbs)
# Khoa hoc Thuc chien AI cho can bo thue - Ngo Hoang Thuy Huong - Zalo 0328 878 297
# Mo bang CHAY.bat. Luu file nay voi encoding UTF-8 with BOM.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$dir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbs  = Join-Path $dir "BOT_TONG_QUAN.vbs"
$mstF = Join-Path $dir "MST_CHAY.txt"
$stop = Join-Path $dir "STOP_TONG_QUAN.txt"
$prog = Join-Path $dir "PROGRESS_TONG_QUAN.txt"
$script:botProc = $null

function Get-DsMst([string]$raw) {
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($x in ($raw -split "[\r\n;,\s]+")) {
        $t = $x.Trim().Trim("'")
        if ($t -match '^(\d{10}|\d{10}-\d{3}|\d{12}|\d{13})$' -and $seen.Add($t)) { [void]$list.Add($t) }
    }
    return ,$list
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "BOT TONG QUAN MAU - BO TRO GIUP"
$form.Size = New-Object System.Drawing.Size(560, 560)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AllowDrop = $true

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "Buoc 1: Dan danh sach MST (moi dong 1 MST) hoac keo tha file .txt vao day"
$lbl.Location = New-Object System.Drawing.Point(15, 12)
$lbl.Size = New-Object System.Drawing.Size(520, 22)
$form.Controls.Add($lbl)

$tb = New-Object System.Windows.Forms.TextBox
$tb.Multiline = $true
$tb.ScrollBars = "Vertical"
$tb.Location = New-Object System.Drawing.Point(15, 38)
$tb.Size = New-Object System.Drawing.Size(515, 260)
$form.Controls.Add($tb)

$btnDan = New-Object System.Windows.Forms.Button
$btnDan.Text = "Dan clipboard"
$btnDan.Location = New-Object System.Drawing.Point(15, 306)
$btnDan.Size = New-Object System.Drawing.Size(140, 32)
$form.Controls.Add($btnDan)

$lblDem = New-Object System.Windows.Forms.Label
$lblDem.Text = "0 MST hop le"
$lblDem.Location = New-Object System.Drawing.Point(170, 312)
$lblDem.Size = New-Object System.Drawing.Size(360, 22)
$form.Controls.Add($lblDem)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "LAY TONG QUAN"
$btnRun.Location = New-Object System.Drawing.Point(15, 350)
$btnRun.Size = New-Object System.Drawing.Size(250, 48)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(20, 33, 61)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "DUNG BOT"
$btnStop.Location = New-Object System.Drawing.Point(280, 350)
$btnStop.Size = New-Object System.Drawing.Size(250, 48)
$btnStop.BackColor = [System.Drawing.Color]::FromArgb(224, 122, 46)
$btnStop.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnStop)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Mo SAP, dang nhap TMS truoc khi bam LAY TONG QUAN. Khi bot chay khong dung vao SAP va Excel."
$lblStatus.Location = New-Object System.Drawing.Point(15, 410)
$lblStatus.Size = New-Object System.Drawing.Size(515, 60)
$form.Controls.Add($lblStatus)

$lblTg = New-Object System.Windows.Forms.Label
$lblTg.Text = "Bo tro giup - Khoa hoc Thuc chien AI cho can bo thue"
$lblTg.ForeColor = [System.Drawing.Color]::Gray
$lblTg.Location = New-Object System.Drawing.Point(15, 485)
$lblTg.Size = New-Object System.Drawing.Size(515, 22)
$form.Controls.Add($lblTg)

$tb.Add_TextChanged({ $lblDem.Text = "$((Get-DsMst $tb.Text).Count) MST hop le (da bo trung)" })
$btnDan.Add_Click({ $tb.Text = [System.Windows.Forms.Clipboard]::GetText() })

$form.Add_DragEnter({ param($s, $e) if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' } })
$form.Add_DragDrop({
    param($s, $e)
    $f = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)[0]
    if ($f) { $tb.Text = [IO.File]::ReadAllText($f) }
})

$btnRun.Add_Click({
    if (-not (Test-Path -LiteralPath $vbs)) {
        [System.Windows.Forms.MessageBox]::Show("Thieu file BOT_TONG_QUAN.vbs trong thu muc:`r`n$dir", "Thieu file") | Out-Null
        return
    }
    $ds = Get-DsMst $tb.Text
    if ($ds.Count -lt 1) {
        [System.Windows.Forms.MessageBox]::Show("Chua co MST hop le.", "Thieu MST") | Out-Null
        return
    }
    if ($script:botProc -and -not $script:botProc.HasExited) {
        [System.Windows.Forms.MessageBox]::Show("Bot dang chay roi.", "Dang chay") | Out-Null
        return
    }
    $ok = [System.Windows.Forms.MessageBox]::Show("Tra $($ds.Count) MST?`r`nDa mo SAP va dang nhap TMS thi bam Yes.", "Xac nhan", "YesNo")
    if ($ok -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    [IO.File]::WriteAllLines($mstF, $ds.ToArray(), [Text.Encoding]::ASCII)
    Remove-Item -LiteralPath $stop, $prog -Force -ErrorAction SilentlyContinue
    $ws = Join-Path $env:SystemRoot "System32\wscript.exe"
    $script:botProc = Start-Process -FilePath $ws -ArgumentList @("//nologo", "`"$vbs`"") -WorkingDirectory $dir -PassThru
    $lblStatus.Text = "Bot dang chay..."
})

$btnStop.Add_Click({
    Set-Content -LiteralPath $stop -Value "DUNG" -Encoding ASCII
    $lblStatus.Text = "Da bao dung. Bot se luu Excel va dung sau MST dang chay."
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    if ($script:botProc -and (Test-Path -LiteralPath $prog)) {
        $t = (Get-Content -LiteralPath $prog -Raw -ErrorAction SilentlyContinue)
        if ($t) { $lblStatus.Text = "Tien do: $t`r`nKet qua: KET_QUA_TONG_QUAN.xlsx | Nhat ky: LOG_TONG_QUAN.txt" }
    }
})
$timer.Start()

$form.Add_FormClosing({
    if ($script:botProc -and -not $script:botProc.HasExited) {
        Set-Content -LiteralPath $stop -Value "DUNG" -Encoding ASCII
    }
})

[void]$form.ShowDialog()
