param([Parameter(ValueFromRemainingArguments=$true)][string[]]$Files)

# ============================================================
#  xlsx -> CSV 変換ツール（先頭2行を削除）
#  Windows標準のPowerShell + Excel(COM) で動作。追加インストール不要。
# ============================================================

$ErrorActionPreference = 'Stop'

# 先頭から削除する行数（ここを変えれば削除行数を調整できます）
$DeleteRows = 2

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# --- 変換対象の収集 -----------------------------------------
# 引数（ドラッグ&ドロップ）があればそれを、無ければ同フォルダの全xlsxを対象にする
if ($Files -and $Files.Count -gt 0) {
    $targets = @($Files)
} else {
    $targets = @(Get-ChildItem -LiteralPath $scriptDir -Filter *.xlsx -File |
                 Where-Object { $_.Name -notlike '~$*' } |
                 Select-Object -ExpandProperty FullName)
}

# xlsx / xlsm のみに絞り込み
$targets = @($targets | Where-Object { $_ -match '\.(xlsx|xlsm)$' })

if ($targets.Count -eq 0) {
    Write-Host "変換するxlsxファイルが見つかりませんでした。" -ForegroundColor Yellow
    Write-Host "このバッチと同じフォルダにxlsxを置くか、xlsxをバッチにドラッグ&ドロップしてください。"
    exit 1
}

# --- Excel起動 ----------------------------------------------
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Host "Microsoft Excel が見つからないため実行できません。" -ForegroundColor Red
    Write-Host "ExcelがインストールされたPCで実行してください。"
    exit 1
}
$excel.Visible = $false
$excel.DisplayAlerts = $false

$ok = 0
$ng = 0

foreach ($f in $targets) {
    $wb = $null
    try {
        $path = (Resolve-Path -LiteralPath $f).Path
        Write-Host ("変換中: {0}" -f $path)

        $wb = $excel.Workbooks.Open($path)
        $ws = $wb.Worksheets.Item(1)   # 1番目のシートをCSV化

        # 先頭 $DeleteRows 行を削除
        [void]$ws.Rows("1:$DeleteRows").Delete()

        $csv = [System.IO.Path]::ChangeExtension($path, ".csv")
        try {
            # 62 = xlCSVUTF8 （UTF-8。Excel2016以降）
            $wb.SaveAs($csv, 62)
        } catch {
            # 古いExcel向けフォールバック 6 = xlCSV （環境の文字コード）
            $wb.SaveAs($csv, 6)
        }
        $wb.Close($false)
        $wb = $null
        Write-Host ("  -> 完了: {0}" -f $csv) -ForegroundColor Green
        $ok++
    } catch {
        Write-Host ("  -> 失敗: {0}" -f $_.Exception.Message) -ForegroundColor Red
        if ($wb) { try { $wb.Close($false) } catch {} }
        $ng++
    }
}

# --- 後片付け -----------------------------------------------
$excel.Quit()
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host ("完了 {0}件 / 失敗 {1}件" -f $ok, $ng)
