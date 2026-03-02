#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    CSVファイルを基にADユーザーのログオン可能なコンピューターを設定する。
.DESCRIPTION
    CSV内の SamAccountName と Computers(セミコロン区切り)を読み取り、
    各ユーザーの LogonWorkstations 属性を設定する。
    PC の存在チェック、実行ログ出力、サマリ表示、失敗一覧の出力を行う。
.PARAMETER CsvPath
    入力CSVファイルのパス。ヘッダー: SamAccountName,Computers
.PARAMETER LogDir
    ログファイルの出力先ディレクトリ。既定はスクリプトと同じフォルダ。
.EXAMPLE
    .\Set-LogonWorkstations.ps1 -CsvPath .\users.csv
    .\Set-LogonWorkstations.ps1 -CsvPath .\users.csv -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [string]$LogDir = $PSScriptRoot
)

# --- ログ設定 ---
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path $LogDir "Set-LogonWorkstations_$timestamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    $line | Out-File -FilePath $logFile -Append -Encoding UTF8
    switch ($Level) {
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Host $Message -ForegroundColor Red }
        default { Write-Host $Message }
    }
}

# --- 初期化 ---
$users = Import-Csv -Path $CsvPath -Encoding UTF8
$successList = [System.Collections.Generic.List[PSCustomObject]]::new()
$failureList = [System.Collections.Generic.List[PSCustomObject]]::new()
$skipCount = 0

Write-Log "===== 処理開始 ====="
Write-Log "入力CSV: $CsvPath"
Write-Log "対象件数: $($users.Count)"

# --- AD上のコンピューター存在チェック(キャッシュ) ---
$allComputerNames = $users.Computers -split ';' |
    Where-Object { $_ -ne '' } |
    Sort-Object -Unique

$validComputers = @{}
foreach ($name in $allComputerNames) {
    try {
        Get-ADComputer -Identity $name -ErrorAction Stop | Out-Null
        $validComputers[$name] = $true
    }
    catch {
        $validComputers[$name] = $false
        Write-Log "コンピューター '$name' がAD上に見つかりません。" -Level WARN
    }
}

# --- メイン処理 ---
foreach ($entry in $users) {
    $sam = $entry.SamAccountName
    $computerList = ($entry.Computers -split ';') | Where-Object { $_ -ne '' }

    # SamAccountName 空チェック
    if (-not $sam) {
        Write-Log "SamAccountName が空の行をスキップしました。" -Level WARN
        $skipCount++
        continue
    }

    # ユーザー存在チェック
    try {
        Get-ADUser -Identity $sam -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Log "ユーザー '$sam' がAD上に見つかりません。スキップします。" -Level WARN
        foreach ($pc in $computerList) {
            $failureList.Add([PSCustomObject]@{
                SamAccountName = $sam
                Computer       = $pc
                Reason         = 'ユーザーが存在しない'
            })
        }
        continue
    }

    # PC 存在チェック
    $validPCs = @()
    foreach ($pc in $computerList) {
        if ($validComputers[$pc]) {
            $validPCs += $pc
        }
        else {
            $failureList.Add([PSCustomObject]@{
                SamAccountName = $sam
                Computer       = $pc
                Reason         = 'コンピューターが存在しない'
            })
            Write-Log "[$sam] コンピューター '$pc' が存在しないためスキップしました。" -Level WARN
        }
    }

    if ($validPCs.Count -eq 0) {
        Write-Log "[$sam] 有効なコンピューターがないためスキップしました。" -Level WARN
        continue
    }

    $computers = $validPCs -join ','

    # LogonWorkstations 設定
    if ($PSCmdlet.ShouldProcess($sam, "LogonWorkstations を '$computers' に設定")) {
        try {
            Set-ADUser -Identity $sam -LogonWorkstations $computers -ErrorAction Stop
            Write-Log "[$sam] LogonWorkstations を設定しました: $computers"
            foreach ($pc in $validPCs) {
                $successList.Add([PSCustomObject]@{
                    SamAccountName = $sam
                    Computer       = $pc
                })
            }
        }
        catch {
            Write-Log "[$sam] LogonWorkstations の設定に失敗しました: $_" -Level ERROR
            foreach ($pc in $validPCs) {
                $failureList.Add([PSCustomObject]@{
                    SamAccountName = $sam
                    Computer       = $pc
                    Reason         = "Set-ADUser 失敗: $_"
                })
            }
        }
    }
}

# --- サマリ出力 ---
Write-Log ''
Write-Log '========== 実行結果サマリ =========='
Write-Log "処理対象ユーザー数 : $($users.Count)"
Write-Log "成功 (ユーザー/PC) : $($successList.Count)"
Write-Log "失敗 (ユーザー/PC) : $($failureList.Count)"
Write-Log "スキップ (空行)    : $skipCount"
Write-Log '====================================='

# --- 失敗一覧の出力 ---
if ($failureList.Count -gt 0) {
    $failureCsv = Join-Path $LogDir "Set-LogonWorkstations_Failures_$timestamp.csv"
    $failureList | Export-Csv -Path $failureCsv -NoTypeInformation -Encoding UTF8
    Write-Log "失敗一覧を出力しました: $failureCsv" -Level WARN

    Write-Log ''
    Write-Log '--- 失敗一覧 ---'
    foreach ($f in $failureList) {
        Write-Log "  $($f.SamAccountName) | $($f.Computer) | $($f.Reason)" -Level WARN
    }
}

Write-Log "ログファイル: $logFile"
Write-Log "===== 処理完了 ====="
