<#
.SYNOPSIS
  Excel で開いているブック（.xlam アドインを含む）の VBA を、テキストファイルとの間で書き出し・読み込みする。

.DESCRIPTION
  既定の配置:
    アドイン本体 : %APPDATA%\Microsoft\AddIns\KoreDeFilter.xlam（Excel が読み込んでいるもの）
    ソース       : <リポジトリ>\src\KoreDeFilter.xlam\
    配布用コピー : <リポジトリ>\bin\KoreDeFilter.xlam（Export 時と -Save 時に本体をコピー）

.EXAMPLE
  .\vba-sync.ps1 -Action Export
  .\vba-sync.ps1 -Action Import
  .\vba-sync.ps1 -Action Import -Save

.NOTES
  - Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしておくこと。
  - ファイルは UTF-8 (BOM 付き)・CRLF で保存する。
  - Import は既存モジュールのコード部分だけを置き換える。フォームのデザインや属性行の変更は反映しない。
  - 新しい .bas / .cls をフォルダに置くと、新規モジュールとして追加する（.frm の新規追加は不可）。
#>
param(
    [Parameter(Mandatory)][ValidateSet('Export', 'Import')][string]$Action,
    [string]$Book = 'KoreDeFilter.xlam',
    [string]$BookDir,
    [string]$SrcDir,
    [string]$BinDir,
    [switch]$Save
)

$ErrorActionPreference = 'Stop'
$sjis = [Text.Encoding]::GetEncoding(932)
$utf8Bom = New-Object Text.UTF8Encoding $true
$repoRoot = Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent
if (-not $BookDir) { $BookDir = Join-Path $env:APPDATA 'Microsoft\AddIns' }
if (-not $SrcDir) { $SrcDir = Join-Path $repoRoot "src\$Book" }
if (-not $BinDir) { $BinDir = Join-Path $repoRoot 'bin' }

# vbext_ComponentType
$TYPE_STD = 1; $TYPE_CLASS = 2; $TYPE_FORM = 3; $TYPE_DOC = 100
$extOf = @{ $TYPE_STD = '.bas'; $TYPE_CLASS = '.cls'; $TYPE_FORM = '.frm'; $TYPE_DOC = '.bas' }

function Get-Workbook {
    try {
        $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        Write-Host '[INFO] Excel が起動していないため起動します。'
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $true
    }
    try {
        return $xl.Workbooks.Item($Book)
    } catch {
        $path = Join-Path $BookDir $Book
        if (-not (Test-Path $path)) { throw "ブックが見つかりません: $path" }
        Write-Host "[INFO] $Book が開かれていないため開きます: $path"
        return $xl.Workbooks.Open($path)
    }
}

function Get-VBProject($wb) {
    try {
        $p = $wb.VBProject
        $null = $p.VBComponents.Count
        return $p
    } catch {
        throw 'VBA プロジェクトにアクセスできません。トラストセンターの「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を確認し、パスワードがある場合は VBE で解除してください。'
    }
}

function Get-ModuleCode($comp) {
    $cm = $comp.CodeModule
    if ($cm.CountOfLines -eq 0) { return '' }
    return $cm.Lines(1, $cm.CountOfLines)
}

# .bas / .cls / .frm の先頭にある VERSION・Begin…End・Attribute VB_ 行を取り除き、コード本体だけ返す
function Get-CodeBody([string]$text) {
    $lines = $text -split "`r?`n"
    $i = 0
    if ($i -lt $lines.Count -and $lines[$i] -match '^VERSION\s') { $i++ }
    if ($i -lt $lines.Count -and $lines[$i] -match '^Begin\b') {
        $depth = 0
        do {
            if ($lines[$i] -match '^\s*Begin\b') { $depth++ }
            elseif ($lines[$i] -match '^\s*End\s*$') { $depth-- }
            $i++
        } while ($depth -gt 0 -and $i -lt $lines.Count)
    }
    while ($i -lt $lines.Count -and $lines[$i] -match '^Attribute VB_') { $i++ }
    if ($i -ge $lines.Count) { return '' }
    return ($lines[$i..($lines.Count - 1)] -join "`r`n")
}

function Find-Component($vbp, [string]$name) {
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $name) { return $c } }
    return $null
}

# アドイン本体（保存済みのファイル）を bin へコピーする
function Copy-ToBin($wb) {
    if (-not $BinDir) { return }
    New-Item -ItemType Directory -Force $BinDir | Out-Null
    $dest = Join-Path $BinDir $wb.Name
    if ([IO.Path]::GetFullPath($wb.FullName) -eq [IO.Path]::GetFullPath($dest)) { return }
    if (-not $wb.Saved) { Write-Host '[WARN] 未保存の変更があります。bin には保存済みの状態がコピーされます。' }
    Copy-Item $wb.FullName $dest -Force
    Write-Host "[COPY] $dest"
}

function Invoke-Export($vbp) {
    New-Item -ItemType Directory -Force $SrcDir | Out-Null
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("vba-sync-" + [Guid]::NewGuid())
    New-Item -ItemType Directory $tmp | Out-Null
    $exported = @()
    try {
        foreach ($comp in $vbp.VBComponents) {
            $ext = $extOf[[int]$comp.Type]
            if (-not $ext) { Write-Host "[SKIP] 未対応の種類: $($comp.Name)"; continue }
            $dest = Join-Path $SrcDir ($comp.Name + $ext)
            if ($comp.Type -eq $TYPE_DOC) {
                # シート / ThisWorkbook はコード本体だけを保存
                $code = Get-ModuleCode $comp
                [IO.File]::WriteAllText($dest, $code + "`r`n", $utf8Bom)
            } else {
                $tmpFile = Join-Path $tmp ($comp.Name + $ext)
                $comp.Export($tmpFile)
                [IO.File]::WriteAllText($dest, [IO.File]::ReadAllText($tmpFile, $sjis), $utf8Bom)
                if ($comp.Type -eq $TYPE_FORM) {
                    $frx = [IO.Path]::ChangeExtension($tmpFile, '.frx')
                    if (Test-Path $frx) { Copy-Item $frx (Join-Path $SrcDir ($comp.Name + '.frx')) -Force }
                }
            }
            $exported += $comp.Name + $ext
            Write-Host "[EXPORT] $($comp.Name)$ext"
        }
    } finally {
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
    # Excel 側に無いモジュールのファイルを知らせる（削除はしない）
    Get-ChildItem $SrcDir -File -Include *.bas, *.cls, *.frm -Recurse:$false -ErrorAction SilentlyContinue |
        Where-Object { $exported -notcontains $_.Name } |
        ForEach-Object { Write-Host "[WARN] Excel 側に無いファイル: $($_.Name)" }
    Write-Host "[DONE] $($exported.Count) 件を書き出しました -> $SrcDir"
}

function Invoke-Import($wb, $vbp) {
    if (-not (Test-Path $SrcDir)) { throw "フォルダが見つかりません: $SrcDir" }
    $files = Get-ChildItem (Join-Path $SrcDir '*') -File -Include *.bas, *.cls, *.frm
    $changed = 0
    foreach ($f in $files) {
        $name = $f.BaseName
        $text = [IO.File]::ReadAllText($f.FullName)   # BOM から UTF-8 を判定
        $comp = Find-Component $vbp $name
        if ($comp) {
            $body = Get-CodeBody $text
            $current = Get-ModuleCode $comp
            if ($body.TrimEnd() -eq $current.TrimEnd()) { continue }
            $cm = $comp.CodeModule
            if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
            if ($body.Trim()) { $cm.AddFromString($body.TrimEnd()) }
            $changed++
            Write-Host "[UPDATE] $($f.Name)"
        } else {
            # 新規モジュール
            if ($f.Extension -eq '.frm') { Write-Host "[SKIP] フォームの新規追加は未対応です（Excel 側で作成してから Export してください）: $($f.Name)"; continue }
            if ($text -notmatch '(?m)^Attribute VB_Name') {
                $header = if ($f.Extension -eq '.cls') {
                    "VERSION 1.0 CLASS`r`nBEGIN`r`n  MultiUse = -1  'True`r`nEND`r`nAttribute VB_Name = `"$name`"`r`nAttribute VB_GlobalNameSpace = False`r`nAttribute VB_Creatable = False`r`nAttribute VB_PredeclaredId = False`r`nAttribute VB_Exposed = False`r`n"
                } else {
                    "Attribute VB_Name = `"$name`"`r`n"
                }
                $text = $header + $text
            }
            $tmpFile = Join-Path ([IO.Path]::GetTempPath()) ("vba-sync-" + [Guid]::NewGuid() + $f.Extension)
            try {
                [IO.File]::WriteAllText($tmpFile, $text, $sjis)
                $null = $vbp.VBComponents.Import($tmpFile)
            } finally {
                Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
            }
            $changed++
            Write-Host "[ADD] $($f.Name)"
        }
    }
    Write-Host "[DONE] $changed 件を反映しました"

    if ($Save) {
        $backupDir = Join-Path $BookDir 'backups'
        New-Item -ItemType Directory -Force $backupDir | Out-Null
        if (Test-Path $wb.FullName) {
            $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $bak = Join-Path $backupDir ("{0}_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($wb.Name), $stamp, [IO.Path]::GetExtension($wb.Name))
            Copy-Item $wb.FullName $bak
            Write-Host "[BACKUP] $bak"
        }
        $wb.Save()
        Write-Host "[SAVE] $($wb.FullName)"
        Copy-ToBin $wb
    } elseif ($changed -gt 0) {
        Write-Host '[INFO] まだ保存していません。確認後に VBE で Ctrl+S するか、-Save を付けて実行してください。'
    }
}

$wb = Get-Workbook
$vbp = Get-VBProject $wb
Write-Host "[INFO] 対象: $($wb.FullName)"
switch ($Action) {
    'Export' { Invoke-Export $vbp; Copy-ToBin $wb }
    'Import' { Invoke-Import $wb $vbp }
}
