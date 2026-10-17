<#
.SYNOPSIS
  KoreDeFilter の動作テスト。画面に表示しない別の Excel を起動し、テスト用のブックでフィルタを実行して結果を確かめる。

.DESCRIPTION
  - 既定では bin\KoreDeFilter.xlam を読み取り専用で開き、VBA を src\KoreDeFilter.xlam\ のコードに差し替えてテストする。
    （AddIns のアドインや、いま開いている Excel には触れない。何も保存しない）
  - MsgBox / WSH.Popup / メモ帳への出力はログに置き換えるので、テスト中に画面には何も出ない。
  - すべて PASS なら終了コード 0、FAIL があれば 1。

.EXAMPLE
  .\tests\run-tests.ps1
  .\tests\run-tests.ps1 -Cases trim,blank-line
  .\tests\run-tests.ps1 -Original        # src に差し替えず、xlam のコードのままテストする（修正前との比較用）
#>
param(
    [string]$SrcDir,
    [string]$Xlam,
    [string[]]$Cases,
    [switch]$Original
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent
if (-not $SrcDir) { $SrcDir = Join-Path $repoRoot 'src\KoreDeFilter.xlam' }
if (-not $Xlam) { $Xlam = Join-Path $repoRoot 'bin\KoreDeFilter.xlam' }
if ($Cases) { $Cases = $Cases -split ',' }

function Get-CodeBody([string]$text) {
    $lines = $text -split "`r?`n"; $i = 0
    if ($lines[$i] -match '^VERSION\s') { $i++ }
    if ($lines[$i] -match '^Begin\b') { $d = 0; do { if ($lines[$i] -match '^\s*Begin\b') { $d++ } elseif ($lines[$i] -match '^\s*End\s*$') { $d-- }; $i++ } while ($d -gt 0) }
    while ($lines[$i] -match '^Attribute VB_') { $i++ }
    ($lines[$i..($lines.Count - 1)] -join "`r`n").TrimEnd()
}

# テスト用に追加するモジュール
$harness = @'
Option Explicit
Public LastMsg As String
Public Sub TestMsgBox(Optional p, Optional b, Optional t, Optional u)
    LastMsg = LastMsg & "[Msg] " & p & " "
End Sub
Public Sub TestNotepad(ByVal s As String)
    LastMsg = LastMsg & "[メモ帳] " & Replace(Replace(s, vbCrLf, "|"), vbTab, " ") & " "
End Sub
Public Function RunCase() As String
    LastMsg = ""
    On Error Resume Next
    m_FilteringByArray.KoreDeFilter ""
    If Err.Number <> 0 Then
        RunCase = "ERR " & Err.Number & ": " & Err.Description & " " & LastMsg
    Else
        RunCase = LastMsg
    End If
End Function
Public Sub SetText(ByVal v As String)
    f_FilterTool.TextBox1.Value = v
End Sub
Public Sub ShowAll()
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub
'@

# 実行結果のログを、比較しやすい形（all / none / empty / missing:… / error）にする
function Get-Outcome([string]$log) {
    if ($log -match '^ERR ') { return 'error' }
    if ($log -match '\[メモ帳\] ない\|(.*?)\|\|') { return 'missing:' + ($matches[1] -replace '\|', ',') }
    if ($log -match 'ぜんぷありました') { return 'all' }
    if ($log -match '該当データがない') { return 'none' }
    if ($log -match 'テキストボックスに何も入っていません') { return 'empty' }
    return "?($log)"
}

# テストケース
#   R, C   : 見出しの位置
#   Table  : テーブルにするか（$false なら通常のオートフィルタ）
#   Data   : フィルタする列（1 行目は見出し）。隣の列には連番が入る
#   Setup  : そのほかの準備（ほかのテーブル、タイトル行など）
#   Tests  : 入力（改行は `n）、期待する結果、フィルタ後に表示される値
$fruits = @('品名', 'りんご', 'みかん', 'ぶどう', 'りんご')
$basic = @(
    @{ In = "りんご`nみかん";   Expect = 'all';           Visible = 'りんご,みかん,りんご' }
    @{ In = 'ばなな';            Expect = 'none';          Visible = 'りんご,みかん,ぶどう,りんご' }
    @{ In = "りんご`nばなな";   Expect = 'missing:ばなな'; Visible = 'りんご,りんご' }
)
$addTable = { param($ws, $addr, $values) for ($k = 0; $k -lt $values.Count; $k++) { $ws.Range($addr).Offset($k, 0).Value2 = $values[$k] }; $null = $ws.ListObjects.Add(1, $ws.Range($addr).Resize($values.Count, 1), $null, 1) }

$all = @(
    # 基本（問題 1・2・3・5 の確認を含む）
    @{ Name = 'range-A1';         R = 1; C = 1; Table = $false; Data = $fruits; Tests = $basic }
    @{ Name = 'table-A1';         R = 1; C = 1; Table = $true;  Data = $fruits; Tests = $basic }
    @{ Name = 'table-C3';         R = 3; C = 3; Table = $true;  Data = $fruits; Tests = $basic }
    @{ Name = 'range-C3-title';   R = 3; C = 3; Table = $false; Data = $fruits; Tests = $basic
       Setup = { param($ws) $ws.Range('A1').Value2 = '商品一覧'; $ws.Range('C1').Value2 = '2026年10月' } }
    @{ Name = 'two-tables';       R = 1; C = 5; Table = $true;  Data = $fruits; Tests = $basic
       Setup = { param($ws) & $addTable $ws 'A1' @('X', 'りんご', 'ぶどう') } }
    @{ Name = 'two-tables-first'; R = 1; C = 1; Table = $true;  Data = $fruits; Tests = $basic
       Setup = { param($ws) & $addTable $ws 'H1' @('品名', 'みかん', 'ばなな') } }
    @{ Name = 'three-tables';     R = 5; C = 3; Table = $true;  Data = $fruits; Tests = $basic
       Setup = { param($ws) & $addTable $ws 'A1' @('X', 'りんご', 'ぶどう'); & $addTable $ws 'H1' @('品名', 'みかん', 'ばなな') } }
    @{ Name = 'one-row-table';    R = 1; C = 1; Table = $true;  Data = @('品名', 'りんご'); Tests = @(
        @{ In = 'りんご'; Expect = 'all';  Visible = 'りんご' }
        @{ In = 'ばなな'; Expect = 'none'; Visible = 'りんご' }) }
    @{ Name = 'date-table';       R = 2; C = 2; Table = $true;  Data = @('生年月日', '1981/5/7', '1981/5/17', '1969/7/12', '1994/1/25'); Tests = @(
        @{ In = "1981/5/7`n1981/5/17`n1969/7/12`n1994/1/25"; Expect = 'all';             Visible = '1981/5/7,1981/5/17,1969/7/12,1994/1/25' }
        @{ In = "1981/5/7`n2000/1/1";                         Expect = 'missing:2000/1/1'; Visible = '1981/5/7' }) }
    @{ Name = 'partial';          R = 1; C = 1; Table = $true;  Data = @('品名', 'ABC', 'XYZ', 'ABC', 'XYZ'); Tests = @(
        @{ In = "AB`nXYZ"; Expect = 'missing:AB'; Visible = 'XYZ,XYZ' }) }
    # 問題 4: 各行の前後の空白（半角・全角）
    @{ Name = 'trim';             R = 1; C = 1; Table = $true;  Data = @('品名', 'ABC', 'XYZ', 'ABC', 'DEF'); Tests = @(
        @{ In = "ABC `n  XYZ";      Expect = 'all'; Visible = 'ABC,XYZ,ABC' }
        @{ In = "ABC　`n　XYZ";     Expect = 'all'; Visible = 'ABC,XYZ,ABC' }) }
    # 問題 6: 空の行（空白のセルが一緒に表示されないこと）
    @{ Name = 'blank-line';       R = 1; C = 1; Table = $true;  Data = @('品名', 'りんご', '', 'みかん', 'ぶどう'); Tests = @(
        @{ In = "りんご`n`nみかん";     Expect = 'all';   Visible = 'りんご,みかん' }
        @{ In = "`n`nりんご`n`n";       Expect = 'all';   Visible = 'りんご' }
        @{ In = "りんご`n `nみかん";    Expect = 'all';   Visible = 'りんご,みかん' }
        @{ In = "`n `n";                Expect = 'empty'; Visible = 'りんご,,みかん,ぶどう' }) }
)
if ($Cases) { $all = $all | Where-Object { $Cases -contains $_.Name } }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$pass = 0; $fail = 0
try {
    $addin = $xl.Workbooks.Open($Xlam, 0, $true)
    $vbp = $addin.VBProject
    foreach ($n in 'm_FilteringByArray.bas', 'c_FilteringByArray.cls', 'f_FilterTool.frm') {
        $cm = $vbp.VBComponents.Item([IO.Path]::GetFileNameWithoutExtension($n)).CodeModule
        $code = if ($Original) { $cm.Lines(1, $cm.CountOfLines) } else { Get-CodeBody ([IO.File]::ReadAllText((Join-Path $SrcDir $n))) }
        $code = [regex]::Replace($code, '(?m)^(\s*)MsgBox\b', '$1zz_Test.TestMsgBox')
        $code = [regex]::Replace($code, '(?m)^(\s*)WSH\.Popup\b', '$1zz_Test.TestMsgBox')
        # メモ帳への出力は、内容をログに残すだけにする
        $code = [regex]::Replace($code, '(?m)^(\s*)myDo\.SetText (\w+)', '$1zz_Test.TestNotepad $2')
        $code = [regex]::Replace($code, '(?m)^(\s*)(myDo\.PutInClipboard|ret = Shell\(|AppActivate|CreateObject\("Wscript\.Shell"\)\.SendKeys).*$', '$1'' (test) $2')
        $cm.DeleteLines(1, $cm.CountOfLines); $cm.AddFromString($code)
    }
    $t = $vbp.VBComponents.Add(1); $t.Name = 'zz_Test'; $t.CodeModule.AddFromString($harness)
    Write-Host ("[code] " + $(if ($Original) { "$Xlam（元のコード）" } else { $SrcDir }))

    $wb = $xl.Workbooks.Add()
    foreach ($case in $all) {
        $ws = $wb.Worksheets.Add(); $ws.Name = $case.Name
        $r = $case.R; $c = $case.C; $n = $case.Data.Count
        for ($k = 0; $k -lt $n; $k++) {
            $ws.Cells.Item($r + $k, $c).Formula = $case.Data[$k]
            $ws.Cells.Item($r + $k, $c + 1).Value2 = $(if ($k) { [string]$k } else { '数' })
        }
        $null = $ws.Columns.Item($c).AutoFit()
        if ($case.Setup) { & $case.Setup $ws }
        $rng = $ws.Range($ws.Cells.Item($r, $c), $ws.Cells.Item($r + $n - 1, $c + 1))
        if ($case.Table) { $null = $ws.ListObjects.Add(1, $rng, $null, 1) } else { $null = $rng.AutoFilter() }
        $null = $ws.Activate(); $null = $ws.Cells.Item($r + 1, $c).Select()

        foreach ($test in $case.Tests) {
            $xl.Run("'KoreDeFilter.xlam'!zz_Test.SetText", ($test.In -replace "`n", "`r`n"))
            $log = $xl.Run("'KoreDeFilter.xlam'!zz_Test.RunCase")
            $visible = @(); for ($k = 1; $k -lt $n; $k++) { if (-not $ws.Rows.Item($r + $k).Hidden) { $visible += $ws.Cells.Item($r + $k, $c).Text } }
            $xl.Run("'KoreDeFilter.xlam'!zz_Test.ShowAll")
            $outcome = Get-Outcome $log
            $ok = ($outcome -eq $test.Expect) -and (($visible -join ',') -eq $test.Visible)
            if ($ok) { $pass++ } else { $fail++ }
            $label = ($test.In -replace "`n", '⏎')
            Write-Host ("{0} [{1}] 入力={2}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $case.Name, $label) -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
            if (-not $ok) {
                Write-Host ("       結果: {0}（期待: {1}）" -f $outcome, $test.Expect)
                Write-Host ("       表示: {0}（期待: {1}）" -f ($visible -join ','), $test.Visible)
            }
        }
    }
} finally {
    foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect()
}
Write-Host ("`n{0} PASS / {1} FAIL" -f $pass, $fail) -ForegroundColor $(if ($fail) { 'Red' } else { 'Green' })
exit $(if ($fail) { 1 } else { 0 })
