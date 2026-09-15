# Loaded after WordDefault.ps1 -LibraryOnly. No settings change on import.
$script:batchRoot=Split-Path -Parent $MyInvocation.MyCommand.Path

function Get-WordTarget([string]$Extension) {
 $known=@{'.asd'='Word.AutoRecover.8';'.doc'='Word.Document.8';'.docm'='Word.DocumentMacroEnabled.12';'.docx'='Word.Document.12';'.dot'='Word.Template.8';'.dotm'='Word.TemplateMacroEnabled.12';'.dotx'='Word.Template.12';'.odt'='Word.OpenDocumentText.12'}
 if (!$known.ContainsKey($Extension)) { throw '不支持的文件格式。' }
 # Check the standard Word ProgID first; do not resolve unrelated app handlers.
 $standardExe=[WordDefaultNative]::Query($known[$Extension],2)
 if ($standardExe -and [IO.Path]::GetFileName($standardExe.Trim('"')) -ieq 'WINWORD.EXE' -and (Test-Path -LiteralPath $standardExe.Trim('"'))) { return $known[$Extension] }
 $ids=@()
 foreach ($hive in @('HKCU:','HKLM:')) {
  foreach ($part in @('Software\RegisteredApplications','Software\WOW6432Node\RegisteredApplications')) {
   $path=$hive+'\'+$part
   if (!(Test-Path $path)) { continue }
   $key=Get-Item $path
   foreach ($name in $key.GetValueNames()) {
    $cap=$hive+'\'+$key.GetValue($name)+'\FileAssociations'
    if (Test-Path $cap) { $ids += (Get-Item $cap).GetValue($Extension) }
   }
  }
 }
 $ids += $known[$Extension]
 foreach ($id in @($ids | Where-Object { $_ -and $_ -match '^Word\.' } | Select-Object -Unique)) {
  $exe=[WordDefaultNative]::Query($id,2)
  if ($exe -and [IO.Path]::GetFileName($exe.Trim('"')) -ieq 'WINWORD.EXE' -and (Test-Path -LiteralPath $exe.Trim('"'))) {
   return [string]$id
  }
 }
 return ''
}

function Save-BatchSnapshot([object[]]$Rows) {
 $reportPath=Save-Report '一键设置前'
 if ([string]::IsNullOrEmpty($reportPath) -or !(Test-Path -LiteralPath $reportPath)) { throw '检测记录未保存成功，已停止修改。' }
 $file=[IO.Path]::ChangeExtension($reportPath,'.xml')
 $Rows | Export-Clixml -LiteralPath $file
 $script:batchSnapshot=$file
}

function Read-VerifiedRows {
 # A fresh Windows PowerShell process avoids relying only on this window's cache.
 $out=Join-Path $env:TEMP ('WordDefaultVerify-'+[Guid]::NewGuid().ToString('N')+'.xml')
 try {
  $ps=Join-Path $PSHOME 'powershell.exe'
  & $ps -NoProfile -STA -ExecutionPolicy Bypass -File (Join-Path $script:batchRoot 'Verify.ps1') -OutputFile $out
  if ($LASTEXITCODE -ne 0 -or !(Test-Path -LiteralPath $out)) { throw '独立进程复查失败，不能确认设置成功。' }
  @(Import-Clixml -LiteralPath $out)
 } finally { if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out } }
}

function Write-WordAssociation([string]$Extension,[string]$ProgId) {
 if ($script:extensions -notcontains $Extension) { throw '格式不在允许列表中。' }
 if ($script:isWin7) {
  $apps=@(Find-RegisteredWord $Extension | Select-Object -Unique)
  if ($apps.Count -ne 1) { throw '未找到唯一 Word 注册项，请使用系统设置。' }
  $hr=[WordDefaultNative]::SetDefault($apps[0],$Extension)
  if ($hr -ne 0) { throw ('系统接口返回 '+$hr) }
 } else {
  # No Internet access, UI automation, drivers, elevation or Office repair.
  Set-FTA -ProgId $ProgId -Extension $Extension -ErrorAction Stop | Out-Null
 }
}

function Invoke-WordBatch([string[]]$Selected) {
 if (!$Selected.Count) { throw '请先勾选需要设置的格式。' }
 foreach ($ext in $Selected) { if ($script:extensions -notcontains $ext) { throw '格式不在允许列表中。' } }
 $version=[version]$script:os.Version
 if (!$script:isWin7 -and $version.Major -ne 10) {
  throw '当前一键设置支持 Win7 / Win10 / Win11。此系统请使用“系统默认应用”。'
 }
 if (!$script:isWin7 -and $PSVersionTable.PSVersion.Major -lt 5) { throw 'Win10 / Win11 一键设置需要 Windows PowerShell 5 或更高版本。' }
 $before=@(Get-DefaultRows)
 Save-BatchSnapshot $before
 $attempts=@()
 foreach ($ext in @($Selected | Select-Object -Unique)) {
  $errorText=''; $target=''
  $old=@($before | Where-Object { $_.Extension -eq $ext })
  if (!$old.Count -or $old[0].Status -ne 'Word（系统关联查询）') {
   try {
    $target=Get-WordTarget $ext
    if (!$target) { throw '本机未找到此格式的有效 Word 注册项，未修改。' }
    Write-WordAssociation $ext $target
   } catch { $errorText=$_.Exception.Message }
  }
  $attempts += New-Object PSObject -Property @{Extension=$ext;Target=$target;Detail=$errorText}
 }
 $verified=@(Read-VerifiedRows)
 foreach ($attempt in $attempts) {
  $actual=@($verified | Where-Object { $_.Extension -eq $attempt.Extension })
  $ok=$actual.Count -eq 1 -and $actual[0].Status -eq 'Word（系统关联查询）'
  $detail=$attempt.Detail
  if (!$ok -and !$detail) { $detail='设置后仍未指向 Word；可能被系统保护或策略阻止。' }
  New-Object PSObject -Property @{Extension=$attempt.Extension;Success=$ok;Target=$attempt.Target;Detail=$detail}
 }
}
