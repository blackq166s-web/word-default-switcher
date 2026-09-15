$ErrorActionPreference='Stop'
$root=Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $root 'WordDefault.ps1') -LibraryOnly
. (Join-Path $root 'Batch.ps1')
# Real, read-only validation of the independent verification process.
$real=@(Read-VerifiedRows)
if ($real.Count -ne 8) { throw 'Expected eight detected formats.' }
foreach($ext in $extensions) { if (@($real | Where-Object {$_.Extension -eq $ext}).Count -ne 1) { throw 'Missing format.' } }
Write-Output 'PASS: independent read-only verification returns all eight formats'
Save-BatchSnapshot $real
if ([string]::IsNullOrEmpty($script:lastReportPath) -or !(Test-Path -LiteralPath $script:lastReportPath)) { throw 'Real report path missing.' }
if ([string]::IsNullOrEmpty($script:batchSnapshot) -or !(Test-Path -LiteralPath $script:batchSnapshot)) { throw 'Real snapshot path missing.' }
if (@(Import-Clixml -LiteralPath $script:batchSnapshot).Count -ne 8) { throw 'Real snapshot incomplete.' }
Write-Output 'PASS: real report returns a nonempty path and real snapshot contains all eight formats'

# Isolated fakes below. No registry writes and no vendor Set-FTA invocation.
$os=New-Object PSObject -Property @{Version='10.0.19045';Caption='Test Windows 10'}
$isWin7=$false
function Get-DefaultRows { foreach($ext in $extensions) { New-Object PSObject -Property @{Extension=$ext;Status='WPS'} } }
function Save-BatchSnapshot($Rows) { $script:saved=$true }
function Get-WordTarget($Extension) { if ($Extension -eq $script:missing) { return '' }; return 'TEST.Word' }
function Write-WordAssociation($Extension,$ProgId) {
 if (!$script:saved) { throw 'Write before snapshot!' }
 $script:writes += $Extension
 if ($Extension -eq $script:denied) { throw 'TEST: write denied' }
}
function Read-VerifiedRows {
 if ($script:verifyThrows) { throw 'TEST: verification failed' }
 foreach($ext in $extensions) {
  $state='Word（系统关联查询）'
  if ($ext -eq $script:missing -or $ext -eq $script:denied -or $script:silentReject) { $state='WPS' }
  New-Object PSObject -Property @{Extension=$ext;Status=$state}
 }
}
$saved=$false; $writes=@(); $missing=''; $denied=''; $verifyThrows=$false; $silentReject=$false
$r=@(Invoke-WordBatch $extensions)
if ($writes.Count -ne 8 -or @($r | Where-Object {!$_.Success}).Count) { throw 'All-success case failed' }
Write-Output 'PASS: saves snapshot before writing and reports all eight verified successes'
$missing='.dot'; $denied='.doc'; $writes=@(); $saved=$false
$r=@(Invoke-WordBatch $extensions)
if ($writes -contains '.dot' -or $writes.Count -ne 7 -or @($r | Where-Object {!$_.Success}).Count -ne 2) { throw 'Partial failure case failed' }
Write-Output 'PASS: missing target and denied write are isolated; other formats continue'
$missing=''; $denied=''; $silentReject=$true
$r=@(Invoke-WordBatch $extensions)
if (@($r | Where-Object {$_.Success}).Count) { throw 'False success on rejected writes' }
Write-Output 'PASS: write returning normally is not treated as success when verification disagrees'
$verifyThrows=$true; $caught=$false
try { Invoke-WordBatch $extensions | Out-Null } catch { $caught=$true }
if (!$caught) { throw 'Failed verification must stop success reporting' }
Write-Output 'PASS: failed independent verification cannot produce success'
$verifyThrows=$false; $silentReject=$false
foreach ($build in @('10.0.22631','10.0.26100','10.0.26200')) {
 $os.Version=$build; $writes=@(); $saved=$false
 $r=@(Invoke-WordBatch $extensions)
 if ($writes.Count -ne 8 -or @($r | Where-Object {!$_.Success}).Count) { throw 'Win11 batch branch blocked' }
}
Write-Output 'PASS: Win11 builds 22631, 26100 and 26200 reach batch setting and verification'
$os.Version='6.2.9200'; $writes=@(); $caught=$false
try { Invoke-WordBatch $extensions | Out-Null } catch { $caught=$true }
if (!$caught -or $writes.Count) { throw 'Unsupported OS not blocked' }
Write-Output 'PASS: unsupported OS rejected before any mutation'
