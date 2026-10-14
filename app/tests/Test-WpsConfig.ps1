$ErrorActionPreference='Stop'
$root=Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $root 'WpsConfig.ps1')
# Regression cases taken from the real WPS accessibility tree.
foreach ($name in @('高级(A)...','高级(&A)...','高级（A）…','高级')) {
 if ((Normalize-WpsLabel $name) -ne '高级') { throw ('Advanced label mismatch: '+$name) }
}
if (Test-WpsControlType ([Windows.Automation.ControlType]::Tab) '兼容设置') { throw 'Tab container mistaken for item' }
if (!(Test-WpsControlType ([Windows.Automation.ControlType]::TabItem) '兼容设置')) { throw 'Tab item rejected' }
Write-Output 'PASS: real advanced label variants and duplicate Tab/TabItem names are handled'
# Exercise the real optional-reader branch with a fake accessibility inventory.
function Get-WpsElements($Root,$Label) {
 $count=if ($Label -eq '其他选项') {$script:tabCount} else {$script:optionCount}
 for($i=0;$i -lt $count;$i++) { 'fake-control' }
}
function Invoke-WpsControl($Root,$Label) { }
function Read-WpsToggle($Root,$Label) {
 if ($script:optionCount -ne 1 -or $script:unreadable) { throw 'Unreadable or duplicate checkbox' }
 return $script:optionState
}
$tabCount=0; $optionCount=0; $unreadable=$false; $optionState=0
if ($null -ne (Read-WpsOptionalOffline 'fake')) { throw 'Missing optional tab should skip' }
$tabCount=1
if ($null -ne (Read-WpsOptionalOffline 'fake')) { throw 'Missing optional checkbox should skip' }
$optionCount=1
foreach($optionState in @(0,1)) {
 if ((Read-WpsOptionalOffline 'fake') -ne $optionState) { throw 'Optional checkbox state lost' }
}
foreach($scenario in @('duplicate-tab','duplicate-option','unreadable')) {
 $tabCount=1; $optionCount=1; $unreadable=$false
 if ($scenario -eq 'duplicate-tab') {$tabCount=2}
 if ($scenario -eq 'duplicate-option') {$optionCount=2}
 if ($scenario -eq 'unreadable') {$unreadable=$true}
 $caught=$false
 try { Read-WpsOptionalOffline 'fake' | Out-Null } catch {$caught=$true}
 if (!$caught) { throw ('Optional control errors must not be treated as absence: '+$scenario) }
}
if (!(Test-WpsControlType ([Windows.Automation.ControlType]::CheckBox) $script:wpsLabels.Offline)) { throw 'Offline checkbox type not recognized' }
Write-Output 'PASS: optional control absence skips; duplicate or unreadable controls fail'
# Transaction tests use an in-memory WPS backend; never launch or change real WPS.
function Get-WpsRunningProcesses { return @() }
function Get-WpsVisibleWindows { return @() }
function Start-Sleep {}
function Open-WpsConfig { $script:opens++; return 'fake-window' }
function Get-WpsControl {
 $obj=New-Object PSObject -Property @{Current=(New-Object PSObject -Property @{IsEnabled=$true})}
 $obj | Add-Member ScriptMethod TryGetCurrentPattern {param($id,$pattern) return $true}
 return $obj
}
function Read-WpsOptions {
 if ($script:readFails) { throw 'Unsupported checkbox' }
 if ($script:opens -eq 2 -and $script:rejectSave) { return New-Object PSObject -Property @{Compatibility=1;UpdatesOff=0} }
 return New-Object PSObject -Property @{Compatibility=$script:compat;UpdatesOff=$script:updatesOff;OfflineAvailable=$script:offlineAvailable;Offline=$(if ($script:opens -eq 2 -and $script:offlineLost) {0} else {$script:offline})}
}
function Get-WpsAssociationSnapshot { return @{'.docx'='unchanged-handler'} }
function Clear-WpsAssociationRequests { $script:actions+='Preserve associations' }
function Set-WpsUpdatesOff {
 if ($script:writeFails) { throw 'Update option rejected' }
 $script:updatesOff=1
}
function Save-WpsRecord { param($Record,$Path) $script:records+=,$Record.Status }
function Invoke-WpsControl {
 param($Root,$Label)
 $script:actions+=,$Label
 if ($Label -eq '确定') { $script:commits++ }
}
function Set-WpsToggle {
 param($Root,$Label,$Desired)
 if (!$script:records.Count) { throw 'No snapshot before write' }
 $script:writes++
 if ($script:writeFails -and $Desired -eq 1) { throw 'Checkbox rejected' }
 if ($Label -eq $script:wpsLabels.Offline) {
  if ($script:offlineWriteFails) { throw 'Offline checkbox rejected' }
  $script:offline=$Desired
 } elseif ($Desired -eq 0) { $script:compat=0 } else { throw 'Unexpected checkbox' }
}
function Reset-Fake {
 $script:offlineAvailable=$false; $script:offline=$null; $script:offlineLost=$false; $script:offlineWriteFails=$false
 $script:opens=0; $script:compat=1; $script:updatesOff=0; $script:records=@(); $script:actions=@()
 $script:writes=0; $script:commits=0; $script:readFails=$false; $script:writeFails=$false; $script:rejectSave=$false
}
Reset-Fake
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($opens -ne 2 -or $commits -ne 1 -or $result.After.UpdatesOff -ne 1 -or $result.After.Compatibility -ne 0) { throw 'Happy path failed' }
Write-Output 'PASS: snapshot precedes writes; two desired states verified after reopen'
Reset-Fake; $readFails=$true
$caught=$false
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $caught=$true }
if (!$caught -or $writes -or $commits) { throw 'Unsupported interface wrote settings' }
Write-Output 'PASS: unknown controls stop before writes'
Reset-Fake; $writeFails=$true
$caught=$false
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $caught=$true }
if (!$caught -or $commits -or $actions[-1] -ne '取消') { throw 'Partial edit was not cancelled' }
Write-Output 'PASS: partial edit failure cancels without save'
Reset-Fake; $rejectSave=$true
$caught=$false
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $caught=$true }
if (!$caught -or $records[-1] -ne '已保存但复查未完成，请人工检查') { throw 'False success after save rejection' }
Write-Output 'PASS: persistence mismatch never reports success'
Reset-Fake
function Get-WpsRunningProcesses { return New-Object PSObject -Property @{Id=42;ProcessName='wps'} }
function Get-WpsVisibleWindows { return New-Object PSObject -Property @{ProcessId=42;Title='Unsaved document'} }
$caught=$false
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $caught=$true }
if (!$caught -or $opens -or $writes) { throw 'Running document apps were not blocked' }
Write-Output 'PASS: existing WPS/Word sessions block the operation before launch'

Reset-Fake
function Get-WpsVisibleWindows { return @() }
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($opens -ne 2 -or $commits -ne 1) { throw 'Background-only process incorrectly blocked' }
Write-Output 'PASS: background-only WPS process does not block configuration'
Reset-Fake
function Get-WpsVisibleWindows { return New-Object PSObject -Property @{ProcessId=99;Title='Unrelated window'} }
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($commits -ne 1) { throw 'Unrelated window blocked configuration' }
Write-Output 'PASS: unrelated application windows do not block configuration'
Reset-Fake
function Get-WpsRunningProcesses { return New-Object PSObject -Property @{Id=42;ProcessName='ksomisc'} }
function Get-WpsVisibleWindows { return New-Object PSObject -Property @{ProcessId=42;Title='WPS configuration'} }
$message=''
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $message=$_.Exception.Message }
if ($opens -or $message -notmatch 'WPS configuration' -or $message -notmatch 'PID 42') { throw 'Configuration window not identified precisely' }
Write-Output 'PASS: existing configuration window is named with its process ID'
Reset-Fake
function Get-WpsRunningProcesses { return New-Object PSObject -Property @{Id=42;ProcessName='wps'} }
$script:closeRequests=@(); $script:pending=$true
function Get-WpsVisibleWindows {
 if ($script:pending) {
  New-Object PSObject -Property @{ProcessId=42;Title='WPS configuration';Handle=123}
  New-Object PSObject -Property @{ProcessId=99;Title='Other application';Handle=456}
 }
}
function Send-WpsWindowClose($Window) { $script:closeRequests+=,$Window.Handle; $script:pending=$false }
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml' -CloseWindows
if ($closeRequests.Count -ne 1 -or $closeRequests[0] -ne 123 -or $commits -ne 1) { throw 'Close and apply flow failed' }
Write-Output 'PASS: only relevant windows receive close requests before settings are applied'
$script:pending=$true; $script:closeRequests=@()
function Send-WpsWindowClose($Window) { $script:closeRequests+=,$Window.Handle }
$message=''
try { Close-WpsWindows -TimeoutSeconds 0 } catch { $message=$_.Exception.Message }
if (!$message -or $closeRequests.Count -ne 1 -or !$script:pending) { throw 'Pending save prompt handling failed' }
Write-Output 'PASS: pending window stops operation; no repeated close request or force kill'
Reset-Fake
$script:pending=$true; $script:configExitCalls=0
function Get-WpsRunningProcesses { return New-Object PSObject -Property @{Id=42;ProcessName='ksomisc'} }
function Get-WpsVisibleWindows {
 if ($script:pending) {
  New-Object PSObject -Property @{ProcessId=42;Title='WPS config launcher';Handle=123}
  New-Object PSObject -Property @{ProcessId=42;Title='WPS config advanced';Handle=124}
 }
}
function Exit-WpsConfiguration($ProcessId) { $script:configExitCalls++; $script:pending=$false }
function Send-WpsWindowClose($Window) { throw 'Config must use its own exit path' }
Close-WpsWindows -TimeoutSeconds 0
if ($script:configExitCalls -ne 1) { throw 'Configuration exit must run once per process' }
Write-Output 'PASS: configuration windows use dedicated exit once per process'
function Get-WpsRunningProcesses { return @() }
function Get-WpsVisibleWindows { return @() }
Reset-Fake
$script:compat=0; $script:updatesOff=1
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($commits -ne 0 -or $writes -ne 0) { throw 'Already-correct configuration must not be saved again' }
Write-Output 'PASS: already-correct settings skip saving; no file associations are requested'
Reset-Fake
function Get-WpsAssociationSnapshot {
 if ($script:commits) { return @{'.docx'='unexpected-handler'} }
 return @{'.docx'='original-handler'}
}
$message=''
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $message=$_.Exception.Message }
if ($message -notmatch '文件打开方式变化' -or $message -notmatch '\.docx') { throw 'Association drift was not reported' }
Write-Output 'PASS: association drift never reports success'
if ($actions.IndexOf('Preserve associations') -gt $actions.IndexOf('确定')) { throw 'Preservation ran after save' }
function Get-WpsAssociationSnapshot { return @{'.docx'='unchanged-handler'} }
Reset-Fake; $offlineAvailable=$true; $offline=0
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($result.After.Offline -ne 1 -or $commits -ne 1) { throw 'Optional offline checkbox not enabled' }
Write-Output 'PASS: present offline option is enabled and verified after reopen'
Reset-Fake; $compat=0; $updatesOff=1; $offlineAvailable=$true; $offline=0
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($result.After.Offline -ne 1 -or $commits -ne 1) { throw 'Offline-only change incorrectly skipped' }
Reset-Fake; $compat=0; $updatesOff=1; $offlineAvailable=$true; $offline=1
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($commits -or $writes) { throw 'Already-enabled offline option triggered unnecessary save' }
Write-Output 'PASS: offline-only change is applied; all-correct configuration skips saving'
Reset-Fake
$result=Invoke-WpsRecommended 'fake.exe' 'fake.xml'
if ($result.After.OfflineAvailable -or $result.Status -notmatch '已跳过' -or $commits -ne 1) { throw 'Absent option should skip without blocking other settings' }
Write-Output 'PASS: absent offline option is recorded as skipped; other settings complete'
Reset-Fake; $offlineAvailable=$true; $offline=0; $offlineWriteFails=$true
$message=''
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $message=$_.Exception.Message }
if (!$message -or $commits -or $actions[-1] -ne '取消') { throw 'Offline write failure must cancel before save' }
Reset-Fake; $offlineAvailable=$true; $offline=0; $offlineLost=$true
$message=''
try { Invoke-WpsRecommended 'fake.exe' 'fake.xml' } catch { $message=$_.Exception.Message }
if ($message -notmatch '离线使用设置保存后未生效' -or $records[-1] -ne '已保存但复查未完成，请人工检查') { throw 'Offline persistence failure falsely succeeded' }
Write-Output 'PASS: offline write and persistence failures never report success'
