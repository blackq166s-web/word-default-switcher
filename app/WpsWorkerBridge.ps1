# The UI process never loads UI Automation or calls the WPS automation backend.
$script:wpsWorkerFile=Join-Path $PSScriptRoot 'WpsWorker.ps1'
function Invoke-WpsWorker([string]$Mode,[string]$RecordPath,[bool]$CloseWindows=$false,[switch]$Elevated) {
 $ErrorActionPreference='Stop'
 $jobDir=Join-Path $env:TEMP ('WordDefaultWorker-'+[Guid]::NewGuid().ToString('N'))
 [void][IO.Directory]::CreateDirectory($jobDir)
 $requestPath=Join-Path $jobDir 'request.xml'
 $responsePath=Join-Path $jobDir 'response.xml'
 [pscustomobject]@{Mode=$Mode;RecordPath=$RecordPath;CloseWindows=$CloseWindows;UserSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value} | Export-Clixml -LiteralPath $requestPath
 $script:wpsWorkerBusy=$true
 try {
  $arguments='-NoProfile -STA -ExecutionPolicy Bypass -File "'+$script:wpsWorkerFile+'" -RequestPath "'+$requestPath+'" -ResponsePath "'+$responsePath+'"'
  $powershell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  $native=Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
  if (Test-Path -LiteralPath $native) { $powershell=$native }
  if ($Elevated) { $worker=Start-Process -FilePath $powershell -ArgumentList $arguments -WindowStyle Hidden -Verb RunAs -PassThru }
  else { $worker=Start-Process -FilePath $powershell -ArgumentList $arguments -WindowStyle Hidden -PassThru }
  while (!$worker.HasExited) {
   if ('Windows.Forms.Application' -as [type]) { [Windows.Forms.Application]::DoEvents() }
   Start-Sleep -Milliseconds 50
   $worker.Refresh()
  }
  if (!(Test-Path -LiteralPath $responsePath)) { throw '后台任务意外结束，未确认设置成功。请重试。' }
  $response=Import-Clixml -LiteralPath $responsePath
  if (!$response.Success) {
   if (!$Elevated -and $response.Error -match 'ELEVATION_REQUIRED' -and ('Windows.Forms.Application' -as [type])) {
    $answer=[Windows.Forms.MessageBox]::Show('此版本的 WPS 配置工具可能需要管理员权限。是否使用当前账号授权后重试？请勿输入其他账号的凭据。','需要授权',[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Information,[Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($answer -eq [Windows.Forms.DialogResult]::Yes) { return Invoke-WpsWorker $Mode $RecordPath $CloseWindows -Elevated }
   }
   throw $response.Error
  }
  return $response.Value
 } finally {
  $script:wpsWorkerBusy=$false
  # Delete only the two per-call IPC files; never touch report or document paths.
  foreach($file in @($requestPath,$responsePath)) { if ([IO.File]::Exists($file)) { [IO.File]::Delete($file) } }
  if ([IO.Directory]::Exists($jobDir)) { [IO.Directory]::Delete($jobDir) }
 }
}
function Assert-WpsWindowsClosed { Invoke-WpsWorker 'Check' }
function Find-WpsConfig { return Invoke-WpsWorker 'Locate' }
function Invoke-WpsRecommended($ConfigPath,[string]$RecordPath,[switch]$CloseWindows) {
 return Invoke-WpsWorker 'Optimize' $RecordPath ([bool]$CloseWindows)
}
