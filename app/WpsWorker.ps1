param([string]$RequestPath,[string]$ResponsePath)
$ErrorActionPreference='Stop'
try {
 . (Join-Path $PSScriptRoot 'WpsConfig.ps1')
 $request=Import-Clixml -LiteralPath $RequestPath
 if ($request.UserSid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value) { throw '请使用当前账号授权，不能切换到其他管理员账号修改配置。' }
 $value=$null
 switch ($request.Mode) {
  'Check' { Assert-WpsWindowsClosed }
  'Locate' { $value=Find-WpsConfig }
  'Optimize' { $value=Invoke-WpsRecommended (Find-WpsConfig) $request.RecordPath -CloseWindows:([bool]$request.CloseWindows) }
  'Probe' {
   try {
    $root=Open-WpsConfig (Find-WpsConfig)
    $value=Read-WpsOptions $root
    Invoke-WpsControl $root '取消'
   } finally { Stop-WpsBackground }
  }
  default { throw '未知的后台任务。' }
 }
 [pscustomobject]@{Success=$true;Value=$value;Error=''} | Export-Clixml -LiteralPath $ResponsePath -Depth 6
} catch {
 [pscustomobject]@{Success=$false;Value=$null;Error=$_.Exception.Message} | Export-Clixml -LiteralPath $ResponsePath
 exit 1
}
