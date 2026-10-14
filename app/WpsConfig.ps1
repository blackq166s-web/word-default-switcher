# Uses supported accessibility patterns only; no guessed registry keys or coordinates.
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$script:wpsLabels=@{Compatibility='WPS Office 兼容第三方系统和软件';UpdatesOff='关闭自动升级';Offline='兼容离线状态下的未登录使用方式'}
$script:wpsBackgroundProcesses=@()
function Initialize-WpsBackground {
 if ('WpsBackgroundWindows' -as [type]) { return }
 Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
public static class WpsBackgroundWindows {
 delegate bool Callback(IntPtr handle, IntPtr param);
 [DllImport("user32.dll")] static extern bool EnumWindows(Callback callback, IntPtr param);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr handle, out uint id);
 [DllImport("user32.dll")] static extern bool ShowWindowAsync(IntPtr handle, int command);
 [DllImport("user32.dll")] static extern bool PostMessage(IntPtr handle, uint message, IntPtr wparam, IntPtr lparam);
 public static IntPtr[] Handles(int pid) {
  var result=new List<IntPtr>();
  EnumWindows(delegate(IntPtr h, IntPtr unused) { uint id; GetWindowThreadProcessId(h,out id); if(id==pid) result.Add(h); return true; },IntPtr.Zero);
  return result.ToArray();
 }
 public static ManualResetEvent Guard(Process process) {
  var stop=new ManualResetEvent(false);
  ThreadPool.QueueUserWorkItem(delegate(object unused) {
   // Only a process launched by this operation is eligible for hiding.
   // A bounded guard cannot continue hiding windows indefinitely after failure.
   var deadline=DateTime.UtcNow.AddMinutes(2);
   try {
    while(!stop.WaitOne(20) && DateTime.UtcNow<deadline && !process.HasExited)
     foreach(var h in Handles(process.Id)) ShowWindowAsync(h,0);
   } catch { }
   finally { if(!stop.WaitOne(0) && !process.HasExited) Reveal(process.Id); }
  });
  return stop;
 }
 public static void Close(int pid) { foreach(var h in Handles(pid)) PostMessage(h,0x0010,IntPtr.Zero,IntPtr.Zero); }
 public static void Reveal(int pid) { foreach(var h in Handles(pid)) ShowWindowAsync(h,4); }
}
'@
}
function Stop-WpsBackground {
 foreach ($entry in $script:wpsBackgroundProcesses) {
  # Graceful close of our own configuration windows only, never document processes.
  if (!$entry.Process.HasExited) {
   foreach ($handle in [WpsBackgroundWindows]::Handles($entry.Process.Id)) {
    try {
     $window=[Windows.Automation.AutomationElement]::FromHandle($handle)
     if (@(Get-WpsElements $window '退出').Count -eq 1) { Invoke-WpsControl $window '退出' }
    } catch { }
   }
   if (!$entry.Process.HasExited) { [WpsBackgroundWindows]::Close($entry.Process.Id) }
  }
  [void]$entry.Stop.Set()
  if (!$entry.Process.WaitForExit(1000)) { [WpsBackgroundWindows]::Reveal($entry.Process.Id) }
 }
 $script:wpsBackgroundProcesses=@()
}
function Find-WpsConfig {
 $paths=@()
 foreach ($hive in @('HKCU','HKLM')) {
  foreach ($view in @('SOFTWARE','SOFTWARE\WOW6432Node')) {
   foreach ($app in @('wps.exe','ksomisc.exe','ksolaunch.exe')) {
    $key=Get-Item -LiteralPath ($hive+':\'+$view+'\Microsoft\Windows\CurrentVersion\App Paths\'+$app) -ErrorAction SilentlyContinue
    if ($key) { $value=[string]$key.GetValue(''); if ($value) { $paths+=Join-Path (Split-Path $value.Trim('"') -Parent) 'ksomisc.exe' } }
   }
  }
 }
 # WPS installs versioned directories, so use the installed configuration shortcut.
 $shell=New-Object -ComObject WScript.Shell
 try {
  foreach ($menu in @([Environment]::GetFolderPath('StartMenu'),[Environment]::GetFolderPath('CommonStartMenu'))) {
   foreach ($link in @(Get-ChildItem -LiteralPath $menu -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue)) {
    if ($link.Name -match '配置|Configuration') {
     $shortcut=$shell.CreateShortcut($link.FullName)
     if ([IO.Path]::GetFileName($shortcut.TargetPath) -ieq 'ksomisc.exe') { $paths+=$shortcut.TargetPath }
    }
   }
  }
 } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
 $found=@($paths | Where-Object {$_ -and (Test-Path -LiteralPath $_ -PathType Leaf)} | Select-Object -Unique)
 if ($found.Count -ne 1) { throw '未找到唯一的 WPS 配置工具。请从 WPS 开始菜单打开配置工具，手动检查相关设置。' }
 return $found[0]
}
function Normalize-WpsLabel([string]$Name) {
 return ($Name.Trim() -replace '&','' -replace '(\.{3}|…)+$','' -replace '[（(][A-Za-z][）)]\s*$','').Trim()
}
function Test-WpsControlType($Type,[string]$Label) {
 if ($Label -in @('兼容设置','其他选项','升级设置')) { return $Type -eq [Windows.Automation.ControlType]::TabItem }
 if ($Label -eq $script:wpsLabels.UpdatesOff) { return $Type -eq [Windows.Automation.ControlType]::RadioButton }
 if ($Label -in @($script:wpsLabels.Compatibility,$script:wpsLabels.Offline)) { return $Type -eq [Windows.Automation.ControlType]::CheckBox }
 return $Type -eq [Windows.Automation.ControlType]::Button
}
function Get-WpsElements($Root,[string]$Label) {
 $matches=@()
 foreach ($element in $Root.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)) {
  $name=Normalize-WpsLabel $element.Current.Name
  # Hidden windows still expose UIA controls. Match type and name, not screen visibility.
  if ($name -eq $Label -and (Test-WpsControlType $element.Current.ControlType $Label)) { $matches+=,$element }
 }
 return $matches
}
function Get-WpsControl($Root,[string]$Label) {
 $items=@(Get-WpsElements $Root $Label)
 if ($items.Count -ne 1) { throw ('无法唯一识别选项：'+$Label+'。此版本界面可能不支持自动操作，请手动设置。') }
 return $items[0]
}
function Invoke-WpsControl($Root,[string]$Label) {
 $element=Get-WpsControl $Root $Label
 if (!$element.Current.IsEnabled) { throw ('控件不可用：'+$Label) }
 $pattern=$null
 if ($element.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern,[ref]$pattern)) { $pattern.Select() }
 elseif ($element.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern,[ref]$pattern)) { $pattern.Invoke() }
 else { throw ('控件没有可用的自动操作接口：'+$Label) }
 Start-Sleep -Milliseconds 250
}
function Read-WpsToggle($Root,[string]$Label) {
 $element=Get-WpsControl $Root $Label
 $pattern=$null
 if (!$element.Current.IsEnabled -or !$element.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern,[ref]$pattern)) { throw ('无法读写复选框：'+$Label) }
 $state=[int]$pattern.Current.ToggleState
 if ($state -notin @(0,1)) { throw ('复选框状态不明确：'+$Label) }
 return $state
}
function Set-WpsToggle($Root,[string]$Label,[int]$Desired) {
 if ((Read-WpsToggle $Root $Label) -eq $Desired) { return }
 $element=Get-WpsControl $Root $Label
 $pattern=$element.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
 $pattern.Toggle()
 if ((Read-WpsToggle $Root $Label) -ne $Desired) { throw ('选项未切换到目标状态：'+$Label) }
}
function Open-WpsConfig([string]$Path) {
 Initialize-WpsBackground
 $process=Start-Process -FilePath $Path -PassThru -WindowStyle Hidden
 $script:wpsBackgroundProcesses+=New-Object PSObject -Property @{Process=$process;Stop=[WpsBackgroundWindows]::Guard($process)}
 $deadline=[DateTime]::UtcNow.AddSeconds(12)
 $advancedOpened=$false
 do {
  $process.Refresh()
  if ($process.HasExited) { throw 'ELEVATION_REQUIRED: WPS 配置工具启动后切换了进程，可能需要管理员权限。' }
  $windows=@()
  foreach ($handle in [WpsBackgroundWindows]::Handles($process.Id)) {
   try { $win=[Windows.Automation.AutomationElement]::FromHandle($handle) }
   catch [Windows.Automation.ElementNotAvailableException] { continue }
   if ($win.Current.Name -match 'WPS.*配置工具') { $windows+=,$win }
  }
  # The advanced dialog may be a separate window or a descendant of the launcher.
  # Re-enumerate until its TabItem exists; never return the launcher prematurely.
  foreach ($root in $windows) {
   if (@(Get-WpsElements $root '兼容设置').Count -eq 1) { return $root }
  }
  if (!$advancedOpened) {
   foreach ($root in $windows) {
    if (@(Get-WpsElements $root '高级').Count -eq 1) {
     Invoke-WpsControl $root '高级'; $advancedOpened=$true; break
    }
   }
  }
  Start-Sleep -Milliseconds 200
 } while ([DateTime]::UtcNow -lt $deadline)
 throw '未能打开 WPS 高级设置页面，本次尚未修改设置。请关闭配置工具后重试。'
}
function Read-WpsOptionalOffline($Root) {
 $tabs=@(Get-WpsElements $Root '其他选项')
 if ($tabs.Count -eq 0) { return $null }
 if ($tabs.Count -ne 1) { throw '无法唯一识别其他选项页面，未确认离线设置。' }
 Invoke-WpsControl $Root '其他选项'
 $options=@(Get-WpsElements $Root $script:wpsLabels.Offline)
 if ($options.Count -eq 0) { return $null }
 # Existing but ambiguous, disabled or unreadable controls are failures, not absence.
 return Read-WpsToggle $Root $script:wpsLabels.Offline
}
function Read-WpsOptions($Root) {
 Invoke-WpsControl $Root '兼容设置'
 $compat=Read-WpsToggle $Root $script:wpsLabels.Compatibility
 Invoke-WpsControl $Root '升级设置'
 $updatesOff=Read-WpsUpdatesOff $Root
 $offline=Read-WpsOptionalOffline $Root
 return New-Object PSObject -Property @{Compatibility=$compat;UpdatesOff=$updatesOff;OfflineAvailable=($null -ne $offline);Offline=$offline}
}
function Read-WpsUpdatesOff($Root) {
 $control=Get-WpsControl $Root $script:wpsLabels.UpdatesOff
 $pattern=$null
 if (!$control.Current.IsEnabled -or !$control.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern,[ref]$pattern)) { throw '此版本无法自动读取升级选项，请在 WPS 配置工具中手动关闭自动升级。' }
 return [int]$pattern.Current.IsSelected
}
function Set-WpsUpdatesOff($Root) {
 if ((Read-WpsUpdatesOff $Root) -eq 1) { return }
 Invoke-WpsControl $Root $script:wpsLabels.UpdatesOff
 if ((Read-WpsUpdatesOff $Root) -ne 1) { throw '关闭自动升级未生效，未保存配置。' }
}
function Get-WpsAssociationSnapshot {
 if (!('WpsAssociationQuery' -as [type])) {
  Add-Type -TypeDefinition @'
using System; using System.Text; using System.Runtime.InteropServices;
public static class WpsAssociationQuery {
 [DllImport("shlwapi.dll",CharSet=CharSet.Unicode)] static extern int AssocQueryString(uint f,uint k,string a,string e,StringBuilder b,ref uint s);
 public static string Read(string extension) { uint n=32768;var b=new StringBuilder((int)n);return AssocQueryString(0,1,extension,"open",b,ref n)==0?b.ToString():""; }
}
'@
 }
 $values=@{}
 foreach($ext in @('.doc','.docx','.docm','.dot','.dotx','.dotm','.rtf','.wps','.wpt','.odt','.xls','.xlsx','.xlsm','.xlt','.xltx','.et','.ett','.ppt','.pptx','.pptm','.pot','.potx','.pps','.ppsx','.dps','.dpt','.pdf','.ofd','.opg','.jpg','.png','.gif','.epub','.mobi','.dwg','.dxf','.eml','.msg','.zip','.rar','.7z','.pos','.posf','.posm','.xmind','.vsd','.vsdx')) { $values[$ext]=[WpsAssociationQuery]::Read($ext) }
 return $values
}
function Clear-WpsAssociationRequests($Root) {
 # WPS says an unchecked file-type box keeps its existing opening method.
 # Leave default-save-format checkboxes and unrelated options untouched.
 Invoke-WpsControl $Root '兼容设置'
 foreach($element in $Root.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)) {
  if ($element.Current.ControlType -ne [Windows.Automation.ControlType]::CheckBox) { continue }
  $name=$element.Current.Name
  if ($name -notmatch '\*\.' -or $name -match '保存|新建') { continue }
  if (!$element.Current.IsEnabled) { continue }
  $pattern=$null
  if (!$element.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern,[ref]$pattern)) { throw ('无法保留文件打开方式，未保存：'+$name) }
  if ([int]$pattern.Current.ToggleState -eq 2) { throw ('文件关联状态不明确，未保存：'+$name) }
  if ([int]$pattern.Current.ToggleState -eq 1) { $pattern.Toggle() }
  if ([int]$pattern.Current.ToggleState -ne 0) { throw ('无法取消文件关联请求，未保存：'+$name) }
 }
}
function Save-WpsRecord($Record,[string]$Path) { $Record | Export-Clixml -LiteralPath $Path -Depth 5 }
function Get-WpsRunningProcesses {
 $session=(Get-Process -Id $PID).SessionId
 return @(Get-Process -Name 'wps','wpp','et','wpspdf','WINWORD','ksomisc' -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session })
}
function Get-WpsVisibleWindows {
 # Enumerate all top-level windows: MainWindowHandle alone can miss secondary windows.
 if (!('WpsWindowInventory' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class WpsVisibleWindow { public int ProcessId; public string Title; public IntPtr Handle; }
public static class WpsWindowInventory {
 delegate bool Callback(IntPtr handle, IntPtr param);
 [DllImport("user32.dll")] static extern bool EnumWindows(Callback callback, IntPtr param);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr handle);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr handle, out uint id);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr handle, StringBuilder title, int max);
 [DllImport("user32.dll", SetLastError=true)] static extern bool PostMessage(IntPtr handle, uint message, IntPtr wparam, IntPtr lparam);
 public static bool RequestClose(IntPtr handle) { return PostMessage(handle,0x0010,IntPtr.Zero,IntPtr.Zero); }
 public static WpsVisibleWindow[] Read() {
  var list=new List<WpsVisibleWindow>();
  EnumWindows(delegate(IntPtr h, IntPtr unused) {
   if (IsWindowVisible(h)) {
    uint id; GetWindowThreadProcessId(h, out id);
    var title=new StringBuilder(1024); GetWindowText(h,title,title.Capacity);
    list.Add(new WpsVisibleWindow { ProcessId=(int)id, Title=title.ToString(), Handle=h });
   }
   return true;
  },IntPtr.Zero);
  return list.ToArray();
 }
}
'@
 }
 return [WpsWindowInventory]::Read()
}
function Assert-WpsWindowsClosed {
 $processes=@(Get-WpsRunningProcesses)
 $windows=@(Get-WpsVisibleWindows)
 $blocking=@(foreach ($process in $processes) {
  foreach ($window in $windows) {
   if ($window.ProcessId -eq $process.Id) {
    $title=$window.Title
    if ([string]::IsNullOrWhiteSpace($title)) { $title=$process.ProcessName }
    $title+'（PID '+$process.Id+'）'
   }
  }
 })
 if ($blocking.Count) {
  throw ('还有 '+$blocking.Count+' 个窗口未关闭，请保存后关闭：'+"`r`n"+($blocking -join "`r`n")+"`r`n"+'最小化窗口也需要关闭。无需结束后台进程；本工具不会强制关闭文档。')
 }
}
function Send-WpsWindowClose($Window) {
 if (![WpsWindowInventory]::RequestClose($Window.Handle)) { throw ('无法关闭窗口：'+$Window.Title+'。请手动关闭后重试。') }
}
function Exit-WpsConfiguration([int]$ProcessId) {
 $process=Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
 if (!$process) { return }
 $expected=Find-WpsConfig
 if (!$process.Path) { throw 'ELEVATION_REQUIRED: 无法访问 WPS 配置工具进程，请使用同一账号的管理员权限重试。' }
 if ($process.ProcessName -ne 'ksomisc' -or $process.Path -ine $expected) { throw '配置工具路径未通过检查，请手动关闭该窗口。' }
 Initialize-WpsBackground
 # The launcher can ignore WM_CLOSE. Use its own Cancel/Exit controls instead.
 foreach ($label in @('取消','退出')) {
  foreach ($handle in [WpsBackgroundWindows]::Handles($ProcessId)) {
   try {
    $root=[Windows.Automation.AutomationElement]::FromHandle($handle)
    if ($root.Current.Name -notmatch 'WPS.*配置工具') { continue }
    if (@(Get-WpsElements $root $label).Count -eq 1) { Invoke-WpsControl $root $label; break }
   } catch [Windows.Automation.ElementNotAvailableException] { continue }
  }
 }
 # Only the verified configuration executable may be terminated after a failed exit.
 # Document applications never enter this function.
 if (!$process.WaitForExit(2000)) {
  $process.Refresh()
  if (!$process.HasExited) { $process.Kill() }
  if (!$process.WaitForExit(3000)) { throw 'WPS 配置工具未能退出，请手动关闭后重试。' }
 }
}
function Close-WpsWindows([int]$TimeoutSeconds=30) {
 $processes=@(Get-WpsRunningProcesses)
 $windows=@(Get-WpsVisibleWindows)
 $configClosed=@()
 # Document windows retain their own save prompts. Configuration dialogs use Cancel/Exit.
 foreach ($window in $windows) {
  $owner=@($processes | Where-Object { $_.Id -eq $window.ProcessId })
  if (!$owner.Count) { continue }
  if ($owner[0].ProcessName -eq 'ksomisc') {
   if ($configClosed -notcontains $window.ProcessId) {
    Exit-WpsConfiguration $window.ProcessId
    $configClosed+=$window.ProcessId
   }
  } else { Send-WpsWindowClose $window }
 }
 $deadline=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
 do {
  try { Assert-WpsWindowsClosed; return } catch { $remaining=$_.Exception.Message }
  Start-Sleep -Milliseconds 250
  if ('Windows.Forms.Application' -as [type]) { [Windows.Forms.Application]::DoEvents() }
 } while ([DateTime]::UtcNow -lt $deadline)
 throw ('窗口尚未全部关闭，未开始修改设置。请处理保存提示，或手动关闭窗口后重试。'+"`r`n"+$remaining)
}
function Invoke-WpsRecommended([string]$ConfigPath,[string]$RecordPath,[switch]$CloseWindows) {
 if ($CloseWindows) { Close-WpsWindows }
 Assert-WpsWindowsClosed
 $root=$null; $saved=$false
 $record=New-Object PSObject -Property @{Time=(Get-Date);ConfigPath=$ConfigPath;Before=$null;After=$null;AssociationsBefore=$null;AssociationsAfter=$null;Status='尚未修改';Error=''}
 try {
  $record.AssociationsBefore=Get-WpsAssociationSnapshot
  $root=Open-WpsConfig $ConfigPath
  $record.Before=Read-WpsOptions $root
  if ($record.Before.Compatibility -eq 0 -and $record.Before.UpdatesOff -eq 1 -and (!$record.Before.OfflineAvailable -or $record.Before.Offline -eq 1)) {
   Invoke-WpsControl $root '取消'; $root=$null
   $record.After=$record.Before
   if ($record.Before.OfflineAvailable) { $record.Status='自动升级和第三方兼容已关闭，离线使用已开启，无需重复设置。' }
   else { $record.Status='自动升级和第三方兼容已关闭；此版本无离线选项，已跳过。' }
   Save-WpsRecord $record $RecordPath
   return $record
  }
  # Confirm save/cancel support before making any changes.
  foreach ($label in @('确定','取消')) {
   $control=Get-WpsControl $root $label; $p=$null
   if (!$control.Current.IsEnabled -or !$control.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern,[ref]$p)) { throw ('无法安全使用按钮：'+$label) }
  }
  Save-WpsRecord $record $RecordPath
  Invoke-WpsControl $root '兼容设置'
  Set-WpsToggle $root $script:wpsLabels.Compatibility 0
  Invoke-WpsControl $root '升级设置'
  Set-WpsUpdatesOff $root
  if ($record.Before.OfflineAvailable) {
   Invoke-WpsControl $root '其他选项'
   Set-WpsToggle $root $script:wpsLabels.Offline 1
  }
  Clear-WpsAssociationRequests $root
  # Mark uncertainty before save: an Invoke error may still have applied changes.
  $record.Status='已请求保存，待重新打开确认'
  Save-WpsRecord $record $RecordPath
  Invoke-WpsControl $root '确定'; $saved=$true; $root=$null
  Start-Sleep -Milliseconds 700
  $root=Open-WpsConfig $ConfigPath
  $record.After=Read-WpsOptions $root
  if ($record.After.Compatibility -ne 0 -or $record.After.UpdatesOff -ne 1) { throw '重新打开后设置与目标不一致，请根据原状态记录检查。' }
  if ($record.Before.OfflineAvailable -ne $record.After.OfflineAvailable) { throw '重新打开后离线选项的可用状态发生变化，未确认成功。' }
  if ($record.After.OfflineAvailable -and $record.After.Offline -ne 1) { throw '离线使用设置保存后未生效，未确认成功。' }
  Invoke-WpsControl $root '取消'; $root=$null
  $record.AssociationsAfter=Get-WpsAssociationSnapshot
  $changed=@($record.AssociationsBefore.Keys | Where-Object { $record.AssociationsBefore[$_] -ne $record.AssociationsAfter[$_] })
  if ($changed.Count) { throw ('检测到文件打开方式变化：'+($changed -join '、')+'。请在系统默认应用中核对；本次不标记为成功。') }
  if ($record.After.OfflineAvailable) { $record.Status='已关闭自动升级和第三方兼容，已开启离线使用；文件打开方式检查通过。' }
  else { $record.Status='已关闭自动升级和第三方兼容；无离线选项，已跳过。文件打开方式检查通过。' }
  Save-WpsRecord $record $RecordPath
  return $record
 } catch {
  $record.Error=$_.Exception.Message
  if ($root) { try { Invoke-WpsControl $root '取消' } catch {} }
  if ($saved) { $record.Status='已保存但复查未完成，请人工检查' }
  try { Save-WpsRecord $record $RecordPath } catch {}
  if (!$record.Before) { throw ('未能读取 WPS 设置，本次未修改配置。'+"`r`n"+$record.Error+"`r`n"+'诊断记录：'+$RecordPath) }
  throw ($record.Error+"`r`n"+'记录：'+$RecordPath+"`r`n"+'未确认成功时，请检查 WPS 配置；按记录中的 Before 原状态可手动恢复。')
 } finally { Stop-WpsBackground }
}
