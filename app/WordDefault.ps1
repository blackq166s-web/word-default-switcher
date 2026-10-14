param([switch]$DetectOnly, [switch]$LibraryOnly)
$ErrorActionPreference = 'Stop'
$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$extensions = @('.asd','.doc','.docm','.docx','.dot','.dotm','.dotx','.odt')
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
[ComImport, Guid("4e530b0a-e611-4c77-a3ac-9031d022281b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAssocRegistration {
 [PreserveSig] int QueryCurrentDefault([MarshalAs(UnmanagedType.LPWStr)] string q, int type, int level, [MarshalAs(UnmanagedType.LPWStr)] out string result);
 [PreserveSig] int QueryAppIsDefault([MarshalAs(UnmanagedType.LPWStr)] string q, int type, int level, [MarshalAs(UnmanagedType.LPWStr)] string app, [MarshalAs(UnmanagedType.Bool)] out bool result);
 [PreserveSig] int QueryAppIsDefaultAll(int level, [MarshalAs(UnmanagedType.LPWStr)] string app, [MarshalAs(UnmanagedType.Bool)] out bool result);
 [PreserveSig] int SetAppAsDefault([MarshalAs(UnmanagedType.LPWStr)] string app, [MarshalAs(UnmanagedType.LPWStr)] string ext, int type);
}
public static class WordDefaultNative {
 [DllImport("shlwapi.dll", CharSet=CharSet.Unicode)]
 static extern int AssocQueryString(uint flags, uint kind, string assoc, string extra, StringBuilder output, ref uint size);
 public static string Query(string assoc, uint kind) {
  uint size=32768; StringBuilder b=new StringBuilder((int)size);
  int hr=AssocQueryString(0,kind,assoc,"open",b,ref size);
  return hr==0 ? b.ToString() : "";
 }
 public static int SetDefault(string app, string ext) {
  object instance=Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("591209c7-767b-42b2-9fba-44ee4615f2c7")));
  try { return ((IAssocRegistration)instance).SetAppAsDefault(app,ext,0); }
  finally { Marshal.ReleaseComObject(instance); }
 }
 [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
 public struct OpenInfo { [MarshalAs(UnmanagedType.LPWStr)] public string file; [MarshalAs(UnmanagedType.LPWStr)] public string type; public uint flags; }
 [DllImport("shell32.dll", CharSet=CharSet.Unicode)]
 static extern int SHOpenWithDialog(IntPtr parent, ref OpenInfo info);
 public static int Choose(string ext) {
  OpenInfo info=new OpenInfo(); info.file="document"+ext; info.flags=2|8;
  return SHOpenWithDialog(IntPtr.Zero,ref info);
 }
}
'@

function Get-DefaultRows {
 foreach ($ext in $script:extensions) {
  $exe = [WordDefaultNative]::Query($ext,2)
  $command = [WordDefaultNative]::Query($ext,1)
  $prog = ''
  $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+$ext+'\UserChoice'
  if (Test-Path $key) { $prog = (Get-ItemProperty $key -ErrorAction SilentlyContinue).ProgId }
  $latestProg=''
  $latestKey='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\'+$ext+'\UserChoiceLatest\ProgId'
  if (Test-Path $latestKey) { $latestProg=(Get-ItemProperty $latestKey -ErrorAction SilentlyContinue).ProgId }
  $status = '未能确认 / 未设置'
  if ($exe) {
   $leaf = [IO.Path]::GetFileName($exe.Trim('"'))
   if ($leaf -ieq 'WINWORD.EXE') {
    if (Test-Path -LiteralPath $exe.Trim('"')) { $status = 'Word（系统关联查询）' }
    else { $status = 'Word 路径已失效' }
   } elseif ($leaf -ieq 'wps.exe') { $status = 'WPS' }
   else { $status = '其他程序 / 需核对' }
  }
  New-Object PSObject -Property @{ Extension=$ext; Status=$status; Executable=$exe; Command=$command; UserChoice=$prog; LatestUserChoice=$latestProg }
 }
}

function Find-Word {
 $paths = @()
 foreach ($hive in @('HKCU:','HKLM:')) {
  foreach ($part in @('Software\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE','Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE')) {
   $key = $hive+'\'+$part
   if (Test-Path $key) { $paths += (Get-Item $key).GetValue('') }
  }
 }
 foreach ($id in @('Word.Document.8','Word.Document.12','Word.Template.8')) { $paths += [WordDefaultNative]::Query($id,2) }
 foreach ($root in @($env:ProgramFiles,${env:ProgramFiles(x86)})) {
  if ($root) {
   foreach ($pattern in @('\Microsoft Office\Office*\WINWORD.EXE','\Microsoft Office\root\Office*\WINWORD.EXE')) {
    $paths += @(Get-Item -Path ($root+$pattern) -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
   }
  }
 }
 @($paths | Where-Object { $_ -and ([IO.Path]::GetFileName($_.Trim('"')) -ieq 'WINWORD.EXE') -and (Test-Path -LiteralPath $_.Trim('"')) } | ForEach-Object { $_.Trim('"') } | Select-Object -Unique)
}

function Find-RegisteredWord([string]$extension) {
 foreach ($hive in @('HKCU:','HKLM:')) {
  foreach ($part in @('Software\RegisteredApplications','Software\WOW6432Node\RegisteredApplications')) {
   $path = $hive+'\'+$part
   if (!(Test-Path $path)) { continue }
   $key = Get-Item $path
   foreach ($name in $key.GetValueNames()) {
    $cap = $hive+'\'+$key.GetValue($name)+'\FileAssociations'
    if (!(Test-Path $cap)) { continue }
    $values = Get-Item $cap
    $valid = $true
    foreach ($ext in @($extension)) {
     $id = [string]$values.GetValue($ext)
     if (!$id) { $valid=$false; break }
     $exe = [WordDefaultNative]::Query($id,2)
     if (!$exe -or [IO.Path]::GetFileName($exe.Trim('"')) -ine 'WINWORD.EXE' -or !(Test-Path -LiteralPath $exe.Trim('"'))) { $valid=$false; break }
    }
    if ($valid) { $name }
   }
  }
 }
}

function Show-Status {
 Write-Host ''
 Get-DefaultRows | Format-List Extension,Status,Executable,UserChoice,Command | Out-Host
 Write-Host '检测到的 Word 安装路径：'
 if ($script:wordPaths.Count -gt 0) { $script:wordPaths | ForEach-Object { Write-Host $_ } }
 else { Write-Host '未检测到。可能未安装或注册异常，请先人工核实。' }
}

function Save-Report([string]$label) {
 $dir = Join-Path $base '检测记录'
 try { if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null } }
 catch { $dir=Join-Path $env:TEMP 'WordDefaultReports'; if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null } }
 $file=Join-Path $dir ((Get-Date -Format 'yyyyMMdd-HHmmss-fff')+'-'+$label+'.txt')
 $lines=@()
 $lines += '电脑：'+$env:COMPUTERNAME
 $lines += '当前用户：'+$env:USERNAME
 $lines += '系统：'+$script:os.Caption+' '+$script:os.Version
 $lines += 'Word 路径：'+($script:wordPaths -join '；')
 foreach ($row in @(Get-DefaultRows)) {
  $lines += ''
  $lines += '格式：'+$row.Extension+'；状态：'+$row.Status
  $lines += '打开程序：'+$row.Executable
  $lines += '旧机制关联标识：'+$row.UserChoice
  $lines += '新机制关联标识：'+$row.LatestUserChoice
  $lines += '打开命令：'+$row.Command
 }
 $lines | Out-File -LiteralPath $file -Encoding UTF8
 $script:lastReportPath=$file
 Write-Host ('检测记录：'+$file)
 return $file
}

function Get-WordSettingsUri {
 $uri='ms-settings:defaultapps'
 if (([version]$script:os.Version).Build -lt 22000) { return $uri }
 foreach ($scope in @(@('HKLM:','registeredAppMachine'),@('HKCU:','registeredAppUser'))) {
  $path=$scope[0]+'\Software\RegisteredApplications'
  if (!(Test-Path $path)) { continue }
  $key=Get-Item $path
  foreach ($name in $key.GetValueNames()) {
   if ($name -notmatch '^Word\.Application(\.\d+)?$') { continue }
   $cap=$scope[0]+'\'+$key.GetValue($name)+'\FileAssociations'
   if ((Test-Path $cap) -and (Get-Item $cap).GetValue('.docx') -eq 'Word.Document.12') {
    return ($uri+'?'+$scope[1]+'='+[Uri]::EscapeDataString($name))
   }
  }
 }
 return $uri
}

function Open-Settings {
 if ($script:isWin7) {
  Start-Process control.exe -ArgumentList '/name Microsoft.DefaultPrograms /page pageFileAssoc'
 } else {
  Start-Process (Get-WordSettingsUri)
 }
}

$os = Get-WmiObject Win32_OperatingSystem
$isWin7 = ([version]$os.Version).Major -eq 6 -and ([version]$os.Version).Minor -eq 1
$wordPaths = @(Find-Word)
if ($LibraryOnly) { return }
Write-Host ('Word 默认打开工具 - '+($extensions -join ' / '))
Write-Host '.asd 为自动恢复文件；关联到 Word 不代表能直接打开或恢复。'
Write-Host ('系统：'+$os.Caption+'；当前用户：'+$env:USERNAME)
Write-Host '请以客户平时使用 PLM 的 Windows 账号运行，不要换管理员账号。'
Show-Status
if ($DetectOnly) { return }
Save-Report '操作前'
while ($true) {
 Write-Host ''
 Write-Host '1 = 切换到 Word（Win7 尝试自动；Win10 打开系统设置）'
 Write-Host '2 = 重新检测并保存结果'
 Write-Host '3 = 打开系统默认程序设置（也可手动恢复原程序）'
 Write-Host '0 = 退出'
 $choice=Read-Host '请输入数字'
 if ($choice -eq '0') { break }
 try {
  if ($choice -eq '2') { $wordPaths=@(Find-Word); Show-Status; Save-Report '复查' }
  elseif ($choice -eq '3') { Open-Settings }
  elseif ($choice -eq '1') {
   $wordPaths=@(Find-Word)
   if ($wordPaths.Count -eq 0) { Write-Host '未检测到 Word，停止自动设置。请先核实安装情况。'; continue }
   Save-Report '切换前'
   if ($isWin7) {
    foreach ($ext in $extensions) {
     $apps=@(Find-RegisteredWord $ext | Select-Object -Unique)
     if ($apps.Count -eq 1) {
      try {
       $hr=[WordDefaultNative]::SetDefault($apps[0],$ext)
       Write-Host ($ext+' 系统设置接口返回：'+$hr+'（以复查结果为准）')
      } catch { Write-Host ($ext+' 自动设置接口不可用，将进入系统选择窗口。') }
     } else { Write-Host ($ext+' 未找到唯一可确认的 Word 注册项，请在系统窗口中选择 Microsoft Word。') }
    }
    foreach ($row in @(Get-DefaultRows)) {
     if ($row.Status -ne 'Word（系统关联查询）') {
      Write-Host ('请为 '+$row.Extension+' 选择 Microsoft Word；没有列出时，浏览到上方 WINWORD.EXE 路径。')
      [void][WordDefaultNative]::Choose($row.Extension)
     }
    }
   } else {
    Write-Host ('在系统设置中找到“按文件类型指定默认应用”，分别将 '+($extensions -join '、')+' 设为 Microsoft Word。')
    Write-Host '如系统支持搜索，可逐个搜索扩展名。某项未列出或没有 Word 时先跳过，继续其余项；完成后回本窗口复查。'
    Open-Settings
    [void](Read-Host '完成选择后按回车复查；未选择也可回车，不会报告为成功')
   }
   Show-Status
   Save-Report '切换后'
   $remaining=@(Get-DefaultRows | Where-Object { $_.Status -ne 'Word（系统关联查询）' })
   if ($remaining.Count -eq 0) { Write-Host '系统关联查询：八种扩展名均指向 Word。请用普通测试文档验证，并单独验证 PLM 签字、保存、重开；.asd 恢复需另行验证。' }
   else { Write-Host '尚未确认全部切换成功。请查看上方结果，使用选项 3 手动设置。' }
  }
 } catch { Write-Host ('操作未完成：'+$_.Exception.Message); Write-Host '可选择 3 进入系统设置，或保留检测记录进一步排查。' }
}
