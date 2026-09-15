param([string]$PreviewPath,[switch]$SmokeTest)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
try {
 . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'WordDefault.ps1') -LibraryOnly
 . (Join-Path $base 'Batch.ps1')
 if (!$isWin7 -and $PSVersionTable.PSVersion.Major -ge 5) {
  $vendorFile=Join-Path $base 'vendor\SFTA.ps1'
  $sha=[Security.Cryptography.SHA256]::Create()
  try { $hash=[BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($vendorFile))).Replace('-','') }
  finally { $sha.Dispose() }
  if ($hash -ne 'B8A1327E8DE2DB110A1674CABFEA965745CFE7F191414516054DA5C6E00B4FA4') { throw '批量设置组件校验失败，请重新解压完整工具包。' }
  . $vendorFile
 }
 $form=New-Object Windows.Forms.Form
 $form.Text='Word 默认打开工具 · 一键版 2.2'
 $form.Size=New-Object Drawing.Size(940,690)
 $form.MinimumSize=New-Object Drawing.Size(940,690)
 $form.StartPosition='CenterScreen'
 $form.Font=New-Object Drawing.Font('Microsoft YaHei',10)
 $form.BackColor=[Drawing.Color]::White
 $form.AutoScaleMode='Dpi'

 function Label-At($text,$x,$y,$w,$h) {
  $control=New-Object Windows.Forms.Label
  $control.Text=$text; $control.Location=New-Object Drawing.Point($x,$y)
  $control.Size=New-Object Drawing.Size($w,$h)
  $form.Controls.Add($control)
  return $control
 }
 $title=Label-At '让文档默认使用 Microsoft Word 打开' 24 20 860 36
 $title.Font=New-Object Drawing.Font('Microsoft YaHei',17,[Drawing.FontStyle]::Bold)
 $intro=Label-At '① 自动检测    →    ② 点击“一键设为 Word”    →    ③ 自动复查并显示结果' 26 65 860 28
 $system=Label-At ('当前用户：'+$env:USERNAME+'    系统：'+$os.Caption) 26 98 860 24
 $wordInfo=Label-At '' 26 126 860 42

 $grid=New-Object Windows.Forms.DataGridView
 $grid.Location=New-Object Drawing.Point(26,176)
 $grid.Size=New-Object Drawing.Size(872,278)
 $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false
 $grid.AllowUserToResizeRows=$false; $grid.RowHeadersVisible=$false
 $grid.SelectionMode='FullRowSelect'; $grid.MultiSelect=$false
 $grid.BackgroundColor=[Drawing.Color]::White
 $grid.AutoSizeColumnsMode='Fill'; $grid.RowTemplate.Height=28
 $grid.ColumnHeadersHeight=32
 $col=New-Object Windows.Forms.DataGridViewCheckBoxColumn
 $col.Name='Selected'; $col.HeaderText='选择'; $col.FillWeight=45
 [void]$grid.Columns.Add($col)
 foreach ($spec in @(@('Extension','格式',65),@('Description','文件类型',140),@('Status','当前状态',170),@('Executable','当前打开程序（完整路径）',365))) {
  $col=New-Object Windows.Forms.DataGridViewTextBoxColumn
  $col.Name=$spec[0]; $col.HeaderText=$spec[1]; $col.FillWeight=$spec[2]; $col.ReadOnly=$true
  [void]$grid.Columns.Add($col)
 }
 foreach ($col in $grid.Columns) { $col.SortMode='NotSortable' }
 $form.Controls.Add($grid)
 $tip=Label-At '' 26 467 872 54
 $tip.ForeColor=[Drawing.Color]::FromArgb(35,75,120)
 $tip.Text='点击后自动保存原状态、批量设置并复查，无需逐项打开系统设置。已指向 Word 的格式会自动跳过。'
 if (([version]$os.Version).Build -ge 22000) { $tip.Text='Win11：已启用一键设置尝试。系统版本或策略可能阻止修改；完成后逐项复查，未成功的格式会明确列出。' }

 function Button-At($text,$x,$width) {
  $b=New-Object Windows.Forms.Button; $b.Text=$text
  $b.Location=New-Object Drawing.Point($x,528); $b.Size=New-Object Drawing.Size($width,38)
  $form.Controls.Add($b); return $b
 }
 $setButton=Button-At '一键设为 Word' 26 142
 $setButton.BackColor=[Drawing.Color]::FromArgb(35,105,185); $setButton.ForeColor=[Drawing.Color]::White
 $refreshButton=Button-At '重新检测' 180 128
 $saveButton=Button-At '保存检测记录' 320 145
 $settingsButton=Button-At '系统默认应用' 477 145
 $selectButton=Button-At '全选 / 全不选' 634 142
 $closeButton=Button-At '关闭' 788 110
 $status=Label-At '' 26 578 872 28
 $foot=Label-At '.asd 为恢复文件，关联成功不代表恢复成功。此工具不修改宏设置；PLM 签名需单独验证。' 26 610 872 26
 $foot.Font=New-Object Drawing.Font('Microsoft YaHei',9)
 $foot.ForeColor=[Drawing.Color]::DimGray
 $descriptions=@{'.asd'='自动恢复文件';'.doc'='Word 文档';'.docm'='含宏文档';'.docx'='Word 文档';'.dot'='Word 模板';'.dotm'='含宏模板';'.dotx'='Word 模板';'.odt'='开放文档文本'}
 function Notice([string]$message) {
  if ($SmokeTest) { $script:smokeNotices += $message; return }
  [void][Windows.Forms.MessageBox]::Show($form,$message,'操作提示',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information)
 }
 function Refresh-View {
  $script:wordPaths=@(Find-Word)
  if ($script:wordPaths.Count) { $wordInfo.Text='已找到 Word：'+$script:wordPaths[0] }
  else { $wordInfo.Text='未找到 Word。请客户 IT 确认本机已安装 Microsoft Word。' }
  $setButton.Enabled=($script:wordPaths.Count -gt 0)
  $grid.Rows.Clear(); $count=0
  foreach ($row in @(Get-DefaultRows)) {
   $good=$row.Status -eq 'Word（系统关联查询）'
   $display=$row.Status
   if ($good) { $display='已指向 Word'; $count++ }
   elseif ($row.Executable -match '(?i)\\OpenWith\.exe$') { $display='尚未明确默认程序' }
   $index=$grid.Rows.Add([object[]]@($true,$row.Extension,$descriptions[$row.Extension],$display,$row.Executable))
   if ($good) { $grid.Rows[$index].Cells['Status'].Style.ForeColor=[Drawing.Color]::DarkGreen }
   else { $grid.Rows[$index].Cells['Status'].Style.ForeColor=[Drawing.Color]::FromArgb(165,80,15) }
  }
  $status.Text=('检测完成：'+$count+' / 8 种格式指向 Word。默认已选择全部八种格式。')
 }
 $refreshButton.Add_Click({ try { Refresh-View } catch { Notice ('检测未完成：'+$_.Exception.Message) } })
 $saveButton.Add_Click({ try { $reportPath=Save-Report '窗口版检测'; Notice ('已保存检测记录：'+"`r`n"+$reportPath) } catch { Notice ('保存失败：'+$_.Exception.Message) } })
 $settingsButton.Add_Click({ try { Open-Settings } catch { Notice ('无法打开系统设置：'+$_.Exception.Message) } })
 $selectButton.Add_Click({
  [void]$grid.EndEdit(); $all=$true
  foreach ($row in $grid.Rows) { if (!$row.Cells['Selected'].Value) { $all=$false } }
  foreach ($row in $grid.Rows) { $row.Cells['Selected'].Value=!$all }
 })
 $closeButton.Add_Click({ $form.Close() })
 $setButton.Add_Click({
  try {
   [void]$grid.EndEdit()
   $selected=@($grid.Rows | Where-Object { $_.Cells['Selected'].Value } | ForEach-Object { [string]$_.Cells['Extension'].Value })
   if (!$selected.Count) { Notice '请先勾选需要设为 Word 的格式。'; return }
   $setButton.Enabled=$false; $form.UseWaitCursor=$true
   $status.Text='正在保存原状态、设置默认程序并复查，请稍候…'; $form.Refresh()
   $results=@(Invoke-WordBatch $selected)
   Refresh-View
   $failed=@($results | Where-Object { !$_.Success })
   $passed=$results.Count-$failed.Count
   if ($failed.Count -eq 0) { $summary='设置成功：所选 '+$results.Count+' 种格式均已指向 Word。' }
   else { $summary='处理完成：'+$passed+' 项已指向 Word，'+$failed.Count+' 项未成功。' }
   $status.Text=$summary
   $details=($failed | ForEach-Object { $_.Extension+'：'+$_.Detail }) -join "`r`n"
   try {
    $reportPath=Save-Report '一键设置后'
    if ([string]::IsNullOrEmpty($reportPath)) { throw '结果记录路径为空。' }
    $results | Export-Clixml -LiteralPath ([IO.Path]::ChangeExtension($reportPath,'.xml'))
   } catch { $details += "`r`n"+'结果记录未保存：'+$_.Exception.Message }
   $form.UseWaitCursor=$false
   Notice ($summary+"`r`n"+$details+"`r`n"+'PLM 签名、保存和重新打开仍需单独验证。')
  } catch { $status.Text='操作未完成，未确认全部成功。'; Notice ('操作未完成：'+$_.Exception.Message) }
  finally { $form.UseWaitCursor=$false; $setButton.Enabled=($script:wordPaths.Count -gt 0) }
 })
 Refresh-View
 if ($SmokeTest) {
  # Actual button handlers and reports; intercept only association writes.
  $script:smokeNotices=@(); $script:smokeWrites=@()
  function Write-WordAssociation([string]$Extension,[string]$ProgId) { $script:smokeWrites += $Extension }
  $form.Show(); [Windows.Forms.Application]::DoEvents()
  $saveButton.PerformClick()
  if (!(Test-Path -LiteralPath $script:lastReportPath)) { throw 'Smoke: report button did not save a report.' }
  $setButton.PerformClick()
  if (@($script:smokeNotices | Where-Object { $_ -match '操作未完成|保存失败|结果记录未保存' }).Count) { throw ($script:smokeNotices -join "`r`n") }
  if (!(Test-Path -LiteralPath $script:batchSnapshot)) { throw 'Smoke: pre-change snapshot missing.' }
  $resultFile=[IO.Path]::ChangeExtension($script:lastReportPath,'.xml')
  $smokeResults=@(Import-Clixml -LiteralPath $resultFile)
  if ($smokeResults.Count -ne 8) { throw 'Smoke: expected eight final results.' }
  Write-Output ('PASS: real GUI save + batch buttons, real before/after TXT/XML, independent verification; intercepted writes='+$script:smokeWrites.Count)
  $form.Close()
 } elseif ($PreviewPath) {
  $form.Show(); [Windows.Forms.Application]::DoEvents()
  $bitmap=New-Object Drawing.Bitmap($form.Width,$form.Height)
  $form.DrawToBitmap($bitmap,(New-Object Drawing.Rectangle(0,0,$form.Width,$form.Height)))
  $bitmap.Save($PreviewPath,[Drawing.Imaging.ImageFormat]::Png)
  $bitmap.Dispose(); $form.Close()
 } else { [void]$form.ShowDialog() }
 $form.Dispose()
} catch {
 if ($SmokeTest) { throw }
 [void][Windows.Forms.MessageBox]::Show(('工具启动失败：'+$_.Exception.Message+"`r`n"+'请把错误内容发给技术人员。'),'启动失败')
 exit 1
}
