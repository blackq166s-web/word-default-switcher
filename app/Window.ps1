param([string]$PreviewPath,[switch]$SmokeTest,[switch]$BackgroundLayoutTest)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
try {
 . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'WordDefault.ps1') -LibraryOnly
 . (Join-Path $base 'Batch.ps1')
 . (Join-Path $base 'WpsWorkerBridge.ps1')
 if (!$isWin7 -and $PSVersionTable.PSVersion.Major -ge 5) {
  $vendorFile=Join-Path $base 'vendor\SFTA.ps1'
  $sha=[Security.Cryptography.SHA256]::Create()
  try { $hash=[BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes($vendorFile))).Replace('-','') }
  finally { $sha.Dispose() }
  if ($hash -ne 'B8A1327E8DE2DB110A1674CABFEA965745CFE7F191414516054DA5C6E00B4FA4') { throw '批量设置组件校验失败，请重新解压完整工具包。' }
  . $vendorFile
 }
 $form=New-Object Windows.Forms.Form
 $form.Text='Word 与 WPS 设置'
 $form.Add_FormClosing({ param($sender,$eventArgs) if ($script:wpsWorkerBusy) { $eventArgs.Cancel=$true } })
 $form.ClientSize=New-Object Drawing.Size(660,410)
 $form.FormBorderStyle='FixedDialog'
 $form.MaximizeBox=$false
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
 $intro=Label-At '① 自动检测    →    ② 点击“设为微软 Word 打开”    →    ③ 自动复查并显示结果' 26 65 860 28
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
 $setButton=Button-At '设为微软 Word 打开' 26 142
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
  param([object[]]$VerifiedRows)
  $script:wordPaths=@(Find-Word)
  if ($script:wordPaths.Count) { $wordInfo.Text='PLM 等公司内部系统要求用 Word 打开文档时使用。'+"`r`n"+'已检测到 Microsoft Word，可以设置。' }
  else { $wordInfo.Text='未找到 Word。请客户 IT 确认本机已安装 Microsoft Word。' }
  $setButton.Enabled=($script:wordPaths.Count -gt 0)
  $grid.Rows.Clear(); $count=0
  if (!$PSBoundParameters.ContainsKey('VerifiedRows')) { $VerifiedRows=@(Get-DefaultRows) }
  foreach ($row in $VerifiedRows) {
   $good=$row.Status -eq 'Word（系统关联查询）'
   $display=$row.Status
   if ($good) { $display='已指向 Word'; $count++ }
   elseif ($row.Executable -match '(?i)\\OpenWith\.exe$') { $display='尚未明确默认程序' }
   $index=$grid.Rows.Add([object[]]@($true,$row.Extension,$descriptions[$row.Extension],$display,$row.Executable))
   if ($good) { $grid.Rows[$index].Cells['Status'].Style.ForeColor=[Drawing.Color]::DarkGreen }
   else { $grid.Rows[$index].Cells['Status'].Style.ForeColor=[Drawing.Color]::FromArgb(165,80,15) }
  }
  if ($count -eq 8) { $status.Text='8 / 8 种文档格式已指向 Word。' }
  else { $status.Text=($count.ToString()+' / 8 种文档格式已指向 Word，其余格式可通过上方按钮设置。') }
 }
 $refreshButton.Add_Click({ try { Refresh-View } catch { Notice ('检测未完成：'+$_.Exception.Message) } })
 $saveButton.Add_Click({ try { $reportPath=Save-Report '窗口版检测'; Notice ('已保存检测记录：'+"`r`n"+$reportPath) } catch { Notice ('保存失败：'+$_.Exception.Message) } })
 function Open-WordSettingsAndWatch {
  $script:settingsPending=$true; $script:settingsLeft=$false
  try { Open-Settings } catch { $script:settingsPending=$false; throw }
  $status.Text='请在系统设置中将 .docx 等未成功格式选择为 Word；返回本窗口后自动复查。'
 }
 $form.Add_Deactivate({ if ($script:settingsPending) { $script:settingsLeft=$true } })
 $form.Add_Activated({
  if (!$script:settingsPending -or !$script:settingsLeft -or $script:wpsWorkerBusy) { return }
  $script:settingsPending=$false
  try {
   $form.UseWaitCursor=$true
   Refresh-View -VerifiedRows @(Read-VerifiedRows)
  } catch { $status.Text='复查未完成，请点击重新检测。' }
  finally { $form.UseWaitCursor=$false }
 })
 $settingsButton.Add_Click({ try { Open-WordSettingsAndWatch } catch { Notice ('无法打开系统设置：'+$_.Exception.Message) } })
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
   if ($failed.Count -gt 0 -and !$SmokeTest) {
    $answer=[Windows.Forms.MessageBox]::Show($form,($summary+"`r`n"+$details+"`r`n`r`n"+'这些格式未能自动修改。打开后找到上面未成功的格式（如 .docx），点击当前应用，选择 Microsoft Word 并确认。返回本窗口会自动复查。现在打开 Word 默认应用设置？'),'部分格式需要手动设置',[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Information)
    if ($answer -eq [Windows.Forms.DialogResult]::Yes) { Open-WordSettingsAndWatch }
   } else { Notice ($summary+"`r`n"+$details+"`r`n"+'PLM 签名、保存和重新打开仍需单独验证。') }
  } catch { $status.Text='操作未完成，未确认全部成功。'; Notice ('操作未完成：'+$_.Exception.Message) }
  finally { $form.UseWaitCursor=$false; $setButton.Enabled=($script:wordPaths.Count -gt 0) }
 })
 Refresh-View
 $wpsTitle=Label-At 'WPS 共存配置（业务系统使用 Microsoft Word）' 26 650 872 25
 $wpsHint=Label-At '推荐：关闭第三方兼容；开启离线未登录使用。自动操作仍需客户版本验证。' 26 680 872 25
 $wpsButton=Button-At '应用 WPS 推荐设置' 26 205
 $wpsButton.Location=New-Object Drawing.Point(26,716)
 $wpsManual=Button-At '打开 WPS 配置工具' 246 205
 $wpsManual.Location=New-Object Drawing.Point(246,716)
 $wpsManual.Add_Click({ try { $path=Find-WpsConfig; Start-Process -FilePath $path } catch { Notice $_.Exception.Message } })
 function Show-WpsClosePrompt([string]$message) {
  $dialog=New-Object Windows.Forms.Form
  $dialog.Text='请先关闭相关窗口'
  $dialog.ClientSize=New-Object Drawing.Size(580,330)
  $dialog.StartPosition='CenterParent'; $dialog.FormBorderStyle='FixedDialog'
  $dialog.MaximizeBox=$false; $dialog.MinimizeBox=$false
  $dialog.Font=$form.Font; $dialog.BackColor=[Drawing.Color]::White
  $heading=New-Object Windows.Forms.Label
  $heading.Text='以下窗口尚未关闭，暂时无法优化 WPS 设置。'
  $heading.Location=New-Object Drawing.Point(22,20); $heading.Size=New-Object Drawing.Size(536,30)
  $dialog.Controls.Add($heading)
  $list=New-Object Windows.Forms.TextBox
  $list.Multiline=$true; $list.ReadOnly=$true; $list.ScrollBars='Vertical'
  $list.Text=$message; $list.BackColor=[Drawing.Color]::White
  $list.Location=New-Object Drawing.Point(22,58); $list.Size=New-Object Drawing.Size(536,200)
  $dialog.Controls.Add($list)
  $autoClose=New-Object Windows.Forms.Button
  $autoClose.Text='一键关闭并优化'; $autoClose.Location=New-Object Drawing.Point(248,278)
  $autoClose.Size=New-Object Drawing.Size(180,36)
  $autoClose.BackColor=[Drawing.ColorTranslator]::FromHtml('#245EDB'); $autoClose.ForeColor=[Drawing.Color]::White
  $autoClose.FlatStyle='Flat'; $autoClose.FlatAppearance.BorderSize=0
  $autoClose.DialogResult=[Windows.Forms.DialogResult]::Yes
  $dialog.Controls.Add($autoClose)
  $cancel=New-Object Windows.Forms.Button
  $cancel.Text='暂不处理'; $cancel.Location=New-Object Drawing.Point(442,278)
  $cancel.Size=New-Object Drawing.Size(116,36); $cancel.DialogResult=[Windows.Forms.DialogResult]::Cancel
  $dialog.Controls.Add($cancel); $dialog.CancelButton=$cancel; $dialog.ActiveControl=$cancel
  try { return ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::Yes) }
  finally { $dialog.Dispose() }
 }
 $wpsButton.Add_Click({
  if ($script:wpsWorkerBusy) { return }
  $closeRequested=$false
  if (!$SmokeTest) {
   try { Assert-WpsWindowsClosed }
   catch {
    if (!(Show-WpsClosePrompt $_.Exception.Message)) { return }
    $answer=[Windows.Forms.MessageBox]::Show($form,('即将关闭 Word、WPS 和 WPS 配置工具窗口，然后优化 WPS 设置。'+"`r`n`r`n"+'如有未保存文档，请在软件弹出的提示中保存。窗口全部关闭后才会继续；不会强制关闭文档；无响应的配置工具会自动退出。'+"`r`n`r`n"+'是否继续？'),'确认关闭并优化',[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Question,[Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
    $closeRequested=$true
   }
  }
  try {
   $wpsButton.Enabled=$false; $setButton.Enabled=$false; $detailsButton.Enabled=$false; $closeButton.Enabled=$false; $form.UseWaitCursor=$true
   if ($closeRequested) { $status.Text='正在关闭窗口；请处理保存提示，随后自动优化…' }
   else { $status.Text='正在后台优化 WPS 设置，请稍候…' }
   $form.Refresh()
   $path=Find-WpsConfig
   $report=Save-Report 'WPS配置前'
   $recordPath=[IO.Path]::ChangeExtension($report,'.wps.xml')
   $result=Invoke-WpsRecommended $path $recordPath -CloseWindows:$closeRequested
   Refresh-View
    $status.Text=$result.Status
  } catch { $status.Text='WPS 配置未确认完成。'; Notice $_.Exception.Message }
  finally { $wpsButton.Enabled=$true; $setButton.Enabled=($script:wordPaths.Count -gt 0); $detailsButton.Enabled=$true; $closeButton.Enabled=$true; $form.UseWaitCursor=$false }
 })
 # Keep the customer page short; move diagnostics into an optional dialog.
 $detailsForm=New-Object Windows.Forms.Form
 $detailsForm.Text='检测详情与辅助工具'
 $detailsForm.ClientSize=New-Object Drawing.Size(920,590)
 $detailsForm.StartPosition='CenterParent'
 $detailsForm.Font=$form.Font
 $detailsForm.BackColor=[Drawing.Color]::White
 $detailsForm.FormBorderStyle='FixedDialog'; $detailsForm.MaximizeBox=$false
 foreach ($control in @($system,$grid,$tip,$refreshButton,$saveButton,$settingsButton,$selectButton,$foot,$wpsManual)) {
  $detailsForm.Controls.Add($control)
 }
 $system.Location=New-Object Drawing.Point(24,20)
 $grid.Location=New-Object Drawing.Point(24,58)
 $tip.Location=New-Object Drawing.Point(24,350)
 $refreshButton.Location=New-Object Drawing.Point(24,414)
 $saveButton.Location=New-Object Drawing.Point(168,414)
 $settingsButton.Location=New-Object Drawing.Point(329,414)
 $selectButton.Location=New-Object Drawing.Point(490,414)
 $wpsManual.Location=New-Object Drawing.Point(648,414)
 $foot.Location=New-Object Drawing.Point(24,476)
 $detailsClose=New-Object Windows.Forms.Button
 $detailsClose.Text='返回'; $detailsClose.Location=New-Object Drawing.Point(788,532)
 $detailsClose.Size=New-Object Drawing.Size(110,36)
 $detailsClose.Add_Click({ $detailsForm.Hide() })
 $detailsForm.Controls.Add($detailsClose)
 $detailsForm.Add_FormClosing({ param($sender,$eventArgs) $eventArgs.Cancel=$true; $sender.Hide() })

 # Main page: two action cards, followed by a persistent detection summary.
 $form.ClientSize=New-Object Drawing.Size(760,560)
 $form.BackColor=[Drawing.ColorTranslator]::FromHtml('#F3F5F9')
 $form.ForeColor=[Drawing.ColorTranslator]::FromHtml('#172B45')
 $title.Text='WPS 一键优化'
 $title.Font=New-Object Drawing.Font('Microsoft YaHei',22,[Drawing.FontStyle]::Bold)
 $title.Location=New-Object Drawing.Point(30,24); $title.Size=New-Object Drawing.Size(520,44)
 $intro.Text='优化 WPS 使用设置，配置 Word 文档打开方式'
 $intro.ForeColor=[Drawing.ColorTranslator]::FromHtml('#65758B')
 $intro.Location=New-Object Drawing.Point(33,76); $intro.Size=New-Object Drawing.Size(670,25)
 function Card-At([int]$y,[int]$height) {
  $panel=New-Object Windows.Forms.Panel
  $panel.Location=New-Object Drawing.Point(30,$y)
  $panel.Size=New-Object Drawing.Size(700,$height)
  $panel.BackColor=[Drawing.Color]::White
  $panel.Add_Paint({ param($sender,$eventArgs)
   $pen=New-Object Drawing.Pen([Drawing.ColorTranslator]::FromHtml('#E1E7EF'))
   try { $eventArgs.Graphics.DrawRectangle($pen,0,0,$sender.Width-1,$sender.Height-1) } finally { $pen.Dispose() }
  })
  $form.Controls.Add($panel)
  return $panel
 }
 function Style-Action($button,[bool]$primary) {
  $button.FlatStyle='Flat'; $button.FlatAppearance.BorderSize=1
  $button.FlatAppearance.BorderColor=[Drawing.ColorTranslator]::FromHtml('#C8D4E5')
  $button.Font=New-Object Drawing.Font('Microsoft YaHei',10,[Drawing.FontStyle]::Bold)
  $button.Cursor=[Windows.Forms.Cursors]::Hand
  $button.Size=New-Object Drawing.Size(156,44)
  if ($primary) {
   $button.BackColor=[Drawing.ColorTranslator]::FromHtml('#245EDB')
   $button.ForeColor=[Drawing.Color]::White
   $button.FlatAppearance.BorderSize=0
   $button.FlatAppearance.MouseOverBackColor=[Drawing.ColorTranslator]::FromHtml('#1D4FBE')
  } else {
   $button.BackColor=[Drawing.Color]::White
   $button.ForeColor=[Drawing.ColorTranslator]::FromHtml('#245EDB')
   $button.FlatAppearance.MouseOverBackColor=[Drawing.ColorTranslator]::FromHtml('#F0F5FF')
  }
 }
 $wpsCard=Card-At 122 126
 foreach($control in @($wpsTitle,$wpsHint,$wpsButton)) { $wpsCard.Controls.Add($control) }
 $wpsTitle.Text='优化 WPS 设置'
 $wpsTitle.Font=New-Object Drawing.Font('Microsoft YaHei',13,[Drawing.FontStyle]::Bold)
 $wpsTitle.Location=New-Object Drawing.Point(24,20); $wpsTitle.Size=New-Object Drawing.Size(300,28)
 $wpsHint.Text='关闭自动升级及第三方兼容；支持时开启离线使用。'+"`r`n"+'保留文件打开方式；业务系统依赖 WPS 时请联系 IT。'
 $wpsHint.ForeColor=[Drawing.ColorTranslator]::FromHtml('#65758B')
 $wpsHint.Font=New-Object Drawing.Font('Microsoft YaHei',9)
 $wpsHint.Location=New-Object Drawing.Point(24,59); $wpsHint.Size=New-Object Drawing.Size(480,48)
 $wpsButton.Text='立即优化'; Style-Action $wpsButton $true
 $wpsButton.Location=New-Object Drawing.Point(520,40)
 $wordCard=Card-At 262 116
 $wordHeading=New-Object Windows.Forms.Label
 $wordHeading.Text='修改默认程序'
 $wordHeading.Font=New-Object Drawing.Font('Microsoft YaHei',13,[Drawing.FontStyle]::Bold)
 $wordHeading.Location=New-Object Drawing.Point(24,18); $wordHeading.Size=New-Object Drawing.Size(420,28)
 $wordCard.Controls.Add($wordHeading)
 foreach($control in @($wordInfo,$setButton)) { $wordCard.Controls.Add($control) }
 $wordInfo.Location=New-Object Drawing.Point(24,54); $wordInfo.Size=New-Object Drawing.Size(480,48)
 $wordInfo.Font=New-Object Drawing.Font('Microsoft YaHei',9)
 $wordInfo.ForeColor=[Drawing.ColorTranslator]::FromHtml('#65758B')
 $setButton.Text='设为微软 Word 打开'; Style-Action $setButton $false
 $setButton.Location=New-Object Drawing.Point(520,35)
 $detectionCard=Card-At 394 88
 $detectionHeading=New-Object Windows.Forms.Label
 $detectionHeading.Text='打开方式检测'
 $detectionHeading.Font=New-Object Drawing.Font('Microsoft YaHei',10,[Drawing.FontStyle]::Bold)
 $detectionHeading.Location=New-Object Drawing.Point(24,14); $detectionHeading.Size=New-Object Drawing.Size(300,24)
 $detectionCard.Controls.Add($detectionHeading)
 $detectionCard.Controls.Add($status)
 $status.Location=New-Object Drawing.Point(24,43); $status.Size=New-Object Drawing.Size(650,38)
 $status.Font=New-Object Drawing.Font('Microsoft YaHei',9)
 $status.ForeColor=[Drawing.ColorTranslator]::FromHtml('#52657E')
 $detailsButton=New-Object Windows.Forms.LinkLabel
 $detailsButton.Text='查看检测详情'; $detailsButton.AutoSize=$true
 $detailsButton.LinkBehavior='HoverUnderline'
 $detailsButton.LinkColor=[Drawing.ColorTranslator]::FromHtml('#245EDB')
 $detailsButton.ActiveLinkColor=$detailsButton.LinkColor
 $detailsButton.Location=New-Object Drawing.Point(548,15)
 $detailsButton.Add_LinkClicked({ [void]$detailsForm.ShowDialog($form) })
 $detectionCard.Controls.Add($detailsButton)
 $usageHint=Label-At '操作前请保存文档；若窗口未关闭，可按提示一键处理。' 32 510 550 26
 $usageHint.Font=New-Object Drawing.Font('Microsoft YaHei',9)
 $usageHint.ForeColor=[Drawing.ColorTranslator]::FromHtml('#65758B')
 Style-Action $closeButton $false
 $closeButton.Text='关闭'; $closeButton.Font=$usageHint.Font
 $closeButton.Size=New-Object Drawing.Size(86,34)
 $closeButton.Location=New-Object Drawing.Point(644,501)
 $wpsButton.TabIndex=0; $setButton.TabIndex=1; $detailsButton.TabIndex=2; $closeButton.TabIndex=3
 if ($SmokeTest) {
  # Actual button handlers and reports; intercept only association writes.
  $script:smokeNotices=@(); $script:smokeWrites=@()
  function Write-WordAssociation([string]$Extension,[string]$ProgId) { $script:smokeWrites += $Extension }
  $form.Show(); [Windows.Forms.Application]::DoEvents()
  $detailsForm.Show($form); [Windows.Forms.Application]::DoEvents()
  $saveButton.PerformClick()
  $detailsForm.Hide()
  if (!(Test-Path -LiteralPath $script:lastReportPath)) { throw 'Smoke: report button did not save a report.' }
  $setButton.PerformClick()
  if (@($script:smokeNotices | Where-Object { $_ -match '操作未完成|保存失败|结果记录未保存' }).Count) { throw ($script:smokeNotices -join "`r`n") }
  if (!(Test-Path -LiteralPath $script:batchSnapshot)) { throw 'Smoke: pre-change snapshot missing.' }
  $resultFile=[IO.Path]::ChangeExtension($script:lastReportPath,'.xml')
  $smokeResults=@(Import-Clixml -LiteralPath $resultFile)
  if ($smokeResults.Count -ne 8) { throw 'Smoke: expected eight final results.' }
  Write-Output ('PASS: real GUI save + batch buttons, real before/after TXT/XML, independent verification; intercepted writes='+$script:smokeWrites.Count)
  function Find-WpsConfig { return 'fake-WPS-config.exe' }
  function Invoke-WpsRecommended($ConfigPath,$RecordPath) {
   $script:wpsSmokeCalled=$true
   return New-Object PSObject -Property @{Status='TEST: WPS settings intercepted'}
  }
  $script:wpsSmokeCalled=$false
  $wpsButton.PerformClick()
  if (!$script:wpsSmokeCalled) { throw 'Smoke: WPS button handler not connected.' }
  Write-Output 'PASS: real WPS button calls configuration backend; real WPS launch and writes intercepted'
  $form.Close()
 } elseif ($BackgroundLayoutTest) {
  $form.Show(); [Windows.Forms.Application]::DoEvents()
  function Get-LayoutState($control) {
   $control.GetType().Name+'|'+$control.Bounds.ToString()+'|'+$control.Font.Size.ToString()+'|'+$control.PointToScreen([Drawing.Point]::Empty).ToString()
   foreach($child in $control.Controls) { Get-LayoutState $child }
  }
  $before=(Get-LayoutState $form) -join "`n"
  $probe=Invoke-WpsWorker 'Probe'
  [Windows.Forms.Application]::DoEvents()
  $after=(Get-LayoutState $form) -join "`n"
  if ($before -ne $after) { throw 'Layout changed after background WPS automation.' }
  Write-Output ('PASS: real hidden WPS read; all UI bounds/fonts/screen positions unchanged; Offline='+$probe.Offline+'; Compatibility='+$probe.Compatibility)
  if ($PreviewPath) {
   $bitmap=New-Object Drawing.Bitmap($form.Width,$form.Height)
   $form.DrawToBitmap($bitmap,(New-Object Drawing.Rectangle(0,0,$form.Width,$form.Height)))
   $bitmap.Save($PreviewPath,[Drawing.Imaging.ImageFormat]::Png); $bitmap.Dispose()
  }
  $form.Close()
 } elseif ($PreviewPath) {
  $form.Show(); [Windows.Forms.Application]::DoEvents()
  $bitmap=New-Object Drawing.Bitmap($form.Width,$form.Height)
  $form.DrawToBitmap($bitmap,(New-Object Drawing.Rectangle(0,0,$form.Width,$form.Height)))
  $bitmap.Save($PreviewPath,[Drawing.Imaging.ImageFormat]::Png)
  $bitmap.Dispose(); $form.Close()
 } else { [void]$form.ShowDialog() }
 $detailsForm.Dispose()
 $form.Dispose()
} catch {
 if ($SmokeTest -or $BackgroundLayoutTest) { throw }
 [void][Windows.Forms.MessageBox]::Show(('工具启动失败：'+$_.Exception.Message+"`r`n"+'请把错误内容发给技术人员。'),'启动失败')
 exit 1
}
