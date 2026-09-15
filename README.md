# Word Default Switcher

带中文界面的 Windows 文档默认打开程序切换工具。检测八种文件格式，尝试批量设为 Microsoft Word，并保存操作记录、独立复查结果。
<img width="2308" height="1703" alt="image" src="https://github.com/user-attachments/assets/a7de7e77-bbb8-4853-9847-fea176435828" />


## 使用

1. 下载并解压整个工具包。
2. 使用需要修改默认应用的 Windows 账号，双击 **启动工具.vbs**。若 VBS 不可用，双击 **Start.cmd**。
3. 确认已检测到本机 Microsoft Word，点击 **一键设为 Word**。
4. 查看逐项结果，并在 Windows 默认应用设置中确认；用测试文档验证实际打开行为。

无需安装本工具或联网；无需以其他管理员账号运行。企业策略禁止脚本时，请由 IT 按组织策略处理。

支持的格式：`.asd`、`.doc`、`.docm`、`.docx`、`.dot`、`.dotm`、`.dotx`、`.odt`。

## 环境与实现

| 环境 | 设置方式 |
| --- | --- |
| Windows 7 | 尝试调用系统默认程序接口 |
| Windows 10 / 11 | 使用随包 PS-SFTA 尝试设置当前用户关联；需要 Windows PowerShell 5+ |

需已安装并正确注册桌面版 Microsoft Word。请通过提供的启动文件运行 Windows PowerShell，不要使用 PowerShell 7。

界面使用 Windows Forms；通过 C# 调用 Windows API 查询关联。设置前保存 TXT/XML 记录，设置后在新的 PowerShell 进程中重新查询。

**已知限制：**“查询结果指向 Word”不保证 Windows 默认应用设置页已同步认可。部分系统上曾出现实际打开程序为 Word、设置页仍为 WPS 的现象。遇到这种情况，通过系统设置完成修改，不要将工具提示视为完整验收。

PS-SFTA 修改当前用户的 `UserChoice` 和校验值，不是微软官方支持的设置接口。系统补丁或组织策略可能阻止修改。工具不会关闭保护驱动、修改宏安全设置、卸载 WPS 或修复 Office。

`.asd` 是自动恢复文件，修改关联不能保证恢复成功。文件默认应用修改也不能证明第三方业务系统的签名、排版或 Office 调用已恢复。

## 记录与恢复

记录保存在工具目录的 `检测记录`；目录不可写时尝试当前用户临时目录中的 `WordDefaultReports`。

记录包含电脑名、用户名和安装路径，不包含业务文档内容。**不要上传原始检测记录或含客户信息的截图。** 如需恢复，参考设置前记录，通过“系统默认应用”按钮重新选择原程序。本版不提供自动回滚。

## 验证状态

- 已验证：中文窗口、八种格式检测、真实按钮调用、TXT/XML 保存、独立进程复查。
- 已测试：模拟批量成功、部分失败、写入被拒绝及复查失败；模拟 Win11 多个内部版本进入设置流程。
- 现场反馈：使用者报告在客户机器上验证成功，尚未记录其 Windows/Office 具体版本及完整测试范围。因此不代表所有 Win7/10/11 环境均已验证。

## 开发验证

在 Windows PowerShell 中运行：

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-Batch.ps1
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Window.ps1 -SmokeTest
```

测试会读取当前关联并生成本地检测记录；批量测试及窗口测试拦截实际关联写入，不修改默认应用。`ExecutionPolicy Bypass` 只作用于启动的进程，不更改持久执行策略。

## 第三方组件

包含 [DanysysTeam/PS-SFTA](https://github.com/DanysysTeam/PS-SFTA)，采用 MIT 许可。固定源代码版本及本地调整见 [来源说明](vendor/来源说明.txt)，许可全文见 [vendor/LICENSE.txt](vendor/LICENSE.txt)。启动时会校验此组件的 SHA256。

本项目自身代码尚未另行指定开源许可证；第三方代码继续遵循其原许可证。
