# Outlook 邮件转 Markdown 1.2.5

- 普通使用者下载 `OutlookMarkdown-1.2.5-Windows-x64.zip`，完整解压后双击 `Start.bat`。
- 需要 Windows x64 与已配置的经典 Outlook。包内含独立 Python；不包含 Microsoft Outlook。
- `OutlookMarkdown-1.2.5-Source.zip` 用于源码保存和维护；`SHA256SUMS.txt` 用于核对两个包。
- 默认增量索引，可选完整重建；保留历史索引，避免逐个重读所有旧邮件的元数据。
- 原版可免费备份自己有权处理的邮件（包括工作邮件）。修改、再发布和收费服务须事先取得书面许可；盈利分成另行书面约定。详见 LICENSE。

发布包由 GitHub Windows 构建环境从本仓库源码生成，经过虚构数据回归测试、ZIP 校验及便携 GUI 启动检查。构建环境未安装用户的 Outlook 邮箱，未执行真实 PST/邮箱导出验证。包中不包含维护者的邮件、PST、个人配置或运行日志。先使用小样本验证后再处理自己的真实邮件。

该公开构建与维护者本机先前打包版本使用相同应用代码；Python 补丁版本及构建环境可能不同，ZIP 的校验值可能不同，以本次附带的校验文件为准。
