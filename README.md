# Outlook 邮件转 Markdown

Windows 本地邮件归档工具，中文图形界面，当前版本 **1.2.6**。

支持经典 Outlook、PST 工作副本、独立 MSG 文件。可选择邮箱、文件夹、年份或日期范围，将邮件保存为 Markdown，并按需保存 MSG 原件、附件和正文图片。

## 使用

便携包发布后，请从本仓库 Releases 获取 Windows x64 ZIP，完整解压，双击 `Start.bat`。源码下载 ZIP 不是包含 Python 的便携包。

需要 Windows 10/11 x64 和已配置的经典 Outlook。新版 Outlook 不能替代经典 Outlook COM。请先用少量样本验证，再导出大范围邮件。

- 仅文字备份：选择“仅 Markdown”。原始 PST 请另行保留。
- PST：先创建或复用经过验证的完整工作副本，需额外磁盘空间。
- 索引：默认增量更新；外部改动归档后可选择完整重建。已有完整索引无需先重建。
- 中断续跑：保留相同保存位置及其中 `state`、`indexes`、`logs`。
- 程序不发送、删除、移动或标记邮件，不上传邮件内容，不下载外链图片。
- 程序只尝试卸载本次新增的 PST 工作副本挂载，不删除副本文件，不卸载用户预先打开的数据文件。

详见[中文使用说明](一页中文使用说明.md)、[版本历史](CHANGELOG.md)和[测试说明](TESTING.md)。

## 从源码运行

安装 Python 3.13 x64，在项目目录执行：

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements-lock.txt
.venv\Scripts\python -m outlook_archiver gui
```

构建便携包：`.venv\Scripts\python packaging\build_release.py`。运行 `python -m unittest discover -s tests -v` 可进行虚构数据测试；COM 集成测试需要单独运行，不能将单元测试等同于真实邮箱验收。

## 权利与商业授权

这是**源码公开、保留权利**的软件，不采用 MIT 等允许自由修改和商业再分发的开源许可证。修改、再发布、售卖或基于本软件提供收费服务，须先联系 [songw8995](https://github.com/songw8995) 取得书面许可。涉及盈利的合作，分成比例、计算方式和结算安排须另行书面约定；本仓库不自动授予商业许可或确定分成金额。

允许免费使用未经修改的原版程序备份自己有权处理的邮件，包括工作邮件。完整条件见 [LICENSE](LICENSE)。GitHub 平台条款允许的查看、Fork 等行为，以及适用法律不得限制的权利，不因本说明被排除。第三方组件适用各自的许可证，本项目不对其主张额外授权限制。

软件按现状提供，不保证适合所有 Outlook 环境；在删除唯一原件前务必检查备份。此说明不是针对具体交易拟定的商业许可或分成合同。

## 隐私与反馈

公开发布只包含程序、测试、文档和依赖，不包含作者邮件、PST、运行日志或个人配置。提交 Issue 时也请删除真实邮箱、主题、正文、附件、访问令牌和个人路径。不要上传真实邮件作为测试材料。
