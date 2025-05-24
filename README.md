# wx-auto-friend 微信自动加好友工具

## 📘 项目简介

本项目是一个基于 Python 的桌面自动化工具，借助 `uiautomation` 与 `pyautogui` 实现自动化添加微信好友。它具备图形化界面，便于用户设置参数，支持批量添加账号、备注、填写验证信息，并模拟人工操作行为。

---

## ✨ 功能特点

* 图形界面配置操作参数
* 自动启动微信客户端
* 批量导入账号与备注信息
* 随机时间间隔操作，模拟人类行为
* 图像识别定位点击“添加到通讯录”按钮
* 操作日志记录并输出
* 倒计时提示与任务终止机制

---

## 🧰 环境依赖

* 操作系统：Windows 10/11
* Python 版本：Python 3.7+
* 第三方库：

  * `pyautogui`
  * `uiautomation`
  * `comtypes`

### 安装依赖

```bash
pip install -r requirements.txt
```

