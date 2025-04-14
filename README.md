# WxConnectorLib
## 写在前面
欢迎各位指出bug提出Issue，也欢迎各位提交PR使项目变得更好
## 项目简介
WxConnectorLib 是一个基于 .Net Core 8.0 开发的微信操作模拟库
面向微信版本 3.9.12.xx 开发

旨在替代 wxauto 作为新的免费开源的微信操作模拟库

~~因为 wxauto 实在是太慢了，还没有类型检查，故开发了此项目~~
## 相关资源
WeChat 3.9.12 安装包 https://www.123912.com/s/trNHjv-Hi9GA

.Net 8.0.15 Runtime-win-x64 https://www.123912.com/s/trNHjv-Ai9GA

## 快速开始
```csharp
// 引入需要的命名空间
using WxConnectorLib.Managers;
using WxConnectorLib.Utils;

// 使用ActionUtil类进行启动微信和打开监听窗口

// LoginAction会使用传入的exe路径启动微信并等待登录
ActionUtil.Get().LoginAction(@"C:\Program Files\Tencent\WeChat\WeChat.exe");

// OpenListenerWindowsAction对传入的用户打开独立监听窗口
// 并返回监听窗口列表（这里传入的字符串是好友备注名和群名）
var listeners = ActionUtil.Get().OpenListenerWindowsAction(["测试用户", "测试机器人群聊"]);

// 设置一些消息的保存路径（如文件、图片等）
MessageUtil.Get().SetSavePath(".data");

// 使用ListenManager类进行消息监听
ListenManager.Get().InitListen(listener)

// 使用EventManager类来挂载事件处理函数
EventManager.Get().OnNewMessage += (message, chatWindow) => Console.WriteLine($"从{chatWindow.Title}收到消息：{message}");

// ActionUtil类提供了发送文字消息与文件消息的方法
// window参数指定了需要向哪一个用户独立监听窗口发送消息
// 可以从事件处理函数中获取到，也可以从OpenListenerWindowsAction中获取
ActionUtil.Get().SendTextMessage("这是一条测试文字消息", window);
```

## 文档
文档在路上了，当前请各位先参考代码中的注释
