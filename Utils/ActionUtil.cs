using System.Drawing;
using System.Text.RegularExpressions;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Input;
using WxConnectorLib.Managers;

namespace WxConnectorLib.Utils;

/// <summary>
///     操作工具类（单例）
/// </summary>
public class ActionUtil
{
    private static ActionUtil? _instance;
    private static readonly object Lock = new();
    private readonly ElementUtil _util = ElementUtil.Get();

    /// <summary>
    ///     获取单例实例
    /// </summary>
    /// <returns>ActionUtil实例</returns>
    public static ActionUtil Get()
    {
        if (_instance != null) return _instance;
        lock (Lock)
        {
            _instance ??= new ActionUtil();
        }

        return _instance;
    }

    /// <summary>
    ///     用于处理需要点击消息对象然后另存为的操作
    /// </summary>
    /// <remarks>会在传入的保存路径中使用微信默认的文件名保存指定的文件并返回完整保存路径</remarks>
    /// <param name="chatWindow">消息来源窗口</param>
    /// <param name="msgElement">消息对象</param>
    /// <param name="savePath">保存路径</param>
    /// <returns>完整保存路径</returns>
    public string SaveAsAction(Window chatWindow, AutomationElement msgElement, string savePath, bool isVedio = false)
    {
        // 初始化可点击范围
        var clickBox = _util.GetFirstElementFromGiveItem(
            msgElement, XPath.ClickAbelBoxBaseOnSelfMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
            msgElement, XPath.ClickAbelBoxBaseOnOtherMsgItem
            );

        ListenManager.Get().PauseListen();

        if (isVedio)
        {
            // 视频要点击下载
            clickBox?.Click();
        }

        var saveAsButton = WaitUtil.WaitUntil(
            () =>
            {
                clickBox?.RightClick();
                // 获取右键菜单中“另存为”
                var res = _util.GetAllElementsFromGiveWindow(
                    chatWindow,
                    XPath.RightClickMenuItems
                ).FirstOrDefault(x => x.Name.Contains("另存为") );
                if (res == null)
                {
                    // 没有找到则再次右键关闭右键菜单
                    clickBox?.RightClick();
                    Thread.Sleep(200);
                    // 还需要点击一下窗口刷新一下
                    Mouse.MovePixelsPerMillisecond = 100;
                    Mouse.MoveTo(new Point(Mouse.Position.X + 200, Mouse.Position.Y));
                    Mouse.Click();
                    return (false, null);
                }
                return (true, res);
            },
            TimeSpan.FromMilliseconds(200),
            TimeSpan.FromMinutes(10)
        );

        // 点击另存为按钮
        saveAsButton!.Click();
        Thread.Sleep(400);

        // 等待文件保存对话框弹出并获取
        var saveAsWindow = WaitUtil.WaitUntil(
            () =>
            {
                var res = (_util.GetFirstElementFromGiveWindow(
                    chatWindow,
                    XPath.SaveAsWindowBaseOnChatWindow
                ).AsWindow() ?? _util.GetFirstElementFromMainWindow(
                    XPath.SaveAsWindowBaseOnChatWindow
                ).AsWindow());
                return res == null ? (false, null) : (true, res);
            },
            TimeSpan.FromMilliseconds(200),
            TimeSpan.FromSeconds(1)
        );

        // 获取需要的组件
        var pathShowBox =
            _util.GetFirstElementFromGiveWindow(
                saveAsWindow!, XPath.SaveAsWindowPathShowBoxBaseOnChatWindow
            );

        // 点击showBox才会能找到pathEditXPath.SaveAsWindowPathEditBaseOnSaveAsWindow
        pathShowBox?.Click();
        var pathEdit = _util.GetFirstElementFromGiveWindow(saveAsWindow ?? _util.MainWindow!, XPath.SaveAsWindowPathEditBaseOnSaveAsWindow);
        var fileNameEdit =
            _util.GetFirstElementFromGiveWindow(
                saveAsWindow!, XPath.SaveAsWindowFileNameEditBaseOnSaveAsWindow
            );
        var saveButton =
            _util.GetFirstElementFromGiveWindow(
                saveAsWindow!, XPath.SaveAsWindowSaveButtonBaseOnSaveAsWindow
            );
        var toButton =
            _util.GetFirstElementFromGiveWindow(
                saveAsWindow!, XPath.SaveAsWindowToPathButtonBaseOnSaveAsWindow
            );

        // 确保路径存在
        if (!Directory.Exists(savePath)) Directory.CreateDirectory(savePath);

        // 设置保存路径
        pathEdit?.Click();
        pathEdit?.Focus();
        pathEdit?.Patterns.Value.Pattern.SetValue(savePath);

        // 转到路径
        toButton?.Click();

        /*Keyboard.Type(savePath);*/

        // 获取文件名
        var fileName = fileNameEdit?.Patterns.Value.Pattern.Value.Value;

        // 点击保存按钮
        saveButton?.Click();

        ListenManager.Get().ResumeListen();

        return Path.Combine(savePath, fileName!);
    }

    /// <summary>
    ///     执行语音转文字操作并返回转换后的文字
    /// </summary>
    /// <param name="msgElement">语音消息对象</param>
    /// <param name="chatWindow">消息来源窗口</param>
    /// <remarks>这里可能会遇到“网络不给力”的问题，会原样捕捉为信息</remarks>
    /// <returns>转换后的文字</returns>
    public string VoiceToTextAction(Window chatWindow, AutomationElement msgElement)
    {
        // 获取一个可点击的范围
        var clickAbleElement = _util.GetFirstElementFromGiveItem(
            msgElement, XPath.VoiceMsgItemUseAbleBoxBaseOnSelfMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
            msgElement, XPath.VoiceMsgItemUseAbleBoxBaseOnOtherMsgItem
            );

        // 先右键点击消息对象，唤起右键菜单
        clickAbleElement?.RightClick();

        // 获取右键菜单中“语音转文字”并点击
        var vttButton = _util.GetAllElementsFromGiveWindow(
            chatWindow,
            XPath.RightClickMenuItems
        ).First(x => x.Name == "语音转文字");
        vttButton.Click();

        // 循环等待并获取内容
        var content = WaitUtil.WaitUntil(
            () =>
            {
                var res = _util.GetFirstElementFromGiveItem(
                    msgElement, XPath.VoiceToTextContentBaseOnSelfMsgItem
                    ) ?? _util.GetFirstElementFromGiveItem(
                    msgElement, XPath.VoiceToTextContentBaseOnOtherMsgItem
                    );
                return res == null ? (false, res) : (true, res);
            },
            TimeSpan.FromMilliseconds(200),
            TimeSpan.MaxValue
        );

        return content!.Name;
    }

    /// <summary>
    ///     打开合并转发的聊天记录获取全部内容并返回
    /// </summary>
    /// <param name="msgElement">消息对象</param>
    /// <returns>聊天记录</returns>
    public List<string> OpenMergeForwardAction(AutomationElement msgElement)
    {
        /*
         * Todo: 后人可以做一下这里的更多消息类型支持
         */
        // 点击聊天记录打开窗口
        msgElement.Click();

        // 获取窗口
        var mergeForwardWindow = _util.GetMergeForwardWindow();

        // 获取List
        var chatMsgList = _util.GetAllElementsFromGiveWindow(
            mergeForwardWindow!,
            XPath.MsgItemOfMergeForwardBaseOnMergeForwardWindow
        );

        // 获取每个消息的内容，并重写格式化消息（这里只支持文本，按照文本推断，非文本的使用re处理）
        var res = new List<string>();
        // 获取内容（文本）
        foreach (var msg in chatMsgList)
        {
            var content = _util.GetFirstElementFromGiveItem(
                msg,
                XPath.MsgContentTextBaseOnMsgItemOfMergeForwardItem
            );
            if (content == null)
            {
                // 则不是文本
                var match = Regex.Match(msg.Name, @"\[(.*?)\]");
                if (!match.Success) continue;
                var contentName = match.Value;
                res.Add(
                    $"用户[{msg.Name.Split(contentName)[0]}] 在[{msg.Name.Split(contentName)[1]}]发送了一个：{contentName}"
                );
            }
            else
            {
                res.Add(
                    $"用户[{msg.Name.Split(content.Name)[0]}] 在[{msg.Name.Split(content.Name)[1]}]说：{content.Name}"
                );
            }
        }

        // 关闭窗口
        _util.GetFirstElementFromGiveWindow(mergeForwardWindow, XPath.CloseButtonBaseOnMergeForwardWindow)?.Click();
        return res;
    }

    /// <summary>
    /// 会进行微信登录并等待登录成功刷新页面
    /// </summary>
    /// <param name="wxPath">微信路径</param>
    public void LoginAction(string wxPath)
    {
        _util.SetWxPath(wxPath);
        _util.LaunchWxMainWindow();
        var loginButton = _util.GetFirstElementFromMainWindow(XPath.LoginButton);
        // 获取的到loginButton则点击，获取不到就是有二维码需要扫码
        loginButton?.Click();
        // 等待获取不到登录sign（也就是等待登录成功）
        WaitUtil.WaitUntil(
            () =>
            {
                // 这里抓报错是因为登录后会刷新，导致我们拿到的窗口不存在，会报错，其实已经登录成功了
                try
                {
                    return _util.GetFirstElementFromMainWindow(XPath.LoginSignBaseOnLoginWindow) != null
                        ? (false, string.Empty)
                        : (true, string.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    return (true, string.Empty);
                }
            },
            TimeSpan.FromMilliseconds(200),
            TimeSpan.MaxValue
        );
        // 刷新页面
        _util.RefreshWxMainWindow();
    }

    /// <summary>
    /// 用于打开监听窗口并返回监听窗口对象
    /// </summary>
    /// <param name="listeners">监听用户的备注名</param>
    /// <returns>监听窗口对象列表</returns>
    public List<Window> OpenListenerWindowsAction(List<string> listeners)
    {
        var listenWindows = new List<Window>();
        foreach (var item in listeners)
        {
            // 获取搜索框
            var searchBox = _util.GetFirstElementFromMainWindow(XPath.SearchBox);
            // 点击搜索框
            searchBox?.Click();
            // 输入用户名
            Keyboard.Type(item);
            // 等待搜索结果出来并点击
            WaitUtil.WaitUntil(
                () =>
                {
                    var res = _util.GetFirstElementFromMainWindow(XPath.SearchUserItem);
                    if (res == null) return (false, string.Empty);
                    if (res.Name != item) return (false, string.Empty);
                    res.Click();
                    return (true, string.Empty);
                },
                TimeSpan.FromMilliseconds(200),
                TimeSpan.FromSeconds(1)
            );
            // 双击刚刚点击的搜索对象，打开到独立窗口
            _util.GetFirstElementFromMainWindow(XPath.UserListFirstItem)?.DoubleClick();
            listenWindows.Add(_util.GetUserChatWindow(item)!);
        }
        return listenWindows;
    }

    /// <summary>
    /// 向指定的聊天窗口发送文字消息
    /// 窗口必须是监听的窗口，也必须是监听产生的
    /// </summary>
    /// <param name="msg">要发送的消息</param>
    /// <param name="chatWindow">用于发送的聊天窗口（监听产生的）</param>
    public void SendTextMessage(string msg, Window chatWindow)
    {
        // 先暂停监听，并激活窗口
        ListenManager.Get().PauseListen();
        chatWindow.Focus();

        // 获取消息输入框和发送按钮，点击输入框，输入内容，点击发送按钮
        var edit = _util.GetFirstElementFromGiveWindow(
            chatWindow, XPath.MsgEditBaseOnChatWindow
        );
        var sendButton = _util.GetFirstElementFromGiveWindow(
            chatWindow, XPath.MsgSendButtonBaseOnChatWindow
        );
        edit?.Click();
        Keyboard.Type(msg);
        sendButton?.Click();

        // 恢复监听
        ListenManager.Get().ResumeListen();
    }

    /// <summary>
    /// 向指定的聊天窗口发送文件消息
    /// 窗口必须是监听的窗口，也必须是监听产生的
    /// </summary>
    /// <param name="path">要发送的文件路径</param>
    /// <param name="chatWindow">聊天窗口</param>
    public void SendFileMessage(string path, Window chatWindow)
    {
        // 确认文件存在
        if (!File.Exists(path)) return;

        // 先暂停监听，并激活窗口
        ListenManager.Get().PauseListen();
        chatWindow.Focus();

        // 获取发送文件按钮，并点击唤起发送窗
        var sendFileButton = _util.GetFirstElementFromGiveWindow(
            chatWindow, XPath.SendFileButtonBaseOnChatWindow
        );
        sendFileButton?.Click();

        // 等待文件发送对话框弹出并获取（这里类似于另存为，故直接使用相同Xpath）
        var sendWindow = WaitUtil.WaitUntil(
            () =>
            {
                var res = (_util.GetFirstElementFromGiveWindow(
                    chatWindow,
                    XPath.SaveAsWindowBaseOnChatWindow
                ).AsWindow() ?? _util.GetFirstElementFromMainWindow(
                    XPath.SaveAsWindowBaseOnChatWindow
                ).AsWindow());
                return res == null ? (false, null) : (true, res);
            },
            TimeSpan.FromMilliseconds(200),
            TimeSpan.FromSeconds(1)
        );

        // 获取需要的组件（这里类似于另存为，故直接使用相同Xpath）
        var pathShowBox =
            _util.GetFirstElementFromGiveWindow(
                sendWindow!, XPath.SaveAsWindowPathShowBoxBaseOnChatWindow
            );

        // 点击showBox才会能找到pathEditXPath.SaveAsWindowPathEditBaseOnSaveAsWindow
        pathShowBox?.Click();
        var pathEdit = _util.GetFirstElementFromGiveWindow(sendWindow ?? _util.MainWindow!, XPath.SaveAsWindowPathEditBaseOnSaveAsWindow);

        // FileNameEdit和OpenButton的Xpath与另存为不同
        var fileNameEdit =
            _util.GetFirstElementFromGiveWindow(
                sendWindow!, XPath.FileNameEditBaseOnFileSendWindow
            );
        var openButton =
            _util.GetFirstElementFromGiveWindow(
                sendWindow!, XPath.OpenFileButtonBaseOnFileSendWindow
            );

        var toButton =
            _util.GetFirstElementFromGiveWindow(
                sendWindow!, XPath.SaveAsWindowToPathButtonBaseOnSaveAsWindow
            );

        // 将文件路径分为路径与名字
        var fileName = Path.GetFileName(path);
        var filePath = Path.GetDirectoryName(path);

        // 输入路径并跳转
        pathEdit?.Click();
        pathEdit?.Focus();
        pathEdit?.Patterns.Value.Pattern.SetValue(filePath!);
        toButton?.Click();

        // 输入文件名
        fileNameEdit?.Click();
        fileNameEdit?.Focus();
        fileNameEdit?.Patterns.Value.Pattern.SetValue(fileName);

        // 点击打开按钮进行发送
        openButton?.Click();

        // 获取发送按钮并点击
        var sendButton = WaitUtil.WaitUntil(
            () =>
            {
                var res = (_util.GetFirstElementFromGiveWindow(
                    chatWindow,
                    XPath.SendButtonBaseOnFileSendCard
                ) ?? _util.GetFirstElementFromMainWindow(
                    XPath.SendButtonBaseOnFileSendCard
                ));
                return res == null ? (false, null) : (true, res);
            },
            TimeSpan.FromMilliseconds(100),
            TimeSpan.FromSeconds(1)
        );
        sendButton!.Click();

        // 恢复监听
        ListenManager.Get().ResumeListen();
    }
}
