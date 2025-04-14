using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;

namespace WxConnectorLib.Utils;

/// <summary>
///     元素工具类（单例）
/// </summary>
public class ElementUtil
{
    private static ElementUtil? _instance;
    private static readonly object Lock = new();
    public Window? MainWindow;
    private string? _wxPath;

    /// <summary>
    ///     获取单例实例
    /// </summary>
    /// <returns>ElementUtil实例</returns>
    public static ElementUtil Get()
    {
        if (_instance != null) return _instance;
        lock (Lock)
        {
            _instance ??= new ElementUtil();
        }

        return _instance;
    }

    /// <summary>
    ///     设置当前实例的微信路径
    /// </summary>
    /// <param name="wxPath">微信exe的路径</param>
    public void SetWxPath(string wxPath)
    {
        _wxPath = wxPath;
    }

    /// <summary>
    ///     用于刷新微信窗口
    /// </summary>
    public void RefreshWxMainWindow()
    {
        if (_wxPath == null) return;
        using var automation = new UIA3Automation();
        MainWindow = Application.Attach(_wxPath).GetMainWindow(automation);
    }

    /// <summary>
    ///     启动微信并获取窗口
    /// </summary>
    public void LaunchWxMainWindow()
    {
        if (_wxPath == null) return;
        using var automation = new UIA3Automation();
        var app = Application.Launch(_wxPath);
        MainWindow = app.GetMainWindow(automation);
    }

    /// <summary>
    ///     用于根据XPath获取微信主窗口的第一个匹配的元素
    /// </summary>
    /// <param name="xPath">需要获取的元素的xpath</param>
    /// <returns>获取到的元素</returns>
    public AutomationElement? GetFirstElementFromMainWindow(string xPath)
    {
        return MainWindow?.FindFirstByXPath(xPath);
    }

    /// <summary>
    ///     用于根据XPath获取微信主窗口的全部匹配的元素
    /// </summary>
    /// <param name="xPath">需要获取的元素的xpath</param>
    /// <returns>获取到的元素</returns>
    public List<AutomationElement> GetAllElementsFromMainWindow(string xPath)
    {
        return MainWindow?.FindAllByXPath(xPath).ToList() ?? [];
    }

    /// <summary>
    ///     用于获取单个用户的独立的聊天窗口
    /// </summary>
    /// <param name="name">需要获取窗口的用户名</param>
    /// <returns>获取到的窗口对象</returns>
    public Window? GetUserChatWindow(string name)
    {
        if (_wxPath == null) return null;
        using var automation = new UIA3Automation();
        return Application.Attach(_wxPath).GetAllTopLevelWindows(automation).First(x => x.Title == name);
    }

    /// <summary>
    ///     从给定的窗口中获取全部匹配的元素
    /// </summary>
    /// <param name="window">给定的窗口</param>
    /// <param name="xPath">xpath</param>
    /// <returns>获取到的全部匹配的元素</returns>
    public List<AutomationElement> GetAllElementsFromGiveWindow(Window window, string xPath)
    {
        return window.FindAllByXPath(xPath).ToList();
    }

    /// <summary>
    ///     从给定的窗口中获取第一个匹配的元素
    /// </summary>
    /// <param name="window">给定的窗口</param>
    /// <param name="xPath">xpath</param>
    /// <returns>获取到的第一个匹配的元素</returns>
    public AutomationElement? GetFirstElementFromGiveWindow(Window window, string xPath)
    {
        return window.FindFirstByXPath(xPath);
    }

    /// <summary>
    ///     用于获取打开的全部独立聊天窗口（双击用户在独立窗口中打开产生的独立聊天窗口）
    /// </summary>
    /// <returns>获取到的窗口列表</returns>
    public List<Window>? GetAllUserChatWindows()
    {
        if (_wxPath == null) return null;
        using var automation = new UIA3Automation();
        return Application.Attach(_wxPath).GetAllTopLevelWindows(automation).Where(x => x.Title != "微信").ToList();
    }

    /// <summary>
    ///     用于从给定的元素中获取第一个匹配到的元素
    /// </summary>
    /// <param name="item">给定的元素</param>
    /// <param name="xPath">xpath</param>
    /// <returns>第一个匹配的元素</returns>
    public AutomationElement? GetFirstElementFromGiveItem(AutomationElement item, string xPath)
    {
        return item.FindFirstByXPath(xPath);
    }

    /// <summary>
    ///     用于从给定的元素中获取全部匹配到的元素
    /// </summary>
    /// <param name="item">给定的元素</param>
    /// <param name="xPath">xpath</param>
    /// <returns>匹配到的元素列表</returns>
    public List<AutomationElement> GetAllElementsFromGiveItem(AutomationElement item, string xPath)
    {
        return item.FindAllByXPath(xPath).ToList();
    }

    /// <summary>
    ///     用于检测给定的头像元素是否是自己的头像
    /// </summary>
    /// <remarks>
    ///     测量原理：
    ///     测量窗口到头像的X轴的距离判断是自己的头像还是别人的头像
    ///     自己的距离是486，别人的是30
    /// </remarks>
    /// <param name="avatarItem">需要检测的头像元素</param>
    /// <param name="chatWindow">头像元素的来源窗口（独立聊天窗口）</param>
    /// <returns>是否是自己的头像</returns>
    public bool IsSelfAvatar(AutomationElement avatarItem, Window chatWindow)
    {
        return avatarItem.BoundingRectangle.Location.X - chatWindow.BoundingRectangle.Location.X != 30;
    }

    /// <summary>
    ///     检测某个消息Item是否是人类发送的消息
    /// </summary>
    /// <remarks>
    ///     因为人类发送的消息能抓到一个头像按钮
    ///     而系统消息不行，故而可以用这个来判断
    /// </remarks>
    /// <param name="msgItem">消息item</param>
    /// <returns>是否是人类的消息</returns>
    public bool IsHumanMsg(AutomationElement msgItem)
    {
        return GetFirstElementFromGiveItem(msgItem, XPath.UserWindowUserAvatarButtonBaseOnMsgItem) != null;
    }

    /// <summary>
    ///     用于从独立的聊天窗口获取另存为窗口
    /// </summary>
    /// <param name="chatWindow">聊天窗口</param>
    /// <returns>另存为窗口</returns>
    public Window? GetSaveAsWindowFromChatWindow(Window chatWindow)
    {
        return chatWindow.FindFirstByXPath(XPath.SaveAsWindowBaseOnChatWindow).AsWindow();
    }

    /// <summary>
    ///     用于获取打开聊天记录后生成的窗口
    /// </summary>
    /// <returns>聊天记录窗口</returns>
    public Window? GetMergeForwardWindow()
    {
        if (_wxPath == null) return null;
        using var automation = new UIA3Automation();
        return Application.Attach(_wxPath).GetAllTopLevelWindows(automation).First(
            x => GetFirstElementFromGiveWindow(x, XPath.MsgItemOfMergeForwardBaseOnMergeForwardWindow) != null
        );
    }
}

/// <summary>
///     这个类用于存放微信主窗口的元素的XPath
/// </summary>
public static class XPath
{
    /// <summary>
    ///     登录界面的的“进入微信”按钮
    /// </summary>
    public const string LoginButton = "/Pane[2]/Pane/Pane[2]/Pane/Pane/Pane[1]/Pane/Pane[2]/Pane/Button";

    /// <summary>
    ///     头像右侧，用户区最顶端的搜索框
    /// </summary>
    public const string SearchBox = "/Pane[2]/Pane/Pane[1]/Pane[1]/Pane[1]/Pane/Edit";

    /// <summary>
    ///     搜索框输入用户后，搜索结果列表的第一个用户Item
    /// </summary>
    public const string SearchUserItem = "/Pane[2]/Pane/Pane[1]/Pane[2]/Pane[2]/List/ListItem[1]";

    /// <summary>
    ///     用户区的用户列表的第一个用户Item
    /// </summary>
    public const string UserListFirstItem = "/Pane[2]/Pane/Pane[1]/Pane[2]/Pane/Pane/Pane/List/ListItem[1]";

    /// <summary>
    ///     右键点击用户Item后(对应UserListFirstItem)，右键菜单的所有菜单项（估计是所有的右键菜单项都可以在这里拿到）
    ///     确定了猜测，每个窗口的右键菜单的所有菜单项都可以用这个拿到
    /// </summary>
    public const string RightClickMenuItems = "/Menu/Pane[2]/List/MenuItem";

    /// <summary>
    ///     用户独立聊天窗口的消息列表的所有消息Item
    /// </summary>
    public const string MsgItems = "/Pane[2]/Pane/Pane[2]/Pane/Pane/Pane[2]/Pane[1]/Pane/List/ListItem";

    /// <summary>
    ///     用户独立聊天窗口的消息输入框元素
    /// </summary>
    public const string UserWindowMsgEdit = "/Pane[2]/Pane/Pane[2]/Pane/Pane/Pane[2]/Pane[2]/Pane[2]/Pane/Pane[1]/Edit";

    /// <summary>
    ///     用户独立聊天窗口的消息发送按钮元素
    /// </summary>
    public const string UserWindowMsgSendButton =
        "/Pane[2]/Pane/Pane[2]/Pane/Pane/Pane[2]/Pane[2]/Pane[2]/Pane/Pane[2]/Pane[3]/Button";

    /// <summary>
    ///     基于消息Item去获取用户头像按钮的XPath
    /// </summary>
    public const string UserWindowUserAvatarButtonBaseOnMsgItem = "/Pane/Button";

    /// <summary>
    ///     基于消息对象获取纯文字消息（猜测，也用于判断是否是文字消息）的文字元素（这个Xpath获取是别人发的消息）
    /// </summary>
    public const string TextOfOtherTextMsgBaseOnMsgItem = "/Pane/Pane[1]/Pane/Pane/Pane/Text";

    /// <summary>
    ///     基于消息对象获取纯文字消息（猜测，也用于判断是否是文字消息）的文字元素（这个Xpath获取是自己发的消息）
    /// </summary>
    public const string TextOfSelfTextMsgBaseOnMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane/Text";

    /// <summary>
    ///     基于自己的消息对象获取小程序卡片消息（猜测/判断）的“小程序”标签（这是一个小程序卡片的标志）
    /// </summary>
    public const string MiniProgramCardSignOfMsgBaseOnSelfMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane/Pane/Pane[2]/Text[2]";
    /// <summary>
    ///     基于别人的消息对象获取小程序卡片消息（猜测/判断）的“小程序”标签（这是一个小程序卡片的标志）
    /// </summary>
    public const string MiniProgramCardSignOfMsgBaseOnOtherMsgItem = "/Pane/Pane[1]/Pane/Pane/Pane/Pane/Pane[2]/Text[2]";

    /// <summary>
    ///     基于别人的消息对象获取转账消息（猜测/判断）的“转账”标签（这是一个转账的标志）
    /// </summary>
    public const string TransferSignOfMsgBaseOnOtherMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane/Pane/Pane[2]/Text[1]";

    /// <summary>
    ///     基于自己的消息对象获取转账消息（猜测/判断）的“转账”标签（这是一个转账的标志）
    /// </summary>
    public const string TransferSignOfMsgBaseOnSelfMsgItem = "/Pane/Pane[1]/Pane/Pane/Pane/Pane/Pane[2]/Text[1]";

    /// <summary>
    ///     基于独立的聊天窗口查找另存为窗口
    /// </summary>
    public const string SaveAsWindowBaseOnChatWindow = "/Window";

    public const string SaveAsWindowPathShowBoxBaseOnChatWindow = "/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar";

    /// <summary>
    ///     基于另存为窗口查找路径编辑框
    /// </summary>
    public const string SaveAsWindowPathEditBaseOnSaveAsWindow = "/Pane[2]/Pane[3]/ProgressBar/ComboBox/Edit";

    /// <summary>
    ///     基于另存为窗口查找文件名编辑框
    /// </summary>
    public const string SaveAsWindowFileNameEditBaseOnSaveAsWindow = "/Pane[1]/ComboBox[1]/Edit";

    /// <summary>
    ///     基于另存为窗口查找保存按钮
    /// </summary>
    public const string SaveAsWindowSaveButtonBaseOnSaveAsWindow = "/Button[1]";

    /// <summary>
    ///     基于另存为窗口查找“转到xxx”的按钮
    /// </summary>
    public const string SaveAsWindowToPathButtonBaseOnSaveAsWindow = "/Pane[2]/Pane[3]/ProgressBar/ToolBar/Button[1]";

    /// <summary>
    ///     自己的语音转文字消息的内容元素
    /// </summary>
    public const string VoiceToTextContentBaseOnSelfMsgItem = "/Pane/Pane[2]/Pane/Pane[2]/Pane/Pane[2]/Pane[2]/Text";

    /// <summary>
    ///     别人的的语音转文字消息的内容元素
    /// </summary>
    public const string VoiceToTextContentBaseOnOtherMsgItem = "/Pane/Pane[1]/Pane/Pane[1]/Pane/Pane[2]/Pane[2]/Text";

    /// <summary>
    ///     自己的语音消息的可用框部分
    /// </summary>
    public const string VoiceMsgItemUseAbleBoxBaseOnSelfMsgItem = "/Pane/Pane[2]/Pane/Pane[2]/Pane";

    /// <summary>
    ///     别人的的语音消息的可用框部分
    /// </summary>
    public const string VoiceMsgItemUseAbleBoxBaseOnOtherMsgItem = "/Pane/Pane[1]/Pane/Pane[1]/Pane";

    /// <summary>
    ///     基于群聊的聊天记录窗口查找所有转发的消息Item
    /// </summary>
    public const string MsgItemOfMergeForwardBaseOnMergeForwardWindow = "/Pane[2]/List/ListItem";

    /// <summary>
    ///     基于转发的消息Item查找消息内容（文字内容）
    /// </summary>
    public const string MsgContentTextBaseOnMsgItemOfMergeForwardItem = "/Pane/Pane[2]/Pane[2]/Text";

    /// <summary>
    ///     基于群聊的聊天记录窗口查找窗口关闭按钮
    /// </summary>
    public const string CloseButtonBaseOnMergeForwardWindow = "/Pane[2]/Pane[1]/Button[2]";
    /// <summary>
    ///     基于小程序卡片消息获取小程序标题
    /// </summary>
    public const string MiniProgramTitleBaseOnOtherMsgItem = "/Pane/Pane[1]/Pane/Pane/Pane/Pane/Pane[1]/Text";
    /// <summary>
    ///     基于小程序卡片消息获取小程序卡片内容
    /// </summary>
    public const string MiniProgramContentBaseOnOtherMsgItem = "/Pane/Pane[1]/Pane/Pane/Pane/Pane/Text";
    /// <summary>
    ///     基于小程序卡片消息获取小程序标题
    /// </summary>
    public const string MiniProgramTitleBaseOnSelfMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane/Pane/Pane[1]/Text";
    /// <summary>
    ///     基于小程序卡片消息获取小程序卡片内容
    /// </summary>
    public const string MiniProgramContentBaseOnSelfMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane/Pane/Text";
    /// <summary>
    ///     基于别人的消息对象获取转账消息的金额
    /// </summary>
    public const string TransferNumberBaseOnOtherMsgItem = "/Pane/Pane[1]/Pane/Pane/Pane/Pane/Pane[1]/Pane/Text[2]";
    /// <summary>
    ///     基于自己的消息对象获取转账消息的金额
    /// </summary>
    public const string TransferNumberBaseOnSelfMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane/Pane/Pane[1]/Pane/Text[2]";
    /// <summary>
    ///  基于消息对象获取引用消息的来源用户
    /// </summary>
    public const string QuoteUserBaseOnMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane[2]/Pane/Pane/Pane/Pane/Text";

    /// <summary>
    ///     用于修复文件右键不到的情况（适用于查找自己的消息）
    /// </summary>
    public const string ClickAbelBoxBaseOnSelfMsgItem = "/Pane/Pane[2]/Pane/Pane/Pane";

    /// <summary>
    ///     用于修复文件右键不到的情况（适用于查找别人的消息）
    /// </summary>
    public const string ClickAbelBoxBaseOnOtherMsgItem = "/Pane/Pane[1]/Pane/Pane/Pane";
    /// <summary>
    ///     基于消息对象获取拍一拍消息对象
    /// </summary>
    public const string PypMsgBaseOnMsgItem = "/Pane/Pane[2]/Pane/ListItem";
    /// <summary>
    ///     基于登录窗口获取登录信号（就是用户名，登录成功后获取不到就是登录完成了）
    /// </summary>
    public const string LoginSignBaseOnLoginWindow = "/Pane[2]/Pane/Pane[1]/Button[1]";
    /// <summary>
    ///     基于聊天窗口获取消息编辑框
    /// </summary>
    public const string MsgEditBaseOnChatWindow =
        "/Pane[2]/Pane/Pane[2]/Pane/Pane/Pane[2]/Pane[2]/Pane[2]/Pane/Pane[1]/Edit";
    /// <summary>
    ///     基于聊天窗口获取消息发送按钮
    /// </summary>
    public const string MsgSendButtonBaseOnChatWindow =
        "/Pane[2]/Pane/Pane[2]/Pane/Pane/Pane[2]/Pane[2]/Pane[2]/Pane/Pane[2]/Pane[3]/Button";
    /// <summary>
    ///     基于聊天窗口获取发送文件按钮
    /// </summary>
    public const string SendFileButtonBaseOnChatWindow =
        "/Pane[2]/Pane/Pane[2]/Pane/Pane/Pane[2]/Pane[2]/Pane[2]/Pane/ToolBar/Button[2]";

    /// <summary>
    ///     基于发送文件窗口获取文件名编辑框
    /// </summary>
    public const string FileNameEditBaseOnFileSendWindow = "ComboBox[1]/Edit";
    /// <summary>
    ///     基于发送文件窗口获取打开文件按钮
    /// </summary>
    public const string OpenFileButtonBaseOnFileSendWindow = "SplitButton";
    /// <summary>
    ///     基于发送文件浮窗获取发送按钮
    /// </summary>
    public const string SendButtonBaseOnFileSendCard = "/Window/Pane[2]/Pane[5]/Button[1]";
}

/*
 * Todo:测试并修复全局的Xpath问题
 * 现在的xpath是基于文件传输助手的比较多，可是有个问题，相同的消息，别人发的和自己发的xpath不一样
 * 这一点需要测试并修复
 */
