using FlaUI.Core.Input;
using WxConnectorLib.Managers;
using WxConnectorLib.Utils;

namespace WxConnectorLib;

public static class WxConnector
{
    public static void Main(string[] args)
    {
        /*// 设置路径，启动，并等待启动完成
        ElementUtil.Get().SetWxPath(@"C:\Program Files\Tencent\WeChat\WeChat.exe");
        ElementUtil.Get().LaunchWxMainWindow();
        ElementUtil.Get().GetFirstElementFromMainWindow(XPath.LoginButton)?.Click();
        Thread.Sleep(6000);
        // 刷新界面
        ElementUtil.Get().RefreshWxMainWindow();*/
        ActionUtil.Get().LoginAction(@"C:\Program Files\Tencent\WeChat\WeChat.exe");
        // 搜索用户并点击
        var listener = ActionUtil.Get().OpenListenerWindowsAction(["群聊测试", "测试机器人群聊"]);
        /*var searchBox = ElementUtil.Get().GetFirstElementFromMainWindow(XPath.SearchBox);
        searchBox?.Click();
        Keyboard.Type("文件传输助手");
        Thread.Sleep(800);
        ElementUtil.Get().GetFirstElementFromMainWindow(XPath.SearchUserItem)?.Click();*/
        // 打开用户到指定独立的聊天窗口
        /*
        第一种方法，右键菜单点击在独立窗口中打开
        ElementUtil.Get().GetFirstElementFromMainWindow(XPath.UserListFirstItem)?.Click();
        Thread.Sleep(200);
        ElementUtil.Get().GetAllElementsFromMainWindow(XPath.RightClickMenuItems).First(x => x.Name == "在独立窗口中打开").Click();
        */
        /*// 第二种方法：double click 用户对象
        ElementUtil.Get().GetFirstElementFromMainWindow(XPath.UserListFirstItem)?.DoubleClick();
        // 获取窗口对象
        var window = ElementUtil.Get().GetUserChatWindow("文件传输助手");
        Console.WriteLine(window?.Title);*/
        // 获取消息对象
        /*var msg = ElementUtil.Get().GetAllElementsFromGiveWindow(window, XPath.MsgItems).LastOrDefault();
        Console.WriteLine(msg != null ? msg.Name : "无消息");*/
        // 这里插入测试头像检测
        /* 头像检测可用
        var isSelf = ElementUtil.Get().IsSelfAvatar(
            ElementUtil.Get()
                .GetFirstElementFromGiveItem(msg, XPath.UserWindowUserAvatarButtonBaseOnMsgItem),
            window
        );
        Console.WriteLine(isSelf ? "是自己" : "不是自己");
        */
        // 激活窗口获取输入框发送信息
        /*
        这里激活窗口获取输入框发送信息可用
        var msgEdit = ElementUtil.Get().GetFirstElementFromGiveWindow(window, XPath.UserWindowMsgEdit);
        msgEdit?.Click();
        Keyboard.Type("这是一条测试消息");
        ElementUtil.Get().GetFirstElementFromGiveWindow(window, XPath.UserWindowMsgSendButton)?.Click();
        Keyboard.Type("测试排版喵\n百奇是可爱猫娘喵！");
        ElementUtil.Get().GetFirstElementFromGiveWindow(window, XPath.UserWindowMsgSendButton)?.Click();
        */
        // 测试监听消息
        //这里监听可用
        /*MessageUtil.Get().SetSavePath();
        var listenManager = ListenManager.Get();
        listenManager.InitListen(listener);
        EventManager.Get().OnNewMessage += (message, chatWindow) => Console.WriteLine($"从{chatWindow.Title}收到消息：{message}");
        EventManager.Get().OnNewMessage += (msg, window) =>
        {
            if (msg.MsgContent![0] == "测试")
            {
                ActionUtil.Get().SendTextMessage("开始测试", window);
                ActionUtil.Get().SendTextMessage("这是一条测试文字消息", window);
                ActionUtil.Get().SendTextMessage("下面是一条测试图片消息", window);
                ActionUtil.Get().SendFileMessage(@"C:\Users\Administrator\Pictures\1.jpg", window);
            }
        };*/
        /*
        发送消息可用
        ActionUtil.Get().SendTextMessage("这是一条测试消息", listener.First(x => x.Title == "文件传输助手"));
        ActionUtil.Get().SendTextMessage("接下来测试文件消息", listener.First(x => x.Title == "文件传输助手"));
        ActionUtil.Get().SendFileMessage(
            @"C:\Users\Administrator\Pictures\1.jpg",
            listener.First(x => x.Title == "文件传输助手")
            );*/
        /*Thread.Sleep(TimeSpan.FromMinutes(10));*/
        // 测试保存文件
        /*
        测试保存文件可用
        ActionUtil.Get().SaveAsAction(
            window, msg, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".testData")
        );
        */
        /*
         * 测试语音转文字可用
         * Todo: 需要能捕捉一个网络不给力的错误，不然会一直等待，还需要设置超时和测试正常语音是否能转
         * ActionUtil.Get().VoiceToTextAction(window, msg);
         */
        /*
         * 测试转发消息处理可用
         * ActionUtil.Get().OpenMergeForwardAction(msg);
         */
        var window = listener.First(x => x.Title == "群聊测试");
        var msg = ElementUtil.Get().GetAllElementsFromGiveWindow(
            window, XPath.MsgItems
        );
        ActionUtil.Get().QuoteMessage(
            msg.Last(x => x.Name == "没开"),
            window,
            true
        );
        // 通过名字 at 可用 ActionUtil.Get().AtByNameInGroup("小米", window);
        /*
        通过头像 at 可用
        var user = ActionUtil.Get().SearchUserAvatarByName(window, "小米");
        ActionUtil.Get().At(user, window);
        ActionUtil.Get().SendTextMessage("爱你宝宝", window);
        */
    }
}
