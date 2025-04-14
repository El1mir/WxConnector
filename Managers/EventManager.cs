using FlaUI.Core.AutomationElements;
using WxConnectorLib.Models;

namespace WxConnectorLib.Managers;

public class EventManager
{
    private static EventManager? _instance;
    private static readonly object Lock = new();
    public delegate void NewMessageHandler(WMessage msg, Window chatWindow);
    public delegate void NewMessageWithoutSelfHandler(WMessage msg, Window chatWindow);

    public event NewMessageHandler? OnNewMessage;
    public event NewMessageWithoutSelfHandler? OnNewMessageWithoutSelf;

    public static EventManager Get()
    {
        if (_instance != null) return _instance;
        lock (Lock)
        {
            _instance ??= new EventManager();
        }

        return _instance;
    }

    /// <summary>
    /// 用于处理新消息的事件
    /// </summary>
    /// <param name="msg">消息对象</param>
    /// <param name="chatWindow">源窗口</param>
    public void InvokeNewMessage(WMessage msg, Window chatWindow)
    {
        OnNewMessage?.Invoke(msg, chatWindow);
        if (msg.MsgSenderType != WMsgSenderType.Self) InvokeNewMessageWithoutSelf(msg, chatWindow);
    }

    /// <summary>
    /// 用于处理新消息的事件（不包含自己发送的消息）
    /// </summary>
    /// <param name="msg">消息对象</param>
    /// <param name="chatWindow">源窗口</param>
    private void InvokeNewMessageWithoutSelf(WMessage msg, Window chatWindow)
    {
        OnNewMessageWithoutSelf?.Invoke(msg, chatWindow);
    }
}
