namespace WxConnectorLib.Models;

public class WMessage
{
    public WMsgType MsgType { get; set; }
    public WMsgSourceType MsgSourceType { get; set; }
    public WMsgSenderType MsgSenderType { get; set; }
    public WMsgOriginType MsgOriginType { get; set; }
    public List<string>? MsgContent { get; set; }
    public string? MsgSenderName { get; set; }
    public string? MsgFromWindow { get; set; }

    public override string ToString()
    {
        var msgContent = MsgContent != null
            ? string.Join(";", MsgContent)
            : "null";
        return $"{MsgType.GetType().Name}.{MsgType};" +
                $"{MsgSourceType.GetType().Name}.{MsgSourceType};" +
                $"{MsgSenderType.GetType().Name}.{MsgSenderType};" +
                $"{MsgOriginType.GetType().Name}.{MsgOriginType};" +
                $"{msgContent};" +
                $"{MsgSenderName};" +
                $"{MsgFromWindow};";
    }
}

/// <summary>
///     用于区分消息生成源——区分是系统消息还是用户消息
/// </summary>
/// <remarks>
///     Human：用户消息
///     System：系统消息
/// </remarks>
public enum WMsgSourceType
{
    Human,
    System
}

/// <summary>
///     用于区分消息发送者——区分是自己还是好友
/// </summary>
/// <remarks>
///     Self：自己
///     Friend：好友（这里的好友指“其他人”，不一定是真正的微信好友，可能是群友，为了方便统一称为好友）
/// </remarks>
public enum WMsgSenderType
{
    Self,
    Friend
}

/// <summary>
///     用于区分消息来源——区分是单人消息（私聊消息）还是群消息
/// </summary>
/// <remarks>
///     Single：单人消息
///     Group：群消息
/// </remarks>
public enum WMsgOriginType
{
    Single,
    Group
}

/// <summary>
///     消息类型
/// </summary>
/// <remarks>
///     Text：普通文本消息
///     Image：图片消息
///     Video：视频消息
///     Emoji：表情消息
///     File：文件消息
///     MiniProgramCard：小程序卡片消息
///     MergeForward：合并转发消息
///     Voice：语音消息
///     Transfer：转账消息
///     Quote：引用消息
/// </remarks>
public enum WMsgType
{
    Text,
    Image,
    Video,
    Emoji,
    File,
    MiniProgramCard,
    MergeForward,
    Voice,
    Transfer,
    Quote
}
