using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using FlaUI.Core.AutomationElements;
using WxConnectorLib.Models;

namespace WxConnectorLib.Utils;

/// <summary>
/// 消息处理工具类（单例）
/// </summary>
public class MessageUtil
{
    private static MessageUtil? _instance;
    private static readonly object Lock = new object();
    private readonly ElementUtil _util = new ElementUtil();
    private string? _savePath;

    /// <summary>
    /// 获取单例实例
    /// </summary>
    /// <returns>MessageUtil实例</returns>
    public static MessageUtil Get()
    {
        if (_instance != null) return _instance;
        lock (Lock)
        {
            _instance ??= new MessageUtil();
        }
        return _instance;
    }

    /// <summary>
    /// 用于设置文件保存路径
    /// </summary>
    /// <param name="path">保存路径</param>
    /// <param name="isRelative">是否是相对路径</param>
    public void SetSavePath(string path = ".\\data", bool isRelative = true)
    {
        _savePath = isRelative ? Path.Combine(Directory.GetCurrentDirectory(), path) : path;
    }

    /// <summary>
    /// 处理消息对象，返回WMsgSourceType（消息生成者类型、也叫源类型——人/系统）
    /// </summary>
    /// <param name="msgElement">消息对象</param>
    /// <returns>消息生成者类型</returns>
    private WMsgSourceType HandleMsgSourceType(AutomationElement msgElement) => _util.IsHumanMsg(msgElement) ? WMsgSourceType.Human : WMsgSourceType.System;

    /// <summary>
    /// 处理消息对象，返回WMsgSenderType（消息发送者类型——自己/好友）
    /// </summary>
    /// <param name="msgElement">消息对象</param>
    /// <param name="chatWindow">来源窗口</param>
    /// <remarks>窗口的来源参考ElementUtil的实现</remarks>
    /// <see cref="ElementUtil"/>
    /// <returns>消息发送者类型</returns>
    private WMsgSenderType HandleMsgSenderType(AutomationElement msgElement, Window chatWindow) => _util.IsSelfAvatar(
        _util.GetFirstElementFromGiveItem(msgElement, XPath.UserWindowUserAvatarButtonBaseOnMsgItem)!,
        chatWindow
        ) ? WMsgSenderType.Self : WMsgSenderType.Friend;

    /// <summary>
    /// 处理消息对象，返回WMsgOriginType（消息来源类型——单人/群）
    /// </summary>
    /// <remarks>
    /// 原理是检测消息带的用户备注名和聊天窗口名是否一样：
    /// 一样则是私聊消息，不一样则是群
    /// 有潜在Bug就是可能群内有人真拿群名当自己名字，或者你给的备注和群名一样
    /// 那就真炸了
    ///
    /// 还有就是，这个方法对自己发的消息是无效的，检测到自己的消息时一直都是Group
    /// </remarks>
    /// <param name="msgElement">消息对象</param>
    /// <param name="chatWindow">来源窗口</param>
    /// <returns>消息来源类型</returns>
    private WMsgOriginType HandleMsgOriginType(AutomationElement msgElement, Window chatWindow) => msgElement.Name == chatWindow.Title ? WMsgOriginType.Single : WMsgOriginType.Group;

    /// <summary>
    /// 用于获取消息类型
    /// </summary>
    /// <param name="msgElement">消息对象</param>
    /// <exception cref="ArgumentException">未知消息类型，或收到了不支持的消息类型</exception>
    /// <returns>消息类型</returns>
    private WMsgType HandleMsgType(AutomationElement msgElement)
    {
        if (_util.GetFirstElementFromGiveItem(
                msgElement,
                XPath.TextOfOtherTextMsgBaseOnMsgItem
                ) is not null || _util.GetFirstElementFromGiveItem(
                    msgElement,
                    XPath.TextOfSelfTextMsgBaseOnMsgItem
                ) is not null
            ) return WMsgType.Text;
        switch (msgElement.Name)
        {
            case "[图片]":
                return WMsgType.Image;
            case "[视频]":
                return WMsgType.Video;
            case "[动画表情]":
                return WMsgType.Emoji;
            case "[文件]":
                return WMsgType.File;
            case "[聊天记录]":
                return WMsgType.MergeForward;
        }
        if (msgElement.Name.Contains("[语音]")) return WMsgType.Voice;
        if (
            (_util.GetFirstElementFromGiveItem(
                msgElement,
                XPath.MiniProgramCardSignOfMsgBaseOnSelfMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
                msgElement,
                XPath.MiniProgramCardSignOfMsgBaseOnOtherMsgItem
            )) is not null && (_util.GetFirstElementFromGiveItem(
                msgElement,
                XPath.MiniProgramCardSignOfMsgBaseOnSelfMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
                msgElement,
                XPath.MiniProgramCardSignOfMsgBaseOnOtherMsgItem
            ))!.Name != "" && msgElement.Name != "微信转账"
        ) return WMsgType.MiniProgramCard;
        if (
            (_util.GetFirstElementFromGiveItem(
                msgElement,
                XPath.TransferSignOfMsgBaseOnSelfMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
                msgElement,
                XPath.TransferSignOfMsgBaseOnOtherMsgItem
            )) is not null
        ) return WMsgType.Transfer;
        if (Regex.Match(msgElement.Name, @".*\n引用.*的消息.*").Success == true)
        {
            return WMsgType.Quote;
        }
        throw new ArgumentException("未知消息类型，或收到了不支持的消息类型");
    }

    /// <summary>
    /// 从小程序卡片中获取内容（标题和描述）
    /// </summary>
    /// <param name="msgElement">消息对象</param>
    /// <returns>List[0]标题、List[1]描述</returns>
    private List<string> GetContentByMiniProgramCard(AutomationElement msgElement)
    {
        var res = new List<string>();
        res.Add(
            (_util.GetFirstElementFromGiveItem(
                msgElement, XPath.MiniProgramTitleBaseOnOtherMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
                msgElement, XPath.MiniProgramTitleBaseOnSelfMsgItem
                ))!.Name
        );
        res.Add(
            (_util.GetFirstElementFromGiveItem(
                msgElement, XPath.MiniProgramContentBaseOnOtherMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
                msgElement, XPath.MiniProgramContentBaseOnSelfMsgItem
                ))!.Name
        );
        return res;
    }

    /// <summary>
    /// 获取转账消息的转账金额
    /// </summary>
    /// <param name="msgElement">信息对象</param>
    /// <returns>转账金额</returns>
    private string GetTransferNumber(AutomationElement msgElement)
    {
        return (
            _util.GetFirstElementFromGiveItem(
                msgElement, XPath.TransferNumberBaseOnSelfMsgItem
            ) ?? _util.GetFirstElementFromGiveItem(
                msgElement, XPath.TransferNumberBaseOnOtherMsgItem
            )
        )!.Name;
    }

    /// <summary>
    /// 获取引用消息的内容
    /// </summary>
    /// <param name="msgElement"></param>
    /// <returns>[0]是消息内容，[1]是引用消息的内容，[3]是发送者的备注名</returns>
    private List<string> GetQuoteCotent(AutomationElement msgElement)
    {
        var quoteUser = _util.GetFirstElementFromGiveItem(
            msgElement, XPath.QuoteUserBaseOnMsgItem
        )!.Name.Split(" : ", count:2)[0];
        var res = new List<string>();
        res.AddRange(
            msgElement.Name.Split("\n引用  的消息 : ")
            );
        res.Add(quoteUser);
        return res;
    }

    /// <summary>
    /// 从表情消息对象中获取位置截图并获取Base64编码
    /// </summary>
    /// <param name="msgElement">消息对象</param>
    /// <returns>截图的Base64</returns>
    private string FromEmojiGetBase64(AutomationElement msgElement)
    {
        using var bitmap = FlaUI.Core.Capturing.Capture.Element(msgElement);
        using var memoryStream = new MemoryStream();
        bitmap.Bitmap.Save(memoryStream, ImageFormat.Png);
        var imageBytes = memoryStream.ToArray();
        return Convert.ToBase64String(imageBytes);
    }

    /// <summary>
    /// 根据消息类型清理消息到方便处理的文本列表内容形式
    /// </summary>
    /// <param name="chatWindow">聊天消息来源窗口</param>
    /// <param name="msgElement">消息对象</param>
    /// <param name="msgType">消息类型</param>
    /// <returns>消息内容列表</returns>
    /// <exception cref="ArgumentOutOfRangeException">不支持的消息类型</exception>
    private List<string> CleanUpMsgByType(Window chatWindow, AutomationElement msgElement, WMsgType msgType)
    {
        return msgType switch
        {
            WMsgType.Text => [msgElement.Name],
            WMsgType.File or WMsgType.Image =>
            [
                ActionUtil.Get().SaveAsAction(chatWindow, msgElement, _savePath!)
            ],
            WMsgType.Video => [ActionUtil.Get().SaveAsAction(chatWindow, msgElement, _savePath!, true)],
            WMsgType.Emoji => [FromEmojiGetBase64(msgElement)],
            WMsgType.MiniProgramCard => GetContentByMiniProgramCard(msgElement),
            WMsgType.MergeForward => ActionUtil.Get().OpenMergeForwardAction(msgElement),
            WMsgType.Voice => [ActionUtil.Get().VoiceToTextAction(chatWindow, msgElement)],
            WMsgType.Transfer => [GetTransferNumber(msgElement)],
            WMsgType.Quote => GetQuoteCotent(msgElement),
            _ => throw new ArgumentOutOfRangeException(nameof(msgType), msgType, null)
        };
    }

    /// <summary>
    /// 基于消息对象获取消息发送者的备注名
    /// </summary>
    /// <param name="msgElement">消息对象</param>
    /// <returns>发送者的备注名</returns>
    private string GetMsgSenderName(AutomationElement msgElement)
    {
        return _util.GetFirstElementFromGiveItem(
            msgElement, XPath.UserWindowUserAvatarButtonBaseOnMsgItem
        )!.Name;
    }

    /// <summary>
    /// 处理消息到WMessage对象
    /// </summary>
    /// <param name="chatWindow">消息来源窗口</param>
    /// <param name="msgElement">消息对象</param>
    /// <returns>WMessage信息</returns>
    public WMessage HandleMsg(Window chatWindow, AutomationElement msgElement)
    {
        var msgSourceType = HandleMsgSourceType(msgElement);
        if (msgSourceType == WMsgSourceType.Human)
        {
            var msgOriginType = HandleMsgOriginType(msgElement, chatWindow);
            var msgSenderType = HandleMsgSenderType(msgElement, chatWindow);
            var msgType = HandleMsgType(msgElement);
            var msgContent = CleanUpMsgByType(chatWindow, msgElement, msgType);
            return new WMessage
            {
                MsgType = msgType,
                MsgOriginType = msgOriginType,
                MsgSenderType = msgSenderType,
                MsgSourceType = msgSourceType,
                MsgContent = msgContent,
                MsgFromWindow = chatWindow.Name,
                MsgSenderName = GetMsgSenderName(msgElement)
            };
        }
        // Todo: 这里可以扩展一下系统消息的处理
        return new WMessage
        {
            MsgType = WMsgType.Text,
            MsgOriginType = WMsgOriginType.Single,
            MsgSenderType = WMsgSenderType.Self,
            MsgSourceType = msgSourceType,
            MsgContent = [msgElement.Name == "" ? FromMsgElementGetPypMsg(msgElement) : msgElement.Name],
            MsgFromWindow = chatWindow.Name,
            MsgSenderName = "系统"
        };
    }

    /// <summary>
    /// 从系统消息对象中获取拍一拍消息内容
    /// </summary>
    /// <param name="msgElement">系统消息对象</param>
    /// <returns>拍一拍消息内容</returns>
    private string FromMsgElementGetPypMsg(AutomationElement msgElement)
    {
        return _util.GetFirstElementFromGiveItem(msgElement, XPath.PypMsgBaseOnMsgItem)!.Name;
    }
}
