using FlaUI.Core.AutomationElements;
using WxConnectorLib.Models;
using WxConnectorLib.Utils;

namespace WxConnectorLib.Managers;

/// <summary>
///     监听器管理类（单例）
/// </summary>
public class ListenManager
{
    private static ListenManager? _instance;
    private static readonly object Lock = new();
    private readonly CancellationToken _cancellationToken = CancellationToken.None;
    private readonly Dictionary<string, List<AutomationElement>> _listenMsgRecord = new();
    private readonly Dictionary<string, Window> _listenRecord = new();
    private Dictionary<string, int> _listenSameIndexRecord = new();
    private bool PauseToken { get; set; } = false;

    /// <summary>
    ///     获取单例实例
    /// </summary>
    /// <returns>ListenManager实例</returns>
    public static ListenManager Get()
    {
        if (_instance != null) return _instance;
        lock (Lock)
        {
            _instance ??= new ListenManager();
        }

        return _instance;
    }

    /// <summary>
    ///     初始化监听
    /// </summary>
    /// <param name="listenWindos">需要监听的用户的独立聊天窗口</param>
    public void InitListen(List<Window> listenWindos)
    {
        // 先记录监听窗口
        foreach (var window in listenWindos)
        {
            _listenRecord.Add(window.Title, window);
            // 初始化一下消息记录
            var msgs = ElementUtil.Get().GetAllElementsFromGiveWindow(window, XPath.MsgItems);
            Console.WriteLine($"window: {window.Title}, msgs count: {msgs.Count}");
            _listenMsgRecord.Add(window.Title, msgs);
            _listenSameIndexRecord.Add(window.Title, msgs.Count - 1);
        }

        // 打开线程监听
        Task.Run(() =>
        {
            while (true)
            {
                try
                {
                    foreach (var item in _listenRecord.Values)
                    {
                        // 如果暂停，则等待暂停结束（因为某些操作会打断监听：如发送信息等）
                        WaitUtil.WaitUntil(
                            () => PauseToken ? (false, string.Empty) : (true, string.Empty),
                            TimeSpan.FromMilliseconds(200),
                            TimeSpan.MaxValue
                        );

                        item.Focus();
                        Thread.Sleep(200);
                        // 获取新消息
                        var msgs = ElementUtil.Get().GetAllElementsFromGiveWindow(item, XPath.MsgItems);
                        List<AutomationElement> newMsg;
                        if (_listenMsgRecord[item.Title].Count == 0)
                        {
                            // 处理如果没有消息的情况
                            newMsg = msgs;
                        }
                        else
                        {
                            // 计算索引
                            var sameMessage = msgs.Last(x => x.Name == _listenMsgRecord[item.Title].Last().Name);
                            var index = msgs.IndexOf(sameMessage);
                            // 获取新消息
                            newMsg = msgs.Skip(index + 1).ToList();

                            // 这里是为了修正上面这个差异分析法监听不到相同内容消息的bug
                            if (newMsg.Count == 0)
                            {
                                var sames = msgs.Where(
                                    x => x.Name == _listenMsgRecord[item.Title].Last().Name
                                ).Select(
                                    x => msgs.IndexOf(x)
                                    ).ToList();
                                var sameIndex = GetConsecutiveRanges(sames).Last();
                                if (sames.Count > 1 && sameIndex.Count > 1 && sameIndex.Last() > _listenSameIndexRecord[item.Title])
                                {
                                    newMsg = msgs.Skip(sameIndex.Last()).ToList();
                                    _listenSameIndexRecord[item.Title] = sameIndex.Last();
                                }
                            }
                        }

                        if (newMsg.Count == 0) continue;
                        // 监听到的消息会在这里处理

                        foreach (var outM in newMsg.Select(o => MessageUtil.Get().HandleMsg(item, o)))
                        {
                            EventManager.Get().InvokeNewMessage(outM, item);
                        }

                        // 更新消息记录
                        _listenMsgRecord[item.Title] = msgs;
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex);
                    throw;
                }

                Thread.Sleep(500);
            }
        }, _cancellationToken);
    }

    /// <summary>
    /// 用于获取连续的数字范围（获取消息中的索引连续）
    /// </summary>
    /// <param name="numbers">索引列表</param>
    /// <returns>索引连续</returns>
    private static List<List<int>> GetConsecutiveRanges(List<int> numbers)
    {
        var sortedNumbers = numbers.OrderBy(n => n).ToList();
        List<List<int>> consecutiveRanges = [];
        var currentRange = new List<int> { sortedNumbers[0] };
        for (var i = 1; i < sortedNumbers.Count; i++)
        {
            if (sortedNumbers[i] == sortedNumbers[i - 1] + 1)
            {
                currentRange.Add(sortedNumbers[i]);
            }
            else
            {
                consecutiveRanges.Add(currentRange);
                currentRange = [sortedNumbers[i]];
            }
        }
        consecutiveRanges.Add(currentRange);
        return consecutiveRanges;
    }

    /// <summary>
    /// 用于设置监听暂停
    /// </summary>
    public void PauseListen()
    {
        PauseToken = true;
    }

    /// <summary>
    /// 用于恢复监听
    /// </summary>
    public void ResumeListen()
    {
        PauseToken = false;
    }
}
