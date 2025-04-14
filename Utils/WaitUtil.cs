namespace WxConnectorLib.Utils;

public static class WaitUtil
{
    /// <summary>
    /// 用于等待直到方法成立
    /// </summary>
    /// <param name="action">传入需要的方法</param>
    /// <param name="waitOnce">单次等待的时间</param>
    /// <param name="timeout">等待超时时间</param>
    /// <typeparam name="T">需要获取的回调内容类型</typeparam>
    /// <returns>需要获取的回调内容</returns>
    public static T WaitUntil<T>(Func<(bool, T)> action, TimeSpan waitOnce, TimeSpan timeout)
    {
        var res = action();
        var startTime = DateTime.Now;
        while (!res.Item1)
        {
            if (DateTime.Now - startTime > timeout)
            {
                throw new TimeoutException("等待超时");
            }
            Thread.Sleep(waitOnce);
            res = action();
        }
        return res.Item2;
    }
}
