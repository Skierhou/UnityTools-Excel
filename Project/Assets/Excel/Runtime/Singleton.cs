public abstract class Singleton<T>
{
    private static T _ServiceContext;
    private readonly static object lockObj = new object();

    /// <summary>
    /// 禁止外部进行实例化
    /// </summary>
    protected Singleton()
    {
        Initialize();
    }

    /// <summary>
    /// 获取唯一实例，双锁定防止多线程并发时重复创建实例
    /// </summary>
    /// <returns></returns>
    public static T Instance
    {
        get
        {
            if (_ServiceContext == null)
            {
                lock (lockObj)
                {
                    if (_ServiceContext == null)
                    {
                        _ServiceContext = (T)System.Activator.CreateInstance(typeof(T));
                    }
                }
            }
            return _ServiceContext;
        }
    }
    protected virtual void Initialize() { }
}