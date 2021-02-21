using Excel;
using System.Collections.Generic;

    /// <summary>
    /// 对应Excel配置类必须继承自IMainKey接口，这个ID为数据主键，不允许重复
    /// </summary>
    [System.Serializable]
    public class TestData : IMainKey
    {
        public int id;
        public string name;
        public TestType testType;
        public List<int> testList;
        public List<GGGG> gggList;

        public int ID => id;
    }

    public enum TestType
    {
        T1,
        T2,
        T3,
        T4
    }

    [System.Serializable]
    public class GGGG
    {
        public int count;
        public int money;
        public string name;
    }

