using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using UnityEngine;

namespace Excel
{
    public partial class ExcelManager: Singleton<ExcelManager>
    {
        private Dictionary<string, IContainer> m_InfosMap = null;
        private IHelper m_Helper;

        protected override void Initialize()
        {
            m_InfosMap = new Dictionary<string, IContainer>();
            m_Helper = new Helper();

            Stopwatch sw = new Stopwatch();
            sw.Start();
            IEnumerator<Type> it = GetInterfaceType((typeof(IMainKey))).GetEnumerator();
            while (it.MoveNext())
            {
                m_Helper.ReadData(it.Current.Name);
            }
            sw.Stop();
            UnityEngine.Debug.Log(string.Format("加载所有Excel用时：{0} ms", sw.ElapsedMilliseconds));
        }
        /// <summary>
        /// 通过key获取对象
        /// </summary>
        public T Access<T>(int inKey) where T : IMainKey
        {
            if (m_InfosMap.TryGetValue(typeof(T).Name, out IContainer container))
            {
                return (T)container.Get(inKey);
            }
            return default(T);
        }
        /// <summary>
        /// 获取该表所有数据
        /// </summary>
        public List<T> AccessAll<T>() where T : IMainKey
        {
            if (m_InfosMap.TryGetValue(typeof(T).Name, out IContainer container))
            {
                return (List<T>)container.GetAll();
            }
            return null;
        }
        /// <summary>
        /// 读取数据后添加至集合
        /// </summary>
        public void AddData(string inKey, IContainer inContainer)
        {
            if (m_InfosMap.ContainsKey(inKey))
                UnityEngine.Debug.LogWarning(string.Format("Excel add fail！ because of contains same key by {0}", inKey));
            else
                m_InfosMap.Add(inKey, inContainer);
        }
    }
}
