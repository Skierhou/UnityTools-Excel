using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace Excel
{
    public interface IMainKey
    {
        int ID { get; }
    }

    public interface IContainer
    {
        /// <summary>
        /// 按主键获得数据
        /// </summary>
        /// <returns>T</returns>
        object Get(int inKey);
        /// <summary>
        /// 获得所有数据
        /// </summary>
        /// <returns>List<T></returns>
        object GetAll();
        /// <summary>
        /// 用于Editor模式下Excel解析合并
        /// </summary>
        void Combine(IContainer container);
    }
    [System.Serializable]
    public class ContainerExcel : IContainer
    {
        public IList m_Infos;
        public IDictionary<int, object> m_InfosMap;

        public ContainerExcel(IList infos)
        {
            m_Infos = infos;
            m_InfosMap = new Dictionary<int, object>();
            for (int i = 0; i < infos.Count; i++)
            {
                IMainKey mainKey = infos[i] as IMainKey;
                m_InfosMap.Add(mainKey.ID, infos[i]);
            }
        }

        public void Combine(IContainer container)
        {
            if (container == null) return;
            IList infos = container.GetAll() as IList;
            foreach (object item in infos)
            {
                IMainKey mainKey = item as IMainKey;
                if (m_InfosMap.ContainsKey(mainKey.ID))
                {
                    Debug.LogWarning(string.Format("导入数据Excel时出现主键相同情况，Type:{0}，已自动过滤！", m_Infos.GetType().GenericTypeArguments[0]));
                }
                else
                {
                    m_Infos.Add(mainKey);
                    m_InfosMap.Add(mainKey.ID, mainKey);
                }
            }
        }

        public object Get(int inKey)
        {
            m_InfosMap.TryGetValue(inKey, out object value);
            return value;
        }

        public object GetAll()
        {
            return m_Infos;
        }
    }
    [System.Serializable]
    public class Container<T> : IContainer where T : IMainKey
    {
        private List<T> m_Infos;
        private Dictionary<int, T> m_InfosMap;

        public Container(List<T> infos)
        {
            m_Infos = infos;
            m_InfosMap = new Dictionary<int, T>();
            for (int i = 0; i < infos.Count; i++)
            {
                m_InfosMap.Add(infos[i].ID, infos[i]);
            }
        }

        public object Get(int inKey)
        {
            m_InfosMap.TryGetValue(inKey, out T t);
            return t;
        }
        public object GetAll()
        {
            return m_Infos;
        }

        public void Combine(IContainer container)
        {
            if (container == null) return;
            IList<T> list = container.GetAll() as IList<T>;
            if (list != null)
            {
                foreach (T item in list)
                {
                    if (m_InfosMap.ContainsKey(item.ID))
                        Debug.LogWarning(string.Format("导入数据Excel时出现主键相同情况，Type:{0}，已自动过滤！", typeof(T)));
                    else
                    {
                        m_Infos.Add(item);
                        m_InfosMap.Add(item.ID, item);
                    }
                }
            }
        }
    }
}
