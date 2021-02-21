using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using UnityEngine;

namespace Excel
{
    public partial class ExcelManager
    {
        private interface IHelper
        {
            void ReadData(string typeName);
        }
        private sealed class Helper:IHelper
        {
            //运行时读取路径
            private const string m_ReadPath = "Assets/Excel/Data/Byte/{0}{1}";
            private const string m_ReadExtension = ".byte";

            /// <summary>
            /// 获取存放路径
            /// </summary>
            private string GetDataPath(string typeName)
            {
                return string.Format(m_ReadPath, typeName, m_ReadExtension);
            }

            public void ReadData(string typeName)
            {
                string assetPath = GetDataPath(typeName);
#if UNITY_EDITOR
                string path = Path.GetFullPath(string.Format("{0}/../{1}", Application.dataPath, assetPath));
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    BinaryFormatter bf = new BinaryFormatter();
                    IContainer value = bf.Deserialize(fs) as IContainer;
                    ExcelManager.Instance.AddData(typeName, value);
                }
#endif
            }
        }
    }
}
