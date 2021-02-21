using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System;
using System.Reflection;
using OfficeOpenXml;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using LitJson;

namespace Excel
{
    public class ExcelHelper
    {
        // Excel路径
        public static string EXCELPATH = Application.streamingAssetsPath + "/../Excel/Data/Excel";
        // Json路径
        public static string JSONPATH = Application.streamingAssetsPath + "/../Excel/Data/Json";
        // Byte路径
        public static string BYTEPATH = Application.streamingAssetsPath + "/../Excel/Data/Byte";
        // Excel后缀
        public static string EXCELEXTENSION = ".xlsx";
        // Xml后缀
        public static string JSONEXTENSION = ".json";
        // byte后缀
        public static string BYTEEXTENSION = ".byte";
        // 区分每列颜色
        private static List<System.Drawing.Color> TitleColors = new List<System.Drawing.Color>
        {
            System.Drawing.Color.MediumPurple,
            System.Drawing.Color.MediumSeaGreen,
            System.Drawing.Color.MidnightBlue,
            System.Drawing.Color.Orange,
            System.Drawing.Color.MediumVioletRed
        };
        private static System.Drawing.Color NoteColor = System.Drawing.Color.DarkRed;
        private static System.Drawing.Color TypeColor = System.Drawing.Color.DarkRed;

        #region 按类型生成Excel
        [MenuItem("Assets/配置表EX/类转Excel")]
        public static void ClassToExcel()
        {
            UnityEngine.Object[] objs = Selection.objects;
            for (int i = 0; i < objs.Length; i++)
            {
                TypeToExcel(GetType(Path.GetFileNameWithoutExtension(objs[i].name)), objs[i].name + "1", EXCELPATH);
            }
            AssetDatabase.Refresh();
        }
        public static bool TypeToExcel(Type inType, string inExcelName, string inPath)
        {
            string excelPath = Path.GetFullPath(string.Format("{0}/{1}@{2}{3}", inPath, inExcelName, inType.Name, EXCELEXTENSION));

            if (File.Exists(excelPath))
            {
                if (!EditorUtility.DisplayDialog("存在重复表！", "是否覆盖", "确认", "取消"))
                {
                    return false;
                }
                File.Delete(excelPath);
            }
            try
            {
                using (FileStream fs = new FileStream(excelPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    using (ExcelPackage package = new ExcelPackage(fs))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(inType.Name);
                        TypeToWorksheet(inType, worksheet);
                        package.Save();
                    }
                }
            }
            catch (Exception e)
            {
                Debug.LogError(string.Format("表：{0}导出失败！Exception：{1}", excelPath, e));
                return false;
            }
            Debug.Log(string.Format("表：{0}导出成功！", excelPath));
            return true;
        }
        public static void TypeToWorksheet(Type inType, ExcelWorksheet inWorksheet, int inStartCol = 1)
        {
            FieldInfo[] fieldInfos = inType.GetFields();

            int colorIndex = 0;
            for (int i = 0; i < fieldInfos.Length; i++)
            {
                // List Class
                if (fieldInfos[i].FieldType.IsGenericType && fieldInfos[i].FieldType.GetGenericTypeDefinition() == typeof(List<>))
                {
                    //重新一张表
                    SetCell(inWorksheet.Cells[1, i + inStartCol], fieldInfos[i].Name, TitleColors[colorIndex++ % TitleColors.Count]);
                    SetCell(inWorksheet.Cells[2, i + inStartCol], "类型：Int", TypeColor);
                    SetCell(inWorksheet.Cells[3, i + inStartCol], string.Format("注释：关联{0}表的ID", fieldInfos[i].Name), NoteColor);

                    ExcelWorksheet listWorksheet = inWorksheet.Workbook.Worksheets[fieldInfos[i].Name];
                    if (listWorksheet == null)
                    {
                        listWorksheet = inWorksheet.Workbook.Worksheets.Add(fieldInfos[i].Name);
                    }
                    ListToWorksheet(fieldInfos[i].FieldType, listWorksheet, string.Format("关联表{0}的{1}", inWorksheet.Name, fieldInfos[i].Name));
                }
                else if (fieldInfos[i].FieldType.IsEnum)
                {
                    var validation = inWorksheet.DataValidations.AddListValidation(inWorksheet.Cells[4, i + inStartCol, 5000, i + inStartCol].Address);
                    validation.Error = "枚举类型请按下拉框！";
                    validation.ErrorTitle = "请勿输入！";
                    validation.ShowErrorMessage = true;
                    validation.ShowInputMessage = true;
                    string[] enums = Enum.GetNames(fieldInfos[i].FieldType);
                    for (int j = 0; j < enums.Length; j++)
                    {
                        validation.Formula.Values.Add(enums[j]);
                    }
                    SetCell(inWorksheet.Cells[1, i + inStartCol], fieldInfos[i].Name, TitleColors[colorIndex++ % TitleColors.Count]);
                    SetCell(inWorksheet.Cells[2, i + inStartCol], string.Format("类型：{0}", fieldInfos[i].FieldType.Name), TypeColor);
                    SetCell(inWorksheet.Cells[3, i + inStartCol], "注释：", NoteColor);
                }
                else
                {
                    SetCell(inWorksheet.Cells[1, i + inStartCol], fieldInfos[i].Name, TitleColors[colorIndex++ % TitleColors.Count]);
                    SetCell(inWorksheet.Cells[2, i + inStartCol], string.Format("类型：{0}", fieldInfos[i].FieldType.Name), TypeColor);
                    SetCell(inWorksheet.Cells[3, i + inStartCol], "注释：", NoteColor);
                }
            }
            using (ExcelRange range = inWorksheet.Cells[1, 1, 1, fieldInfos.Length])
            {
                range.AutoFilter = true;
                range.Style.Font.Bold = true;
                range.Style.Font.Size += 5;
            }
        }
        private static void ListToWorksheet(Type inType, ExcelWorksheet inWorksheet, string inTip = "", int inStartCol = 1)
        {
            int colorIndex = 0;

            SetCell(inWorksheet.Cells[1, inStartCol], "ID", TitleColors[colorIndex++ % TitleColors.Count]);
            SetCell(inWorksheet.Cells[2, inStartCol], "类型：Int", TypeColor);
            SetCell(inWorksheet.Cells[3, inStartCol], string.Format("注释：{0}", inTip), NoteColor);

            Type genericType = inType.GenericTypeArguments[0];
            if (genericType.IsClass && genericType != typeof(string))
            {
                TypeToWorksheet(genericType, inWorksheet, 2);
            }
            else
            {
                SetCell(inWorksheet.Cells[1, 1 + inStartCol], genericType.Name, TitleColors[colorIndex++ % TitleColors.Count]);
                SetCell(inWorksheet.Cells[2, 1 + inStartCol], string.Format("类型：{0}", genericType.Name), TypeColor);
                SetCell(inWorksheet.Cells[3, 1 + inStartCol], "注释：", NoteColor);
            }
        }
        private static void SetCell(ExcelRange inCell, string inContent, System.Drawing.Color inColor)
        {
            inCell.Value = inContent;
            inCell.Style.Font.Color.SetColor(inColor);
            inCell.AutoFitColumns(30);
        }
        #endregion

        #region Excel转bytes,json

        private static Dictionary<ExcelWorksheet, Dictionary<int, object>> m_CacheMap;
        public static bool ExcelToJson(Type inType, string inJsonPath, string inBytePath)
        {
            if (Directory.Exists(EXCELPATH))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(EXCELPATH);
                FileInfo[] fileInfos = directoryInfo.GetFiles();

                List<IContainer> containers = new List<IContainer>();
                foreach (FileInfo fileInfo in fileInfos)
                {
                    if (fileInfo.Name.Contains(inType.Name))
                    {
                        containers.Add(ExcelToJson(fileInfo.FullName));
                    }
                }
                //合并container
                for (int i = 1; i < containers.Count; ++i)
                {
                    containers[0].Combine(containers[i]);
                }
                inJsonPath = string.Format("{0}/{1}{2}", inJsonPath, inType.Name, JSONEXTENSION);
                inBytePath = string.Format("{0}/{1}{2}", inBytePath, inType.Name, BYTEEXTENSION);
                JsonSerialize(inJsonPath, containers[0]);
                BinarySerialize(inBytePath, containers[0]);
                Debug.Log(string.Format("Json：{0}，Byte：{1}导出成功！", inJsonPath, inBytePath));
                return true;
            }
            return false;
        }

        private static IContainer ExcelToJson(string inExcelPath)
        {
            if (Path.GetExtension(inExcelPath) != EXCELEXTENSION) return null;
            string excelName = Path.GetFileNameWithoutExtension(inExcelPath);
            string typeStr = excelName;
            if (typeStr.Contains("@"))
            {
                typeStr = typeStr.Split('@')[1];
            }
            Type type = GetType(typeStr);

            Type containerType = typeof(Container<>);
            containerType = containerType.MakeGenericType(new Type[] { type });
            //Type containerType = typeof(ContainerExcel);
            object container = null;
            if (type != null && containerType != null)
            {
                try
                {
                    using (FileStream fs = new FileStream(inExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (ExcelPackage package = new ExcelPackage(fs))
                        {
                            ExcelWorksheets excelWorksheets = package.Workbook.Worksheets;

                            m_CacheMap = new Dictionary<ExcelWorksheet, Dictionary<int, object>>();
                            object instances = WorksheetToList(excelWorksheets[type.Name], type);
                            m_CacheMap.Clear();
                            m_CacheMap = null;

                            container = System.Activator.CreateInstance(containerType, new object[] { instances });
                        }
                    }
                }
                catch (Exception e)
                {
                    Debug.LogError(string.Format("表：{0}导出失败！Exception：{1}", inExcelPath, e));
                }
            }

            return container == null ? null : container as IContainer;
        }

        private static object WorksheetToList(ExcelWorksheet worksheet, Type inType, int inStartCol = 1, int inID = 0)
        {
            FieldInfo[] fieldInfos = inType.GetFields();

            //判断当前为子表
            bool isChildList = inStartCol == 2;
            if (isChildList)
            {
                if (m_CacheMap.TryGetValue(worksheet, out Dictionary<int, object> tMap)
                    && tMap != null && tMap.TryGetValue(inID, out object cacheValue))
                {
                    return cacheValue;
                }
            }

            object instances = CreateList(inType);
            MethodInfo addMethod = instances.GetType().GetMethod("Add", BindingFlags.Instance | BindingFlags.Public);

            //第4行开始读 前三行是注释等内容
            for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
            {
                object instance = null;
                if (inType.IsClass)
                    instance = System.Activator.CreateInstance(inType);
                for (int col = 0; col < fieldInfos.Length; col++)
                {
                    ExcelRange range = worksheet.Cells[row, col + inStartCol];
                    if (range != null && range.Value != null)
                    {
                        string valueStr = range.Value.ToString().Trim();
                        object value;
                        if (fieldInfos[col].FieldType.IsGenericType && fieldInfos[col].FieldType.GetGenericTypeDefinition() == typeof(List<>))
                        {
                            ExcelWorksheet newWorksheet = worksheet.Workbook.Worksheets[fieldInfos[col].Name];
                            int referenceId = int.Parse(valueStr);
                            value = WorksheetToList(newWorksheet, fieldInfos[col].FieldType.GenericTypeArguments[0], 2, referenceId);
                        }
                        else
                        {
                            if (fieldInfos[col].FieldType.IsEnum)
                            {
                                value = Enum.Parse(fieldInfos[col].FieldType, valueStr);
                            }
                            else
                            {
                                if (fieldInfos[col].FieldType != typeof(string))
                                    value = System.Convert.ChangeType(valueStr, fieldInfos[col].FieldType);
                                else
                                    value = valueStr;
                            }
                        }

                        if (inType.IsClass)
                            fieldInfos[col].SetValue(instance, value);
                        else
                            instance = value;
                    }
                }

                //是子表则缓存一份数据
                if (isChildList)
                {
                    if (!m_CacheMap.ContainsKey(worksheet))
                        m_CacheMap.Add(worksheet, new Dictionary<int, object>());
                    if (m_CacheMap.TryGetValue(worksheet, out Dictionary<int, object> tMap)
                        && !tMap.ContainsKey(inID))
                        tMap.Add(inID, CreateList(inType));

                    if (m_CacheMap.TryGetValue(worksheet, out tMap)
                        && tMap.TryGetValue(inID, out object cacheList))
                    {
                        IList list = cacheList as IList;
                        list.Add(instance);
                    }
                    if (inID != int.Parse(worksheet.Cells[row, 1].Value.ToString().Trim()))
                        continue;
                }
                if (addMethod != null)
                    addMethod.Invoke(instances, new object[] { instance });
            }
            return instances;
        }
        #endregion

        #region 序列化
        public static void JsonSerialize(string path, object instance)
        {
            try
            {
                string jsonStr = JsonMapper.ToJson(((IContainer)instance).GetAll() as IList);
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        sw.Write(jsonStr);
                    }
                }
            }
            catch (Exception e)
            {
                Debug.LogError("找不到路径：" + path + "， 错误：" + e);
            }
        }
        public static bool BinarySerialize(string path, System.Object obj)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    BinaryFormatter bf = new BinaryFormatter();
                    bf.Serialize(fs, obj);
                }
                return true;
            }
            catch (Exception e)
            {
                Debug.LogError("无法完成Binary序列化：" + path + "，错误：" + e);
                if (File.Exists(path)) File.Delete(path);
            }
            return false;
        }
        #endregion

        #region Running 读取bytes,json
        public static T ReadByte<T>(Type inType) where T : class
        {
            T t = null;
            string path = string.Format("{0}/{1}{2}", BYTEPATH, inType.Name, BYTEEXTENSION);
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                BinaryFormatter bf = new BinaryFormatter();
                t = (T)bf.Deserialize(fs);
            }
            return t;
        }
        public static T ReadJson<T>(Type inType) where T : class
        {
            T t = null;
            string path = string.Format("{0}/{1}{2}", JSONPATH, inType.Name, JSONEXTENSION);
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (StreamReader sr = new StreamReader(fs))
                {
                    t = JsonUtility.FromJson<T>(sr.ReadToEnd());
                }
            }
            return t;
        }
        //[MenuItem("Assets/配置表EX/测试读取")]
        //public static void ReadTest()
        //{
        //    ExcelComponent.ContainerExcel testData = ReadByte<ExcelComponent.ContainerExcel>(typeof(TestData));
        //    Debug.Log(testData.m_Infos.Count);

        //    //testData = ReadJson<TestData_Container>();
        //    //Debug.Log(testData.Infos.Count);
        //}
        #endregion


        /// <summary>
        /// 获取类型
        /// </summary>
        private static Type GetType(string name)
        {
            Type type = null;
            Assembly[] assemblies = AppDomain.CurrentDomain.GetAssemblies();   //获取当前所有运行的程序集
            foreach (Assembly temp in assemblies)
            {
                //遍历每个程序集，尝试获取是否有name对应的类
                Type tempType = temp.GetType(name);
                if (tempType != null && tempType != tempType.BaseType)
                {
                    type = tempType;
                    break;
                }
            }
            return type;
        }
        private static object CreateList(Type type)
        {
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
            {
                return Activator.CreateInstance(type);
            }
            Type listType = typeof(List<>);
            Type newType = listType.MakeGenericType(new Type[] { type });
            object list = Activator.CreateInstance(newType);
            return list;
        }
    }
}
