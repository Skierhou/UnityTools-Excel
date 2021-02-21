using System;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using UnityEngine;

namespace Excel
{
    public class ExportExcel : ScriptableWizard
    {
        private readonly List<Type> dataTypes = null;
        private readonly string[] dataShowNames = null;
        private readonly int[] dataOptions = null;
        private string outJsonPath = null;
        private string outBytePath = null;
        private int selectIndex = 0;
        public ExportExcel()
        {
            IEnumerator<Type> it = ExcelManager.GetInterfaceType(typeof(IMainKey)).GetEnumerator();
            dataTypes = new List<Type>();
            while (it.MoveNext())
            {
                dataTypes.Add(it.Current);
            }
            dataShowNames = new string[dataTypes.Count];
            dataOptions = new int[dataTypes.Count];
            for (int i = 0; i < dataTypes.Count; i++)
            {
                dataShowNames[i] = dataTypes[i].Name;
                dataOptions[i] = i;
            }
            helpString = "数据类型需实现ExcelComponent.IMainKey接口！";
            minSize = new Vector2(600, 500);
            outJsonPath = Path.GetFullPath(ExcelHelper.JSONPATH);
            outBytePath = Path.GetFullPath(ExcelHelper.BYTEPATH);
        }

        //开启窗口或数据更新时调用
        void OnWizardUpdate()
        {
        }
        //当用户按下"应用"时被调用,保存设置但不关闭窗口
        void OnWizardOtherButton()
        {
            //导出所有
            for (int i = 0; i < dataTypes.Count; i++)
            {
                EditorUtility.DisplayProgressBar("正在导出...", string.Format("数据类型：{0}", dataTypes[i].Name), i * 1.0f / dataTypes.Count);
                ExcelHelper.ExcelToJson(dataTypes[i], outJsonPath, outBytePath);
            }
            AssetDatabase.Refresh();
            EditorUtility.ClearProgressBar();
        }
        //点击"确定"时调用，关闭窗口并保存设置
        void OnWizardCreate()
        {
        }

        protected override bool DrawWizardGUI()
        {
            bool bSelectPath = false;
            EditorGUILayout.Space();
            EditorGUILayout.BeginVertical();
            if (dataShowNames != null)
            {
                EditorGUI.indentLevel = 1;
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField("选择数据：");
                selectIndex = EditorGUILayout.IntPopup(selectIndex, dataShowNames, dataOptions);
                EditorGUILayout.EndHorizontal();

                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField(string.Format("导出Json路径：{0}", outJsonPath));
                if (bSelectPath |= EditorGUILayout.DropdownButton(new GUIContent("点击选择Json目录"), FocusType.Keyboard))
                {
                    string newPath = EditorUtility.SaveFolderPanel("请选择保存路径", outJsonPath, "");
                    if (!string.IsNullOrEmpty(newPath))
                    {
                        outJsonPath = newPath;
                    }
                }
                if (!bSelectPath)
                    EditorGUILayout.EndHorizontal();

                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField(string.Format("导出Bytes路径：{0}", outBytePath));
                if (!bSelectPath && (bSelectPath |= EditorGUILayout.DropdownButton(new GUIContent("点击选择Bytes目录"), FocusType.Keyboard)))
                {
                    string newPath = EditorUtility.SaveFolderPanel("请选择保存路径", outBytePath, "");
                    if (!string.IsNullOrEmpty(newPath))
                    {
                        outBytePath = newPath;
                    }
                }
                if (!bSelectPath)
                    EditorGUILayout.EndHorizontal();

                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField(string.Format("解析单个Excel：{0}", dataTypes[selectIndex]));
                if (EditorGUILayout.DropdownButton(new GUIContent("点击解析"), FocusType.Keyboard))
                {
                    bool bSuccess = ExcelHelper.ExcelToJson(dataTypes[selectIndex], outJsonPath, outBytePath);
                    string message = bSuccess ? string.Format("导出成功：{0}\nJson：{1}\nByte：{2}", dataTypes[selectIndex].Name, outJsonPath, outBytePath) : "导出失败，请查看Log！";
                    EditorUtility.DisplayDialog(bSuccess ? "导出成功！" : "导出失败！", message, "确定");
                    AssetDatabase.Refresh();
                }
                EditorGUILayout.EndHorizontal();
            }
            if (!bSelectPath)
                EditorGUILayout.EndVertical();
            EditorGUILayout.Space();
            return base.DrawWizardGUI();
        }
    }

}