using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System;
using System.IO;

namespace Excel
{
    public class CreateExcel : ScriptableWizard
    {
        private readonly List<Type> dataTypes = null;
        private readonly string[] dataShowNames = null;
        private readonly int[] dataOptions = null;
        private string outPath = null;
        private int selectIndex = 0;
        private string saveName = null;
        public CreateExcel()
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
            outPath = Path.GetFullPath(ExcelHelper.EXCELPATH);
        }

        //开启窗口或数据更新时调用
        void OnWizardUpdate()
        {
        }
        //当用户按下"应用"时被调用,保存设置但不关闭窗口
        void OnWizardOtherButton()
        {
            bool bSuccess = ExcelHelper.TypeToExcel(dataTypes[selectIndex], saveName, outPath);
            string message = bSuccess ?
                string.Format("导出成功：{0}\\{1}@{2}{3}", outPath, saveName, dataTypes[selectIndex].Name, ExcelHelper.EXCELEXTENSION)
                : "导出失败，请查看Log！";
            EditorUtility.DisplayDialog(bSuccess ? "导出成功！" : "导出失败！", message, "确定");
            AssetDatabase.Refresh();
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
                EditorGUILayout.LabelField(string.Format("Excel创建路径：{0}", outPath));
                if (bSelectPath |= EditorGUILayout.DropdownButton(new GUIContent("目录"), FocusType.Keyboard))
                {
                    string newPath = EditorUtility.SaveFolderPanel("请选择保存路径", outPath, "");
                    if (!string.IsNullOrEmpty(newPath))
                    {
                        outPath = newPath;
                    }
                }
                if (!bSelectPath)
                    EditorGUILayout.EndHorizontal();
                saveName = EditorGUILayout.TextField("Excel名称：", saveName == null ? dataShowNames[selectIndex] : saveName);
            }
            if (!bSelectPath)
                EditorGUILayout.EndVertical();
            EditorGUILayout.Space();
            return base.DrawWizardGUI();
        }
    }
}
