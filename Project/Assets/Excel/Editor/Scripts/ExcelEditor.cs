using System;
using UnityEditor;

namespace Excel
{
    public class ExcelEditor
    {

        [MenuItem("Tools/Excel/Create Table...")]
        //弹出窗口调用
        static void CreateWizard()
        {
            ScriptableWizard.DisplayWizard<CreateExcel>("Create Excel", "取消", "创建");
        }

        [MenuItem("Tools/Excel/Export Table...")]
        //弹出窗口调用
        static void ExportWizard()
        {
            ScriptableWizard.DisplayWizard<ExportExcel>("Export Excel", "取消", "导出所有Excel");
        }
    }
}