using UnityEngine;
using UnityEditor;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using UnityQuickSheet;

///
/// !!! Machine generated code !!!
///
[CustomEditor(typeof(ItemSO))]
public class ItemSOEditor : BaseExcelEditor<ItemSO>
{	    
    public override bool Load()
    {
        ItemSO targetData = target as ItemSO;

        string path = targetData.SheetName;
        if (!File.Exists(path))
            return false;

        string sheet = targetData.WorksheetName;

        ExcelQuery query = new ExcelQuery(path, 0);
        if (query != null && query.IsValid())
        {
            targetData.dataArray = query.Deserialize<ItemExcelData>(3).ToArray();
            EditorUtility.SetDirty(targetData);
            AssetDatabase.SaveAssets();
            return true;
        }
        else
            return false;
    }
}
