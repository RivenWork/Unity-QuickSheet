using UnityEngine;
using UnityEditor;
using System.Collections;
using System.IO;
using UnityQuickSheet;

///
/// !!! Machine generated code !!!
///
public class ItemAssetPostprocessor : AssetPostprocessor 
{
    private static readonly string filePath = "Assets/Excel/Source/Item.xlsx";
    private static readonly string assetFilePath = "Assets/Excel/Data/Item.asset";
    private static readonly string sheetName = "ItemSO";
    
    static void OnPostprocessAllAssets (string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
    {
        foreach (string asset in importedAssets) 
        {
            if (!filePath.Equals (asset))
                continue;
                
            ItemSO data = (ItemSO)AssetDatabase.LoadAssetAtPath (assetFilePath, typeof(ItemSO));
            if (data == null) {
                data = ScriptableObject.CreateInstance<ItemSO> ();
                data.SheetName = filePath;
                data.WorksheetName = sheetName;
                AssetDatabase.CreateAsset ((ScriptableObject)data, assetFilePath);
                //data.hideFlags = HideFlags.NotEditable;
            }
            
            //data.dataArray = new ExcelQuery(filePath, sheetName).Deserialize<ItemExcelData>().ToArray();		

            //ScriptableObject obj = AssetDatabase.LoadAssetAtPath (assetFilePath, typeof(ScriptableObject)) as ScriptableObject;
            //EditorUtility.SetDirty (obj);

            ExcelQuery query = new ExcelQuery(filePath, 0);
            if (query != null && query.IsValid())
            {
                data.dataArray = query.Deserialize<ItemExcelData>(3).ToArray();
                ScriptableObject obj = AssetDatabase.LoadAssetAtPath (assetFilePath, typeof(ScriptableObject)) as ScriptableObject;
                EditorUtility.SetDirty (obj);
            }
        }
    }
}
