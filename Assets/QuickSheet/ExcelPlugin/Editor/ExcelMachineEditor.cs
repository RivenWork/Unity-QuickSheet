///////////////////////////////////////////////////////////////////////////////
///
/// ExcelMachineEditor.cs
///
/// (c)2014 Kim, Hyoun Woo
///
///////////////////////////////////////////////////////////////////////////////
using UnityEngine;
using UnityEditor;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;

namespace UnityQuickSheet
{
    /// <summary>
    /// Custom editor script class for excel file setting.
    /// </summary>
    [CustomEditor(typeof(ExcelMachine))]
    public class ExcelMachineEditor : BaseMachineEditor
    {
        protected override void OnEnable()
        {
            base.OnEnable();

            machine = target as ExcelMachine;
            if (machine != null && ExcelSettings.Instance != null)
            {
                if (string.IsNullOrEmpty(ExcelSettings.Instance.RuntimePath) == false)
                    machine.RuntimeClassPath = ExcelSettings.Instance.RuntimePath;
                if (string.IsNullOrEmpty(ExcelSettings.Instance.EditorPath) == false)
                    machine.EditorClassPath = ExcelSettings.Instance.EditorPath;
            }
        }

        public override void OnInspectorGUI()
        {
            base.OnInspectorGUI();

            ExcelMachine machine = target as ExcelMachine;

            GUILayout.Label("Excel Settings:", headerStyle);

            GUILayout.BeginHorizontal();
            GUILayout.Label("Folder:", GUILayout.Width(50));

            string path = string.Empty;
            if (string.IsNullOrEmpty(machine.fileFolder))
                path = Application.dataPath;
            else
                path = machine.fileFolder;

            machine.fileFolder = GUILayout.TextField(path, GUILayout.Width(250));
            if (GUILayout.Button("...", GUILayout.Width(20)))
            {
                string folder = Path.GetDirectoryName(path);
#if UNITY_EDITOR_WIN
                path = EditorUtility.OpenFolderPanel("Open Excel file folder", folder, string.Empty);
#else // for UNITY_EDITOR_OSX
                path = EditorUtility.OpenFolderPanel("Open Excel file folder", folder, string.Empty);
#endif
                if (path.Length != 0)
                {
                    var dirInfo = new DirectoryInfo(path);
                    if (dirInfo.Exists)
                    {
                        int index = path.IndexOf("Assets");
                        if (index >= 0)
                        {
                            // set relative path
                            machine.fileFolder = path.Substring(index);
                        }
                        else
                        {
                            EditorUtility.DisplayDialog("Error",
                                @"Wrong folder is selected.
                        Set a folder under the 'Assets' folder! \n
                        The excel file should be anywhere under  the 'Assets' folder", "OK");
                            return;
                        }
                    }
                }
            }
            GUILayout.EndHorizontal();

            EditorGUILayout.Separator();
            
            GUILayout.Label("Path Settings:", headerStyle);

            machine.TemplatePath = EditorGUILayout.TextField("Template: ", machine.TemplatePath);
            machine.RuntimeClassPath = EditorGUILayout.TextField("Runtime: ", machine.RuntimeClassPath);
            machine.EditorClassPath = EditorGUILayout.TextField("Editor:", machine.EditorClassPath);
            machine.DataFilePath = EditorGUILayout.TextField("Data:", machine.DataFilePath);

            EditorGUILayout.Separator();

            GUILayout.BeginHorizontal();

            if (GUILayout.Button("Import"))
            {
                var targetPath = Path.GetDirectoryName(AssetDatabase.GetAssetPath(target));
                if (!AssetDatabase.IsValidFolder(targetPath + "/Runtime"))
                    AssetDatabase.CreateFolder(targetPath, "Runtime");
                if (!AssetDatabase.IsValidFolder(targetPath + "/Editor"))
                    AssetDatabase.CreateFolder(targetPath, "Editor");
                if (!AssetDatabase.IsValidFolder(targetPath + "/Data"))
                    AssetDatabase.CreateFolder(targetPath, "Data");

                CreateScriptableObjectClassScript();
                CreateScriptableObjectEditorClassScript();
                //CreateDataClassScript();
                //CreateAssetCreationScript();

                Debug.Log("Excel Done!");
            }

            if (GUILayout.Button("Generate"))
            {

            }

            GUILayout.EndHorizontal();

            if (GUI.changed)
            {
                EditorUtility.SetDirty(machine);
            }
        }

        void CreateScriptableObjectClassScript()
        {
            ExcelMachine machine = target as ExcelMachine;

#if UNITY_EDITOR_WIN
            var files = new DirectoryInfo(machine.fileFolder).GetFiles().Where(x => x.Extension == ".xlsx");
#else // for UNITY_EDITOR_OSX
            var files = new DirectoryInfo(machine.fileFolder).GetFiles().Where(x => x.Extension == ".xls");
#endif
            foreach (var fileInfo in files)
            {
                var className = Path.GetFileNameWithoutExtension(fileInfo.Name);
                // check the directory path exists
                string fullPath = TargetPathForClassScript(className);
                string folderPath = Path.GetDirectoryName(fullPath);
                if (!Directory.Exists(folderPath))
                {
                    EditorUtility.DisplayDialog(
                        "Warning",
                        "The folder for runtime script files does not exist. Check the path " + folderPath + " exists.",
                        "OK"
                    );
                    return;
                }

                var sp = new ScriptPrescription()
                {
                    className = className + "Editor",
                    worksheetClassName = className,
                    dataClassName = className + "ExcelData",
                    template = GetTemplate("ScriptableObjectClass"),
                };

                StreamWriter writer = null;
                try
                {
                    // write a script to the given folder.		
                    writer = new StreamWriter(fullPath);
                    writer.Write(new ScriptGenerator(sp).ToString());
                }
                catch (System.Exception e)
                {
                    Debug.LogError(e.Message);
                }
                finally
                {
                    if (writer != null)
                    {
                        writer.Close();
                        writer.Dispose();
                    }
                }
            }
        }
        void CreateScriptableObjectEditorClassScript()
        {
            ExcelMachine machine = target as ExcelMachine;

#if UNITY_EDITOR_WIN
            var files = new DirectoryInfo(machine.fileFolder).GetFiles().Where(x => x.Extension == ".xlsx");
#else // for UNITY_EDITOR_OSX
            var files = new DirectoryInfo(machine.fileFolder).GetFiles().Where(x => x.Extension == ".xls");
#endif
            foreach (var fileInfo in files)
            {
                var className = Path.GetFileNameWithoutExtension(fileInfo.Name);
                // check the directory path exists
                string fullPath = TargetPathForClassScript(className);
                string folderPath = Path.GetDirectoryName(fullPath);
                if (!Directory.Exists(folderPath))
                {
                    EditorUtility.DisplayDialog(
                        "Warning",
                        "The folder for runtime script files does not exist. Check the path " + folderPath + " exists.",
                        "OK"
                    );
                    return;
                }

                var sp = new ScriptPrescription()
                {
                    className = className + "Editor",
                    worksheetClassName = className,
                    dataClassName = className + "ExcelData",
                    template = GetTemplate("ScriptableObjectEditorClass"),
                };

                StreamWriter writer = null;
                try
                {
                    // write a script to the given folder.		
                    writer = new StreamWriter(fullPath);
                    writer.Write(new ScriptGenerator(sp).ToString());
                }
                catch (System.Exception e)
                {
                    Debug.LogError(e);
                }
                finally
                {
                    if (writer != null)
                    {
                        writer.Close();
                        writer.Dispose();
                    }
                }
            }
        }
        void CreateDataClassScript()
        {

        }
        void CreateAssetCreationScript()
        {

        }

        /// <summary>
        /// Import the specified excel file and prepare to set type of each cell.
        /// </summary>
        protected override void Import(bool reimport = false)
        {
            ExcelMachine machine = target as ExcelMachine;

            string path = machine.excelFilePath;
            string sheet = machine.WorkSheetName;

            if (string.IsNullOrEmpty(path))
            {
                string msg = "You should specify spreadsheet file first!";
                EditorUtility.DisplayDialog("Error", msg, "OK");
                return;
            }

            if (!File.Exists(path))
            {
                string msg = string.Format("File at {0} does not exist.", path);
                EditorUtility.DisplayDialog("Error", msg, "OK");
                return;
            }

            int startRowIndex = 0;
            string error = string.Empty;
            var titles = new ExcelQuery(path, sheet).GetTitle(startRowIndex, ref error);
            if (titles == null || !string.IsNullOrEmpty(error))
            {
                EditorUtility.DisplayDialog("Error", error, "OK");
                return;
            }
            else
            {
                // check the column header is valid
                foreach (string column in titles)
                {
                    if (!IsValidHeader(column))
                    {
                        error = string.Format(@"Invalid column header name {0}. Any c# keyword should not be used for column header. Note it is not case sensitive.", column);
                        EditorUtility.DisplayDialog("Error", error, "OK");
                        return;
                    }
                }
            }

            List<string> titleList = titles.ToList();

            if (machine.HasColumnHeader() && reimport == false)
            {
                var headerDic = machine.ColumnHeaderList.ToDictionary(header => header.name);

                // collect non-changed column headers
                var exist = titleList.Select(t => GetColumnHeaderString(t))
                    .Where(e => headerDic.ContainsKey(e) == true)
                    .Select(t => new ColumnHeader { name = t, type = headerDic[t].type, isArray = headerDic[t].isArray, OrderNO = headerDic[t].OrderNO });


                // collect newly added or changed column headers
                var changed = titleList.Select(t => GetColumnHeaderString(t))
                    .Where(e => headerDic.ContainsKey(e) == false)
                    .Select(t => ParseColumnHeader(t, titleList.IndexOf(t)));

                // merge two list via LINQ
                var merged = exist.Union(changed).OrderBy(x => x.OrderNO);

                machine.ColumnHeaderList.Clear();
                machine.ColumnHeaderList = merged.ToList();
            }
            else
            {
                machine.ColumnHeaderList.Clear();
                if (titleList.Count > 0)
                {
                    int order = 0;
                    machine.ColumnHeaderList = titleList.Select(e => ParseColumnHeader(e, order++)).ToList();
                }
                else
                {
                    string msg = string.Format("An empty workhheet: [{0}] ", sheet);
                    Debug.LogWarning(msg);
                }
            }

            EditorUtility.SetDirty(machine);
            AssetDatabase.SaveAssets();
        }

        protected void ImportAll()
        {
            ExcelMachine machine = target as ExcelMachine;

            foreach (var fileInfo in new DirectoryInfo(machine.fileFolder).GetFiles())
            {
                var path = fileInfo.FullName;

                if (string.IsNullOrEmpty(path))
                {
                    string msg = "You should specify spreadsheet file first!";
                    EditorUtility.DisplayDialog("Error", msg, "OK");
                    return;
                }

                if (!File.Exists(path))
                {
                    string msg = string.Format("File at {0} does not exist.", path);
                    EditorUtility.DisplayDialog("Error", msg, "OK");
                    return;
                }

                string error = string.Empty;
                var titles = new ExcelQuery(path, 0).GetTitle(2, ref error);
                if (titles == null || !string.IsNullOrEmpty(error))
                {
                    EditorUtility.DisplayDialog("Error", error, "OK");
                    return;
                }
                else
                {
                    // check the column header is valid
                    foreach (string column in titles)
                    {
                        if (!IsValidHeader(column))
                        {
                            error = string.Format(@"Invalid column header name {0}. Any c# keyword should not be used for column header. Note it is not case sensitive.", column);
                            EditorUtility.DisplayDialog("Error", error, "OK");
                            return;
                        }
                    }
                }

                List<string> titleList = titles.ToList();

                if (titleList.Count == 0)
                {
                    string msg = string.Format("An empty workhheet: [{0}] ", 0);
                    Debug.LogWarning(msg);
                }
            }

            EditorUtility.SetDirty(machine);
            AssetDatabase.SaveAssets();
        }

        /// <summary>
        /// Generate AssetPostprocessor editor script file.
        /// </summary>
        protected override void CreateAssetCreationScript(BaseMachine m, ScriptPrescription sp)
        {
            ExcelMachine machine = target as ExcelMachine;

            sp.className = machine.WorkSheetName;
            sp.dataClassName = machine.WorkSheetName + "EData";
            sp.worksheetClassName = machine.WorkSheetName;

            // where the imported excel file is.
            sp.importedFilePath = machine.fileFolder;

            // path where the .asset file will be created.
            string path = Path.GetDirectoryName(machine.fileFolder).Replace("\\", "/");
            path += "/" + machine.WorkSheetName + ".asset";
            sp.assetFilepath = path;
            sp.assetPostprocessorClass = machine.WorkSheetName + "AssetPostprocessor";
            sp.template = GetTemplate("PostProcessor");

            // write a script to the given folder.
            using (var writer = new StreamWriter(TargetPathForAssetPostProcessorFile(machine.WorkSheetName)))
            {
                writer.Write(new ScriptGenerator(sp).ToString());
                writer.Close();
            }
        }
    }
}