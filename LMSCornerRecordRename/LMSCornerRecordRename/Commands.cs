using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using CivilDB = Autodesk.Civil.DatabaseServices;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(CornerRecordUpdate.Commands))]
[assembly: ExtensionApplication(null)]
namespace CornerRecordUpdate
{
    #region Commands
    public class Commands
    {
        [CommandMethod("OCPWRN")]
        public void ListAttributes()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            var doc = Application.DocumentManager.MdiActiveDocument;

            // read input parameters from JSON file
            //InputParams inputParams = JsonConvert.DeserializeObject<InputParams>(File.ReadAllText("params.json"));
            dynamic dynamicResultObject = JsonConvert.DeserializeObject(File.ReadAllText("params.json"));

            try
            {
                var acDB = doc.Database;

                using (var trans = acDB.TransactionManager.StartTransaction())
                {
                    // Capture Layouts to be Checked
                    DBDictionary layoutPages = (DBDictionary)trans.GetObject(acDB.LayoutDictionaryId, OpenMode.ForRead);

                    // List of Layout names 
                    List<string> layoutNamesChecked = new List<string>();

                    // Test Json Dictionary
                    //var Doc_Num_Test = new Dictionary<string, string>()
                    //{
                    //    { "cr1", "CR 2020-9888" },
                    //    { "cr2", "CR 2020-9889" },
                    //    { "cr3", "CR 2020-9890" },
                    //    { "cr4", "CR 2020-9891" }
                    //};

                    // Cast dynamicResultObject to dictionary
                    Dictionary<string, string> docNumber = dynamicResultObject as Dictionary<string, string>;

                    foreach (DBDictionaryEntry layoutPage in layoutPages)
                    {
                        var layoutUnchecked = layoutPage.Value.GetObject(OpenMode.ForWrite) as Layout;
                        var isModelSpace = layoutUnchecked.ModelType;

                        ObjectIdCollection textObjCollection = new ObjectIdCollection();

                        if (isModelSpace != true)
                        {
                            Match layoutNameMatch = Regex.Match(layoutUnchecked.LayoutName, "^(\\s*cr\\s*\\d\\d*)$",
                               RegexOptions.IgnoreCase);

                            if (layoutNameMatch.Success)
                            {
                                string layoutChecked = layoutUnchecked.LayoutName.Trim().ToString().ToLower().Replace(" ", "");

                                // for each Form_Result.Value in Doc_Num_Test.json file, 
                                // if the file, if the SF_Json_Key matches layoutPage,
                                // rename the layoutPage.LayoutName with the SF_Json_Key.Key["Doc_Num"]
                                if (dynamicResultObject.ContainsKey(layoutChecked.ToString()))
                                {
                                    //ed.WriteMessage("\n" + dynamicResultObject[layoutChecked.ToString()]);
                                    layoutUnchecked.LayoutName = dynamicResultObject[layoutChecked.ToString()];
                                }
                            }
                        }
                    }
                    trans.Commit();
                }
                acDB.SaveAs("outputFile.dwg", DwgVersion.Current);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("Exception: " + ex.Message);
            }
        }
    }
    #endregion
}
