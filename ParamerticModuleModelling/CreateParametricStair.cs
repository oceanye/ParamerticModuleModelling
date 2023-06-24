using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using System.IO;
using Autodesk.Revit.DB.Structure;
using System.Data.SQLite;
using System.Data;
using Autodesk.Revit.DB.Analysis;
using System.Windows.Forms;


namespace ParamerticModuleModelling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.UsingCommandData)]
    class CreateParametricStair : IExternalCommand
    {
        //public string PlanName;
        //public string Flag;







        // public string FamilyPath = @"Y:\数字化课题\族库\载入族\";

        public List<string> FList = new List<string>();
        public string DataFlag = null;

        Dictionary<string, int> BeamDic = new Dictionary<string, int>();
        Dictionary<string, int> ColumnDic = new Dictionary<string, int>();

        List<string> PlanList = new List<string>();

        List<string> LayList = new List<string>();

        // public string DbPath = @"Y:\数字化课题\数据库\CenterData.db";
        public static string f_path;
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {


            string f_path = @" Y:\数字化课题";


            if (Directory.Exists(f_path) == false)
            {
                f_path = @"C:\ProgramData\Autodesk\Revit\Addins\2018";
            }



            var from = new InputParameter_stair();
            from.ShowDialog();

            UIDocument Uidoc = commandData.Application.ActiveUIDocument;
            Document Doc = Uidoc.Document;
            Document docNew = null;
            ViewFamilyType viewPlanType = null;

            string f_name = InputParameter_stair.Modelselect;
            LoadFam(Doc, f_path,f_name);

            var dup_symbol = new Transaction(Doc, "dup_symbol");

            FamilySymbol familySymbol = null;
            IList<ElementId> TypeL = new FilteredElementCollector(Doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel).ToElementIds().ToList();


            IList<ElementId> FamilyL = new FilteredElementCollector(Doc).OfClass(typeof(Family)).ToElementIds().ToList();

            Family fa = null;
            //string fam = fa.Name;

            FamilySymbol fs = null;
            FamilyInstance familyins=null;

            FamilySymbol firstSymbol = null;



            dup_symbol.Start();

            foreach(var item in FamilyL)
            {
                fa = Doc.GetElement(item) as Family;
                if (fa != null)
                {
                    if(fa.Name.Contains(InputParameter_stair.Modelselect))
                    {
                        try
                        {
                            firstSymbol = fa.GetFamilySymbolIds().Select(id => Doc.GetElement(id)).OfType<FamilySymbol>().FirstOrDefault();

                            //IList<string> FamilySymbolL = fa.GetFamilySymbolIds().Select(id => Doc.GetElement(id)).OfType<FamilySymbol>().OfType<Name>;

                    

                            string name1 = InputParameter_stair.ST_name;
                            
                            familySymbol = firstSymbol.Duplicate(name1) as FamilySymbol;
                            XYZ insertXYZ = new XYZ(0, 0, 0);
                            //familyins = Doc.Create.NewFamilyInstance(insertXYZ, familySymbol, StructuralType.NonStructural);

                            familyins = Doc.Create.NewFamilyInstance(insertXYZ, familySymbol, StructuralType.NonStructural);


                            break;



                        }
                        catch
                        {

                        }
                    }
                }
            }


            dup_symbol.Commit();


            //foreach (var item in TypeL)
            //{
            //    fs = Doc.GetElement(item) as FamilySymbol;

            //    if (fs != null)
            //        if (fs.FamilyName == "槽钢型钢楼梯（双跑）" && fs.Name == "默认") 
            //        {
            //            try
            //            {



            //                string name1 = InputParameter_stair.ST_name;
            //                familySymbol = fs.Duplicate(name1) as FamilySymbol;
            //                XYZ insertXYZ = new XYZ(0, 0, 0);
            //                familyins = Doc.Create.NewFamilyInstance(insertXYZ, familySymbol, StructuralType.NonStructural);
            //                break;
            //            }
            //            catch
            //            { }
            //        }
            //}

            //dup_symbol.Commit();




            try
            {
                // 更新参数的值
                using (Transaction transaction = new Transaction(Doc, "更新参数值"))
                {
                    transaction.Start();

                    // 设置参数的值
                    familyins.LookupParameter("AA").SetValueString(InputParameter_stair.AA);
                    familyins.LookupParameter("AB").SetValueString(InputParameter_stair.AB);
                    //familyins.LookupParameter("AC").SetValueString(InputParameter_stair.AC);
                    familyins.LookupParameter("AD").SetValueString(InputParameter_stair.AD);
                    familySymbol.LookupParameter("AE").SetValueString(InputParameter_stair.AE);
                    familySymbol.LookupParameter("AF").SetValueString(InputParameter_stair.AF);

                    familyins.LookupParameter("BA").SetValueString(InputParameter_stair.BA);
                    familyins.LookupParameter("BB").SetValueString(InputParameter_stair.BB);
                    familyins.LookupParameter("CA").SetValueString(InputParameter_stair.CA);
                    familyins.LookupParameter("CB").SetValueString(InputParameter_stair.CB);
                    //parameter.SetValueString("1501");


                    familySymbol.LookupParameter("DA").SetValueString(InputParameter_stair.DA);
                    familySymbol.LookupParameter("DB").SetValueString(InputParameter_stair.DB);
                    familySymbol.LookupParameter("DC").SetValueString(InputParameter_stair.DC);
                    familySymbol.LookupParameter("DD").SetValueString(InputParameter_stair.DD);
                    familySymbol.LookupParameter("DE").SetValueString(InputParameter_stair.DE);
                    familySymbol.LookupParameter("DF").SetValueString(InputParameter_stair.DF);
                    familySymbol.LookupParameter("DG").SetValueString(InputParameter_stair.DG);



                    transaction.Commit();
                }
            }
            catch
            { }

            return Result.Succeeded;

        }


        public void LoadFam(Document Doc,string filePath,string f_name)
        {
            //filePath = file_direction_exist(filePath);


            string f_path = filePath;
            if (Directory.Exists(f_path) == false)
            {
                f_path = @"C:\ProgramData\Autodesk\Revit\Addins\2018";
            }

            var dirsFirst = new DirectoryInfo(f_path + @"\族库\楼梯系统模块");

            var fileFirst = dirsFirst.GetFiles();
            var loadTr = new Transaction(Doc, "loadTr");

            loadTr.Start();
            for (var i = 0; i < fileFirst.Length; i++)
            {
                if (fileFirst[i].FullName.Contains(f_name))
                {
                                    var op = new MyFamilyLoadOptions();
                Family family = null;
                Doc.LoadFamily(fileFirst[i].FullName, op, out family);
                }

            }

            loadTr.Commit();
        }



    }
}
