using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using System.Reflection;
using System.Windows.Media.Imaging;


namespace ParamerticModuleModelling
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.UsingCommandData)]
    class ParametricModelUI : IExternalApplication
    {


        string dllPath = Assembly.GetExecutingAssembly().Location;
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string TitleName = "数字化集成模型";
            application.CreateRibbonTab(TitleName);


            string imagePath = @"C:\ProgramData\Autodesk\Revit\Addins\2018\ICON\";


    



            RibbonPanel newPanel1 = application.CreateRibbonPanel(TitleName, "楼梯系统");

            PushButton CreateParametricStair = newPanel1.AddItem(new PushButtonData("生成参数化楼梯", "生成楼梯", dllPath, "ParamerticModuleModelling.CreateParametricStair")) as PushButton;
            CreateParametricStair.ToolTip = "生成参数化楼梯";
            Uri link1 = new Uri(imagePath + "createstair1.png");
            CreateParametricStair.LargeImage = new BitmapImage(link1);



            PushButton DataToSqlite_Stair = newPanel1.AddItem(new PushButtonData("提取楼梯模型", "上传楼梯模型", dllPath, "ParamerticModuleModelling.DataToSqlite_Stair")) as PushButton;
            DataToSqlite_Stair.ToolTip = "楼梯模型上传";

            DataToSqlite_Stair.LargeImage = new BitmapImage(new Uri(imagePath + "SqliteStair1.png"));



            RibbonPanel newPanel2 = application.CreateRibbonPanel(TitleName, "内墙系统");

            PushButton UploadModuletoSQL = newPanel2.AddItem(new PushButtonData("提取模块数据", "上传模块模型", dllPath, "ParamerticModuleModelling.DataToSqlite_Module")) as PushButton;
            DataToSqlite_Stair.ToolTip = "模块模型上传";
            DataToSqlite_Stair.LargeImage = new BitmapImage(new Uri(imagePath + "uploadmodule1.png"));




            PushButton CreateModulebySQL = newPanel2.AddItem(new PushButtonData("生成标准模型", "根据模块信息，生成模型", dllPath, "ParamerticModuleModelling.CreateModulebySQL")) as PushButton;
            DataToSqlite_Stair.ToolTip = "模块模型生成";
            DataToSqlite_Stair.LargeImage = new BitmapImage(new Uri(imagePath + "createmodule1.png"));




            return Result.Succeeded;
        }
    }
}
