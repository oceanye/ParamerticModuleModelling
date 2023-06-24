using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows.Forms;
using System.Numerics;
using System.IO;
using Autodesk.Revit;
using Autodesk.Revit.DB.Analysis;

namespace ParamerticModuleModelling
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.UsingCommandData)]
    public class CreateModulebySQL : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            string f_path = @" Y:\数字化课题";


            if (Directory.Exists(f_path) == false)
            {
                f_path = @"C:\ProgramData\Autodesk\Revit\Addins\2018";
            }


            using (SQLiteConnection conn = new SQLiteConnection("Data Source = " + f_path + @"\数据库\RevitData.db"))
            {
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand();

                cmd.Connection = conn;


                #region 板




                string sqlPlate = "SELECT*FROM ModuleInfo";
                DataTable Platetable = getData(sqlPlate, cmd);


                for (int x = 0; x < Platetable.Rows.Count; x++)
                {
                    string point_info = Platetable.Rows[x].ItemArray[2].ToString().Replace('(', ' ').Replace(')', ' ');
                    string plate_thick = Platetable.Rows[x].ItemArray[3].ToString();
                    string plate_name = Platetable.Rows[x].ItemArray[6].ToString();

                    //int np = Convert.ToInt32(point_info.Split('*')[0]);



                    // 获取当前文档对象
                    UIDocument Uidoc = commandData.Application.ActiveUIDocument;
                    Document doc = Uidoc.Document;

                    if (plate_name == "结构楼板")
                    {
                        CreateSlab(point_info, plate_thick, plate_name,doc);
                    }
                    else if (plate_name == "基墙")
                    {
                        CreateWall(point_info, plate_thick, plate_name, doc);
                    }
                    else if (plate_name == "门窗")
                    {
                        CreateDoor(point_info, plate_thick, plate_name, doc);
                    }
                    else if (plate_name == "楼地面构造层")
                    {
                        CreateSlab(point_info, plate_thick, plate_name, doc);
                    }
                    else if (plate_name == "翻边")
                    {
                        CreateWall(point_info, plate_thick, plate_name, doc);
                    }
                    else if (plate_name == "墙体面层")
                    {
                        CreateWall(point_info, plate_thick, plate_name, doc);
                    }
                }








                #endregion




                conn.Close();




                return Result.Succeeded;




            }



        }

        public DataTable getData(string sql, SQLiteCommand cmd)
        {

            cmd.CommandText = sql;

            SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
            DataSet ds = new DataSet();
            adapter.Fill(ds);

            DataTable table = ds.Tables[0];

            return table;
        }


        public void CreateSlab(string point_info,string thickness ,string platename, Document doc)
        {
            
            // 将字符串格式的点信息解析为 XYZ 对象列表
            List<string> pointStrings = new List<string>();

            pointStrings = point_info.Split('*').ToList();
            pointStrings.RemoveAt(0);

            List<XYZ> points = new List<XYZ>();

            double plateThickness = Convert.ToDouble(thickness);

            foreach (string pointString in pointStrings)
            {
                string[] coordinates = pointString.Split(',');
                double x = UnitUtils.Convert(Convert.ToDouble(coordinates[0]), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_DECIMAL_FEET);
                double y = UnitUtils.Convert(Convert.ToDouble(coordinates[1]), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_DECIMAL_FEET);
                double z = UnitUtils.Convert(Convert.ToDouble(coordinates[2]), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_DECIMAL_FEET);
                XYZ point = new XYZ(x, y, z);
                points.Add(point);
            }


            // 开始事务
            Transaction trans = new Transaction(doc, "Create Slab");
            trans.Start();

            try
            {
                // 创建多边形的边界
                CurveArray boundary = new CurveArray();
                for (int i = 0; i < points.Count - 1; i++)
                {
                    Line line = Line.CreateBound(points[i], points[i + 1]);
                    boundary.Append(line);
                }
                // 添加多边形的最后一条边
                Line lastLine = Line.CreateBound(points[points.Count - 1], points[0]);
                boundary.Append(lastLine);

                // 创建结构楼板

                FloorType floorType = null;
                IList<Element> floorTypeList = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).ToElements();

                foreach (Element ele in floorTypeList)
                {
                    FloorType Ft = ele as FloorType;
                    if (ele.Name.Contains("常规"))
                    {
                        floorType = Ft;

                    }
                }


                Floor floor = doc.Create.NewFloor(boundary, floorType, null, true);

                // 提交事务
                trans.Commit();

                // 如果需要，可以返回新创建的楼板对象
                // return floor;
            }
            catch (Exception ex)
            {
                // 发生错误时回滚事务
                trans.RollBack();
                throw new Exception("创建楼板时发生错误：" + ex.Message);
            }
        }
        public void CreateWall(string point_info, string thickness, string platename, Document doc)
        {

            // 将字符串格式的点信息解析为 XYZ 对象列表
            List<string> pointStrings = new List<string>();

            pointStrings = point_info.Split('*').ToList();
            pointStrings.RemoveAt(0);

            List<XYZ> points = new List<XYZ>();

            double plateThickness = Convert.ToDouble(thickness);

            foreach (string pointString in pointStrings)
            {
                string[] coordinates = pointString.Split(',');
                double x = UnitUtils.Convert(Convert.ToDouble(coordinates[0]), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_DECIMAL_FEET);
                double y = UnitUtils.Convert(Convert.ToDouble(coordinates[1]), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_DECIMAL_FEET);
                double z = UnitUtils.Convert(Convert.ToDouble(coordinates[2]), DisplayUnitType.DUT_MILLIMETERS, DisplayUnitType.DUT_DECIMAL_FEET);
                XYZ point = new XYZ(x, y, z);
                points.Add(point);
            }


            double minZ = double.MaxValue;
            double maxZ = double.MinValue;

            foreach (XYZ point in points)
            {
                if (point.Z < minZ)
                {
                    minZ = point.Z;
                }

                if (point.Z > maxZ)
                {
                    maxZ = point.Z;
                }
            }

            XYZ start_p = new XYZ();
            XYZ end_p = new XYZ();


            for (int i = 0; i < points.Count - 1; i++)
            {
                start_p = new XYZ(points[0].X, points[0].Y, minZ);
                end_p = new XYZ(points[1].X, points[1].Y, minZ);
                if (start_p.DistanceTo(end_p) < 1e-3)
                {
                    end_p = new XYZ(points[2].X, points[2].Y, minZ);

                }


            }


            // 指定的标高高度
            //double specifiedElevation_bottom = minZ;
            

            // 查找最接近的标高
            Level level_bottom = GetClosestLevel(doc, minZ);
            Level level_top = GetClosestLevel(doc,  maxZ );

            Transaction trans = new Transaction(doc, "Create Wall");
            trans.Start();

            try
            {
                // 开始事务



            if (level_bottom == null || level_top ==null)
            {
                return;
            }
                // 获取墙类型
                WallType wallType = null;
                IList<Element> wallTypeList = new FilteredElementCollector(doc).OfClass(typeof(WallType)).ToElements();

                foreach (Element ele in wallTypeList)
                {
                    WallType Wt = ele as WallType;
                    if (ele.Name.Contains("常规"))
                    {
                        wallType = Wt;

                    }
                }


                wallType = GetWallType(doc, "基本墙", plateThickness);

                // 创建墙的线段
                Line line = Line.CreateBound(start_p, end_p);

                // 创建墙
                Wall wall = Wall.Create(doc, line, level_bottom.Id, false);
                wall.WallType = wallType;

                // 获取墙的底部偏移参数
                Parameter baseOffsetParam = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);

                if (baseOffsetParam != null && baseOffsetParam.HasValue)
                {
                    double baseOffset = baseOffsetParam.AsDouble();
                    double newBaseOffset = minZ - level_bottom.Elevation;
                    double offsetDifference = newBaseOffset - baseOffset;

                    // 将墙的底部偏移参数设置为新的偏移值
                    baseOffsetParam.Set(newBaseOffset);

                    // 平移墙以保持其顶部位置不变
                    //ElementTransformUtils.MoveElement(doc, wall.Id, new XYZ(0, 0, offsetDifference));
                }

                // 设置墙的高度
                Parameter heightParameter = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                heightParameter.Set((maxZ-minZ) ); // 将毫米转换为英尺


                //Parameter topOffsetParam = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);

                //if (topOffsetParam.AsString()!="未连接")


                //    if (topOffsetParam != null && topOffsetParam.HasValue)
                //    {
                //        double topOffset = topOffsetParam.AsDouble();
                //        double newTopOffset = maxZ * 304 - level_top.Elevation;
                //        double offsetDifference = newTopOffset - topOffset;

                //        // 将墙的底部偏移参数设置为新的偏移值
                //        topOffsetParam.Set(newTopOffset);

                //        // 平移墙以保持其顶部位置不变
                //        //ElementTransformUtils.MoveElement(doc, wall.Id, new XYZ(0, 0, offsetDifference));
                //    }


                // 完成事务
                doc.Regenerate(); // 重新生成文档以更新墙的几何信息


            // 提交事务
            trans.Commit();



            }
            catch (Exception ex)
            {
                // 发生错误时回滚事务
                trans.RollBack();
                throw new Exception("创建墙单元时发生错误：" + ex.Message);
            }
        }

        public void CreateDoor(string point_info, string thickness, string platename, Document doc)
        {
            string point_coord = point_info.Split('*')[1];
            string door_size = point_info.Split('*')[0];
            List<string> point_str = point_coord.Substring(1, point_coord.Length - 2).Split(',').ToList();
            // 选择一个点作为门的布置点
            XYZ point = new XYZ(Convert.ToDouble(point_str[0])/304.8, Convert.ToDouble(point_str[1])/304.8, Convert.ToDouble(point_str[2])/304.8);



            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            IList<Element> walls = collector.ToElements();

            // 初始化最小距离和最近墙的变量
            double minDistance = double.MaxValue;
            Wall closestWall = null;

            // 遍历所有墙，找到距离最近的墙
            foreach (Wall wall in walls)
            {
                // 获取墙中心线
                LocationCurve location = wall.Location as LocationCurve;
                if (location != null)
                {
                    Curve curve = location.Curve;

                    // 计算点到墙中心线的投影距离
                    //double distance = curve.Distance(point);
                    XYZ project_p = curve.Project(point).XYZPoint;
                    double distance = project_p.DistanceTo(point);

                    // 更新最小距离和最近墙
                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        closestWall = wall;
                    }
                }
            }




            try
            {
                // 选择一个点作为门的布置点

                if (point == null)
                    return;

                // 在门的位置创建一个门实例
                using (Transaction tx = new Transaction(doc, "Place Door"))
                {
                    tx.Start();

                    // 在指定的 XYZ 坐标点处创建门的实例
                    FamilySymbol doorSymbol = GetDoorSymbol(doc, "single_door", Convert.ToDouble(door_size.Split('-')[0]) / 304.8, Convert.ToDouble(door_size.Split('-')[1]) / 304.8);
                    if (doorSymbol != null)
                    {
                        Level level = GetClosestLevel(doc, point.Z/304.8);
                        Wall ref_wall = closestWall;
                        ;
                        FamilyInstance door = doc.Create.NewFamilyInstance(point, doorSymbol,ref_wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                        // 设置门的其他属性（尺寸、类型等）
                        // ...
                        // 假设 door 是要布置的门族的实例对象

                        // 获取高度参数
                        Parameter heightParameter = door.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);

                        // 将底部高度设置为0
                        double bottomHeight = 0.0;
                        heightParameter.Set(bottomHeight);


                        tx.Commit();
                        
                        return;
                    }
                }

                TaskDialog.Show("Error", "Failed to place door at XYZ coordinate.");
                return;
            }
            catch (Exception ex)
            {
                ;
                return;
            }

        }




        private FamilySymbol GetDoorSymbol(Document doc, string familyName, double width, double height)
        {

            string new_name = Math.Round(width * 304.8, 0) + "x" + Math.Round(height * 304.8, 0);

            // 在指定的门族中查找名称匹配的族类型
            Family family = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => f.Name == familyName);

            if (family != null)
            {
                // 获取族中的所有类型的 ElementId
                ICollection<ElementId> symbolIds = family.GetFamilySymbolIds();


                foreach (ElementId elem in symbolIds)
                {
                    FamilySymbol iSymbol = doc.GetElement(elem) as FamilySymbol;
                    if (iSymbol.Name == new_name)
                        return iSymbol;
                }



                if (symbolIds.Count > 0)
                {
                    // 获取第一个类型的 ElementId
                    ElementId firstSymbolId = symbolIds.First();

                    
                    

                    FamilySymbol firstSymbol = doc.GetElement(firstSymbolId) as FamilySymbol;
                    
                    FamilySymbol newDoorSymbol = firstSymbol.Duplicate(new_name) as FamilySymbol;

                    if (newDoorSymbol != null)
                    {
                        // 设置新门类型的高度和宽度参数
                        Parameter heightParameter = newDoorSymbol.LookupParameter("高度");
                        Parameter widthParameter = newDoorSymbol.LookupParameter("宽度");

                        if (heightParameter != null && widthParameter != null)
                        {
                            // 将高度和宽度参数设置为指定的值
                            heightParameter.Set(height);
                            widthParameter.Set(width);

  

                            return newDoorSymbol;
                        }
                    }


                }
            }

            return null; // 未找到符合条件的门族符号或类型
        }


        private FamilySymbol GetWallSymbol(Document doc, string familyName, double thick)
        {

            string new_name = Math.Round(thick * 304.8, 0).ToString();

            // 在指定的门族中查找名称匹配的族类型
            Family family = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => f.Name == familyName);

            if (family != null)
            {
                // 获取族中的所有类型的 ElementId
                ICollection<ElementId> symbolIds = family.GetFamilySymbolIds();


                foreach (ElementId elem in symbolIds)
                {
                    FamilySymbol iSymbol = doc.GetElement(elem) as FamilySymbol;
                    if (iSymbol.Name == new_name)
                        return iSymbol;
                }



                if (symbolIds.Count > 0)
                {
                    // 获取第一个类型的 ElementId
                    ElementId firstSymbolId = symbolIds.First();




                    FamilySymbol firstSymbol = doc.GetElement(firstSymbolId) as FamilySymbol;

                    FamilySymbol newWallSymbol = firstSymbol.Duplicate(new_name) as FamilySymbol;

                    if (newWallSymbol != null)
                    {
                        // 设置新门类型的高度和宽度参数
                        Parameter thickParameter = newWallSymbol.LookupParameter("厚度");


                        if (thickParameter != null)
                        {
                            // 将高度和宽度参数设置为指定的值
                            thickParameter.Set(thick);




                            return newWallSymbol;
                        }
                    }


                }
            }

            return null; // 未找到符合条件的族符号或类型
        }


        public Level GetClosestLevel(Document doc, double elevation)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            Level closestLevel = null;
            double closestElevationDifference = double.MaxValue;

            foreach (Element element in collector)
            {
                Level level = element as Level;
                double elevationDifference = Math.Abs(level.Elevation - elevation);

                if (elevationDifference < closestElevationDifference)
                {
                    closestLevel = level;
                    closestElevationDifference = elevationDifference;
                }
            }

            return closestLevel;
        }

        public WallType GetWallType(Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (Element element in collector)
            {
                WallType wallType = element as WallType;
                if (wallType.Name.Equals(wallTypeName, StringComparison.OrdinalIgnoreCase))
                {
                    return wallType;
                }
            }

            return null;
        }

        public WallType GetWallType(Document doc,string wall_name, double thickness)
        {
            string new_wallname = "W" + thickness;


            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            IList<Element> wallTypeList2 = new FilteredElementCollector(doc).OfClass(typeof(WallType)).ToElements();

            foreach (Element elem in collector)
            {
                WallType wallType = elem as WallType;
                if (wallType.Name == new_wallname)
                {
                    return wallType;
                }
            }



            // 获取族类型 "基本墙" 的 FamilySymbol
            //FilteredElementCollector symbolCollector = new FilteredElementCollector(doc);
            //symbolCollector.OfClass(typeof(FamilySymbol));
            IList<Element> wallTypeList = new FilteredElementCollector(doc).OfClass(typeof(WallType)).ToElements();

            WallType basicWallType = wallTypeList.Cast<WallType>().FirstOrDefault(x => x.FamilyName == wall_name);

            // 创建一个新的墙类型
            WallType newWallType = basicWallType.Duplicate(new_wallname) as WallType;
            Parameter newThicknessParam = newWallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
            if (newThicknessParam != null && newThicknessParam.StorageType == StorageType.Double)
            {
                //newThicknessParam.Set(thickness / 304.8); // 将厚度从毫米转换为英尺
            }



            return newWallType;
        }


    }
}
