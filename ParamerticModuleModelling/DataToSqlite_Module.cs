using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows.Forms;
using System.Numerics;
using System.IO;

//using Excel = Microsoft.Office.Interop.Excel;
namespace ParamerticModuleModelling
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.UsingCommandData)]
    public class DataToSqlite_Module : IExternalCommand
    {
        public string DbPath = @"Y:\数字化课题";





        public void ClearTable(SQLiteCommand cmd)
        {
            cmd.CommandText = " delete from ModuleInfo";
            cmd.ExecuteNonQuery();
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            if (Directory.Exists(DbPath) == false)
            {
                DbPath = @"C:\ProgramData\Autodesk\Revit\Addins\2018";
            }

            UIDocument Uidoc = commandData.Application.ActiveUIDocument;
            Document Doc = Uidoc.Document;
            FamilySymbol fs = null;

            List<string> Obj_List = new List<string> { "基墙", "结构楼板", "门窗"  };//, "楼地面构造层"

            //FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            //folderBrowserDialog1.ShowDialog();
            //DbPath = folderBrowserDialog1.SelectedPath + @"\RevitData.db";







            FilteredElementCollector collector = new FilteredElementCollector(Doc);
            var columns = collector.OfClass(typeof(FamilyInstance));


            string info = "";


            foreach (FamilyInstance item in columns)
            {
                ExtractIDTree(item);
            }





            try
            {



                string f_path = DbPath;
                if (Directory.Exists(f_path) == false)
                {
                    f_path = @"C:\ProgramData\Autodesk\Revit\Addins\2018";

                }



                SQLiteConnection conn = new SQLiteConnection("Data Source=" + f_path + @"\数据库\RevitData.db");

                var cmd = Sql_Clean_Data(conn);


                #region 楼板

  

                #endregion

                #region 拉伸体（楼梯）
                foreach (FamilyInstance item in columns)
                {
                    string id = item.Id.ToString();
                    string name = item.Symbol.FamilyName;
                    string syname = item.Symbol.Name;

                    string xyzs = null;

                    string elem_hash = item.UniqueId;

                    info += id + "," + name + "," + syname + "\r\n";

                    Options options = new Options();

                    GeometryElement geometry = item.get_Geometry(options);

                    int i = 0;

                    foreach (GeometryObject obj in geometry)
                    {
                        GeometryInstance instance = obj as GeometryInstance;

                        GeometryElement geometryElement = instance.GetInstanceGeometry();

                        foreach (GeometryObject elem in geometryElement)
                        {



                            Solid solid = elem as Solid;

                            //string gstyleName = gStyle.GraphicsStyleCategory.Name;

                            string gStyleId = "";
                            string gstyleName = "";

                            GraphicsStyle gStyle = Doc.GetElement(elem.GraphicsStyleId) as GraphicsStyle;
                            gStyleId = elem.GraphicsStyleId.ToString(); // detailed id



                            if (gStyle != null)
                            {
                                gstyleName = gStyle.GraphicsStyleCategory.Name.ToString();// detailedName
                            }
                            else
                            {
                                continue;
                            }

                            if (Obj_List.Contains(gstyleName) == false)
                                continue;




                            if (gstyleName == "结构楼板" || gstyleName == "楼地面构造层")
                            {
                                Extract_Slab_Gemo(conn, solid, id, name, syname, gStyleId, gstyleName);
                                    continue;
                            }


                            if (gstyleName == "基墙" || gstyleName == "翻边" || gstyleName == "墙体面层")
                            {
                                Extract_Wall_Gemo(conn, solid, id, name, syname, gStyleId, gstyleName);
                                continue;
                            }

                            if (gstyleName == "门窗" )

                            {
                                
                                Extract_DoorWindow_Gemo(Doc,conn, solid, id, name, syname, gStyleId, gstyleName);

                                
                                continue;
                            }

                            // 摘录实体solid

                            if (solid != null)
                            {
                                Extract_Solid_Gemo(conn, solid, id, name, syname, gStyleId, gstyleName);
                            }


                        }
                    }





                }

                MessageBox.Show("完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            #endregion


            return Result.Succeeded;
        }

        public void ExtractFamilyParameters(Element nestedElement)
        {
            FamilyInstance nestedInstance = nestedElement as FamilyInstance;
            if (nestedInstance != null && nestedInstance.Symbol != null)
            {
                FamilySymbol symbol = nestedInstance.Symbol;

                foreach (Parameter parameter in symbol.Parameters)
                {
                    if (parameter != null && parameter.HasValue)
                    {
                        string paramName = parameter.Definition.Name;
                        string paramValue = parameter.AsValueString();

                        // 使用 paramName 和 paramValue 进行后续操作
                    }
                }
            }
        }


        private static void ExtractIDTree(FamilyInstance Fam)
        {


            FilteredElementCollector collector = new FilteredElementCollector(Fam.Document);
            var Fam1 = collector.OfClass(typeof(FamilyInstance));
            var Sweep1 = collector.OfClass(typeof(Sweep));
            var Extrusion1 = collector.OfClass(typeof(Extrusion));

            string id = Fam.Id.ToString();
            int a = Fam1.ToList().Count();

            foreach (FamilyInstance Fam2 in Fam1)
            {
                Fam2.GetHashCode();

                ExtractIDTree(Fam2);
            }

            foreach (Sweep Sweep2 in Sweep1)
            {
                Sweep2.GetHashCode();
            }

            foreach (Sweep Extrusion2 in Extrusion1)
            {
                Extrusion2.GetHashCode();
            }

        }


        private void Extract_Wall_Gemo(SQLiteConnection conn, Solid solid, string id, string name, string syname, string gStyleId, string gstyleName)
        {
            FaceArray fArray = solid.Faces;
            List<string> point_list = new List<string>();
            double thk = 0;
            string point_str = "";


            // Create a list to store face center points
            List<Tuple<XYZ, Face>> faceCenters = new List<Tuple<XYZ, Face>>();
            List<Face> fArray_vertical = new List<Face>();
            // Iterate over each face in the array and calculate the center point
            foreach (Face face in fArray)
            {
                XYZ faceNormal = face.ComputeNormal(new UV());
                if (Math.Abs(faceNormal.Z) < 0.0001)
                {
                    //XYZ center = GetFaceCenter(face);
                    //faceCenters.Add(new Tuple<XYZ, Face>(center, face));
                    fArray_vertical.Add(face);
                    
                }
            }

            

            // Sort the face centers by descending order of Z coordinate
            //faceCenters.Sort((a, b) => b.Item1.Z.CompareTo(a.Item1.Z));

            fArray_vertical.Sort((a, b) => b.Area.CompareTo(a.Area));



            // Get the two faces with the largest area
            Face largestFace1 = fArray_vertical[0];
            Face largestFace2 = fArray_vertical[1];





            double thk1 = largestFace2.Project(largestFace1.Triangulate().Vertices[0]).Distance;

            thk = Math.Round(UnitUtils.Convert(thk1, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS));

            List<string> Srf_mid_point2 = new List<string>();

            FaceArray fArray2 = new FaceArray();
            fArray2.Append(fArray_vertical[2]);

            //获得中心面  Srfmid2
            Srf_mid_point2 = CalSrfThk_with_Srf(largestFace1, fArray2, thk);
            List<string> Srf_mid_outline_point = Outline_Rec(Srf_mid_point2);
            

            point_str = Srf_mid_outline_point.Count.ToString() + "*" + string.Join("*", Srf_mid_outline_point.ToList());

            Sql_Write_Data(conn, id, name, syname, point_str, thk, gStyleId, gstyleName);



 

        }



        private void Extract_Solid_Gemo(SQLiteConnection conn, Solid solid, string id, string name, string syname, string gStyleId, string gstyleName)
        {
            EdgeArray eArray = solid.Edges;
            FaceArray fArray = solid.Faces;
            List<string> point_list = new List<string>();
            double thk = 0;
            string point_str = "";

            foreach (Edge ed in eArray)
            {
                Curve cu = ed.AsCurve();
                if (cu != null)
                {
                    XYZ xyz1 = cu.GetEndPoint(0);
                    XYZ xyz2 = cu.GetEndPoint(1);
                    //xyzs+=xyz1 +","+ xyz2;  
                    point_list.Add(ConvertCoord2Mill(xyz1.ToString()));
                    point_list.Add(ConvertCoord2Mill(xyz2.ToString()));
                    //point_list.Add("("+  xyz1.X.ToString("f5")+","+ xyz1.Y.ToString("f5") + "," + xyz1.Y.ToString("f5")+")");
                    //point_list.Add("(" + xyz2.X.ToString("f5") + "," + xyz2.Y.ToString("f5") + "," + xyz2.Y.ToString("f5") + ")");
                    // Math.Round(UnitUtils.Convert(FinalX.X, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS)
                }
            }

            var point_list_distinct = point_list.Distinct().ToList();

            if (point_list_distinct.Count < 0)
            {
                List<List<string>> Srf_info_list = CalSrfThk(point_list_distinct);

                thk = Convert.ToDouble(Srf_info_list[0][0]);
                List<string> Srf_mid_point = new List<string>();
                for (int pi = 0; pi < Srf_info_list.Count(); pi = pi + 1)
                {
                    Srf_mid_point.Add(Srf_info_list[pi][3]);
                }

                point_str = Srf_mid_point.Count.ToString() + "*" + string.Join("*", Srf_mid_point.ToArray());
            }
            else if (point_list_distinct.Count <= 8 && point_list_distinct.Count > 2)
            {
                List<XYZ> fCorner = new List<XYZ>();
                Face f_side = null;
                double f_area = 0;
                foreach (Face f in fArray)
                {
                    if (f.Area > f_area) // 三角形或四边形，选取最大面积为拉伸基础面，并后续计算厚度
                    {
                        f_area = f.Area;
                        f_side = f;
                    }
                }

                EdgeArrayArray eLoopArray = f_side.EdgeLoops;
                foreach (EdgeArray eloopA1 in eLoopArray)
                {
                    foreach (Edge eloop1 in eloopA1)
                    {
                        List<string> fCornerS = ConvertXYZ2String(fCorner);

                        if (fCornerS.Contains(Convert.ToString(eloop1.Tessellate()[0])) == false)
                            fCorner.Add(eloop1.Tessellate()[0]);
                        if (fCornerS.Contains(Convert.ToString(eloop1.Tessellate()[1])) == false)
                            fCorner.Add(eloop1.Tessellate()[1]);
                    }
                }

                fCorner = point_sort_CW(fCorner);


                // 计算厚度


                foreach (string p in point_list_distinct)

                {
                    string p1 = p.Replace('(', ' ').Replace(')', ' ');
                    XYZ pt = new XYZ(
                        UnitUtils.Convert(Convert.ToDouble(p1.Split(',')[0]), DisplayUnitType.DUT_MILLIMETERS,
                            DisplayUnitType.DUT_DECIMAL_FEET),
                        UnitUtils.Convert(Convert.ToDouble(p1.Split(',')[1]), DisplayUnitType.DUT_MILLIMETERS,
                            DisplayUnitType.DUT_DECIMAL_FEET),
                        UnitUtils.Convert(Convert.ToDouble(p1.Split(',')[2]), DisplayUnitType.DUT_MILLIMETERS,
                            DisplayUnitType.DUT_DECIMAL_FEET));
                    thk = dist_point2srf(fCorner[0], fCorner[1], fCorner[2], pt);
                    if (thk > 1e-3)
                        break;
                }


                //求得中心面
                List<string> Srf_mid_point2 = new List<string>();

                Srf_mid_point2 = CalSrfThk_with_Srf(f_side, fArray, thk);
                point_str = Srf_mid_point2.Count.ToString() + "*" + string.Join("*", Srf_mid_point2.ToArray());
            }
            //根据节点数等于多边形数，筛选出拉伸基础面的复杂多边形顺序点位
            else if (point_list_distinct.Count > 8)
            {
                List<XYZ> fCorner = new List<XYZ>();
                Face f_side = null;
                foreach (Face f in fArray)
                {
                    EdgeArrayArray eLoopArray = f.EdgeLoops;
                    fCorner = new List<XYZ>(); // 每次点位清零

                    foreach (EdgeArray eloopA1 in eLoopArray)
                    {
                        foreach (Edge eloop1 in eloopA1)
                            fCorner.Add(eloop1.Tessellate()[0]);
                    }

                    if (fCorner.Count == point_list_distinct.Count / 2)
                    {
                        if (f != null)
                        {
                            f_side = f;
                            break;
                        }
                    }
                }


                // 计算厚度



                foreach (string p in point_list_distinct)

                {
                    string p1 = p.Replace('(', ' ').Replace(')', ' ');
                    XYZ pt = new XYZ(
                        UnitUtils.Convert(Convert.ToDouble(p1.Split(',')[0]), DisplayUnitType.DUT_MILLIMETERS,
                            DisplayUnitType.DUT_DECIMAL_FEET),
                        UnitUtils.Convert(Convert.ToDouble(p1.Split(',')[1]), DisplayUnitType.DUT_MILLIMETERS,
                            DisplayUnitType.DUT_DECIMAL_FEET),
                        UnitUtils.Convert(Convert.ToDouble(p1.Split(',')[2]), DisplayUnitType.DUT_MILLIMETERS,
                            DisplayUnitType.DUT_DECIMAL_FEET));
                    thk = dist_point2srf(fCorner[0], fCorner[1], fCorner[2], pt);
                    if (thk > 1e-3)

                        break;
                }


                //求得中心面
                List<string> Srf_mid_point2 = new List<string>();

                Srf_mid_point2 = CalSrfThk_with_Srf(f_side, fArray, thk);
                point_str = Srf_mid_point2.Count.ToString() + "*" + string.Join("*", Srf_mid_point2.ToArray());
            }


            //

            Sql_Write_Data(conn, id, name, syname, point_str, thk, gStyleId, gstyleName);
        }

        private void Extract_foldplate_Gemo(SQLiteConnection conn, Solid solid, string id, string name, string syname, string gStyleId, string gstyleName)
        {
            EdgeArray eArray = solid.Edges;
            FaceArray fArray = solid.Faces;
            List<string> point_list = new List<string>();
            string point_str = "";
            double thk = 1e20;

            List<XYZ> fCorner = new List<XYZ>();
            Face f_side = null;
            double f_area = 1e20;
            double l_e = 1e20;

            List<Face> f_list = new List<Face>();
            List<Edge> e_list = new List<Edge>();
            List<double> d_list = new List<double>();
            List<int> c_list = new List<int> { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };


            //检索折板侧边顶点，并计算板厚中点3point，最终计算中心面的3point
            double len_edge = 0;
            List<List<string>> p_face_list = new List<List<string>>();

            List<string> p_face_list_string = new List<string>();
            string p_face_string;

            foreach (Face f in fArray)
            {
                foreach (EdgeArray eloop in f.EdgeLoops)
                {
                    int i = 0;
                    //List<XYZ> p_face_list_temp =new List<XYZ>();
                    foreach (Edge e in eloop)
                    {
                        i++;
                    }

                    if (i == 6) //找到边长
                    {
                        f_list.Add(f);
                        //p_face_list.Add( f.Triangulate().Vertices.ToList());
                    }
                }
            }


            foreach (Face f in f_list)
            {
                double z_bottom = 1e20;
                List<XYZ> p_face = new List<XYZ>();

                foreach (EdgeArray eloop in f.EdgeLoops)
                {
                    foreach (Edge e in eloop)
                    {
                        string pm = CalMidPoint(Convert.ToString(e.Tessellate()[0]),
                            Convert.ToString(e.Tessellate()[1]));
                        double pm_z = Convert.ToDouble(pm.Substring(1, pm.Length - 2).Split(',')[2]);
                        if (pm_z < z_bottom)
                        {
                            z_bottom = pm_z;
                            p_face = e.Tessellate().ToList();
                        }
                    }


                    foreach (Edge e in eloop)
                    {
                        double len_e = UnitUtils.Convert(e.ApproximateLength, DisplayUnitType.DUT_DECIMAL_FEET,
                            DisplayUnitType.DUT_MILLIMETERS);
                        if (len_e < 50)// 厚度按50预估
                        {
                            continue;

                        }
                        if (Math.Abs(e.Tessellate()[0].Z - p_face[0].Z) < 0.01)
                        {
                            //if (p_face.Contains(e.Tessellate()[1]) == false)
                            {
                                p_face.Add(e.Tessellate()[1]);
                                continue;
                            }
                        }
                        else if (Math.Abs(e.Tessellate()[1].Z - p_face[0].Z) < 0.01)
                        {
                            //if (p_face.Contains(e.Tessellate()[0]) == false)
                            {
                                p_face.Add(e.Tessellate()[0]);
                                continue;
                            }
                        }

                    }

                    List<string> p_face_dics = ConvertXYZ2String(point_sort_CW(p_face));

                    p_face_dics.RemoveAt(1);
                    //p_face_list.Add(p_face_dics);
                    for (int i = 0; i < 3; i++)
                    {
                        p_face_dics[i] = ConvertCoord2Mill(p_face_dics[i]);
                    }
                    p_face_list_string.Add(string.Join(",", p_face_dics));
                }

            }

            p_face_string = string.Join("*", p_face_list_string);



            EdgeArray EdgeArray1 = solid.Edges;



            foreach (Edge e1 in EdgeArray1)
            {
                //string p1 = ConvertCoord2Mill(Convert.ToString(v1));

                //    string p2 = ConvertCoord2Mill(Convert.ToString(v2));
                thk = Math.Round(UnitUtils.Convert(e1.ApproximateLength, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS), 0);

                Console.WriteLine("edge_length " + thk);
                //CalDist(p1, p2)
                if (thk == 0)
                    continue;

                if (d_list.Contains(thk) == false)
                {
                    d_list.Add(thk);
                }
                else
                {
                    c_list[d_list.IndexOf(thk)]++;
                }



            }

            //踏板长度，根据六条边等长原则筛选而得
            len_edge = d_list[c_list.IndexOf(6)];

            //壁厚，最小数值为板厚
            thk = d_list.Min();



            point_str = len_edge + "*" + p_face_string;

            //
            //point_str = "2*" + p_s + "*" + p_e;
            Sql_Write_Data(conn, id, name, syname, point_str, thk, gStyleId, gstyleName);
        }

        private void Extract_Beam_Gemo(SQLiteConnection conn, Line BeamLine, string id, string name, string syname, string gStyleId, string gstyleName)
        {
            string point_str = "";
            double thk = 0;

            string p_s = ConvertCoord2Mill(BeamLine.Tessellate()[0].ToString());
            string p_e = ConvertCoord2Mill(BeamLine.Tessellate()[1].ToString());

            //
            point_str = "2*" + p_s + "*" + p_e;
            Sql_Write_Data(conn, id, name, syname, point_str, thk, gStyleId, gstyleName);
        }


        private void Extract_Rebar_Gemo(SQLiteConnection conn, Solid solid, string id, string name, string syname, string gStyleId, string gstyleName)
        {
            string point_str = "";
            double thk = 0;
            List<XYZ> p_list = new List<XYZ>();

            FaceArray F_array = solid.Faces;
            foreach (Face f in F_array)
            {
                PlanarFace planarFace = f as PlanarFace;
                if (null != planarFace)
                {
                    p_list.Add(planarFace.Origin);
                }
            }

            double l0 = CalDist(p_list[0].ToString(), p_list[1].ToString());
            thk = Math.Round(UnitUtils.Convert(Math.Sqrt(solid.Volume / l0 / Math.PI * 4), DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS), 0);

            string p_s = ConvertCoord2Mill(p_list[0].ToString());
            string p_e = ConvertCoord2Mill(p_list[1].ToString());

            //
            point_str = "2*" + p_s + "*" + p_e;
            Sql_Write_Data(conn, id, name, syname, point_str, thk, gStyleId, gstyleName);
        }


        private void Extract_Slab_Gemo(SQLiteConnection conn, Solid solid, string id, string name, string syname, string gStyleId, string gstyleName)
        {


            FaceArray fArray = solid.Faces;
            List<string> point_list = new List<string>();
            double thk = 0;
            string point_str = "";


            // Create a list to store face center points
            List<Tuple<XYZ, Face>> faceCenters = new List<Tuple<XYZ, Face>>();

            // Iterate over each face in the array and calculate the center point
            foreach (Face face in fArray)
            {
                XYZ center = GetFaceCenter(face);
                faceCenters.Add(new Tuple<XYZ, Face>(center, face));
            }

            // Sort the face centers by descending order of Z coordinate
            faceCenters.Sort((a, b) => b.Item1.Z.CompareTo(a.Item1.Z));

            // Get the two faces with the largest Z coordinates
            Face largestFace1 = faceCenters[0].Item2;//最上板
            Face largestFace2 = faceCenters[faceCenters.Count-1].Item2;//最下板

            XYZ faceNormal = largestFace1.ComputeNormal(new UV());
            if (Math.Abs(faceNormal.X) > 0.0001 || Math.Abs(faceNormal.Y) > 0.0001) //判定仅保留垂直于Z周的结构楼板
            {
                return;
            }

            XYZ largestFace1Center = faceCenters[0].Item1;
            XYZ largestFace2Center = faceCenters[faceCenters.Count - 1].Item1;


            thk = Math.Round(UnitUtils.Convert(largestFace1Center.Z - largestFace2Center.Z, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS));





            // The variable "largestFace1" now contains the face with the largest area,
            // and potentially the face with the larger z-coordinate if the area difference is within the threshold.

            //EdgeArray eArray = largestFace2.EdgeLoops.get_Item(0);
            



                    //EdgeArrayArray eLoopArray = largestFace2.EdgeLoops;
                List<string> fCorner = new List<string>(); // 每次点位清零

                    //foreach (EdgeArray eloopA1 in largestFace2.EdgeLoops)
                    //{
                    //    foreach (Edge eloop1 in eloopA1)
                    //        fCorner.Add(ConvertCoord2Mill(Convert.ToString(eloop1.Tessellate()[0])));
                    //}

                foreach(XYZ p_corner in largestFace1.Triangulate().Vertices)
                    fCorner.Add(ConvertCoord2Mill(Convert.ToString(p_corner)));


            point_str = fCorner.Count.ToString() + "*" + string.Join("*", fCorner.ToList());


            //

            Sql_Write_Data(conn, id, name, syname, point_str, thk, gStyleId, gstyleName);
        }

        private void Extract_DoorWindow_Gemo(Document Doc,SQLiteConnection conn, Solid solid, string id, string name, string syname, string gStyleId, string gstyleName)
        {
            



            XYZ boundary1 = solid.GetBoundingBox().Max;
            XYZ boundary2 = solid.GetBoundingBox().Min;

            List<XYZ> boundaryList = new List<XYZ>();
            boundaryList.Add(boundary1);
            boundaryList.Add(boundary2);

            XYZ d = boundary1 - boundary2;

            // 计算长方体的长、宽和高
            double length = Math.Abs(d.X) * 304.8;
            double width = Math.Abs(d.Y) * 304.8;
            double height = Math.Abs(d.Z)*304.8;

            List<double> SizeList = new List<double>{length, width };

            SizeList.Sort();

            for( int i =0;i<SizeList.Count;i++)
            {
                SizeList[i]= Math.Round(SizeList[i],0);
            }

            //double sim_thick = Math.Round(UnitUtils.Convert(SizeList[2], DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS)); ;
            double sim_thick = SizeList[0];

            if (sim_thick<50  || sim_thick %50 !=0)
                return;



            // 假设 solid 是你要查询的 Solid 对象

            // 获取 Solid 的所有面
            FaceArray faces = solid.Faces;

            // 创建一个列表来存储所有的顶点坐标
            List<XYZ> vertexCoordinates = new List<XYZ>();
            List<XYZ> centerPoint = new List<XYZ>();
            // 遍历每个面，收集顶点坐标
            foreach (Face face in faces)
            {
                // 获取面的顶点
                List<XYZ> vertices = face.Triangulate().Vertices.ToList();

                // 将顶点坐标添加到列表中
                foreach (XYZ vertex in vertices)
                {
                    vertexCoordinates.Add(vertex);
                }
            }

            // 计算所有顶点的平均值，即为中心点坐标
            centerPoint.Add( new XYZ(
                vertexCoordinates.Average(v => v.X),
                vertexCoordinates.Average(v => v.Y),
                vertexCoordinates.Average(v => v.Z)
            ));




            string point_str = SizeList[1] + "-" + height + "*"+ConvertCoord2Mill(ConvertXYZ2String(centerPoint)[0]);//Srf_mid_outline_point.Count.ToString() + "*" + string.Join("*", Srf_mid_outline_point.ToList());

            Sql_Write_Data(conn, id, name, syname, point_str, sim_thick, gStyleId, gstyleName);


        }

        

        // Helper method to calculate the approximate center point of a face
        private XYZ GetFaceCenter(Face face)
        {
            EdgeArrayArray edgeLoops = face.EdgeLoops;

            XYZ sum = XYZ.Zero;
            int count = 0;

            foreach (EdgeArray edgeLoop in edgeLoops)
            {
                foreach (Edge edge in edgeLoop)
                {
                    XYZ startPoint = edge.Evaluate(0);
                    XYZ endPoint = edge.Evaluate(1);

                    sum += (startPoint + endPoint) / 2;
                    count++;
                }
            }

            return sum / count;
        }

        private static void Sql_Write_Data(SQLiteConnection conn, string id, string name, string syname, string point_str, double thk, string gStyleId,
            string gstyleName)
        {

            SQLiteCommand cmd = new SQLiteCommand();
            conn.Open();
            cmd.Connection = conn;

            cmd.CommandText =
                "insert into ModuleInfo(ID,FamilyName,CatalogName,PointCoord,thickness,detailId,detailName) values(@ID,@FamilyName,@CatalogName,@PointCoord,@thickness,@detailId,@detailName)";


            cmd.Parameters.AddWithValue("@ID", id);
            cmd.Parameters.AddWithValue("@FamilyName", name);
            cmd.Parameters.AddWithValue("@CatalogName", syname);
            cmd.Parameters.AddWithValue("@PointCoord", point_str);
            cmd.Parameters.AddWithValue("@thickness", Convert.ToString(Math.Round(thk)));
            cmd.Parameters.AddWithValue("@detailId", gStyleId);
            cmd.Parameters.AddWithValue("@detailName", gstyleName);

            cmd.ExecuteNonQuery();

            conn.Close();
        }

        private SQLiteCommand Sql_Clean_Data(SQLiteConnection conn)
        {
            conn.Open();

            SQLiteCommand cmd = new SQLiteCommand();

            cmd.Connection = conn;

            ClearTable(cmd); //清空数据表
            conn.Close();
            return cmd;
        }

        /// <summary>
        /// 根据点位，计算配对的两点中点，并筛选最短距离即为厚度方向，寻找匹配的点位，求的中点坐标
        /// </summary>
        /// <param name="point_list"></param>
        /// <returns></returns>
        private List<List<string>> CalSrfThk(List<string> point_list)//out double thk, out List<string>point_mid)
        {
            double thk = 1e5;
            List<List<string>> SrfInfo = new List<List<string>>();

            for (int i = 0; i < point_list.Count(); i = i + 1)
            {
                string p_i = point_list[i];
                for (int j = i + 1; j < point_list.Count(); j = j + 1)
                {
                    string p_j = point_list[j];
                    double d1 = CalDist(p_i, p_j);
                    //System.Diagnostics.Debug.WriteLine(Convert.ToString(d1) + ":" + Convert.ToString(thk));
                    if (d1 < thk * 1.05)
                    {
                        thk = d1;
                        List<String> info = new List<String>();
                        info.Add(d1.ToString());
                        info.Add(p_i);
                        info.Add(p_j);
                        SrfInfo.Add(info); // SrfInfo[i]=[dist,p1,p2]
                    }
                }

            }
            for (int i = 0; i < SrfInfo.Count(); i = i + 1)
            {
                if (Convert.ToDouble(SrfInfo[i][0]) < thk * 1.05)
                {
                    string midpoint = CalMidPoint(SrfInfo[i][1], SrfInfo[i][2]);
                    SrfInfo[i].Add(midpoint);
                }
                else
                {
                    SrfInfo[i] = null;
                }
            }

            RemoveNull(SrfInfo);


            return SrfInfo;// [0]--厚度 ;[1][2]起始终点; [3]midpoint
        }


        private List<string> CalSrfThk_with_Srf(Face f1, FaceArray fArray, double thk0)//out double thk, out List<string>point_mid)
        {

            //可考虑删除，已有Srfmid2 替代
            List<string> Srfmid = new List<string>();


            foreach (EdgeArray eloop1 in f1.EdgeLoops)
            {
                foreach (Edge e1 in eloop1)
                {
                    string ps = ConvertCoord2Mill(Convert.ToString(e1.Tessellate()[0]));
                    int Flag = 0;
                    foreach (Face f in fArray)
                    {

                        foreach (EdgeArray eloop in f.EdgeLoops)
                        {
                            foreach (Edge e in eloop)
                            {
                                Curve curve_edge = e.AsCurve();
                                double d0 = Math.Abs(UnitUtils.Convert(curve_edge.Length, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS));
                                string p1 = ConvertCoord2Mill(Convert.ToString(curve_edge.Tessellate()[0]));
                                string p2 = ConvertCoord2Mill(Convert.ToString(curve_edge.Tessellate()[1]));

                                double select_edge = Math.Abs(CalDist(p1, ps) * CalDist(p2, ps));

                                if (select_edge < 1e-3)
                                    Flag++;
                                if (Math.Abs(d0 - thk0) < 0.1)
                                {

                                    string midpoint = CalMidPoint(p1, p2);
                                    if (Srfmid.Contains(midpoint) == false)
                                        Srfmid.Add(midpoint);
                                }
                            }
                        }
                    }
                }
            }
            //



            List<string> Srfmid2 = new List<string>();
            List<XYZ> verticeList_XYZ = GetFaceVertices(f1);
            List<XYZ> verticeList_XYZ_ecc = new List<XYZ>();

            XYZ p1_project = new XYZ(); // 取另一个面上的点，以做中心偏移向量；
            foreach (Face f2 in fArray)
            {
                //if (f2.Triangulate().Vertices.Count != f1.Triangulate().Vertices.Count)
                if (Math.Abs(f2.Area - f1.Area) > 1e-3) // 在fArray中找到一个非f1，面积不等
                {
                    foreach (XYZ f2_p in f2.Triangulate().Vertices)
                    {
                        if (f1.Project(f2_p).Distance > 1e-3) //f2改f1.project
                        {
                            p1_project = f2_p;
                            break;
                        }
                    }

                    break;
                }
            }

            XYZ vec_ecc = Vector_ecc(f1, p1_project);// 求出中心偏移向量

            foreach (XYZ i_XYZ in verticeList_XYZ)
            {
                verticeList_XYZ_ecc.Add(i_XYZ + vec_ecc);
            }

            List<string> verticeList_string = ConvertXYZ2String(verticeList_XYZ_ecc);


            foreach (string i_vertice in verticeList_string)
            {
                Srfmid2.Add(ConvertCoord2Mill(i_vertice));
            }

            return Srfmid2;
        }

        private XYZ Vector_ecc(Face f1, XYZ p1)
        {
            //判断 p1 是否属于f1，如果不是，则计算投影，取投影的1/2 * -1,可得到中心面偏移向量，在已有的f1上全部加上该向量，则为中心面
            XYZ p1_project = f1.Project(p1).XYZPoint;
            List<XYZ> vec_ecc = new List<XYZ>();
            vec_ecc.Add((p1 - p1_project) * 0.5);


            return vec_ecc[0];


        }

        public List<XYZ> GetFaceVertices(Face face)
        {
            List<XYZ> vertices = new List<XYZ>();

            // 获取面的几何信息
            PlanarFace planarFace = face as PlanarFace;
            if (planarFace == null) return vertices;
            XYZ normal = planarFace.FaceNormal;

            // 如果法向量的 Z 坐标不等于 0，则无法确定顶点的顺序
            if (Math.Abs(normal.Z) > double.Epsilon)
            {
                return vertices;
            }

            // 获取面的边界
            EdgeArray edgeArray = face.EdgeLoops.get_Item(0);
            if (edgeArray == null) return vertices;

            // 遍历边界上的所有边
            while (vertices.Count <= edgeArray.Size)
            {
                foreach (Edge edge in edgeArray)
                {
                    // 获取边的几何信息
                    Curve curve = edge.AsCurve();
                    if (curve == null) continue;
                    XYZ startPoint = curve.GetEndPoint(0);
                    XYZ endPoint = curve.GetEndPoint(1);

                    // 判断起点和终点的顺序，保证起点在前，终点在后



                    if (vertices.Count == 0)
                    {
                        vertices.Add(startPoint);
                        //vertices.Add(endPoint);

                    }

                    //if (vertices.Count > edgeArray.Size)
                    //{
                    //    break;
                    //}


                    if (vertices.Count >= 1)
                    {
                        if (startPoint.IsAlmostEqualTo(vertices[vertices.Count - 1]))
                        {
                            vertices.Add(endPoint);
                        }
                        else if (endPoint.IsAlmostEqualTo(vertices[vertices.Count - 1]))
                        {
                            vertices.Add(startPoint);
                        }

                    }



                }

            }

            //// 计算重心
            //XYZ centroid = new XYZ(
            //    vertices.Average(v => v.X),
            //    vertices.Average(v => v.Y),
            //    vertices.Average(v => v.Z)
            //);

            // 根据重心对顶点进行逆时针排序
            //vertices = vertices.OrderBy(v => Math.Atan2(v.Y - centroid.Y, v.X - centroid.X)).ToList();

            // 如果法向量的 X 坐标为正，则顶点顺序应该翻转
            if (normal.X > double.Epsilon)
            {
                vertices.Reverse();
            }

            return vertices;
        }

        //public List<XYZ> GetSortedVertices(Face face)
        //{
        //    // 获取面对象的顶点列表
        //    List<XYZ> vertices = face.Triangulate().Vertices.ToList();

        //    // 获取面对象的法向量
        //    XYZ normal = face.ComputeNormal(new UV());

        //    // 将顶点按照相对于面法向量的逆时针顺序排序
        //    vertices = SortVerticesByNormal(vertices, normal);

        //    return vertices;
        //}

        //private List<XYZ> SortVerticesByNormal(List<XYZ> vertices, XYZ normal)
        //{
        //    // 计算顶点列表的中心点
        //    XYZ centroid = new XYZ(0, 0, 0);
        //    foreach (XYZ vertex in vertices)
        //    {
        //        centroid += vertex;
        //    }
        //    centroid /= vertices.Count;

        //    // 将顶点列表中的每个顶点与中心点进行排序
        //    vertices = vertices.OrderBy(v => Math.Atan2(v.Y - centroid.Y, v.X - centroid.X)).ToList();

        //    // 如果法向量与 Z 轴平行，则直接返回排序后的顶点列表
        //    if (Math.Abs(normal.Z - 1) < double.Epsilon || Math.Abs(normal.Z + 1) < double.Epsilon)
        //    {
        //        return vertices;
        //    }

        //    // 如果法向量与 Z 轴不平行，则通过判断每个顶点与前一个顶点的叉乘来判断排序是否正确
        //    double crossProduct;
        //    XYZ firstVector = vertices[1] - vertices[0];
        //    XYZ secondVector = vertices[2] - vertices[0];
        //    crossProduct = (firstVector.X * secondVector.Y) - (firstVector.Y * secondVector.X);
        //    if ((crossProduct * normal.Z) < 0)
        //    {
        //        vertices.Reverse();
        //    }

        //    return vertices;
        //}



        //求点到平面的距离
        //已知3点坐标，求平面ax+by+cz+d=0;

        double dist_point2srf(XYZ p1, XYZ p2, XYZ p3, XYZ pt)
        {
            double a;
            double b;
            double c;
            double d;
            double dist;

            a = (p2.Y - p1.Y) * (p3.Z - p1.Z) - (p2.Z - p1.Z) * (p3.Y - p1.Y);

            b = (p2.Z - p1.Z) * (p3.X - p1.X) - (p2.X - p1.X) * (p3.Z - p1.Z);

            c = (p2.X - p1.X) * (p3.Y - p1.Y) - (p2.Y - p1.Y) * (p3.X - p1.X);

            d = 0 - (a * p1.X + b * p1.Y + c * p1.Z);

            if ((a * a + b * b + c * c) == 0)

                dist = 0;
            else
                dist = Math.Abs(a * pt.X + b * pt.Y + c * pt.Z + d) / Math.Sqrt(a * a + b * b + c * c);

            //return dist;
            return UnitUtils.Convert(dist, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);
        }







        private double CalDist(string p1, string p2)
        {
            double p1x, p1y, p1z, p2x, p2y, p2z;
            var temp = p1.Substring(1, p1.Length - 1).Split(',');
            p1x = Convert.ToDouble(p1.Substring(1, p1.Length - 2).Split(',')[0]);
            p1y = Convert.ToDouble(p1.Substring(1, p1.Length - 2).Split(',')[1]);
            p1z = Convert.ToDouble(p1.Substring(1, p1.Length - 2).Split(',')[2]);

            p2x = Convert.ToDouble(p2.Substring(1, p2.Length - 2).Split(',')[0]);
            p2y = Convert.ToDouble(p2.Substring(1, p2.Length - 2).Split(',')[1]);
            p2z = Convert.ToDouble(p2.Substring(1, p2.Length - 2).Split(',')[2]);

            double dist = Math.Sqrt((p1x - p2x) * (p1x - p2x) + (p1y - p2y) * (p1y - p2y) + (p1z - p2z) * (p1z - p2z));

            return dist;
        }

        private string CalMidPoint(string pa, string pb)
        {
            double pax, pay, paz, pbx, pby, pbz;
            pax = Convert.ToDouble(pa.Substring(1, pa.Length - 2).Split(',')[0]);
            pay = Convert.ToDouble(pa.Substring(1, pa.Length - 2).Split(',')[1]);
            paz = Convert.ToDouble(pa.Substring(1, pa.Length - 2).Split(',')[2]);

            pbx = Convert.ToDouble(pb.Substring(1, pb.Length - 2).Split(',')[0]);
            pby = Convert.ToDouble(pb.Substring(1, pb.Length - 2).Split(',')[1]);
            pbz = Convert.ToDouble(pb.Substring(1, pb.Length - 2).Split(',')[2]);

            string pm = "(" + Math.Round((pax + pbx) / 2, 3) + "," + Math.Round((pay + pby) / 2, 3) + "," + Math.Round((paz + pbz) / 2, 3) + ")";

            return pm;
        }

        static void RemoveNull<T>(List<T> list)
        {
            // 找出第一个空元素 O(n)
            int count = list.Count;
            for (int i = 0; i < count; i++)
                if (list[i] == null)
                {
                    // 记录当前位置
                    int newCount = i++;

                    // 对每个非空元素，复制至当前位置 O(n)
                    for (; i < count; i++)
                        if (list[i] != null)
                            list[newCount++] = list[i];

                    // 移除多余的元素 O(n)
                    list.RemoveRange(newCount, count - newCount);
                    break;
                }
        }

        static string ConvertCoord2Mill(string point)
        {
            double px, py, pz;
            string converted_point;
            px = UnitUtils.Convert(Math.Round(Convert.ToDouble(point.Substring(1, point.Length - 2).Split(',')[0]), 10), DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);
            py = UnitUtils.Convert(Math.Round(Convert.ToDouble(point.Substring(1, point.Length - 2).Split(',')[1]), 10), DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);
            pz = UnitUtils.Convert(Math.Round(Convert.ToDouble(point.Substring(1, point.Length - 2).Split(',')[2]), 10), DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

            converted_point = "(" + Convert.ToString(px) + "," + Convert.ToString(py) + "," + Convert.ToString(pz) + ")";
            return converted_point;
        }


        static List<string> ConvertXYZ2String(List<XYZ> fCorner)
        {
            List<string> fCornerS = new List<string>();
            foreach (XYZ p in fCorner)
            {
                fCornerS.Add(Convert.ToString(p));
            }

            return fCornerS;
        }

        static List<XYZ> point_sort_CW(List<XYZ> pointList)
        {
            //中点作为参考点
            double Xm = 0;
            double Ym = 0;
            double Zm = 0;
            int N = pointList.Count;

            foreach (XYZ p in pointList)
            {
                Xm = Xm + p.X;
                Ym = Ym + p.Y;
                Zm = Zm + p.Z;
            }

            XYZ pm = new XYZ(Xm / N, Ym / N, Zm / N);


            List<Vector3> vectorList = new List<Vector3>();
            List<double> angleList = new List<double>();
            //Vector3 v1 = new Vector3((float)(pointList[0].X - pm.X), (float)(pointList[0].Y - pm.Y), (float)(pointList[0].Z - pm.Z));

            //vectorList.Add(v1);

            //形成中点到各个点的向量
            foreach (XYZ p in pointList)
            {
                Vector3 vp = new Vector3((float)(p.X - pm.X), (float)(p.Y - pm.Y), (float)(p.Z - pm.Z));
                vectorList.Add(vp);
            }
            //计算向量1与各个向量的夹角 (0,2pi)

            Vector3 v1 = vectorList[0];
            Vector3 v2 = vectorList[1];
            foreach (Vector3 vp in vectorList)
                angleList.Add(vector_angle_2pi(v1, v2, vp));

            //按角度进行排序


            for (int i = 0; i < pointList.Count - 1; i++)  //外层循环控制排序趟数
            {
                for (int j = 0; j < pointList.Count - 1 - i; j++)  //内层循环控制每一趟排序多少次
                {
                    if (angleList[j] > angleList[j + 1])
                    {
                        var temp = pointList[j];
                        pointList[j] = pointList[j + 1];
                        pointList[j + 1] = temp;
                    }
                }
            }

            return pointList;
        }

        static double vector_angle_2pi(Vector3 v1, Vector3 v2, Vector3 v3)//求v1，v3夹角；方向参考v1->v2，同向为正，反向为负
        {
            double angle_2pi;

            double angle = Math.Acos(Math.Round(Vector3.Dot(v1, v3) / (v1.Length() * v3.Length()), 3));//此处round,避免出现大于1的奇异值

            //double angle2 = Math.Acos(1);
            Vector3 dir_v1_v2 = Vector3.Cross(v1, v2);
            Vector3 dir_v1_v3 = Vector3.Cross(v1, v3);

            if (Vector3.Dot(dir_v1_v2, dir_v1_v3) < 0) //V1->V2 逆时针为正
                angle_2pi = 2 * Math.PI - angle;
            else
                angle_2pi = angle;
            //计算theta = arccos()

            //计算叉乘向量,并调整theta
            return angle_2pi / Math.PI * 180;
        }

        public static List<string> Outline_Rec(List<string> point_List)
        {
            List<List<double>> pointList = new List<List<double>>();

            // 将点坐标转换为数值类型
            foreach (string pointStr in point_List)
            {
                string[] coordinates = pointStr.Substring(1,pointStr.Length-2).Split(',');
                double x = double.Parse(coordinates[0]);
                double y = double.Parse(coordinates[1]);
                double z = double.Parse(coordinates[2]);
                List<double> point = new List<double> { x, y, z };
                pointList.Add(point);
            }

            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double minZ = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            double maxZ = double.MinValue;

            // 计算多边形的外包矩形
            foreach (List<double> point in pointList)
            {
                double x = point[0];
                double y = point[1];
                double z = point[2];
                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                minZ = Math.Min(minZ, z);
                maxX = Math.Max(maxX, x);
                maxY = Math.Max(maxY, y);
                maxZ = Math.Max(maxZ, z);
            }

            List<string> outlineRectangle = new List<string>();

            // 外包矩形的八个角点
            outlineRectangle.Add("(" + minX + "," + minY + "," + minZ + ")");
            outlineRectangle.Add("(" + maxX + "," + maxY + "," + minZ + ")");
            outlineRectangle.Add("(" + minX + "," + minY + "," + maxZ + ")");
            outlineRectangle.Add("(" + maxX + "," + maxY + "," + maxZ + ")");
            //outlineRectangle.Add(new List<double> { maxX, maxY, minZ });
            //outlineRectangle.Add(new List<double> { minX, minY, maxZ });
            //outlineRectangle.Add(new List<double> { maxX, maxY, maxZ });

            return outlineRectangle;
        }



    }
}
