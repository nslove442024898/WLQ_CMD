using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Interop;
using System.IO;
using System.Diagnostics;
using Autodesk.AutoCAD.Geometry;
using WLQ_CMD;

namespace WLQ_CMD
{
    public class myPart
    {
        public double Thk { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
        public string Material { get; set; }
        public int Qty { get; set; }
        public double Area { get; set; }
        public string PanelName { get; set; }
        public string PieceMark { get; set; }

        public Database acDb { get; set; }
        public myPart(Polyline partBoundary, List<DBText> listTexts, string strPanelName, Database _db)
        {
            this.PanelName = strPanelName;

            this.acDb = _db;
            using (Transaction trans = this.acDb.TransactionManager.StartTransaction())
            {
                var pnts = helper.GetMaxMinPoint(partBoundary, this.acDb);
                this.Area = (partBoundary as Curve).Area;
                #region//提取长宽尺寸
                var L = (pnts[0] - pnts[2]);
                var w = pnts[1] - pnts[3];
                if (L >= w)
                {
                    this.Length = L;
                    this.Width = w;
                }
                else
                {
                    this.Length = w;
                    this.Width = L;
                }
                #endregion
                var partName = listTexts.Where(c => c.Layer == "0" && c.ColorIndex == 6
                  && c.Position.X > pnts[2] && c.Position.X < pnts[0]
                  && c.Position.Y > pnts[3] && c.Position.Y < pnts[1]).ToList();
                var materialInfor = listTexts.Where(c => c.Layer == "0" && c.ColorIndex.ToString() == "205"
                   && c.Position.X > pnts[2] && c.Position.X < pnts[0]
                   && c.Position.Y > pnts[3] && c.Position.Y < pnts[1]).ToList();
                if (partName.Count == 1) this.PieceMark = partName[0].TextString.Replace("%%U", "");
                if (materialInfor.Count == 3)
                {
                    for (int i = materialInfor.Count - 1; i >= 0; i--)
                    {
                        if (materialInfor[i].TextString.Contains("=")) { this.Qty = int.Parse(materialInfor[i].TextString.Split('=')[1]); }
                        double x = 0;
                        if (double.TryParse(materialInfor[i].TextString, out x)) { this.Thk = x; }
                        if (!materialInfor[i].TextString.Contains("=") && x == 0) { this.Material = materialInfor[i].TextString; }
                    }
                }
                trans.Commit();
            }
        }
    }

    public class MyPanel
    {
        public int 单元件数量 { get; set; }
        public string 单元件名称 { get; set; }
        public string 单元件编号 { get; set; }
        public List<myPart> listMyParts { get; set; }

        public Database acDb;
        public MyPanel(Polyline panelTitleBoundary, List<DBText> lsitAllText, List<Polyline> partBoundarys, Database db)
        {
            this.acDb = db;
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                #region//获取边界
                var pnts = helper.GetMaxMinPoint(panelTitleBoundary, this.acDb);

                #endregion
                //var panelInfors = lsitAllText.Where(c => c.Layer == "A3" && c.ColorIndex.ToString() == "171"
                // && c.Position.X > minPnt.X && c.Position.X < maxPnt.X
                // && c.Position.Y > minPnt.Y && c.Position.Y < maxPnt.Y).ToList();
                var panelInfors = lsitAllText.Where(c => c.Layer == "A3" && c.ColorIndex.ToString() == "171"
                 && c.Position.X > pnts[2] && c.Position.X < pnts[0]
                 && c.Position.Y > pnts[3] && c.Position.Y < pnts[1]).ToList();
                if (panelInfors.Count == 3)
                {
                    for (int i = panelInfors.Count - 1; i >= 0; i--)
                    {
                        if (!panelInfors[i].TextString.Contains("单元件") && !panelInfors[i].TextString.Contains(" 件")) this.单元件编号 = panelInfors[i].TextString;

                        if (panelInfors[i].TextString.Contains(" 件")) this.单元件数量 = int.Parse(panelInfors[i].TextString.Substring(0, panelInfors[i].TextString.Length - 1));

                        if (panelInfors[i].TextString.Contains("单元件")) this.单元件名称 = panelInfors[i].TextString;
                    }
                }
                //提取当前图框的的文字
                var textInsideCurrentTitleblok = lsitAllText.Where(c => c.Layer == "0" && (c.ColorIndex == 6 || c.ColorIndex.ToString() == "205")
                 && c.Position.X > pnts[2] && c.Position.X < pnts[0]
                 && c.Position.Y > pnts[3] && c.Position.Y < pnts[1]).ToList();
                ///提取当前图框的的零件
                var partsInsideCurrentTitleblok = partBoundarys.Where(c => c.Layer == "0" && c.ColorIndex == 1
             && helper.GetMaxMinPoint(c, this.acDb)[2] > pnts[2] && helper.GetMaxMinPoint(c, this.acDb)[0] < pnts[0]
             && helper.GetMaxMinPoint(c, this.acDb)[3] > pnts[3] && helper.GetMaxMinPoint(c, this.acDb)[1] < pnts[1]).ToList();
                this.listMyParts = new List<myPart>();
                foreach (var item in partsInsideCurrentTitleblok)
                {
                    myPart part = new myPart(item, textInsideCurrentTitleblok, this.单元件编号, this.acDb);
                    this.listMyParts.Add(part);
                }
                trans.Commit();
            }
        }
    }

    public class MyDrawing
    {
        public List<Polyline> TitleBlocks { get; set; }
        public List<Polyline> ListPartsBoundary { get; set; }
        public List<DBText> ListAllUsefullTexts { get; set; }
        public List<MyPanel> CurDwgPanels { get; set; }
        public string FileName { get; set; }
        public MyDrawing(string _fileName)
        {
            this.TitleBlocks = new List<Polyline>();

            this.ListPartsBoundary = new List<Polyline>();

            this.ListAllUsefullTexts = new List<DBText>();

            this.CurDwgPanels = new List<MyPanel>();

            this.FileName = _fileName;

            #region//读取图纸的内容
            using (Database db = new Database(false, true))
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                db.ReadDwgFile(this.FileName, FileShare.Read, true, null);
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                //打开数据库的模型空间块表记录对象
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                //循环遍历模型空间中的实体
                foreach (ObjectId id in btr)
                {
                    Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                    //显示实体的类名
                    if (ent != null)
                    {
                        if (ent is Polyline)
                        {
                            Debug.Print(ent.ColorIndex.ToString());
                            if ((ent as Polyline).Layer == "A3" && ent.ColorIndex.ToString() == "171") this.TitleBlocks.Add(ent as Polyline);//图框
                            if ((ent as Polyline).Layer == "0" && ent.ColorIndex == 1) this.ListPartsBoundary.Add(ent as Polyline);//零件
                        }
                        else if (ent is DBText)
                        {
                            Debug.Print(ent.ColorIndex.ToString());
                            if ((ent as DBText).Layer == "A3" && ent.ColorIndex.ToString() == "171") this.ListAllUsefullTexts.Add(ent as DBText);//图框文字
                            if ((ent as DBText).Layer == "0" && (ent.ColorIndex.ToString() == "205" || ent.ColorIndex == 6)) this.ListAllUsefullTexts.Add(ent as DBText);//零件信息文字
                        }
                        else continue;
                    }

                }
                foreach (var item in this.TitleBlocks)//看每一个图框内的内容
                {
                    MyPanel tempPanel = new MyPanel(item, this.ListAllUsefullTexts, this.ListPartsBoundary, db);
                    this.CurDwgPanels.Add(tempPanel);
                }

                trans.Commit();
                db.Dispose();
            }
            #endregion
        }



    }
}
