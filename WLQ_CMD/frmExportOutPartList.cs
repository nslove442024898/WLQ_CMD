using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace WLQ_CMD
{
    public partial class frmExportOutPartList : Form
    {
        public List<string> dwgfileNames { get; set; }
        public List<MyDrawing> AllParts { get; set; }

        public frmExportOutPartList()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 第一步打开文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            this.listBox1.Items.Clear();
            this.dwgfileNames = new List<string>();
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("选择零件图", null, "dwg;", "ok", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in ofd.GetFilenames()) { this.dwgfileNames.Add(item); }
                if (this.dwgfileNames.Count > 0) foreach (var item in this.dwgfileNames) this.listBox1.Items.Add(item);
            }
        }
        /// <summary>
        /// 第二部提取信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (this.dwgfileNames.Count > 0)
            {
                this.AllParts = new List<MyDrawing>();
                foreach (var item in this.dwgfileNames)
                {
                    MyDrawing curDwg = new MyDrawing(item);

                    this.AllParts.Add(curDwg);
                }
                if (this.AllParts.Count > 0)
                {
                    MemoryStream ms = new MemoryStream();
                    IWorkbook wb = new HSSFWorkbook();//创建Workbook对象
                    var firstRows = new string[] { "单元件名称", "单元件编号", "单元件数量", "编号", "规格", "材料", "面积", "数量" };
                    foreach (var item in this.AllParts)
                    {
                        ISheet ws = wb.CreateSheet(Path.GetFileNameWithoutExtension(item.FileName));//创建工作表
                        IRow row = ws.CreateRow(0);//在工作表中添加一行
                        int i = 1;
                        for (int j = 0; j < firstRows.Length; j++) row.CreateCell(j).SetCellValue(firstRows[i]);//写抬头内容

                        foreach (var itemDwg in item.CurDwgPanels)
                        {
                            foreach (var ItemPart in itemDwg.listMyParts)
                            {
                                string[] strItemContents = new string[] { itemDwg.单元件名称, itemDwg.单元件编号, itemDwg.单元件数量.ToString(), ItemPart.PieceMark,ItemPart.Thk+"x"+ Math.Round(ItemPart.Length,0)+"x"+ Math.Round(ItemPart.Width,0),
                                    ItemPart.Material,ItemPart.Area.ToString(), ItemPart.Qty.ToString() };
                                row = ws.CreateRow(i);
                                for (int k = 0; k < strItemContents.Length; k++) row.CreateCell(k).SetCellValue(strItemContents[k]);//写抬头内容
                                i++;
                            }
                        }
                    }
                    wb.Write(ms);
                    ms.Flush();
                    ms.Position = 0;
                    string excelPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase) + ".xls";
                    helper.SaveToFile(ms, excelPath);
                    helper.OpenFile(excelPath);
                }
            }

        }

        /// <summary>
        /// 将数据填入本图形
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (this.AllParts.Count > 0)
            {
                //提起本图纸的图框
                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database acDb = acDoc.Database;
                Editor acEd = acDoc.Editor;
                List<Polyline> listAllTitleBlocks = new List<Polyline>();
                using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    using (Transaction trans = acDb.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)trans.GetObject(acDb.BlockTableId, OpenMode.ForWrite);
                        //打开数据库的模型空间块表记录对象
                        BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        //循环遍历模型空间中的实体
                        foreach (ObjectId id in btr)
                        {
                            Entity ent = trans.GetObject(id, OpenMode.ForWrite) as Entity;
                            if (ent is Polyline && ent.ColorIndex.ToString() == "171") listAllTitleBlocks.Add(ent as Polyline);
                        }
                        //提取单元件名称，编号，数量
                        PromptSelectionResult psr = null;
                        if (listAllTitleBlocks.Count > 0)
                        {
                            foreach (var item in listAllTitleBlocks)
                            {
                                var pnts = item.GeometricExtents;
                                psr = acEd.SelectWindow(pnts.MaxPoint, pnts.MinPoint, new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, RXClass.GetClass(typeof(DBText)).DxfName),
                            new TypedValue((int)DxfCode.Color, 171) ,new TypedValue((int) DxfCode.LayerName,"A3")}));
                                List<DBText> listText = new List<DBText>();
                                if (psr.Status == PromptStatus.OK) foreach (SelectedObject txtid in psr.Value) listText.Add(trans.GetObject(txtid.ObjectId, OpenMode.ForWrite) as DBText);
                                List<myPart> partBelongCurTb = new List<myPart>();
                                if (listText.Count > 0)
                                {
                                    string 单元件名称 = "";
                                    foreach (var partDwg in this.AllParts)
                                    {
                                        foreach (var titleBlk in partDwg.CurDwgPanels)
                                        {
                                            var t = listText.Where(c => c.TextString == titleBlk.单元件编号).ToList();
                                            if (t.Count == 1)
                                            {
                                                partBelongCurTb.AddRange(titleBlk.listMyParts);
                                                单元件名称 = t[0].TextString;
                                            }
                                        }
                                    }
                                    if (partBelongCurTb.Count > 0)
                                    {
                                        using (Transaction trans1 = acDb.TransactionManager.StartTransaction())
                                        {
                                            var tbScale = (int)(pnts.MaxPoint.Y - pnts.MinPoint.Y) / 297;
                                            var oid = TableTools.AddTableStyle(acDb, "myTableStyle", tbScale * 2);
                                            Table table = new Table();
                                            table.TableStyle = oid;
                                            table.NumColumns = 9;
                                            table.NumRows = partBelongCurTb.Count + 3;
                                            table.Position = new Point3d(pnts.MaxPoint.X - tbScale * 3 * 8 * 9, pnts.MinPoint.Y + tbScale * 3 * (table.NumRows+6), 0);
                                            table.SetRowHeight(tbScale * 3);
                                            table.SetColumnWidth(tbScale * 3 * 7.5);
                                            table.SetTextString(0, 0, 单元件名称 + " 单元件配套表（单件)");
                                            string[] arr = new string[] { "序号", " 零件名称", "零件编号", "材质", "零件规格（mm）", "数量", "单重（Kg）", "总重（Kg）", "备注" };
                                            for (int i = 0; i < arr.Length; i++) table.SetTextString(1, i, arr[i]);
                                            //Z-DWFB01 单元件配套表（单件）

                                            for (int i = 2; i < partBelongCurTb.Count+2; i++)
                                            {
                                                var part = partBelongCurTb[i-2];
                                                table.SetTextString(i, 0, (i -1).ToString());
                                                table.SetTextString(i, 1, "N.A.");
                                                table.SetTextString(i, 2, part.PieceMark);
                                                table.SetTextString(i, 3, part.Material);
                                                table.SetTextString(i, 4, part.Thk + "x" + (int)part.Width + "x" + (int)part.Length);
                                                table.SetTextString(i, 5, part.Qty.ToString());
                                                table.SetTextString(i, 6, (part.Area * part.Thk * 7.85 / 1000).ToString("0"));
                                                table.SetTextString(i, 7, (part.Qty * part.Area * part.Thk * 7.85 / 1000).ToString("0"));
                                                table.SetTextString(i, 8, "N.A.");
                                            }
                                            btr.AppendEntity(table);
                                            trans1.AddNewlyCreatedDBObject(table, true);
                                            trans1.Commit();
                                        }
                                    }
                                    else
                                    {
                                        Circle cir = new Circle();
                                        cir.Center = new Point3d(0.5 * (pnts.MaxPoint.X + pnts.MinPoint.X), 0.5 * (pnts.MaxPoint.Y + pnts.MinPoint.Y), 0.5 * (pnts.MaxPoint.Z + pnts.MinPoint.Z));
                                        cir.Radius = Math.Sqrt((pnts.MaxPoint.X - pnts.MinPoint.X) * (pnts.MaxPoint.X - pnts.MinPoint.X) + (pnts.MaxPoint.Y - pnts.MinPoint.Y) * (pnts.MaxPoint.Y - pnts.MinPoint.Y)) / 2;
                                        cir.ColorIndex = 2;
                                        btr.AppendEntity(cir);
                                        trans.AddNewlyCreatedDBObject(cir, true);
                                    }
                                }
                            }
                            trans.Commit();
                            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("自动生成零件汇总表格完成！！！！");
                        }
                        else Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请给图框外轮廓设置成A3图层，颜色171 号颜色！");
                    }
                }
            }
            else Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请先执行第二步操作，谢谢！！");
        }
    }

}

