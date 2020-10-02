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
using Microsoft.Win32;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;

namespace WLQ_CMD
{


    public class MainClass : IExtensionApplication
    {
        public void Initialize()
        {
            helper.Register2HKCR();
            string path = Assembly.GetExecutingAssembly().Location;
            helper.AddCmdtoMenuBar(helper.GetDllCmds(path));
        }
        public void Terminate()
        {
            throw new NotImplementedException();
        }
    }
    public static class helper


    {/// <summary>
     ///提取所有的命令
     /// </summary>
     /// <param name="dllFiles">dll的路径</param>
     /// <returns></returns>
        public static List<gcadDllcmd> GetDllCmds(params string[] dllFiles)
        {
            List<gcadDllcmd> res = new List<gcadDllcmd>();
            List<gcadCmds> cmds = new List<gcadCmds>();
            #region 提取所以的命令
            for (int i = 0; i < dllFiles.Length; i++)
            {
                Assembly ass = Assembly.LoadFile(dllFiles[i]);//反射加载dll程序集
                var clsCollection = ass.GetTypes().Where(t => t.IsClass && t.IsPublic).ToList();
                if (clsCollection.Count > 0)
                {
                    foreach (var cls in clsCollection)
                    {
                        var methods = cls.GetMethods().Where(m => m.IsPublic && m.GetCustomAttributes(true).Length > 0).ToList();
                        if (methods.Count > 0)
                        {
                            foreach (MethodInfo mi in methods)
                            {
                                var atts = mi.GetCustomAttributes(true).Where(c => c is CommandMethodAttribute).ToList();
                                if (atts.Count == 1)
                                {
                                    gcadCmds cmd = new gcadCmds(cls.Name, mi.Name, (atts[0] as CommandMethodAttribute).GlobalName, ass.ManifestModule.Name.Substring(0, ass.ManifestModule.Name.Length - 4));
                                    cmds.Add(cmd);
                                }
                            }
                        }
                    }
                }

            }
            #endregion
            if (cmds.Count > 0)
            {
                List<string> dllName = new List<string>();
                foreach (var item in cmds)
                {
                    if (!dllName.Contains(item.dllName)) dllName.Add(item.dllName);
                }
                foreach (var item in dllName) res.Add(new gcadDllcmd(item, cmds));
            }
            return res;
            //
        }
        public static void AddCmdtoMenuBar(List<gcadDllcmd> cmds)
        {
            var dllName = Assembly.GetExecutingAssembly().ManifestModule.Name.Substring(0, Assembly.GetExecutingAssembly().ManifestModule.Name.Length - 4);
            var gcadApp = Application.AcadApplication as AcadApplication;
            AcadMenuGroup mg = null;

            for (int i = 0; i < gcadApp.MenuGroups.Count; i++) if (gcadApp.MenuGroups.Item(i).Name == "ACAD") mg = gcadApp.MenuGroups.Item(i);

            for (int i = 0; i < mg.Menus.Count; i++) if (mg.Menus.Item(i).Name == dllName) mg.Menus.Item(i).RemoveFromMenuBar();
            AcadPopupMenu popMenu = mg.Menus.Add(dllName);
            for (int i = 0; i < cmds.Count; i++)
            {
                var dllPopMenu = popMenu.AddSubMenu(popMenu.Count + 1, cmds[i].DllName);
                for (int j = 0; j < cmds[i].clsCmds.Count; j++)
                {
                    var clsPopMenu = dllPopMenu.AddSubMenu(dllPopMenu.Count + 1, cmds[i].clsCmds[j].clsName);
                    for (int k = 0; k < cmds[i].clsCmds[j].curClscmds.Count; k++)
                    {
                        var methodPopMenu = clsPopMenu.AddMenuItem(clsPopMenu.Count + 1, cmds[i].clsCmds[j].curClscmds[k].cmdName, cmds[i].clsCmds[j].curClscmds[k].cmdMacro + " ");
                    }
                }
            }
            popMenu.InsertInMenuBar(mg.Menus.Count + 1);
        }

        /// <summary>
        /// 将菜单加载到AutoCAD
        /// </summary>
        public static void Register2HKCR()
        {
            string hkcrKey = HostApplicationServices.Current.RegistryProductRootKey;
            var assName = Assembly.GetExecutingAssembly().CodeBase;
            var apps_Acad = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(Path.Combine(hkcrKey, "Applications"));
            if (apps_Acad.GetSubKeyNames().Count(c => c == Path.GetFileNameWithoutExtension(assName)) == 0)
            {
                var myNetLoader = apps_Acad.CreateSubKey(Path.GetFileNameWithoutExtension(assName), RegistryKeyPermissionCheck.Default);
                myNetLoader.SetValue("DESCRIPTION", "加载自定义dll文件", Microsoft.Win32.RegistryValueKind.String);
                myNetLoader.SetValue("LOADCTRLS", 2, Microsoft.Win32.RegistryValueKind.DWord);
                myNetLoader.SetValue("LOADER", assName, Microsoft.Win32.RegistryValueKind.String);
                myNetLoader.SetValue("MANAGED", 1, Microsoft.Win32.RegistryValueKind.DWord);
                Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "程序自动加载完成，重启AutoCAD 生效！");
            }
            else Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "程序已经启动！欢迎使用");

        }

        public static bool CheckFileReadOnly(string fileName)
        {
            bool inUse = true;
            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None);
                inUse = false;
            }
            catch { }
            return inUse;//true表示正在使用,false没有使用  
        }

        public static List<double> GetMaxMinPoint(Polyline pline, Database db)
        {
            List<Point2d> lsitPnt = new List<Point2d>();
            List<double> resPnt = new List<double>();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < pline.NumberOfVertices; i++) lsitPnt.Add(pline.GetPoint2dAt(i));
                var Xmax = lsitPnt.Max(c => c.X);
                var Ymax = lsitPnt.Max(c => c.Y);
                var Xmin = lsitPnt.Min(c => c.X);
                var Ymin = lsitPnt.Min(c => c.Y);
                resPnt.AddRange(new double[] { Xmax, Ymax, Xmin, Ymin });
            }
            return resPnt;
        }

        public static void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();

                fs.Write(data, 0, data.Length);
                fs.Flush();

                data = null;
            }
        }
        public static void OpenFile(string fileName)
        {
            System.Diagnostics.Process.Start(fileName);
            System.Diagnostics.Process pro = new System.Diagnostics.Process();
            pro.EnableRaisingEvents = false;
            pro.StartInfo.FileName = "rundll32.exe";
            pro.StartInfo.Arguments = "shell32,OpenAs_RunDLL" + fileName;
            pro.Start();
        }

        public static ObjectId AddTableStyle(string style)
        {
            ObjectId styleId; // 存储表格样式的Id
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 打开表格样式字典
                DBDictionary dict = (DBDictionary)db.TableStyleDictionaryId.GetObject(OpenMode.ForRead);
                if (dict.Contains(style)) // 如果存在指定的表格样式
                    styleId = dict.GetAt(style); // 获取表格样式的Id
                else
                {
                    TableStyle ts = new TableStyle(); // 新建一个表格样式
                    // 设置表格的标题行为灰色
                    ts.SetBackgroundColor(Color.FromColorIndex(ColorMethod.ByAci, 8), (int)RowType.TitleRow);
                    // 设置表格所有行的外边框的线宽为0.30mm
                    ts.SetGridLineWeight(LineWeight.LineWeight030, (int)GridLineType.OuterGridLines, TableTools.AllRows);
                    // 不加粗表格表头行的底部边框
                    ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalBottom, (int)RowType.HeaderRow);
                    // 不加粗表格数据行的顶部边框
                    ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalTop, (int)RowType.DataRow);
                    // 设置表格中所有行的文本高度为1
                    ts.SetTextHeight(1, TableTools.AllRows);
                    // 设置表格中所有行的对齐方式为正中
                    ts.SetAlignment(CellAlignment.MiddleCenter, TableTools.AllRows);
                    dict.UpgradeOpen();//切换表格样式字典为写的状态
                    // 将新的表格样式添加到样式字典并获取其Id
                    styleId = dict.SetAt(style, ts);
                    // 将新建的表格样式添加到事务处理中
                    trans.AddNewlyCreatedDBObject(ts, true);
                    trans.Commit();
                }
            }
            return styleId; // 返回表格样式的Id
        }
    }
    /// <summary>
    /// 储存自定义的cad命令的信息的类
    /// </summary>
    public class gcadCmds
    {
        public string clsName { get; set; }
        public string cmdName { get; set; }
        public string cmdMacro { get; set; }
        public string dllName { get; set; }

        public gcadCmds(string _clsName, string _cmdName, string _macro, string _dllName)
        {
            this.dllName = _dllName;
            this.clsName = _clsName;
            this.cmdMacro = _macro;
            this.cmdName = _cmdName;
        }

    }
    /// <summary>
    /// 储存包含自定命令的类
    /// </summary>
    public class gcadClscmd
    {
        public string clsName { get; set; }

        public string dllName { get; set; }

        public bool HasGcadcmds { get; set; }

        public List<gcadCmds> curClscmds { get; set; }

        public gcadClscmd(string _clsName, List<gcadCmds> cmds)
        {
            this.clsName = _clsName;
            this.dllName = cmds.First().dllName;
            var clsCmds = cmds.Where(c => c.clsName == this.clsName).ToList();
            if (clsCmds.Count > 0)
            {
                this.HasGcadcmds = true;
                this.curClscmds = new List<gcadCmds>();
                foreach (var item in clsCmds)
                {
                    if (item.clsName == this.clsName) this.curClscmds.Add(item);
                }

            }
            else this.HasGcadcmds = false;
        }
    }
    /// <summary>
    /// 储存每个dll类的
    /// </summary>
    public class gcadDllcmd
    {
        public string DllName { get; set; }
        public bool HasGcadcls { get; set; }
        public List<gcadClscmd> clsCmds { get; set; }
        public List<gcadCmds> curDllcmds { get; set; }
        public gcadDllcmd(string _dllname, List<gcadCmds> cmds)
        {
            this.DllName = _dllname;
            var curDllcmds = cmds.Where(c => c.dllName == this.DllName).ToList();
            if (curDllcmds.Count > 0)
            {
                this.HasGcadcls = true;
                this.curDllcmds = curDllcmds;
                List<string> listClsName = new List<string>();
                foreach (gcadCmds item in this.curDllcmds)
                {
                    if (!listClsName.Contains(item.clsName)) listClsName.Add(item.clsName);
                }
                this.clsCmds = new List<gcadClscmd>();
                foreach (var item in listClsName)
                {
                    gcadClscmd clsCmds = new gcadClscmd(item, this.curDllcmds.Where(c => c.clsName == item).ToList());
                    this.clsCmds.Add(clsCmds);
                }


            }
            else this.HasGcadcls = false;
        }


    }

    /// <summary>
    /// 表格操作类
    /// </summary>
    public static class TableTools
    {
        /// <summary>
        /// 所有行的标志位（包括标题行、数据行）
        /// </summary>
        public static int AllRows
        {
            get
            {
                return (int)(RowType.DataRow | RowType.HeaderRow | RowType.TitleRow);
            }
        }

        /// <summary>
        /// 创建表格
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="position">表格位置</param>
        /// <param name="numRows">表格行数</param>
        /// <param name="numCols">表格列数</param>
        /// <returns>返回创建的表格的Id</returns>
        public static ObjectId CreateTable(this Database db, Point3d position, int numRows, int numCols)
        {
            ObjectId tableId;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                Table table = new Table();
                //设置表格的行数和列数
                table.SetSize(numRows, numCols);
                //设置表格放置的位置
                table.Position = position;
                //非常重要，根据当前样式更新表格，不加此句，会导致AutoCAD崩溃
                table.GenerateLayout();
                //表格添加到模型空间
                tableId = db.AddToModelSpace(table);
                trans.Commit();
            }
            return tableId;
        }

        /// <summary>
        /// 设置单元格中文本的高度
        /// </summary>
        /// <param name="table">表格对象</param>
        /// <param name="height">文本高度</param>
        /// <param name="rowType">行的标志位</param>
        public static void SetTextHeight(this Table table, double height, RowType rowType)
        {
            table.SetTextHeight(height, rowType);
        }

        /// <summary>
        /// 设置表格中所有单元格中文本为同一高度
        /// </summary>
        /// <param name="table">表格对象</param>
        /// <param name="height">文本高度</param>
        public static void SetTextHeight(this Table table, double height)
        {
            table.SetTextHeight(height, AllRows);
        }

        /// <summary>
        /// 设置单元格中文本的对齐方式
        /// </summary>
        /// <param name="table">表格对象</param>
        /// <param name="align">单元格对齐方式</param>
        /// <param name="rowType">行的标志位</param>
        public static void SetAlignment(this Table table, CellAlignment align, RowType rowType)
        {
            table.SetAlignment(align, (int)rowType);
        }

        /// <summary>
        /// 设置表格中所有单元格中文本为同一对齐方式
        /// </summary>
        /// <param name="table">表格对象</param>
        /// <param name="align">单元格对齐方式</param>
        public static void SetAlignment(this Table table, CellAlignment align)
        {
            table.SetAlignment(align, AllRows);
        }

        /// <summary>
        /// 一次性按行设置单元格文本
        /// </summary>
        /// <param name="table">表格对象</param>
        /// <param name="rowIndex">行号</param>
        /// <param name="data">文本内容</param>
        /// <returns>如果设置成功，则返回true，否则返回false</returns>
        public static bool SetRowTextString(this Table table, int rowIndex, params string[] data)
        {
            if (data.Length > table.NumColumns) return false;
            for (int j = 0; j < data.Length; j++)
            {
                table.SetTextString(rowIndex, j, data[j]);
            }
            return true;
        }

        /// <summary>
        /// 为图形添加一个新的表格样式
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="styleName">表格样式的名称</param>
        /// <returns>返回表格样式的Id</returns>
        public static ObjectId AddTableStyle(this Database db, string styleName,double txtHeight)
        {
            ObjectId styleId;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开表格样式字典
                DBDictionary dict = (DBDictionary)trans.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead);
                //判断是否存在指定的表格样式
                if (dict.Contains(styleName))
                    styleId = dict.GetAt(styleName);//如果存在则返回表格样式的Id
                else
                {
                    //新建一个表格样式
                    TableStyle style = new TableStyle();
                    style.SetTextHeight(txtHeight, TableTools.AllRows);
                    style.SetColor(Color.FromColorIndex(ColorMethod.ByAci, 6), TableTools.AllRows);
                    style.SetGridColor(Color.FromColorIndex(ColorMethod.ByAci, 1), (int)GridLineType.AllGridLines, TableTools.AllRows);
                    dict.UpgradeOpen();//切换表格样式字典为写的状态
                                       //将新的表格样式添加到样式字典并获取其 Id
                    styleId = dict.SetAt(styleName, style);
                    //将新建的表格样式添加到事务处理中
                    trans.AddNewlyCreatedDBObject(style, true);
                    trans.Commit();
                }
            }
            return styleId;//返回表格样式的Id
        }
        /// <summary>
        /// 将实体添加到模型空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到模型空间中的实体ObjectId</returns>
        public static ObjectId AddToModelSpace(this Database db, Entity ent)
        {
            ObjectId entId;//用于返回添加到模型空间中的实体ObjectId
                           //定义一个指向当前数据库的事务处理，以添加直线
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //以读方式打开块表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                //以写方式打开模型空间块表记录.
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                entId = btr.AppendEntity(ent);//将图形对象的信息添加到块表记录中
                trans.AddNewlyCreatedDBObject(ent, true);//把对象添加到事务处理中
                trans.Commit();//提交事务处理
            }
            return entId; //返回实体的ObjectId
        }

        /// <summary>
        /// 将实体添加到模型空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ents">要添加的多个实体</param>
        /// <returns>返回添加到模型空间中的实体ObjectId集合</returns>
        public static ObjectIdCollection AddToModelSpace(this Database db, params Entity[] ents)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
            foreach (var ent in ents)
            {
                ids.Add(btr.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            btr.DowngradeOpen();
            return ids;
        }

        /// <summary>
        /// 将实体添加到图纸空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到图纸空间中的实体ObjectId</returns>
        public static ObjectId AddToPaperSpace(this Database db, Entity ent)
        {
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockPaperSpaceId(db), OpenMode.ForWrite);
            ObjectId id = btr.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            btr.DowngradeOpen();
            return id;
        }
    }
}
