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


namespace WLQ_CMD
{
   public class 导出零件表
    {
        [CommandMethod("WEOP")]
        public void 零件图_单元件()
        {
            //frmExportOutPartList frm = new frmExportOutPartList();
            UserControl1 uc1 = new UserControl1();
            Application.ShowModelessDialog(uc1);
        }
    }

    
}
