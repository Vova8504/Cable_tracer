using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RASH_DAN
{
#if NANOCAD
    using Teigha.DatabaseServices;
    using Teigha.Runtime;
    using Teigha.Geometry;
    using HostMgd.ApplicationServices;
    using HostMgd.EditorInput;
#elif AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.EditorInput;
#endif

    public class CommandMethods
    {
        public struct TPodk
        {
            public string IND;
            public double VisT;
            public Point3d Koord;
            public string KoordMod;
            public int Vstr;
            public string Sist;
            public string BlNOD;
            public void NIND(string i) { IND = i; }
            public void NSist(string i) { Sist = i; }
            public void NBlNOD(string i) { BlNOD = i; }
            public void NVisT(double i) { VisT = i; }
            public void NKoor(Point3d i) { Koord = i; }
            public void NKoorMod(string i) { KoordMod = i; }
            public void NVstr(int i) { Vstr = i; }
        }
        public struct Noda
        {
            public string Nom, Otkuda, Hoz;
            public string SpSmNod;
            public double Ves;
            public double DlinP;
            public double VisT;
            public double Param;
            public Point3d Koord;
            public Point3d Koord3D;
            public void NomNod(string i) { Nom = i; }
            public void NomNOtk(string i) { Otkuda = i; }
            public void NomHoz(string i) { Hoz = i; }
            public void NSpSmNod(string i) { SpSmNod = i; }
            public void NVes(double i) { Ves = i; }
            public void NDlinP(double i) { DlinP = i; }
            public void NVisT(double i) { VisT = i; }
            public void NParam(double i) { Param = i; }
            public void NKoor(Point3d i) { Koord = i; }
            public void NKoor3D(Point3d i) { Koord3D = i; }
        }
        public struct PLOS
        {
            public string Vid, Mas, Spravka, List, Osi;
            public Point3d Psk, Msk, max, min;
            public void NOsi(string i) { Osi = i; }
            public void NVid(string i) { Vid = i; }
            public void NMas(string i) { Mas = i; }
            public void NSpravka(string i) { Spravka = i; }
            public void NList(string i) { List = i; }
            public void NPsk(Point3d i) { Psk = i; }
            public void NMsk(Point3d i) { Msk = i; }
            public void Nmax(Point3d i) { max = i; }
            public void Nmin(Point3d i) { min = i; }
        };

        [CommandMethod("SV_SP_DB")]
        static public void RemoveXdata()
        {
            Form1 form1 = new Form1();
            form1.Show();
        }//Чтение расширеных данных выбранного примитива 
        [CommandMethod("OBN")]
        static public void SoedinUZL()
        {
            Form2 form2 = new Form2();
            form2.Show();
        }//Чтение расширеных данных выбранного примитива 
        [CommandMethod("OK")]
        static public void OdinKab()
        {
            Form3 form2 = new Form3();
            form2.Show();
        }//трасировка отдельного кабеля
        [CommandMethod("TP")]
        static public void OBORUD()
        {
            Form5 form2 = new Form5();
            form2.Show();

        }//трасировка отдельного кабеля
        [CommandMethod("UT")]
        static public void BigNOD()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку");
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d Toch = pPtRes.Value;
                Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
                PromptSelectionResult acSSPrompt;
                TypedValue[] acTypValAr = new TypedValue[5];
                acTypValAr.SetValue(new TypedValue(0, "CIRCLE"), 0);
                acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 1);
                acTypValAr.SetValue(new TypedValue(-4, "="), 2);
                acTypValAr.SetValue(new TypedValue(40, 5.0), 3);
                acTypValAr.SetValue(new TypedValue(62, 6.0), 4);
                // создаем фильтр
                SelectionFilter filter = new SelectionFilter(acTypValAr);
                acSSPrompt = acDocEd.SelectCrossingWindow(new Point3d(Toch.X + 50, Toch.Y + 50, 0),
                                               new Point3d(Toch.X - 50, Toch.Y - 50, 0), filter);
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;
                    Application.ShowAlertDialog(acSSet.Count.ToString());
                }
            }
        }//
        [CommandMethod("HS")]
        static public void SVpoH()
        {
            string strNAME = "-";
            string strSVOI = "-";
            string strNeSVI = "-";
            string strDlinK = "-";
            string strHOZ = "-";
            string strVIS = "-";
            string strDlin = "-";
            int Schet = 0;
            string Handl = "2B87E";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            int sist10 = Convert.ToInt32(Handl, 16);
            Handle H = new Handle(sist10);
            using (tr)
            {
                //ObjectId id = db.GetObjectId(false, H, -1);
                //Circle Uzel = tr.GetObject(id, OpenMode.ForWrite) as Circle;
                ObjectId ID = db.GetObjectId(false, H, -1);
                if (ID.IsErased == true) { Application.ShowAlertDialog("Удален"); return; }
                Entity ent = (Entity)tr.GetObject(ID, OpenMode.ForWrite);
                ResultBuffer buffer = ent.GetXDataForApplication("LAUNCH01");
                // Если есть расширенные данные – удалим их.
                // Для этого в качестве расширенных данных
                // передаём только имя приложения.
                // Только связанные с ним данные будут удалены.
                //string strDan="";
                if (buffer != null)
                {
                    foreach (TypedValue value in buffer)
                    {
                        if (Schet == 1) { strNAME = value.Value.ToString(); }
                        if (Schet == 2) { strVIS = value.Value.ToString(); }
                        if (Schet == 3) { strSVOI = value.Value.ToString(); }
                        if (Schet == 4) { strDlinK = value.Value.ToString(); }
                        if (Schet == 7) { strNeSVI = value.Value.ToString(); }
                        if (Schet == 8) { strHOZ = value.Value.ToString(); }
                        if (Schet == 9) { strDlin = value.Value.ToString(); }
                        Schet = Schet + 1;
                    }
                    Application.ShowAlertDialog(strNAME + strVIS + strSVOI + strDlinK + strNeSVI + strHOZ + strDlin);
                }
            }
        }
        [CommandMethod("UK")]
        static public void UdalKab()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue(8, "Kabeli"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
                    foreach (ObjectId sobj in acSSet.GetObjectIds())
                    {
                        Polyline ln = tr.GetObject(sobj, OpenMode.ForWrite) as Polyline;
                        if (sobj != null)
                        {
                            //Application.ShowAlertDialog(ln.Handle.ToString());
                            ln.Erase();
                        }
                    }
                    tr.Commit();
                }
            }
        }
        [CommandMethod("ZM")]
        static public void ZoomLimits()
        {
            Zoom(new Point3d(), new Point3d(), new Point3d(), 1.01075);
        }
        [CommandMethod("MAR_DL_NN")]
        public void MAR_DL_NN()
        {
            List<string> Vkrazv = new List<string>();
            SozdSpKV(ref Vkrazv);   
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\МАРШРУТ\КабелиРазв.txt", false, Encoding.GetEncoding("Windows-1251"))) { foreach (string line in Vkrazv) { file.WriteLine(line); } }
        }
        [CommandMethod("STO")]
        static public void vSTO()
        {
            Form6 form2 = new Form6();
            form2.Show();
        }//трасировка отдельного кабеля
        [CommandMethod("PlehcKab")]
        static public void PlehiNMCab() 
        {
            List<PLOS> SpPLOS = new List<PLOS>();
            List<string> Vkrazv = new List<string>();
            SBOR_PLOS(ref SpPLOS);
            SozdSpKVKoor(ref Vkrazv, SpPLOS);
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\МАРШРУТ\КАБЕЛИ_КООРДИНАТЫ.txt", false, Encoding.GetEncoding("Windows-1251"))) { foreach (string line in Vkrazv) { file.WriteLine(line); } }
        }
        [CommandMethod("PlehcOBor")]
        static public void PlehiNMObor()
        {
            List<PLOS> SpPLOS = new List<PLOS>();
            List<string> Vkrazv = new List<string>();
            SBOR_PLOS(ref SpPLOS);
            SpVistTP(ref Vkrazv, SpPLOS);
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\МАРШРУТ\ТОЧКИ_КООРДИНАТЫ.txt", false, Encoding.GetEncoding("Windows-1251"))) { foreach (string line in Vkrazv) { file.WriteLine(line); } }
        }
        [CommandMethod("SKRUGL")]
        //static public void FillF() 
        //{
        //    List<ObjectId> ObjId = new List<ObjectId>();
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;
        //    Point3d T1 = new Point3d(0, 0, 0);
        //    Point3d T2 = new Point3d(0, 1000, 0);
        //    Point3d T3 = new Point3d(1000, 0, 0);
        //    LINE(T1, T2, 3, ref ObjId);
        //    LINE(T1, T3, 3, ref ObjId);
        //    doc.Database.Filletrad = 10.0;
        //    ed.Command("_.fillet", ObjId[0], ObjId[1]);
        //}

        [CommandMethod("EditSpline")]
        public static void EditSpline()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                // Create a Point3d Collection
                Point3dCollection acPt3dColl = new Point3dCollection();
                acPt3dColl.Add(new Point3d(1, 1, 0));
                acPt3dColl.Add(new Point3d(5, 5, 0));
                acPt3dColl.Add(new Point3d(10, 0, 0));

                // Set the start and end tangency
                Vector3d acStartTan = new Vector3d(0.5, 0.5, 0);
                Vector3d acEndTan = new Vector3d(0.5, 0.5, 0);

                // Create a spline
                Spline acSpline = new Spline(acPt3dColl,
                                             acStartTan,
                                             acEndTan, 4, 0);

                acSpline.SetDatabaseDefaults();

                // Set a control point
                acSpline.SetControlPointAt(0, new Point3d(0, 3, 0));

                // Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline);
                acTrans.AddNewlyCreatedDBObject(acSpline, true);

                // Save the new objects to the database
                acTrans.Commit();
            }
        }



        public void SweepAlongPath()
        {
            Document doc =Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            // Ask the user to select a region to extrude
            PromptEntityOptions peo1 =
              new PromptEntityOptions( "\nSelect profile or curve to sweep: ");
            peo1.SetRejectMessage("\nEntity must be a region, curve or planar surface.");
            peo1.AddAllowedClass( typeof(Region), false);
            peo1.AddAllowedClass( typeof(Curve), false);
            peo1.AddAllowedClass(typeof(PlaneSurface), false);
            PromptEntityResult per = ed.GetEntity(peo1);
            if (per.Status != PromptStatus.OK)
                return;
            ObjectId regId = per.ObjectId;
            // Ask the user to select an extrusion path
            PromptEntityOptions peo2 =new PromptEntityOptions("\nSelect path along which to sweep: ");
            peo2.SetRejectMessage("\nEntity must be a curve.");
            peo2.AddAllowedClass( typeof(Curve), false);
            per = ed.GetEntity(peo2);
            if (per.Status != PromptStatus.OK)
                return;
            ObjectId splId = per.ObjectId;
            PromptKeywordOptions pko = new PromptKeywordOptions("\nSweep a solid or a surface?");
            pko.AllowNone = true;
            pko.Keywords.Add("SOlid");
            pko.Keywords.Add("SUrface");
            pko.Keywords.Default = "SOlid";
            PromptResult pkr = ed.GetKeywords(pko);
            bool createSolid = (pkr.StringResult == "SOlid");
            if (pkr.Status != PromptStatus.OK)
                return;
            // Now let's create our swept surface
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                try
                {
                    Entity sweepEnt = tr.GetObject(regId, OpenMode.ForRead) as Entity;
                    Curve pathEnt =tr.GetObject(splId, OpenMode.ForRead) as Curve;
                    if (sweepEnt == null || pathEnt == null)
                    {
                        ed.WriteMessage("\nProblem opening the selected entities." );
                        return;
                    }
                    // We use a builder object to create
                    // our SweepOptions
                    SweepOptionsBuilder sob =new SweepOptionsBuilder();
                    // Align the entity to sweep to the path
                    sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath;
                    // The base point is the start of the path
                    sob.BasePoint = pathEnt.StartPoint;
                    // The profile will rotate to follow the path
                    sob.Bank = true;
                    // Now generate the solid or surface...
                    Entity ent;
                    if (createSolid)
                    {
                        Solid3d sol = new Solid3d();
                        sol.CreateSweptSolid( sweepEnt, pathEnt, sob.ToSweepOptions());
                        ent = sol;
                    }
                    else
                    {
                        SweptSurface ss = new SweptSurface();
                        ss.CreateSweptSurface( sweepEnt,pathEnt,sob.ToSweepOptions());
                        ent = ss;
                    }

                    // ... and add it to the modelspace

                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId,OpenMode.ForRead);
                    BlockTableRecord ms =(BlockTableRecord)tr.GetObject( bt[BlockTableRecord.ModelSpace],OpenMode.ForWrite);
                    ms.AppendEntity(ent);
                    tr.AddNewlyCreatedDBObject(ent, true);
                    tr.Commit();
                }
                catch
                { }
            }
        }

        [CommandMethod("Sol")]
        public void SweepAlongPath1()
        {
            Document doc =Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Ask the user to select a region to extrude

            PromptEntityOptions peo1 =
              new PromptEntityOptions( "\nSelect profile or curve to sweep: ");

            peo1.SetRejectMessage( "\nEntity must be a region, curve or planar surface.");
            peo1.AddAllowedClass( typeof(Region), false);
            peo1.AddAllowedClass(typeof(Curve), false);
            peo1.AddAllowedClass( typeof(PlaneSurface), false);
            PromptEntityResult per = ed.GetEntity(peo1);

            if (per.Status != PromptStatus.OK)
                return;
            ObjectId regId = per.ObjectId;
            // Ask the user to select an extrusion path
            PromptEntityOptions peo2 =new PromptEntityOptions("\nSelect path along which to sweep: " );
            peo2.SetRejectMessage("\nEntity must be a curve." );
            peo2.AddAllowedClass(typeof(Curve), false);
            per = ed.GetEntity(peo2);
            if (per.Status != PromptStatus.OK)
                return;
            ObjectId splId = per.ObjectId;
            // Now let's create our swept surface
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                try
                {
                    Entity sweepEnt =tr.GetObject(regId, OpenMode.ForRead) as Entity;
                    Curve pathEnt = tr.GetObject(splId, OpenMode.ForRead) as Curve;
                    if (sweepEnt == null || pathEnt == null)
                    {
                        ed.WriteMessage( "\nProblem opening the selected entities." );
                        return;
                    }
                    // We use a builder object to create
                    // our SweepOptions
                    SweepOptionsBuilder sob = new SweepOptionsBuilder();
                    // Align the entity to sweep to the path
                    sob.Align =SweepOptionsAlignOption.AlignSweepEntityToPath;
                    // The base point is the start of the path
                    sob.BasePoint = pathEnt.StartPoint;
                    // The profile will rotate to follow the path
                    sob.Bank = true;
                    // Now generate the surface...
                    SweptSurface ss = new SweptSurface();
                    ss.CreateSweptSurface( sweepEnt, pathEnt, sob.ToSweepOptions() );
                    // ... and add it to the modelspace
                    BlockTable bt =(BlockTable)tr.GetObject( db.BlockTableId,OpenMode.ForRead );
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    ms.AppendEntity(ss);
                    tr.AddNewlyCreatedDBObject(ss, true);
                    tr.Commit();
                }
                catch
                { }
            }
        }

        //public void Fil()
        //{
        //    List<ObjectId> ObjId = new List<ObjectId>();
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;
        //    using (DocumentLock docLock = doc.LockDocument())
        //    {
        //        Point3d T1 = new Point3d(0, 0, 0);
        //        Point3d T2 = new Point3d(0, 1000, 0);
        //        Point3d T3 = new Point3d(1000, 0, 0);
        //        LINE(T1, T2, 3, ref ObjId);
        //        LINE(T1, T3, 3, ref ObjId);
        //        doc.Database.Filletrad = 10.0;
        //        ed.Command("_.fillet", ObjId[0], ObjId[1]);
        //    }
        //}

        static public void LINE(Point3d Toch1, Point3d Toch2, int Zvet,ref List<ObjectId> ObjId)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            Editor ed = doc.Editor;
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Line KrugPOZVn = new Line(Toch1, Toch2);
                KrugPOZVn.SetDatabaseDefaults();
                KrugPOZVn.ColorIndex = Zvet;
                btr.AppendEntity(KrugPOZVn);
                tr1.AddNewlyCreatedDBObject(KrugPOZVn, true);
                tr1.Commit();
            }
            PromptSelectionResult acSSPrompt = ed.SelectLast();
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    ObjId.Add(acSSObj.ObjectId);
                    Tx.Commit();
                }
            }
        }

        static public void SBOR_PLOS( ref List<PLOS> SpPLOS)
        {
            string Vid = "", Mash = "", Sprav_plos = "", List = "", Osi = "";
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            Point3d BP = new Point3d();
            //Point3d T1 = new Point3d();
            //Point3d T2 = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(8, "Плоскости"), 1);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей изображенных блоками...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    //для плоскостей
                    Vid = "";
                    Mash = "";
                    Sprav_plos = "";
                    List = "";
                    Osi = "";
                    x1 = 0;
                    y1 = 0;
                    x2 = 0;
                    y2 = 0;
                    xMir = 0;
                    yMir = 0;
                    zMir = 0;
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                    BP = bref.Position;
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            object[] values = prop.GetAllowedValues();
                            if (prop.PropertyName == "Положение4 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение4 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение3 X") { x2 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение3 Y") { y2 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Видимость1") { Osi = prop.Value.ToString(); }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "X") { xMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Y") { yMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "Z") { zMir = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "ВИД") { Vid = atrRef.TextString; }
                                if (atrRef.Tag == "МАСШТАБ") { Mash = atrRef.TextString; }
                                if (atrRef.Tag == "ПРИМЕЧАНИЕ") { Sprav_plos = atrRef.TextString; }
                                if (atrRef.Tag == "ЛИСТ") { List = atrRef.TextString; }
                            }
                        }
                    }
                    if (Vid != "")
                    {
                        Point3d Psk = new Point3d(BP.X + x1, BP.Y + y1, 0);
                        Point3d Msk = new Point3d(xMir, yMir, zMir);
                        PLOS tPOZ = new PLOS();
                        tPOZ.Nmax(bref.GeometricExtents.MaxPoint);
                        tPOZ.Nmin(bref.GeometricExtents.MinPoint);
                        tPOZ.NPsk(Psk);
                        tPOZ.NMsk(Msk);
                        tPOZ.NVid(Vid);
                        tPOZ.NMas(Mash);
                        tPOZ.NList(List);
                        tPOZ.NOsi(Osi);
                        tPOZ.NSpravka(Sprav_plos);
                        SpPLOS.Add(tPOZ);
                    }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных блоками
        static void Zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            int nCurVport = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));

            // Get the extents of the current space no points 
            // or only a center point is provided
            // Check to see if Model space is current
            if (acCurDb.TileMode == true)
            {
                if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                {
                    pMin = acCurDb.Extmin;
                    pMax = acCurDb.Extmax;
                }
            }
            else
            {
                // Check to see if Paper space is current
                if (nCurVport == 1)
                {
                    // Get the extents of Paper space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Pextmin;
                        pMax = acCurDb.Pextmax;
                    }
                }
                else
                {
                    // Get the extents of Model space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Extmin;
                        pMax = acCurDb.Extmax;
                    }
                }
            }

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Get the current view
                using (ViewTableRecord acView = acDoc.Editor.GetCurrentView())
                {
                    Extents3d eExtents;

                    // Translate WCS coordinates to DCS
                    Matrix3d matWCS2DCS;
                    matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                                                   acView.ViewDirection,
                                                   acView.Target) * matWCS2DCS;

                    // If a center point is specified, define the min and max 
                    // point of the extents
                    // for Center and Scale modes
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pMin = new Point3d(pCenter.X - (acView.Width / 2),
                                           pCenter.Y - (acView.Height / 2), 0);

                        pMax = new Point3d((acView.Width / 2) + pCenter.X,
                                           (acView.Height / 2) + pCenter.Y, 0);
                    }

                    // Create an extents object using a line
                    using (Line acLine = new Line(pMin, pMax))
                    {
                        eExtents = new Extents3d(acLine.Bounds.Value.MinPoint,
                                                 acLine.Bounds.Value.MaxPoint);
                    }

                    // Calculate the ratio between the width and height of the current view
                    double dViewRatio;
                    dViewRatio = (acView.Width / acView.Height);

                    // Tranform the extents of the view
                    matWCS2DCS = matWCS2DCS.Inverse();
                    eExtents.TransformBy(matWCS2DCS);

                    double dWidth;
                    double dHeight;
                    Point2d pNewCentPt;

                    // Check to see if a center point was provided (Center and Scale modes)
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        dWidth = acView.Width;
                        dHeight = acView.Height;

                        if (dFactor == 0)
                        {
                            pCenter = pCenter.TransformBy(matWCS2DCS);
                        }

                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                    }
                    else // Working in Window, Extents and Limits mode
                    {
                        // Calculate the new width and height of the current view
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;

                        // Get the center of the view
                        pNewCentPt = new Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5),
                                                 ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5));
                    }

                    // Check to see if the new width fits in current window
                    if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;

                    // Resize and scale the view
                    if (dFactor != 0)
                    {
                        acView.Height = dHeight * dFactor;
                        acView.Width = dWidth * dFactor;
                    }

                    // Set the center of the view
                    acView.CenterPoint = pNewCentPt;

                    // Set the current view
                    acDoc.Editor.SetCurrentView(acView);
                }

                // Commit the changes
                acTrans.Commit();
            }
        }

        static public void SozdSpKVKoor(ref List<string> Vkrazv, List<PLOS> SpPLOS)
        {
            string stIND = "";

            double Dlin = 0;
            double Param=0;
            double ZParam = 0;

            Point3d Zentr = new Point3d();
            List<TPodk> SpVin = new List<TPodk>();
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Kabeli"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    foreach (SelectedObject sobj in acSSet)
                    {
                        Polyline ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Polyline;
                        //var KolV = ln.NumberOfVertices;
                        Dlin = ln.Length;
                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stIND = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                            Param = ln.EndParam;
                            ZParam = Param / 2;
                            Zentr = ln.GetPointAtParameter(ZParam);
                            foreach (PLOS iPlos in SpPLOS)
                            {
                                if (Zentr.X > iPlos.min.X & Zentr.X < iPlos.max.X & Zentr.Y > iPlos.min.Y & Zentr.Y < iPlos.max.Y)
                                {
                                    double delX = Zentr.X - iPlos.Psk.X;
                                    double delY = Zentr.Y - iPlos.Psk.Y;
                                    double delZ = Zentr.Z - iPlos.Psk.Z;
                                    double Xnow = iPlos.Msk.X;
                                    double Ynow = iPlos.Msk.Y;
                                    double Znow = iPlos.Msk.Z;
                                    if (iPlos.Osi == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; Znow = Znow + delZ; }
                                    if (iPlos.Osi == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                                    if (iPlos.Osi == "ZY") { Xnow = Xnow + delZ; Ynow = Ynow + delX; Znow = Znow + delY; }
                                    if (iPlos.Osi == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                                    if (iPlos.Osi == "YZ") { Xnow = Xnow + delZ; Ynow = Ynow - delX; Znow = Znow + delY; }
                                    Point3d KoorMod = new Point3d(Xnow, Ynow, Znow);
                                    Vkrazv.Add(stIND + ":" + Xnow.ToString() + " " + Ynow.ToString() + " " + Znow.ToString() + ":" + Dlin.ToString("00"));
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создания списка кабелей
        static public void SpVistTP(ref List<string> Vkrazv, List<PLOS> SpPLOS)
        {
            Point3d Zentr = new Point3d();
            string stIND = "";
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[6];
            acTypValAr.SetValue(new TypedValue(0, "CIRCLE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Kabeli"), 2);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            acTypValAr.SetValue(new TypedValue(40, 15.0), 5);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
                    TPodk TNOD = new TPodk();
                    foreach (ObjectId sobj in acSSet.GetObjectIds())
                    {
                        Circle ln = tr.GetObject(sobj, OpenMode.ForWrite) as Circle;
                        if (ln != null)
                        {
                            //TNOD.NKoor(ln.Center);
                            Zentr = ln.Center;
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { stIND=value.Value.ToString(); }
                                    Schet = Schet + 1;
                                    if (Schet > 1) { break; }
                                }
                                foreach (PLOS iPlos in SpPLOS)
                                {
                                    if (Zentr.X > iPlos.min.X & Zentr.X < iPlos.max.X & Zentr.Y > iPlos.min.Y & Zentr.Y < iPlos.max.Y)
                                    {
                                        double delX = Zentr.X - iPlos.Psk.X;
                                        double delY = Zentr.Y - iPlos.Psk.Y;
                                        double delZ = Zentr.Z - iPlos.Psk.Z;
                                        double Xnow = iPlos.Msk.X;
                                        double Ynow = iPlos.Msk.Y;
                                        double Znow = iPlos.Msk.Z;
                                        if (iPlos.Osi == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; Znow = Znow + delZ; }
                                        if (iPlos.Osi == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                                        if (iPlos.Osi == "ZY") { Xnow = Xnow + delZ; Ynow = Ynow + delX; Znow = Znow + delY; }
                                        if (iPlos.Osi == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                                        if (iPlos.Osi == "YZ") { Xnow = Xnow + delZ; Ynow = Ynow - delX; Znow = Znow + delY; }
                                        Point3d KoorMod = new Point3d(Xnow, Ynow, Znow);
                                        Vkrazv.Add(stIND + ":" + Xnow.ToString("00.00") + " " + Ynow.ToString("00.00") + " " + Znow.ToString("00.00"));
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создание списка точек подключения
        public void SozdSpKV(ref List<string> Vkrazv)
        {
            string stIND = "";
            string stDlin = "";
            string stStar = "";
            string stNov = "";
            string stNom = "";
            string SpVinKab = "";

            int nom;

            double Dlin = 0;
            double DlinN = 0;

            List<TPodk> SpVin = new List<TPodk>();
            SpVistVin(ref SpVin);

            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Kabeli"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    foreach (SelectedObject sobj in acSSet)
                    {
                        Polyline ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Polyline;
                        //var KolV = ln.NumberOfVertices;
                        Dlin = ln.Length;
                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Noda Nnoda = new Noda();
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stIND = value.Value.ToString(); }
                                if (Schet == 5) { stNom = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                            if (Vkrazv.Exists(x => x.Split(':')[0] == stIND) == true)
                            {
                                nom = Vkrazv.FindIndex(x => x.Split(':')[0] == stIND);
                                stStar = Vkrazv.Find(x => x.Split(':')[0] == stIND);
                                DlinN = Convert.ToDouble(stStar.Split(':')[1]) + Dlin;
                                SpVinKab = stStar.Split(':')[2];
                                fSpVin(ln, SpVin, ref SpVinKab);
                                stNov = stIND + ":" + DlinN.ToString("#0.##") + ":" + SpVinKab;
                                Vkrazv[nom] = stNov;
                            }
                            else
                            {
                                SpVinKab = "";
                                fSpVin(ln, SpVin, ref SpVinKab);
                                Vkrazv.Add(stIND + ":" + Dlin.ToString("#0.##") + ":" + SpVinKab);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создания списка выносо
        public void fSpVin(Polyline ln, List<TPodk> LSpVin, ref string SpVin)
        {
            //string SpVin = "";
            double DistA_B;
            double DistA_C;
            double DistC_B;
            double DELT;
            var KolV = ln.NumberOfVertices;
            //for (int i = 0; i <= KolV - 2; i++)
            for (int i = KolV - 1; i > 0; i--)
            {
                Point3d TOtr1 = ln.GetPointAtParameter(i);
                Point3d TOtr2 = ln.GetPointAtParameter(i - 1);
                DistA_B = TOtr1.DistanceTo(TOtr2);
                foreach (TPodk TV in LSpVin)
                {
                    DistA_C = TOtr1.DistanceTo(TV.Koord);
                    DistC_B = TV.Koord.DistanceTo(TOtr2);
                    DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                    if (DELT < 0.5) { if (SpVin.Contains(TV.IND + "*") == false) SpVin = SpVin + TV.IND + "*"; }
                }
            }
        }//
        public void SpVistVin(ref List<TPodk> SpVIN)
        {
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 1);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
                    TPodk TNOD = new TPodk();
                    foreach (ObjectId sobj in acSSet.GetObjectIds())
                    {
                        BlockReference ln = tr.GetObject(sobj, OpenMode.ForWrite) as BlockReference;
                        if (ln != null)
                        {
                            TNOD.NKoor(ln.Position);
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { TNOD.NIND(value.Value.ToString()); }
                                    Schet = Schet + 1;
                                    if (Schet > 1) { break; }
                                }
                            }
                            SpVIN.Add(TNOD);
                        }
                    }
                    tr.Commit();
                }
            }
        }//создание списка выносок
    }
}
