using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace RASH_DAN
{
#if NANOCAD
    using Teigha.DatabaseServices;
    using Teigha.Geometry;
    using HostMgd.ApplicationServices;
    using HostMgd.EditorInput;
    using DbS = Teigha.DatabaseServices;
    using EdI = HostMgd.EditorInput;
#elif AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.EditorInput;
    using DbS = Autodesk.AutoCAD.DatabaseServices;
    using EdI = Autodesk.AutoCAD.EditorInput;
#endif



    public partial class Form6 : Form
    {
        public struct VINOS
        {
            public string Nazv,Spravka;
            public Point3d Koord;
            public void NNazv(string i) { Nazv = i; }
            public void NSpravka(string i) { Spravka = i; }
            public void NKoor(Point3d i) { Koord = i; }
        };
        public struct PEREXOD
        {
            public string Nazv, Spravka;
            public double Dlin;
            public Point3d Koord1, Koord2;
            public void NNazv(string i) { Nazv = i; }
            public void NSpravka(string i) { Spravka = i; }
            public void NDlin(double i) { Dlin = i; }
            public void NKoor1(Point3d i) { Koord1 = i; }
            public void NKoor2(Point3d i) { Koord2 = i; }
        };
        public struct PLOS
        {
            public string Vid, Mas, Spravka, List;
            public Point3d Psk, Msk,max,min;
            public void NVid(string i) { Vid = i; }
            public void NMas(string i) { Mas = i; }
            public void NSpravka(string i) { Spravka = i; }
            public void NList(string i) { List = i; }
            public void NPsk(Point3d i) { Psk = i; }
            public void NMsk(Point3d i) { Msk = i; }
            public void Nmax(Point3d i) { max = i; }
            public void Nmin(Point3d i) { min = i; }
        };
        public List<VINOS> Vinoski = new List<VINOS>();
        public List<PEREXOD> PEREX = new List<PEREXOD>();
        public List<PLOS> SpPLOS = new List<PLOS>();
        public string Handl1;
        public string Handl2;
        public Document doc = Application.DocumentManager.MdiActiveDocument;
        public Form6()
        {
            InitializeComponent();

        }
        public void Form6_Load(object sender, EventArgs e)
        {
            SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS);
            foreach (VINOS TV in Vinoski) { this.dataGridView1.Rows.Add(TV.Nazv, TV.Spravka); }
            foreach (PEREXOD TV in PEREX) { this.dataGridView2.Rows.Add(TV.Nazv, TV.Dlin.ToString(), TV.Nazv); }
            foreach (PLOS TV in SpPLOS) { this.dataGridView3.Rows.Add(TV.Vid, TV.List, TV.Spravka, TV.Msk.ToString()); }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            //InsBlockRef(@"D:\PROBA\TRASER\Переход.dwg","Переход");
            InsBlockRef(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\Переход.dwg", "Переход");
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()){SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS);}
            this.dataGridView2.Rows.Clear();
            foreach (PEREXOD TV in PEREX) { this.dataGridView2.Rows.Add(TV.Nazv, TV.Dlin.ToString(), TV.Nazv); }
            this.Show();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            //InsBlockRef(@"D:\PROBA\TRASER\Выноска.dwg", "Выноска");
            InsBlockRef(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\Выноска.dwg", "Выноска");
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()) { SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS); }
            this.dataGridView1.Rows.Clear();
            foreach (VINOS TV in Vinoski) { this.dataGridView1.Rows.Add(TV.Nazv, TV.Spravka); }
            this.Show();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            //InsBlockRef(@"D:\PROBA\TRASER\Плоскость_пр.dwg", "Плоскость_пр");
            InsBlockRef(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\Плоскость_пр.dwg", "Плоскость_пр");
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()) { SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS); }
            this.dataGridView3.Rows.Clear();
            foreach (PLOS TV in SpPLOS) { this.dataGridView3.Rows.Add(TV.Vid, TV.List, TV.Spravka, TV.Msk.ToString()); }
            this.Show();
        }


        //public void SPISKI(ref List<VINOS> SpVIN, ref List<PEREXOD> SpPEREX)
        //{
        //    int Schet = 0;
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;
        //    ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
        //    ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
        //    TypedValue[] acTypValAr = new TypedValue[9];
        //    acTypValAr.SetValue(new TypedValue(0, "CIRCLE"), 0);
        //    acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
        //    acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 2);
        //    acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 3);
        //    acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
        //    acTypValAr.SetValue(new TypedValue(-4, "<or"), 5);
        //    acTypValAr.SetValue(new TypedValue(40, 4.0), 6);
        //    acTypValAr.SetValue(new TypedValue(40, 6.0), 7);
        //    acTypValAr.SetValue(new TypedValue(-4, "or>"), 8);
        //    // создаем фильтр
        //    SelectionFilter filter = new SelectionFilter(acTypValAr);
        //    PromptSelectionResult selRes = ed.SelectAll(filter);
        //    if (selRes.Status == PromptStatus.OK)
        //    {
        //        SelectionSet acSSet = selRes.Value;
        //        Transaction tr = db.TransactionManager.StartTransaction();
        //        using (tr)
        //        {
        //            Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
        //            foreach (ObjectId sobj in acSSet.GetObjectIds())
        //            {
        //                Circle ln = tr.GetObject(sobj, OpenMode.ForWrite) as Circle;
        //                if (ln != null)
        //                {
        //                    if (ln.Radius == 6)
        //                    {
        //                        VINOS tVin=new VINOS();
        //                        tVin.NKoor(ln.Center);
        //                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
        //                        if (buffer != null)
        //                        {
        //                            Schet = 0;
        //                            foreach (TypedValue value in buffer)
        //                            {
        //                                if (Schet == 1) { tVin.NNazv(value.Value.ToString()); }
        //                                Schet = Schet + 1;
        //                                if (Schet > 1) { break; }
        //                            }
        //                        }
        //                        SpVIN.Add(tVin);
        //                    }
        //                    if (ln.Radius == 4)
        //                    {
        //                        PEREXOD tVin = new PEREXOD();
        //                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
        //                        if (buffer != null)
        //                        {
        //                            Schet = 0;
        //                            foreach (TypedValue value in buffer)
        //                            {
        //                                if (Schet == 1) { tVin.NNazv(value.Value.ToString()); }
        //                                Schet = Schet + 1;
        //                                if (Schet > 1) { break; }
        //                            }
        //                        }
        //                        if (SpPEREX.Exists(x => x.Nazv == tVin.Nazv) == false) {SpPEREX.Add(tVin); } 
        //                    }
        //                }
        //            }
        //            tr.Commit();
        //        }
        //    }
        //}

        static public void SBOR_OBOR(ref List<VINOS> SpVIN, ref List<PEREXOD> SpPEREX, ref List<PLOS> SpPLOS)
        {
            string Nazv_Vin = "", Sprav_vin = "";
            string Nazv_per = "", Sprav_per = "";
            string Vid = "", Mash = "", Sprav_plos = "", List = "";
            double Dlin = 0;
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
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 2);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
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
                    Nazv_Vin = "";
                    Sprav_vin = "";
                    //для переходов
                    Nazv_per = "";
                    Sprav_per = "";
                    //для плоскостей
                    Vid = ""; 
                    Mash = ""; 
                    Sprav_plos = ""; 
                    List = "";
                    Dlin = 0;
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
                            if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение4 X") { x2 = Convert.ToDouble(prop.Value.ToString()); }
                            if (prop.PropertyName == "Положение4 Y") { y2 = Convert.ToDouble(prop.Value.ToString()); }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "НОМЕР_ВЫНОСКИ") { Nazv_Vin = atrRef.TextString; }
                                if (atrRef.Tag == "Справочная_информация") { Sprav_vin = atrRef.TextString; }
                                //переходы
                                if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА1") { Nazv_per = atrRef.TextString; }
                                if (atrRef.Tag == "СПРАВКА_ПЕРЕХОДА1") { Sprav_per = atrRef.TextString; }
                                if (atrRef.Tag == "ДЛИНА_ПЕРЕХОДА") { Dlin = Convert.ToDouble(atrRef.TextString); }
                                //плоскости
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
                    //Application.ShowAlertDialog(bref.Name);
                    if (Nazv_Vin != "")
                    {
                        VINOS tPOZ = new VINOS();
                        tPOZ.NNazv(Nazv_Vin);
                        tPOZ.NSpravka(Sprav_vin);
                        tPOZ.NKoor(BP);
                        SpVIN.Add(tPOZ);
                    }
                    if (Nazv_per != "")
                    {
                        Point3d T1 = new Point3d(BP.X + x1, BP.Y + y1, 0);
                        Point3d T2 = new Point3d(BP.X + x2, BP.Y + y2, 0);
                        PEREXOD tPOZ = new PEREXOD();
                        tPOZ.NNazv(Nazv_per);
                        tPOZ.NSpravka(Sprav_vin);
                        tPOZ.NDlin(Dlin);
                        tPOZ.NKoor1(T1);
                        tPOZ.NKoor2(T2);
                        SpPEREX.Add(tPOZ);
                    }
                    //Application.ShowAlertDialog(bref.Name);
                    if (Vid != "")
                    {
                        Point3d Psk = new Point3d(BP.X + x1, BP.Y + y1, 0);
                        Point3d Msk = new Point3d(xMir, yMir, zMir);
                        PLOS tPOZ = new PLOS();
                        tPOZ.NPsk(Psk);
                        tPOZ.NMsk(Msk);
                        tPOZ.NVid(Vid);
                        tPOZ.NMas(Mash);
                        tPOZ.NList(List);
                        tPOZ.NSpravka(Sprav_plos);
                        SpPLOS.Add(tPOZ);
                    }
                }
                Tx.Commit();
            }
        }//Создание списков деталей изображенных блоками


//        public void VisVin() 
//        {
//            Document doc = Application.DocumentManager.MdiActiveDocument;
//            Database db = doc.Database;
//            Transaction tr = db.TransactionManager.StartTransaction();
//            PromptPointResult pPtRes;
//            PromptPointResult pPtRes1;
//            PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку на трассе");
//            PromptPointOptions pPtOpts1 = new PromptPointOptions("Укажи точку для выноски");
//            Point3d Toch1;
//            ObjectId ID1 = new ObjectId();
//            CreateLayer("ТРАССЫ");
//            using (tr)
//            {
//                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
//                pPtRes = doc.Editor.GetPoint(pPtOpts);
//                Toch1 = pPtRes.Value;
//                Circle KrugPOZVn = new Circle();
//                KrugPOZVn.SetDatabaseDefaults();
//                KrugPOZVn.Center = Toch1;
//                KrugPOZVn.Radius = 6;
//                KrugPOZVn.Layer = "ТРАССЫ";
//                btr.AppendEntity(KrugPOZVn);
//                tr.AddNewlyCreatedDBObject(KrugPOZVn, true);
//                tr.Commit();
//            }
//            PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
//            SelectionSet acSSet = acSSPrompt.Value;
//            using (Transaction Tx = db.TransactionManager.StartTransaction())
//            {
//                foreach (SelectedObject acSSObj in acSSet)
//                {
//                    Circle bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Circle;
//                    ID1 = bref.ObjectId;
//                    Handl1 = bref.Handle.ToString();
//                    Tx.Commit();
//                }
//            }
//            Transaction tr14 = db.TransactionManager.StartTransaction();
//            using (tr14){RAS_D(this.textBox3.Text, "999999999", "", "", "", "0", ID1);tr14.Commit();}
//            Transaction tr1 = db.TransactionManager.StartTransaction();
//             using (tr1)
//             {
//                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
//                pPtOpts1.UseBasePoint = true;
//                pPtOpts1.BasePoint = Toch1;
//                pPtRes1 = doc.Editor.GetPoint(pPtOpts1);
//                Point3d Toch2 = pPtRes1.Value;
//                DBText TNPol = new DBText();
//                TNPol.Position = new Point3d(Toch2.X + 10, Toch2.Y + 5, Toch2.Z);
//                TNPol.SetDatabaseDefaults();
//                TNPol.Layer = "ТРАССЫ";
//                TNPol.Height = 50;
//                TNPol.TextString = this.textBox3.Text;
//                btr.AppendEntity(TNPol);
//                tr1.AddNewlyCreatedDBObject(TNPol, true);

//                Line acLine1 = new Line();
//                acLine1.StartPoint = Toch1;
//                acLine1.EndPoint = Toch2;
//                acLine1.SetDatabaseDefaults();
//                acLine1.Layer = "ТРАССЫ";
//                acLine1.SetDatabaseDefaults();
//                btr.AppendEntity(acLine1);
//                tr1.AddNewlyCreatedDBObject(acLine1, true);

//                Line acLine2 = new Line();
//                acLine2.StartPoint = Toch2;
//                acLine2.EndPoint = new Point3d(Toch2.X + 100, Toch2.Y, Toch2.Z);
//                acLine2.SetDatabaseDefaults();
//                acLine2.Layer = "ТРАССЫ";
//                acLine2.SetDatabaseDefaults();
//                btr.AppendEntity(acLine2);
//                tr1.AddNewlyCreatedDBObject(acLine2, true);
//                tr1.Commit();
//            }
//            //this.listBox1.Items.Add(this.textBox3.Text);
//        }
//        public void VisPer()
//        {
//            PromptPointResult pPtRes;
//            PromptPointResult pPtRes1;
//            PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку на трассе");
//            PromptPointOptions pPtOpts1 = new PromptPointOptions("Укажи точку для выноски");
//            Point3d Toch1 = new Point3d();
//            Point3d Toch2 = new Point3d();
//            ObjectId ID1=new ObjectId();
//            ObjectId ID2=new ObjectId();
//            Database db = doc.Database;
///////////////////////////////первая точка перехода
//            CreateLayer("ТРАССЫскрытые");
//            Transaction tr = db.TransactionManager.StartTransaction();
//            using (tr)
//            {                                    
//                pPtRes = doc.Editor.GetPoint(pPtOpts);
//                Toch1 = pPtRes.Value;
//                KRUG(Toch1);
//                tr.Commit(); 
//            }
//            PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
//            SelectionSet acSSet = acSSPrompt.Value;
//            using (Transaction Tx = db.TransactionManager.StartTransaction())
//            {
//                foreach (SelectedObject acSSObj in acSSet) 
//                { 
//                    Circle bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Circle;
//                    ID1 = bref.ObjectId;
//                    Handl1 = bref.Handle.ToString();
//                    Tx.Commit();
//                } 
//            }        
//            Transaction tr11 = db.TransactionManager.StartTransaction();
//            using (tr11)
//            {
//                pPtOpts1.UseBasePoint = true;
//                pPtOpts1.BasePoint = Toch1;
//                pPtRes1 = doc.Editor.GetPoint(pPtOpts1);
//                Toch2 = pPtRes1.Value;
//                LINE(Toch1, Toch2);
//                LINE(Toch2, new Point3d(Toch2.X + 50, Toch2.Y, Toch2.Z));
//                TEXT(Toch2, this.textBox1.Text);
//                tr11.Commit();
//            }
///////////////////////////////вторая точка перехода
//            Transaction tr12 = db.TransactionManager.StartTransaction();
//            using (tr12)
//            {
//                pPtRes = doc.Editor.GetPoint(pPtOpts);
//                Toch1 = pPtRes.Value;
//                KRUG(Toch1);
//                tr12.Commit();
//            }
//            PromptSelectionResult acSSPrompt1 = doc.Editor.SelectLast();
//            SelectionSet acSSet1 = acSSPrompt1.Value;
//            using (Transaction Tx1 = db.TransactionManager.StartTransaction())
//            {
//                foreach (SelectedObject acSSObj in acSSet1)
//                {
//                    Circle bref = Tx1.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Circle;
//                    ID2 = bref.ObjectId;
//                    Handl2 = bref.Handle.ToString();
//                    Tx1.Commit();
//                }
//            }
//            Transaction tr13 = db.TransactionManager.StartTransaction();
//            using (tr13)
//            {
//                pPtOpts1.UseBasePoint = true;
//                pPtOpts1.BasePoint = Toch1;
//                pPtRes1 = doc.Editor.GetPoint(pPtOpts1);
//                Toch2 = pPtRes1.Value;
//                LINE(Toch1, Toch2);
//                LINE(Toch2, new Point3d(Toch2.X + 50, Toch2.Y, Toch2.Z));
//                TEXT(Toch2, this.textBox1.Text);
//                tr13.Commit();
//            }
//            Transaction tr14 = db.TransactionManager.StartTransaction();
//            using (tr14)
//            {
//                RAS_D(this.textBox1.Text, "0", Handl2, "", "переход", this.textBox2.Text, ID1);
//                RAS_D(this.textBox1.Text, "0", Handl1, "", "переход", this.textBox2.Text, ID2);
//                tr14.Commit();
//            }
//            //this.listBox2.Items.Add(this.textBox1.Text);
//        }
        public void InsBlockRef(string BlockPath, string NAME)
        {
            // Активный документ в редакторе AutoCAD
            Document doc = Application.DocumentManager.MdiActiveDocument;
            // База данных чертежа (в данном случае - активного документа)
            Database db = doc.Database;
            // Редактор базы данных чертежа
            // Запускаем транзакцию
            using (DocumentLock docLock = doc.LockDocument())
            {
                CreateLayer("ТРАССЫ");
                CreateLayer("Плоскости");
                using (DbS.Transaction tr = db.TransactionManager.StartTransaction())
                {
                    EdI.Editor ed = doc.Editor;
                    EdI.PromptPointOptions pPtOpts;
                    pPtOpts = new EdI.PromptPointOptions("\nУкажите точку вставки блока: ");
                    // Выбор точки пользователем
                    var pPtRes = doc.Editor.GetPoint(pPtOpts);
                    if (pPtRes.Status != EdI.PromptStatus.OK)
                        return;
                    var ptStart = pPtRes.Value;

                    DbS.BlockTable bt = tr.GetObject(db.BlockTableId, DbS.OpenMode.ForRead) as DbS.BlockTable;
                    DbS.BlockTableRecord model = tr.GetObject(bt[DbS.BlockTableRecord.ModelSpace], DbS.OpenMode.ForWrite) as DbS.BlockTableRecord;
                    // Создаем новую базу
                    using (DbS.Database db1 = new DbS.Database(false, false))
                    {
                        // Получаем базу чертежа-донора
                        db1.ReadDwgFile(BlockPath, System.IO.FileShare.Read, true, null);
                        // Получаем ID нового блока
                        DbS.ObjectId BlkId = db.Insert(NAME, db1, false);
                        DbS.BlockReference bref = new DbS.BlockReference(ptStart, BlkId);
                        // Дефолтные свойства блока (слой, цвет и пр.)
                        bref.SetDatabaseDefaults();
                        // Добавляем блок в модель
                        model.AppendEntity(bref);
                        // Добавляем блок в транзакцию
                        tr.AddNewlyCreatedDBObject(bref, true);
                        // Расчленяем блок     
                        bref.ExplodeToOwnerSpace();
                        bref.Erase();
                        // Закрываем транзакцию
                        tr.Commit();
                    }
                }
                SetDynamicBlkProperty(NAME);
            }
        }//Выставмить дин блок
        static public void SetDynamicBlkProperty(string NAME)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    RegAppTable regTable = (RegAppTable)Tx.GetObject(db.RegAppTableId, OpenMode.ForWrite);
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    if (NAME == "Плоскость_пр")
                    bref.Layer = "Плоскости";
                    else
                    bref.Layer = "ТРАССЫ";
                    //foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    //{
                    //    using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                    //    {
                    //        if (atrRef != null)
                    //        {
                    //            if (atrRef.Tag == "Помещение") { atrRef.TextString = Pom; }
                    //            if (atrRef.Tag == "Высота_установки") { atrRef.TextString = Visota.ToString(); }
                    //            if (atrRef.Tag == "Раздел_спецификации") { atrRef.TextString = RazdelSp; }
                    //        }
                    //    }
                    //}
                }
                Tx.Commit();
            }
        }//расширеные данные


        public void KRUG(Point3d Toch1)
        {
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Circle KrugPOZVn = new Circle();
                KrugPOZVn.SetDatabaseDefaults();
                KrugPOZVn.Center = Toch1;
                KrugPOZVn.Radius = 4;
                KrugPOZVn.Layer = "ТРАССЫскрытые";
                btr.AppendEntity(KrugPOZVn);
                tr1.AddNewlyCreatedDBObject(KrugPOZVn, true);
                tr1.Commit();
            }
        }
        public void TEXT(Point3d Toch1,string strStr)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr13 = db.TransactionManager.StartTransaction();
            using (tr13)
            {
                BlockTableRecord btr = (BlockTableRecord)tr13.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                DBText TNPol = new DBText();
                TNPol.Position = new Point3d(Toch1.X + 2, Toch1.Y + 2, Toch1.Z);
                TNPol.SetDatabaseDefaults();
                TNPol.Layer = "ТРАССЫскрытые";
                TNPol.Height = 5;
                TNPol.TextString = strStr;
                //TNPol.TextString = this.textBox1.Text;
                btr.AppendEntity(TNPol);
                tr13.AddNewlyCreatedDBObject(TNPol, true);
                tr13.Commit();
            }
        }
        public void LINE(Point3d Toch1, Point3d Toch2) 
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr12 = db.TransactionManager.StartTransaction();
            using (tr12) 
            {
                BlockTableRecord btr = (BlockTableRecord)tr12.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Line acLine1 = new Line();
                acLine1.StartPoint = Toch1;
                acLine1.EndPoint = Toch2;
                acLine1.SetDatabaseDefaults();
                // Add the line to the drawing
                btr.AppendEntity(acLine1);
                tr12.AddNewlyCreatedDBObject(acLine1, true);
                tr12.Commit();
            }
        }
        public void CreateLayer(string SloiName) 
        {
            ObjectId layerID;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (Transaction Trans = db.TransactionManager.StartTransaction()) 
            {
                LayerTable LT = (LayerTable)Trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
                if (LT.Has(SloiName))
                { layerID = LT[SloiName]; }
                else
                {
                    LayerTableRecord LTR = new LayerTableRecord();
                    LTR.Name = SloiName;
                    layerID = LT.Add(LTR);
                    Trans.AddNewlyCreatedDBObject(LTR, true);
                }
                Trans.Commit();
            }
        }
        public void RAS_D(string strNAME, string strVIS, string strSVOI, string strNeSVI, string strHOZ, string strDlin, ObjectId ID)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                try
                {
                    RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                    Entity ent = (Entity)tr.GetObject(ID, OpenMode.ForWrite);
                    if (!regTable.Has("LAUNCH01"))
                    {
                        regTable.UpgradeOpen();
                        // Добавляем имя приложения, которое мы будем
                        // использовать в расширенных данных
                        RegAppTableRecord app =
                                new RegAppTableRecord();
                        app.Name = "LAUNCH01";
                        regTable.Add(app);
                        tr.AddNewlyCreatedDBObject(app, true);
                    }
                    // Добавляем расширенные данные к примитиву
                    ent.XData = new ResultBuffer(
                        new TypedValue(1001, "LAUNCH01"),
                        new TypedValue(1000, strNAME),
                        new TypedValue(1040, Convert.ToDouble(strVIS)),
                        new TypedValue(1000, strSVOI),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, ""),
                        new TypedValue(1000, strNeSVI),
                        new TypedValue(1000, strHOZ),
                        new TypedValue(1040, Convert.ToDouble(strDlin))
                        );
                  tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
            }
        }


       
    }
}
