using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace RASH_DAN
{
#if NANOCAD
    using Teigha.DatabaseServices;
    using Teigha.Geometry;
    using HostMgd.ApplicationServices;
    using HostMgd.EditorInput;
#elif AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.EditorInput;
#endif

    public partial class Form3 : Form
    {
        public struct Noda
        {
            public string Nom, Otkuda,Hoz, Vin;
            public string SpSmNod;
            public double Ves;
            public double DlinP;
            public double VisT;
            public Point3d Koord;
            public Point3d Koord3D;
            public void NomVin(string i) { Vin = i; }
            public void NomNod(string i) { Nom = i; }
            public void NomNOtk(string i) { Otkuda = i; }
            public void NomHoz(string i) { Hoz = i; }
            public void NSpSmNod(string i) { SpSmNod = i; }
            public void NVes(double i) { Ves = i; }
            public void NDlinP(double i) { DlinP = i; }
            public void NVisT(double i) { VisT = i; }
            public void NKoor(Point3d i) { Koord = i; }
            public void NKoor3D(Point3d i) { Koord3D = i; }
        };
        public struct DUGA
        {
            public string NOD1, NOD2;
            public double Ves;
            public void NomNOD1(string i) { NOD1 = i; }
            public void NomNOD2(string i) { NOD2 = i; }
            public void NVes(double i) { Ves = i; }
        };
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

        public string strSist = "";
        public List<Noda> spNod0 = new List<Noda>();
        public int ZvetLin;


        public Form3()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            double dZvet=6;
            string strInd=this.textBox1.Text;
            string strDiam = this.textBox2.Text;
            string strBLOKNOD = this.textBox3.Text;
            List<Noda> SpNOD = new List<Noda>();
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument()){UkazT(strInd, strDiam, strBLOKNOD, dZvet);}
            this.Show();
        }
        public void UkazT(string strInd, string strDiam, string strBLOKNOD, double Zvet) 
        {
            List<Noda> spNod = new List<Noda>();
            spNod= spNod0.ToList();
            string Rez = "Есть";
            List<DUGA> DUGI = new List<DUGA>();
            List<Noda> OPEN = new List<Noda>();
            List<Noda> CLOSE = new List<Noda>();
            Noda T1 = new Noda();
            Noda T2 = new Noda();
            Noda TNOD = new Noda();
            Document doc =Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            PromptPointResult pPtRes;
            PromptPointResult pPtRes1;
            PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку");
            List<TPodk> SpVin = new List<TPodk>();
            List<TPodk> SpBlVin = new List<TPodk>();
            if (strBLOKNOD != "")
            {
                string[] strBLOKNODm = strBLOKNOD.Split(',');
                foreach (string TVin in strBLOKNODm) UdalSvasi(ref spNod, TVin);
            }
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d Toch1 = pPtRes.Value;
                pPtRes1 = doc.Editor.GetPoint(pPtOpts);
                Point3d Toch2 = pPtRes1.Value;
                BlizhNOD(Toch1, spNod, ref T1);
                BlizhNOD(Toch2, spNod, ref T2);
                T1.NVes(0);
                T1.NomNOtk("-");
                TNOD = T1;
                string strTNOD = TNOD.Nom;
                CLOSE.Add(TNOD);
                if (T1.Nom != T2.Nom)
                {
                    while (Rez == "Есть")
                    {
                        spDUG(ref DUGI, ref TNOD, ref OPEN, ref CLOSE,Zvet, spNod);
                        if (DUGI.Count > 0)
                        {
                            Relaks(ref DUGI, ref TNOD,ref OPEN);
                            SlNOD(ref DUGI, ref TNOD,ref OPEN, ref CLOSE, ref strTNOD);
                            if (strTNOD == "") { Rez = "Нет"; }
                        }
                        else 
                        {
                            SlNOD(ref DUGI, ref TNOD,ref OPEN, ref CLOSE, ref strTNOD);
                            if (strTNOD == "") { Rez = "Нет"; }
                        }
                        //Application.ShowAlertDialog(TNOD.Nom);
                        DUGI.Clear();
                    }
                    NodiVKTXTfail(CLOSE, "NODI");
                    POSTR(ref CLOSE, T1, T2, Toch1, Toch2, strInd, strDiam);
                }
                else
                {
                    Application.ShowAlertDialog("Точки совподают");
                }
                tr.Commit();
            }
        }//указать точки вручную
        public void spDUG(ref List<DUGA> DUGI, ref Noda TNOD, ref List<Noda> OPEN, ref List<Noda> CLOSE, double Zvet , List<Noda> spNod)
        {
            string[] SpSmNOD = TNOD.SpSmNod.Split(',');
            Point3d KoorTN = TNOD.Koord3D;    
            foreach (string NTSmNoda in SpSmNOD)
            {
                string[] tNoda = NTSmNoda.Split('*');
                if ( CLOSE.Exists(x => x.Nom == tNoda[0]) == false)
                    {
                        Noda TSMNod = spNod.Find(x => x.Nom == tNoda[0]);
                        DUGA TDuga = new DUGA();
                        TDuga.NomNOD1(TNOD.Nom);
                        TDuga.NomNOD2(tNoda[0]);
                        TDuga.NVes(Convert.ToDouble(tNoda[1]));
                        DUGI.Add(TDuga);
                        OPEN.Add(TSMNod);
                }
            }
        }//список дуг и пополнение списка ОПЕН
        public void Relaks(ref List<DUGA> DUGI, ref Noda TNOD, ref List<Noda> OPEN)
        {
        //string strOP = "";
            foreach (DUGA TDUGA in DUGI) 
            {
                Noda TSMNod = OPEN.Find(x => x.Nom == TDUGA.NOD2);
                if (TSMNod.Ves > TNOD.Ves + TDUGA.Ves)
                {
                    OPEN.Remove(TSMNod);
                    Noda RelNod = TSMNod;
                    RelNod.NomNOtk(TNOD.Nom);
                    RelNod.NVes(TNOD.Ves + TDUGA.Ves);
                    OPEN.Add(RelNod);
                }
            }
        }//релаксация нод в списке ОПЕН
        public void SlNOD(ref List<DUGA> DUGI, ref Noda TNOD,  ref List<Noda> OPEN, ref List<Noda> CLOSE, ref   string strTNOD)
        {
            double minVes = 9999999999999.0;
            strTNOD = "";
            foreach (Noda TVNod in OPEN)
            {
                if (TVNod.Ves  <= minVes )
                {
                    minVes = TVNod.Ves;
                    TNOD = TVNod;
                    strTNOD = TNOD.Nom;
                }
            }
            CLOSE.Add(TNOD);
            OPEN.Remove(TNOD);
        }//поиск следующего нода
        public void POSTR(ref List<Noda> CLOSE, Noda T1, Noda T2, Point3d Toch1, Point3d Toch2, string strInd, string strDiam)
        {
            Point3dCollection TkoorPL = new Point3dCollection();
            Point3d pToh;
            TkoorPL.Add(Toch2);
            string strHOZ;
            double dDlinKab=0;
            string NTTohc = T2.Nom;
            Noda TNod = T1;
            pToh = Toch2;
            //Application.ShowAlertDialog("Дошла до построения");
            if (CLOSE.Exists(x => x.Nom == NTTohc) == true)
            {
                TNod = CLOSE.Find(x => x.Nom == NTTohc);
                TkoorPL.Add(TNod.Koord);
                //dDlinKab = dDlinKab + TNod.Koord3D.DistanceTo(pToh);
                //pToh = TNod.Koord3D;
                NTTohc = TNod.Otkuda;
                strHOZ = TNod.Nom;
            }
            else 
            {
                //dDlinKab = Toch1.DistanceTo(Toch2);
                TkoorPL.Add(Toch1);
                //FPoly(TkoorPL, strInd, strDiam, 0, dDlinKab);
                FPoly2d(TkoorPL, strInd, strDiam, 0, dDlinKab);
                return;
            }
            while (NTTohc != T1.Nom)
            {  
                    TNod = CLOSE.Find(x => x.Nom == NTTohc);
                    //if (TNod.Hoz == "переход" & strHOZ == "переход")
                    if (TNod.Nom.Contains("Переход") == true & strHOZ.Contains("Переход") == true)
                    {
                        //FPoly(TkoorPL, strInd, strDiam, TNod.DlinP, dDlinKab);
                        FPoly2d(TkoorPL, strInd, strDiam, TNod.DlinP, dDlinKab);
                        dDlinKab = 0;
                        TkoorPL.Clear();
                        TkoorPL.Add(TNod.Koord);
                        //pToh = TNod.Koord3D;
                        NTTohc = TNod.Otkuda;
                        strHOZ = TNod.Nom;
                    }
                    else
                    {
                        TkoorPL.Add(TNod.Koord);
                        //dDlinKab = dDlinKab + TNod.Koord3D.DistanceTo(pToh);
                        //pToh = TNod.Koord3D;
                        NTTohc = TNod.Otkuda;
                        strHOZ = TNod.Nom;
                    }
            }
            TNod = CLOSE.Find(x => x.Nom == NTTohc);
            TkoorPL.Add(TNod.Koord);
            //dDlinKab = dDlinKab + TNod.Koord3D.DistanceTo(pToh);
            pToh = TNod.Koord;
            TkoorPL.Add(Toch1);
            //dDlinKab = dDlinKab + Toch1.DistanceTo(pToh);
            //FPoly(TkoorPL, strInd, strDiam, 0, dDlinKab);
            FPoly2d(TkoorPL, strInd, strDiam, 0, dDlinKab);
        }//
        public void FPoly(Point3dCollection TkoorPL, string strInd, string strDiam, double dDlin, double dDlinKab)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            // Append the point to the database
            using (tr1)
            {
                ed.WriteMessage("Началось построение \n");
                Polyline3d poly = new Polyline3d();
                //Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = 5;
                poly.Layer = "Kabeli";
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);
                foreach (Point3d pt in TkoorPL)
                {
                    PolylineVertex3d vex3d = new PolylineVertex3d(pt);
                    poly.AppendVertex(vex3d);
                    tr1.AddNewlyCreatedDBObject(vex3d, true);
                }
                poly.XData = new ResultBuffer(
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, strInd),
                new TypedValue(1000, strDiam),
                new TypedValue(1040, dDlin),
                new TypedValue(1040, dDlinKab)
                //new TypedValue(1040, Convert.ToDouble(strVISk)),
                );
                btr.Dispose();
                tr1.Commit();
            }
        }//
        public void FPoly2d(Point3dCollection TkoorPL, string strInd, string strDiam, double dDlin, double dDlinKab)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            // Append the point to the database
            using (tr1)
            {
                ed.WriteMessage("Началось построение " + strInd + " " + strDiam  + "\n");
                Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = ZvetLin;
                poly.Layer = "Kabeli";           
                int i = 0;
                foreach (Point3d pt in TkoorPL)
                {
                    poly.AddVertexAt(i, new Point2d(pt.X, pt.Y),0,0,0);
                    i = i + 1;
                }
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);
                poly.XData = new ResultBuffer(
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, strInd),
                new TypedValue(1000, strDiam)
                    //new TypedValue(1040, dDlin),
                    //new TypedValue(1040, dDlinKab)
                    //new TypedValue(1040, Convert.ToDouble(strVISk)),
                );
                btr.Dispose();
                tr1.Commit();
            }
        }//
        public void BlizhNOD(Point3d Toch, List<Noda> spNod, ref Noda blNOD) //поиск ближайшего нода
        {
            double MinDist = 9999999999999999;
            foreach (Noda TNOD in spNod) { if (Toch.DistanceTo(TNOD.Koord) < MinDist) { blNOD = TNOD; MinDist = Toch.DistanceTo(TNOD.Koord);}}
        }
        private void Form3_Load(object sender, EventArgs e)
        {   
            CreateLayer("Kabeli");
            CreateXRec();
            string strAdrRis = "";
            
            string[] Grupp = { "I", "II", "III", "IV", "V", "Силовые" };
            foreach (string Zvet in Grupp)
            {
                //strAdrRis = @"C:\МАРШРУТ\Трассы " + Zvet + " группы.txt";
                //if (System.IO.File.Exists(strAdrRis))
                string DOC = HCtenSlov("Трассы" + Zvet + "группы", "");
                if (DOC != "")
                this.comboBox1.Items.Add("Трассы " + Zvet + " группы");
            }

        }
        public void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            spNod0.Clear();
            strSist = this.comboBox1.SelectedItem.ToString().Replace(" ","");
            double[] KolZV = { 20, 2, 110, 6, 230, 4 };
            if (strSist == "ТрассыIгруппы") { ZvetLin = 1; }
            if (strSist == "ТрассыIIгруппы") { ZvetLin = 51; }
            if (strSist == "ТрассыIIIгруппы") { ZvetLin = 3; }
            if (strSist == "ТрассыIVгруппы") { ZvetLin = 231; }
            if (strSist == "ТрассыVгруппы") { ZvetLin = 20; }
            if (strSist == "Трассысиловыегруппы") { ZvetLin = 136; }
            //ShtenTXT(ref spNod, strSist);
            ShtenSLOV_NOD(ref spNod0, strSist);
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
        }//Создание сслоя
        public void CreateXRec() 
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            using (tr1)
            {
                RegAppTable regTable = (RegAppTable)tr1.GetObject(db.RegAppTableId, OpenMode.ForRead);
                if (!regTable.Has("LAUNCH01"))
                {
                    regTable.UpgradeOpen();
                    // Добавляем имя приложения, которое мы будем
                    // использовать в расширенных данных
                    RegAppTableRecord app =
                            new RegAppTableRecord();
                    app.Name = "LAUNCH01";
                    regTable.Add(app);
                    tr1.AddNewlyCreatedDBObject(app, true);
                }
              btr.Dispose();
              tr1.Commit();
            }
        
        }//Создание словоря
        public void ShtenTXT(ref List<Noda> spNod,string File)
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\МАРШРУТ\" + File + ".txt");
            foreach (string Strok in lines)
            {
                //Application.ShowAlertDialog(Strok);
                Noda Nod = new Noda();
                string[] NODm = Strok.Split(':');
                Nod.NomNod(NODm[0]);
                Nod.NomNOtk("-");
                string[] stKoor = NODm[3].TrimStart('(').TrimEnd(')').Split(',');
                double X = Convert.ToDouble(stKoor[0]);
                double Y = Convert.ToDouble(stKoor[1]);
                double Z = Convert.ToDouble(stKoor[2]);
                Point3d Koor = new Point3d(X, Y, Z);
                //Application.ShowAlertDialog(Koor.ToString());
                Nod.NKoor(Koor);
                Nod.NVes(9999999999999.0);
                Nod.NSpSmNod(NODm[4]);
                if (NODm.Length > 5)
                {
                    string[] stKoorMir = NODm[5].TrimStart('(').TrimEnd(')').Split(',');
                    double XMir = Convert.ToDouble(stKoor[0]);
                    double YMir = Convert.ToDouble(stKoor[1]);
                    double ZMir = Convert.ToDouble(stKoor[2]);
                    Point3d KoorMir = new Point3d(XMir, YMir, ZMir);
                    Nod.NKoor3D(KoorMir);
                    Nod.NomVin(NODm[6]);
                    if (NODm[6]=="") Nod.NomVin(NODm[1]);
                }
                spNod.Add(Nod);
            }
        }
        public void ShtenSLOV_NOD(ref List<Noda> spNod0, string File)
        {
            List<string> DOC = HCtenSlovNod(File);
            if (DOC.Count == 0) return;
            foreach (string Strok in DOC)
            {
                Noda Nod = new Noda();
                if (Strok != "")
                {
                    string[] NODm = Strok.Split(':');
                    Nod.NomHoz(NODm[1]);
                    Nod.NomNod(NODm[0]);
                    Nod.NomNOtk("-");
                    string[] stKoor = NODm[3].TrimStart('(').TrimEnd(')').Split(',');
                    double X = Convert.ToDouble(stKoor[0]);
                    double Y = Convert.ToDouble(stKoor[1]);
                    double Z = Convert.ToDouble(stKoor[2]);
                    Point3d Koor = new Point3d(X, Y, Z);
                    Nod.NKoor(Koor);
                    Nod.NVes(9999999999999.0);
                    Nod.NSpSmNod(NODm[4]);
                    if (NODm.Length > 5)
                    {
                        string[] stKoorMir = NODm[5].TrimStart('(').TrimEnd(')').Split(',');
                        double XMir = Convert.ToDouble(stKoor[0]);
                        double YMir = Convert.ToDouble(stKoor[1]);
                        double ZMir = Convert.ToDouble(stKoor[2]);
                        Point3d KoorMir = new Point3d(XMir, YMir, ZMir);
                        Nod.NKoor3D(KoorMir);
                        Nod.NomVin(NODm[6]);
                        if (NODm[6] == "") Nod.NomVin(NODm[1]);
                    }
                    spNod0.Add(Nod);
                }
            }
        }//чтение текстового файла с нодами
        public void SpVistVinDB(ref List<TPodk> SpVIN)
        {
            string Nazv_Vin = "", Sprav_vin = "";
            string Nazv_per = "", Sprav_per = "";
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            Point3d BP = new Point3d();
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Переходы"), 2);
            acTypValAr.SetValue(new TypedValue(8, "Выноски"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
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
                        //Circle ln = tr.GetObject(sobj, OpenMode.ForWrite) as Circle;
                        //TPodk TNOD = new TPodk();
                        BlockReference bref = tr.GetObject(sobj, OpenMode.ForRead) as BlockReference;
                        if (bref != null)
                        {
                            Nazv_Vin = "";
                            Sprav_vin = "";
                            //для переходов
                            Nazv_per = "";
                            Sprav_per = "";
                            //для плоскостей
                            y1 = 0;
                            x2 = 0;
                            y2 = 0;
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
                                    }
                                }
                            }
                            if (Nazv_Vin != "")
                            {
                                TPodk TNOD = new TPodk();
                                TNOD.NIND(Nazv_Vin);
                                TNOD.NKoor(BP);
                                SpVIN.Add(TNOD);
                            }
                            if (Nazv_per != "")
                            {
                                Point3d T1 = new Point3d(BP.X + x1, BP.Y + y1, 0);
                                Point3d T2 = new Point3d(BP.X + x2, BP.Y + y2, 0);
                                TPodk TNOD1 = new TPodk();
                                TNOD1.NIND(Nazv_per);
                                TNOD1.NKoor(T1);
                                SpVIN.Add(TNOD1);
                                TPodk TNOD2 = new TPodk();
                                TNOD2.NIND(Nazv_per);
                                TNOD2.NKoor(T2);
                                SpVIN.Add(TNOD2);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создание списка выносок динБлоками
        public void UdalSvasi(ref List<Noda> spNod, string TVin)
        {
            string UkorSpSmNod1 = "";
            string UkorSpSmNod2 = "";
            if (spNod.Exists(x => x.Vin == TVin) == false) return;
            Noda BlNod = spNod.Find(x => x.Vin== TVin);
            string[] SpSmNOD = BlNod.SpSmNod.Split(',');
                foreach (string NTSmNoda in SpSmNOD)
                {
                    string[] tNoda = NTSmNoda.Split('*');
                    Noda TSMNod = spNod.Find(x => x.Nom == tNoda[0]);
                        UkorSpSmNod1 = UkoroSpNod(BlNod.SpSmNod, tNoda[0]);
                        UkorSpSmNod2 = UkoroSpNod(TSMNod.SpSmNod, BlNod.Nom);
                        Noda NNOD1 = BlNod;
                        NNOD1.NSpSmNod(UkorSpSmNod1);
                        TSMNod.NSpSmNod(UkorSpSmNod2);
                        spNod.RemoveAll(x => x.Nom == BlNod.Nom);
                        spNod.RemoveAll(x => x.Nom == TSMNod.Nom);
                        spNod.Add(NNOD1);
                        spNod.Add(TSMNod);
                        return;
                }
        }//удаление связей между нодами
        public string UkoroSpNod(string SpNod, string Nod)
        {
            string UkorSP = "";
            string[] SpSmNOD = SpNod.Split(',');
            foreach (string tNod in SpSmNOD)
            {
                if (tNod.Split('*')[0] != Nod) UkorSP = UkorSP + "," + tNod;
            }
            return UkorSP.TrimStart(',');
        }//удаление связи у конкретного нода
        public void NodiVKTXTfail(List<Noda> spNod, string File)
        {
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\" + File + ".txt"))
            {
                foreach (Noda line in spNod)
                {
                    string stSpSMNod = "";
                    if (line.SpSmNod != null) { stSpSMNod = line.SpSmNod.TrimStart(','); }
                    file.WriteLine(line.Nom + ":" + line.Otkuda + ":" + line.Hoz + ":" + line.Koord.ToString() + ":" + stSpSMNod);
                }
            }
        }//функция записи в файл нод
        static string HCtenSlov(string Slov, string PoUmol)
        {
            string ZNACH = PoUmol;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForRead);
                            foreach (DBDictionaryEntry Pom in PomD)
                            {
                                Xrecord xRec = (Xrecord)tr.GetObject(Pom.Value, OpenMode.ForRead, false);
                                TypedValue[] rez = xRec.Data.AsArray();
                                ZNACH = rez[0].Value.ToString();
                            }
                        }
                    }

                }
            }
            return ZNACH;
        }//чтение словоря
        static List<string> HCtenSlovNod(string Slov)
        {
            List<string> ZNACH = new List<string>();
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForRead);
                            foreach (DBDictionaryEntry Pom in PomD)
                            {
                                Xrecord xRec = (Xrecord)tr.GetObject(Pom.Value, OpenMode.ForRead, false);
                                TypedValue[] rez = xRec.Data.AsArray();
                                foreach (TypedValue valSl in rez)
                                {
                                    ZNACH.Add(valSl.Value.ToString());
                                    //Application.ShowAlertDialog(ZNACH);
                                }
                            }
                        }
                    }

                }
            }
            return ZNACH;
        }//чтение словоря

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
