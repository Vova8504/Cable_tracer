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
#elif AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.EditorInput;
#endif

    public partial class Form2 : Form
    {
        public struct Noda
        {
            public string Nom, Otkuda, Hoz, Handl, SpSmNod;
            public double Ves;
            public double DlinP;
            public double VisT;
            public double Param;
            public Point3d Koord;
            public Point3d Koord3D;
            public void NomHandl(string i) { Handl = i; }
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
        };
        public struct Kriv 
        {
            public string Name;
            public List<Noda> SpNod;
            public void NomName(string i) { Name = i; }
            public void NSpNod(List<Noda> i) { SpNod = i; }
        };
        public Document doc = Application.DocumentManager.MdiActiveDocument;
        public Form2()
        {
            InitializeComponent();
        }

     private void button5_Click(object sender, EventArgs e)
        {
            string File = "";
            int i = -1;
            Document doc = Application.DocumentManager.MdiActiveDocument;     
            using (DocumentLock docLock = doc.LockDocument())
            {
                UdalKrug();
                double[] KolZV={20,2,110,6,230,4};
                string[] Grupp = { "I", "II", "III", "IV", "V", "Силовые" };
                foreach (double Zvet in KolZV)
                {
                    i = i + 1;
                    if (EstLiSist(Zvet))
                    {
                        List<Noda> spNod = new List<Noda>();
                        List<Noda> spPer = new List<Noda>();
                        List<Noda> spNodFin = new List<Noda>();
                        List<string> spLINI = new List<string>();
                        SozdSpNODPer(ref spPer, "INSERT");
                        SozdSpNODPer(ref spPer, "CIRCLE");
                        SozdSpNOD1(Zvet, ref spNod, ref spLINI);
                        DOP_spNOD(ref spNod, ref spLINI, 10);
                        SvazLin(ref spNod, ref spLINI, ref spNodFin);
                        foreach (Noda TnodEt in spPer) { spNodFin.Add(TnodEt); }
                        Skon(10, ref spNodFin, ref spLINI, ref spPer, Zvet);
                        NodiVKTXTfail(spNodFin, "Трассы " + Grupp[i] + " группы");
                        this.listBox1.Items.Add("Трассы " + Grupp[i] + " группы " + spNodFin.Count.ToString() + " узлов");
                    }
                }
            }
        }//обновить трассы
     public void UdalKrug()
      {
          Document doc = Application.DocumentManager.MdiActiveDocument;
          Database db = doc.Database;
          Editor ed = doc.Editor;
          ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
          ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
          TypedValue[] acTypValAr = new TypedValue[2];
          acTypValAr.SetValue(new TypedValue(0, "CIRCLE"), 0);
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
                      Circle ln = tr.GetObject(sobj, OpenMode.ForWrite) as Circle;
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
     public bool EstLiSist(double Zvet)
      {
          bool Y_N = false;
          Document doc = Application.DocumentManager.MdiActiveDocument;
          Database db = doc.Database;
          Editor ed = doc.Editor;
          TypedValue[] acTypValAr = new TypedValue[6];
          acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
          acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
          acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "ТРАССЫ"), 2);
          acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "ТРАССЫскрытые"), 3);
          acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
          acTypValAr.SetValue(new TypedValue(62, Zvet), 5);
          List<Kriv> SpSkKriv = new List<Kriv>();
          // создаем фильтр
          SelectionFilter filter = new SelectionFilter(acTypValAr);
          PromptSelectionResult selRes = ed.SelectAll(filter);
          if (selRes.Status == PromptStatus.OK)
              Y_N = true;
          else
              Y_N = false;
          return Y_N;
      }//проверка есть ли линии
     public void KRUG(Point3d Toch1,int Zvet)
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
                KrugPOZVn.Layer = "Kabeli";
                KrugPOZVn.ColorIndex = Zvet;
                btr.AppendEntity(KrugPOZVn);
                tr1.AddNewlyCreatedDBObject(KrugPOZVn, true);
                tr1.Commit();
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
     public void SozdSpNODPer(ref List<Noda> spPer,string TIP)
        {
            string strPalub = "-";
            string strBort = "-";
            string strShpan = "-";
            string strPom1 = "-";
            string strPom2 = "-";
            string strNazv = "-";
            double douDlin = 0;
            string strNNod = "-";
            string SpDoPNOD = "";
            Point3d BazT=new Point3d();
            int Schet = 0;
            int i = 0;
            double DDist = Convert.ToDouble(this.textBox1.Text);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            TypedValue[] acTypValAr1 = new TypedValue[6];
            acTypValAr1.SetValue(new TypedValue(0, TIP), 0);
            acTypValAr1.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr1.SetValue(new TypedValue(8, "ТРАССЫ"), 2);
            acTypValAr1.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 3);
            acTypValAr1.SetValue(new TypedValue(8, "Переходы"), 4);
            acTypValAr1.SetValue(new TypedValue(-4, "or>"), 5);
            // создаем фильтр
            SelectionFilter filter1 = new SelectionFilter(acTypValAr1);
            PromptSelectionResult selRes = ed.SelectAll(filter1);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
                    acObjIdColl1 = new ObjectIdCollection(acSSet.GetObjectIds());
                    foreach (ObjectId sobj in acObjIdColl1)
                    {
                        if (TIP == "INSERT")
                        {
                            BlockReference ln = tr.GetObject(sobj, OpenMode.ForWrite) as BlockReference;
                            if (ln != null)
                            {
                                i = i + 1;
                                BazT = ln.Position;
                                strNNod = ln.Handle.ToString();
                                ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                                if (buffer != null)
                                {
                                    Noda Nnoda = new Noda();
                                    Schet = 0;
                                    foreach (TypedValue value in buffer)
                                    {
                                        if (Schet == 1) { strNazv = value.Value.ToString(); }
                                        if (Schet == 2) { douDlin = Convert.ToDouble(value.Value.ToString()); }
                                        if (Schet == 3) { strPom1 = value.Value.ToString(); }
                                        if (Schet == 4) { strPom2 = value.Value.ToString(); }
                                        if (Schet == 5) { strShpan = value.Value.ToString(); }
                                        if (Schet == 6) { strBort = value.Value.ToString(); }
                                        if (Schet == 7) { strPalub = value.Value.ToString(); }
                                        Schet = Schet + 1;
                                    }
                                    Nnoda.NomNod("ПереходБлок" + i);
                                    Nnoda.NomHoz(strNazv);
                                    Nnoda.NParam(i);
                                    Nnoda.NKoor(BazT);
                                    Nnoda.NDlinP(douDlin);
                                    if (douDlin > 0 & douDlin < 999999999)
                                    {
                                        if (spPer.Exists(x => x.Hoz == Nnoda.Hoz))
                                        {
                                            int indTnodEt = spPer.FindIndex(x => x.Hoz == Nnoda.Hoz);
                                            Noda sNod = spPer.Find(x => x.Hoz == Nnoda.Hoz);
                                            SpDoPNOD = Nnoda.Nom + "*" + douDlin;
                                            DoboV_lSM_NOD(ref spPer, SpDoPNOD, sNod, indTnodEt);
                                            SpDoPNOD = sNod.Nom + "*" + douDlin;
                                            Nnoda.NSpSmNod(SpDoPNOD);
                                            spPer.Add(Nnoda);
                                        }
                                        else { spPer.Add(Nnoda); }
                                    }
                                }
                            }
                        }
                        else
                        {
                         Circle ln = tr.GetObject(sobj, OpenMode.ForWrite) as Circle;
                         if (ln != null)
                         {
                             i = i + 1;
                             BazT = ln.Center;
                             strNNod = ln.Handle.ToString();
                             ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                             if (buffer != null)
                             {
                                 Noda Nnoda = new Noda();
                                 Schet = 0;
                                 foreach (TypedValue value in buffer)
                                 {
                                     if (Schet == 1) { strNazv = value.Value.ToString(); }
                                     if (Schet == 2) { douDlin = Convert.ToDouble(value.Value.ToString()); }
                                     if (Schet == 3) { strPom1 = value.Value.ToString(); }
                                     if (Schet == 4) { strPom2 = value.Value.ToString(); }
                                     if (Schet == 5) { strShpan = value.Value.ToString(); }
                                     if (Schet == 6) { strBort = value.Value.ToString(); }
                                     if (Schet == 7) { strPalub = value.Value.ToString(); }
                                     Schet = Schet + 1;
                                 }
                                 Nnoda.NomNod("ПереходКруг" + i);
                                 Nnoda.NomHoz(strNazv);
                                 Nnoda.NParam(i);
                                 Nnoda.NKoor(BazT);
                                 Nnoda.NDlinP(douDlin);
                                 if (douDlin > 0 & douDlin < 999999999) 
                                 {
                                     if (spPer.Exists(x => x.Hoz == Nnoda.Hoz))
                                     {
                                         int indTnodEt = spPer.FindIndex(x => x.Hoz == Nnoda.Hoz);
                                         Noda sNod = spPer.Find(x => x.Hoz == Nnoda.Hoz);
                                         SpDoPNOD = Nnoda.Nom + "*" + douDlin;
                                         DoboV_lSM_NOD(ref spPer, SpDoPNOD, sNod, indTnodEt);
                                         SpDoPNOD = sNod.Nom + "*" + douDlin;
                                         Nnoda.NSpSmNod(SpDoPNOD);
                                         spPer.Add(Nnoda);
                                     }
                                     else { spPer.Add(Nnoda); } 
                                 }
                             }
                         }
                        }
                    }
                    ed.WriteMessage("Количество точек переходов " + spPer.Count);
                    tr.Commit();
                }
            }
        }//создание списка переходов
     public void SozdSpNOD1(double Zvet, ref List<Noda> spNod, ref List<string> spLINI) 
        {
            string smNOD1 = "";
            string smNOD2 = "";
            string spSmNod = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[6];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "ТРАССЫ"), 2);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "ТРАССЫскрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            acTypValAr.SetValue(new TypedValue(62, Zvet), 5);
            List<Kriv> SpSkKriv = new List<Kriv>();
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            CreateLayer("ТРАССЫскрытые");
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    foreach (SelectedObject sobj in acSSet)
                    {
                        Polyline ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Polyline;
                        spLINI.Add(ln.Handle.ToString());
                        var KolV = ln.NumberOfVertices;
                        for (int i = 0; i <= KolV - 1; i++)
                        {
                            smNOD1 = "";
                            smNOD2 = "";
                            spSmNod = "";
                            Noda TNOD = new Noda();
                            TNOD.NomNod(ln.Handle.ToString() + "-" + Convert.ToString(i));
                            TNOD.NomHoz(ln.Handle.ToString());
                            TNOD.NKoor(ln.GetPointAtParameter(i));
                            TNOD.NParam(i);
                            TNOD.NSpSmNod(spSmNod);
                            spNod.Add(TNOD);
                        }
                    }
                    tr.Commit();
                }
            }
        }//создания списка нод
     public void DOP_spNOD(ref List<Noda> spNod, ref List<string> spLINI, double RAD)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            int i1 = 0;
            double DistA_B=0;
            double DistA_C = 0;
            double DistC_B = 0;
            double DELT = 0;
            List<Noda> spNNod = new List<Noda>();
            foreach (string Lin in spLINI)
            {
                List<Noda> spNodEtal = spNod.FindAll(x => x.Hoz == Lin);
                Noda Nod1 = spNodEtal[0];
                int ind1NOD = spNod.FindIndex(x => x.Nom == Nod1.Nom);
                Noda Nod2 = spNodEtal[spNodEtal.Count - 1];
                int ind2NOD = spNod.FindIndex(x => x.Nom == Nod2.Nom);
                foreach (string Lin1 in spLINI)
                {
                    if (Lin != Lin1) 
                    {
                        List<Noda> spNodSR = spNod.FindAll(x => x.Hoz == Lin1);
                        for (int i = 0; i < spNodSR.Count-1;i++ )
                        {
                            Noda TOtr1=spNodSR[i];
                            Noda TOtr2 = spNodSR[i+1];
                            DistA_B = TOtr1.Koord.DistanceTo(TOtr2.Koord);
                            DistA_C = TOtr1.Koord.DistanceTo(Nod1.Koord);
                            DistC_B = Nod1.Koord.DistanceTo(TOtr2.Koord);
                            DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                            if ((DELT < 0.5) & (DistA_C > RAD) & (DistC_B > RAD)) 
                                {
                                    Noda NNoda = new Noda();
                                    NNoda.NomNod(TOtr1.Hoz + "-" + (i + (DistA_C / DistA_B)).ToString());
                                    NNoda.NomHoz(TOtr1.Hoz);
                                    NNoda.NKoor(Nod1.Koord);
                                    NNoda.NParam(i + (DistA_C / DistA_B));
                                    spNNod.Add(NNoda);
                                }
                             DistA_C = TOtr1.Koord.DistanceTo(Nod2.Koord);
                             DistC_B = Nod2.Koord.DistanceTo(TOtr2.Koord);
                             DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                             if ((DELT < 0.5) & (DistA_C > RAD) & (DistC_B > RAD)) 
                                {
                                    Noda NNoda = new Noda();
                                    NNoda.NomNod(TOtr1.Hoz + "-" + (i + (DistA_C / DistA_B)).ToString());
                                    NNoda.NomHoz(TOtr1.Hoz);
                                    NNoda.NKoor(Nod2.Koord);
                                    NNoda.NParam(i + (DistA_C / DistA_B));
                                    spNNod.Add(NNoda);
                                }
                        }
                    }
                }
                i1 = i1 + 1;
                ed.WriteMessage("Обработано " + i1 + " из " + spLINI.Count + " линий \n");
            }
            //this.progressBar1.Value = 0;
            foreach (Noda Lin1 in spNNod) { spNod.Add(Lin1);}
        }//соеденение концов линий с отрезками других линий
     public void SvazLin(ref List<Noda> spNod, ref List<string> spLINI, ref List<Noda> spNodFin) 
     {
         string smNOD1 = "";
         string smNOD2 = "";
         string spSmNod = "";
         foreach (string Lin in spLINI) 
         {
           List<Noda> spNodEtal = spNod.FindAll(x => x.Hoz == Lin);
           spNodEtal.Sort(delegate(Noda x, Noda y){ return x.Param.CompareTo(y.Param);});
           var KolV = spNodEtal.Count;
           for (int i = 0; i <= KolV - 1; i++)
           {
               smNOD1 = "";
               smNOD2 = "";
               spSmNod = "";
               if (i > 0) { smNOD1 = spNodEtal[i].Hoz + "-" + spNodEtal[i-1].Param + "*" + spNodEtal[i].Koord.DistanceTo(spNodEtal[i - 1].Koord).ToString(); }
               if (i < KolV - 1) { smNOD2 = spNodEtal[i].Hoz + "-" + spNodEtal[i + 1].Param + "*" + spNodEtal[i].Koord.DistanceTo(spNodEtal[i + 1].Koord).ToString(); }
               if (smNOD1 != "") { spSmNod = smNOD1; }
               if (smNOD2 != "") { spSmNod = spSmNod + "," + smNOD2; }
               spNodEtal[i].NSpSmNod(spSmNod);
               Noda NNod = spNodEtal[i];
               NNod.NSpSmNod(spSmNod);
               spNodFin.Add(NNod);
           }         
         }
     }//связывание точек линии
     public void Skon(double RAD, ref List<Noda> spNod, ref List<string> spLINI, ref List<Noda> spPer,double Zvet) 
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            int i1 = 0;
            string SpDoPNOD = "";
            foreach (string Lin in spLINI) 
            {
                SpDoPNOD = "";
                List<Noda> spNodEtal = spNod.FindAll(x => x.Hoz == Lin);
                Noda Nod1 = spNodEtal[0];
                int ind1NOD = spNod.FindIndex(x => x.Nom == Nod1.Nom);
                Noda Nod2 = spNodEtal[spNodEtal.Count-1];
                int ind2NOD = spNod.FindIndex(x => x.Nom == Nod2.Nom);
                List<Noda> spNodOst = spNod.FindAll(x => x.Hoz != Lin);
                    foreach (Noda TnodEt in spNodOst) 
                    {
                        if (Nod1.Koord.DistanceTo(TnodEt.Koord) <= RAD) 
                        {
                           int indTnodEt = spNod.FindIndex(x => x.Nom == TnodEt.Nom);
                           SpDoPNOD = TnodEt.Nom + "*" + Nod1.Koord.DistanceTo(TnodEt.Koord).ToString();
                           DoboV_lSM_NOD(ref spNod, SpDoPNOD, Nod1, ind1NOD);
                           SpDoPNOD = Nod1.Nom + "*" + Nod1.Koord.DistanceTo(TnodEt.Koord).ToString();
                           DoboV_lSM_NOD(ref spNod, SpDoPNOD, TnodEt, indTnodEt);
                           KRUG(Nod1.Koord, Convert.ToInt16(Zvet));
                        }
                        if (Nod2.Koord.DistanceTo(TnodEt.Koord) <= RAD) 
                        {
                            int indTnodEt = spNod.FindIndex(x => x.Nom == TnodEt.Nom);
                            SpDoPNOD = TnodEt.Nom + "*" + Nod2.Koord.DistanceTo(TnodEt.Koord).ToString();
                            DoboV_lSM_NOD(ref spNod, SpDoPNOD, Nod2, ind2NOD);
                            SpDoPNOD = Nod2.Nom + "*" + Nod2.Koord.DistanceTo(TnodEt.Koord).ToString();
                            DoboV_lSM_NOD(ref spNod, SpDoPNOD, TnodEt, indTnodEt);
                            KRUG(Nod2.Koord, Convert.ToInt16(Zvet));
                        }
                    }
                    foreach (Noda TnodEt in spPer)
                    {
                        if (Nod1.Koord.DistanceTo(TnodEt.Koord) <= RAD)
                        {
                            int indTnodEt = spNod.FindIndex(x => x.Nom == TnodEt.Nom);
                            SpDoPNOD = TnodEt.Nom + "*" + Nod1.Koord.DistanceTo(TnodEt.Koord).ToString();
                            DoboV_lSM_NOD(ref spNod, SpDoPNOD, Nod1, ind1NOD);
                            SpDoPNOD = Nod1.Nom + "*" + Nod1.Koord.DistanceTo(TnodEt.Koord).ToString();
                            DoboV_lSM_NOD(ref spNod, SpDoPNOD, TnodEt, indTnodEt);
                            KRUG(Nod1.Koord, Convert.ToInt16(Zvet));
                        }
                        if (Nod2.Koord.DistanceTo(TnodEt.Koord) <= RAD)
                        {
                            int indTnodEt = spNod.FindIndex(x => x.Nom == TnodEt.Nom);
                            SpDoPNOD = TnodEt.Nom + "*" + Nod2.Koord.DistanceTo(TnodEt.Koord).ToString();
                            DoboV_lSM_NOD(ref spNod, SpDoPNOD, Nod2, ind2NOD);
                            SpDoPNOD = Nod2.Nom + "*" + Nod2.Koord.DistanceTo(TnodEt.Koord).ToString();
                            DoboV_lSM_NOD(ref spNod, SpDoPNOD, TnodEt, indTnodEt);
                            KRUG(Nod2.Koord, Convert.ToInt16(Zvet));
                        }
                    }
                i1 = i1 + 1;
                ed.WriteMessage("Соедино " + i1 + " из " + spLINI.Count + " линий \n");
             }
        }//соеденение концов линий с вершинами линий
     public void DoboV_lSM_NOD(ref List<Noda> spNod, string stSMNOD, Noda Nod, int ind1NOD)
    {
        Noda StarNODA1_1 = spNod.Find(x => x.Nom == Nod.Nom);
        if (StarNODA1_1.SpSmNod != null)
        {
            //Application.ShowAlertDialog(StarNODA1_1.SpSmNod + ":" + stSMNOD + ":" + StarNODA1_1.SpSmNod.Contains(stSMNOD).ToString());
            if (StarNODA1_1.SpSmNod.Contains(stSMNOD)==false)
            {
                string StSPSmNod = StarNODA1_1.SpSmNod + "," + stSMNOD;
                StarNODA1_1.NSpSmNod(StSPSmNod);
                spNod[ind1NOD] = StarNODA1_1;
            }
        }
        else
        {
          string StSPSmNod = StarNODA1_1.SpSmNod + "," + stSMNOD;
          StarNODA1_1.NSpSmNod(StSPSmNod);
          spNod[ind1NOD] = StarNODA1_1;
        }
    }//добовление смежных нод для линии
     public void DoboV_lSM_NOD_per(ref List<Noda> spPer, string stSMNOD, Noda Nod, int ind1NOD)
     {
         Noda StarNODA1_1 = spPer.Find(x => x.Nom == Nod.Nom);
         if (StarNODA1_1.SpSmNod != null)
         {
             //Application.ShowAlertDialog(StarNODA1_1.SpSmNod + ":" + stSMNOD + ":" + StarNODA1_1.SpSmNod.Contains(stSMNOD).ToString());
             if (StarNODA1_1.SpSmNod.Contains(stSMNOD) == false)
             {
                 string StSPSmNod = StarNODA1_1.SpSmNod + "," + stSMNOD;
                 StarNODA1_1.NSpSmNod(StSPSmNod);
                 spPer[ind1NOD] = StarNODA1_1;
             }
         }
         else
         {
             string StSPSmNod = StarNODA1_1.SpSmNod + "," + stSMNOD;
             StarNODA1_1.NSpSmNod(StSPSmNod);
             spPer[ind1NOD] = StarNODA1_1;
         }
     }//добовление смежных нод для перехода
     public void NodiVKTXTfail(List<Noda> spNod,string File)
        {
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\" + File + ".txt"))
            {
                foreach (Noda line in spNod)
                {
                    string stSpSMNod = "";
                    if (line.SpSmNod != null) { stSpSMNod = line.SpSmNod.TrimStart(','); }
                    file.WriteLine(line.Nom + ":" + line.Hoz + ":" + line.Param + ":" + line.Koord.ToString() + ":" + stSpSMNod);
                }
            }
        }

     private void Form2_Load(object sender, EventArgs e)
     {

     }//функция записи в файл нод
    }
}
