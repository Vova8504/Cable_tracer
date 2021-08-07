using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;
using System.Data.SQLite;

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


    public partial class Form5 : Form
    {
        public struct Noda
        {
            public string Nom, Otkuda, Hoz,Vin;
            public string SpSmNod;
            public double Ves;
            public double DlinP;
            public double VisT;
            public double Param;
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
            public void NParam(double i) { Param = i; }
            public void NKoor(Point3d i) { Koord = i; }
            public void NKoor3D(Point3d i) { Koord3D = i; }  
        }
        public struct DUGA
        {
            public string NOD1, NOD2;
            public double Ves;
            public void NomNOD1(string i) { NOD1 = i; }
            public void NomNOD2(string i) { NOD2 = i; }
            public void NVes(double i) { Ves = i; }
        }
        public struct TPodk
        {
            public string IND, Naim;
            public double VisT;
            public Point3d Koord, Koord1, Povorot;
            public string KoordMod;
            public int Vstr;
            public string Sist;
            public string BlNOD;
            public string Shem;
            public string Pom;
            public void NIND(string i) { IND = i; }
            public void NSist(string i) { Sist = i; }
            public void NBlNOD(string i) { BlNOD = i; }
            public void NVisT(double i) { VisT = i; }
            public void NKoor(Point3d i) { Koord = i; }
            public void NKoor1(Point3d i) { Koord1 = i; }
            public void NPovorot(Point3d i) { Povorot = i; }
            public void NKoorMod(string i) { KoordMod = i; }
            public void NVstr(int i) { Vstr = i; }
            public void NNaim(string i) { Naim = i; }
            public void NShem(string i) { Shem = i; }
            public void NPom(string i) { Pom = i; }
        }
        public struct PodvKab 
        {
            public string IND;
            public double DIam;
            public string INDobor;
            public string Massa;
            public Point3d KoordObor;
            public void NIND(string i) {IND = i;}
            public void NDIam(double i) {DIam = i;}
            public void NMassa(string i) { Massa = i; }
            public void NINDobor(string i) {INDobor = i;}
            public void NKoordObor(Point3d i) { KoordObor = i; }
        }
        public struct ZaPis
        {
            public TPodk INDotk;
            public string Sist;
            public string BlokNOD;
            public Point3d KoordObor;
            public List<PodvKab> SpVixKab;
            public void NINDotk(TPodk i) { INDotk = i; }
            public void NSist(string i) { Sist = i;}
            public void NSpVixKab(List<PodvKab> i) { SpVixKab = i; }
            public void NBlokNOD(string i) { BlokNOD = i; }
            public void NKoordObor(Point3d i) { KoordObor = i; }
        }
        public struct Kab
        {
            public string IND;
            public double DIam;
            public string Massa;
            public double DIamKor;
            public string INDOtk;
            public string INDKud;
            public string Sist;
            public string BlokNOD;
            public string KoorOt;
            public string KoorKud;
            public string spVin;
            public string Shem;
            public string PomOt;
            public string PomKud;
            public void NIND(string i) {IND = i;}
            public void NDIam(double i) {DIam = i;}
            public void NMassa(string i) { Massa = i; }
            public void NDIamKor(double i) { DIamKor = i; }
            public void NINDOtk(string i) {INDOtk = i;}
            public void NINDKud(string i) {INDKud = i; }
            public void NSist(string i) { Sist = i; }
            public void NBlokNOD(string i) { BlokNOD = i; }
            public void NKoorOt(string i) { KoorOt = i; }
            public void NKoorKud(string i) { KoorKud = i; }
            public void NspVin(string i) { spVin = i; }
            public void NShem(string i) { Shem = i; }
            public void NPomOt(string i) { PomOt = i; }
            public void NPomKud(string i) { PomKud = i; }
        }
        public struct Kriv 
        {
            public string Name;
            public List<Noda> SpNod;
            public void NomName(string i) { Name = i; }
            public void NSpNod(List<Noda> i) { SpNod = i; }
        }
        public struct VINOS
        {
            public string Nazv, Spravka, PolnNazv, SUF;
            public Point3d Koord;
            public void NNazv(string i) { Nazv = i; }
            public void NSpravka(string i) { Spravka = i; }
            public void NPolnNazv(string i) { PolnNazv = i; }
            public void NSUF(string i) { SUF = i; }
            public void NKoor(Point3d i) { Koord = i; }
        };
        public struct PEREXOD
        {
            public string Nazv, Spravka, PolnNazv, SUF;
            public double Dlin;
            public Point3d Koord1, Koord2;
            public void NNazv(string i) { Nazv = i; }
            public void NSpravka(string i) { Spravka = i; }
            public void NPolnNazv(string i) { PolnNazv = i; }
            public void NSUF(string i) { SUF = i; }
            public void NDlin(double i) { Dlin = i; }
            public void NKoor1(Point3d i) { Koord1 = i; }
            public void NKoor2(Point3d i) { Koord2 = i; }
        };
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
        public struct DiamVsehc
        {
            public string ind;
            public int row, column, ID;
            public double rad;
            public double radK;
            public double Massa;
            public Point3d Zentr;
            public void NZentr(Point3d i) { Zentr = i; }
            public void Nrow(int i) { row = i; }
            public void Ncolumn(int i) { column = i; }
            public void NID(int i) { ID = i; }
            public void Nrad(double i) { rad = i; }
            public void NradK(double i) { radK = i; }
            public void Nind(string i) { ind = i; }
            public void NMassa(double i) { Massa = i; }

        }
        public struct Zag 
        {
            public string Zagal;
            public int Graf;
            public void NomZagal(string i) { Zagal = i; }
            public void NomGraf(int i) { Graf = i; }
        }
        public struct PERprim 
        {
            public string Nazv,Sprav1,Sprav2;
            public ObjectId Prim1, Prim2;
            public void NNazv(string i) { Nazv = i; }
            public void NSprav1(string i) { Sprav1 = i; }
            public void NSprav2(string i) { Sprav2 = i; }
            public void NPrim1(ObjectId i) { Prim1 = i; }
            public void NPrim2(ObjectId i) { Prim2 = i; }

        }

        public string Mas;
        public List<VINOS> Vinoski = new List<VINOS>();
        public List<PEREXOD> PEREX = new List<PEREXOD>();
        public List<VINOS> PovtVinoski = new List<VINOS>();
        public List<PEREXOD> PovtPEREX = new List<PEREXOD>();
        public List<PLOS> SpPLOS = new List<PLOS>();//Список плоскостей 
        public List<TPodk> Tochki = new List<TPodk>();//Список точек
        public List<TPodk> TochkiSort = new List<TPodk>();//Список точек для гиртвювера без повторов
        public List<TPodk> TochkiRazv = new List<TPodk>();//список точек для разводки
        public List<ZaPis> SgrupVK = new List<ZaPis>();
        public List<Kab> VK = new List<Kab>();
        public List<string> Vkrazv = new List<string>();
        public List<string> ITOGI = new List<string>();
        public string Rezult;
        public List<Noda> spNod = new List<Noda>();
        public List<Noda>[] MSpNOD = new List<Noda>[6];

        public Document doc = Application.DocumentManager.MdiActiveDocument;
        public static string DBPath;
        public static SQLiteConnection connection;
        public static SQLiteCommand command;
        public DataSet DS1 = new DataSet();
        public System.Data.DataTable DT1 = new System.Data.DataTable();

        //public Document doc = Application.DocumentManager.MdiActiveDocument;
        //public Database db = doc.Database;
        public Form5()
        {
            InitializeComponent();
        }
        public void Form5_Load(object sender, EventArgs e)
        {
            Database db = doc.Database;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                RegAppTable regTable = (RegAppTable)Tx.GetObject(db.RegAppTableId, OpenMode.ForWrite);
                if (!regTable.Has("LAUNCH01"))
                {
                    regTable.UpgradeOpen();
                    // Добавляем имя приложения, которое мы будем
                    // использовать в расширенных данных
                    RegAppTableRecord app =
                            new RegAppTableRecord();
                    app.Name = "LAUNCH01";
                    regTable.Add(app);
                    Tx.AddNewlyCreatedDBObject(app, true);
                }
             Tx.Commit();
            }
            CreateLayer("Kabeli");
            CreateLayer("KabeliError");
            CreateLayer("ТРАССЫ");
            CreateLayer("ТРАССЫскрытые");
            CreateLayer("Выноски");
            CreateLayer("ВыноскиНаПолеЧертежа");
            List<String> TochkiSTR = new List<String>();
            List<String> VKisSlov = HCtenSlovNod("VK");
            SozdSpKab(ref Vkrazv);
            ZagrDannNODisSLOV();
            DGVsist();
            Tochki.Clear();
            SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS);
            SpVistTP(ref Tochki);
            SpVistTP_BL(ref Tochki);
            string SpPom = HCtenSlov("spPOM", "");
            string Proj = HCtenSlov("PROG", "");
            this.textBox18.Text = SpPom;
            this.textBox19.Text = Proj;
            obnTP(VKisSlov, SpPom);
            foreach (string Strok in Vkrazv)
            {
                if (VK.FindIndex(x => Strok.Split('*')[0] == x.IND)==-1) 
                {
                    Kab nkab = new Kab();
                    nkab.NIND(Strok.Split('*')[0]);
                    VK.Add(nkab);
                }
            }
            ZapTablT(TochkiSort);
            ZapTablK(VK);
            int i = -1;
            foreach (VINOS TV in Vinoski) 
            {
                i = i + 1;
                PovtVinoski = Vinoski.FindAll(x => x.Nazv == TV.Nazv & x.SUF == TV.SUF);
                PovtPEREX = PEREX.FindAll(x => x.Nazv == TV.Nazv & x.SUF == TV.SUF);
                this.dataGridView3.Rows.Add(TV.Nazv, TV.PolnNazv, TV.Spravka, TV.SUF);
                if (PovtVinoski.Count > 1 | PovtPEREX.Count>0)
                {
                    this.dataGridView3.Rows[i].Cells[0].Style.BackColor = Color.LightPink;
                    this.dataGridView3.Rows[i].Cells[1].Style.BackColor = Color.LightPink;
                    this.dataGridView3.Rows[i].Cells[2].Style.BackColor = Color.LightPink;
                }
            }
            i = -1;
            foreach (PEREXOD TV in PEREX) 
            {
                i = i + 1;
                PovtPEREX = PEREX.FindAll(x => x.Nazv == TV.Nazv & x.SUF==TV.SUF);
                PovtVinoski = Vinoski.FindAll(x => x.Nazv == TV.Nazv & x.SUF == TV.SUF);
                this.dataGridView4.Rows.Add(TV.Nazv, TV.PolnNazv, TV.Dlin.ToString(), TV.SUF);
                if (PovtPEREX.Count > 1 | PovtVinoski.Count > 0)
                {
                    this.dataGridView4.Rows[i].Cells[0].Style.BackColor = Color.LightPink;
                    this.dataGridView4.Rows[i].Cells[1].Style.BackColor = Color.LightPink;
                    this.dataGridView4.Rows[i].Cells[2].Style.BackColor = Color.LightPink;
                }
            }
            foreach (PLOS TV in SpPLOS) { this.dataGridView5.Rows.Add(TV.Vid, TV.List, TV.Osi, TV.Msk.ToString()); }
            string DOC = HCtenSlov("DOC", "");
            if (DOC != "") 
            {
                string[] DOCm = DOC.Split('&');
                this.textBox5.Text = DOCm[0];
                this.textBox6.Text = DOCm[1];
                this.textBox7.Text = DOCm[2];
                this.textBox8.Text = DOCm[3];
                this.textBox9.Text = DOCm[4];
                this.textBox10.Text = DOCm[5];
                this.textBox11.Text = DOCm[6];
                this.textBox12.Text = DOCm[7];
            }
            string File= HCtenSlov("Istok", "-");
            this.label33.Text = File;
            if (System.IO.File.Exists(@File)==false) 
            {
                this.button27.Enabled=false;
                this.button23.Enabled=false;
                //this.button35.Enabled = false;
            }
            string FileVin = HCtenSlov("ADRblVin", "");
            if(FileVin!="") this.label34.Text = FileVin;
            string FilePer = HCtenSlov("ADRblPer", "");
            if (FilePer != "") this.label35.Text = FilePer;
            string FilePlos = HCtenSlov("ADRblPlos", "");
            if (FilePlos != "") this.label36.Text = FilePlos;
            string FileLn = HCtenSlov("ADRblLn", "");
            if (FileLn != "") this.label21.Text = FileLn;
            string FileL1 = HCtenSlov("ADRblL1", "");
            if (FileL1 != "") this.label22.Text = FileL1;
            string FileLreg = HCtenSlov("ADRblLreg", "");
            if (FileLreg != "") this.label23.Text = FileLreg;
            string FileDB = HCtenSlov("AdrBD", "");
            if (FileDB != "") this.label53.Text = FileDB;
            string SUF = HCtenSlov("SUF", "");
            if (SUF != "") this.textBox37.Text = SUF;
            string KolST = HCtenSlov("KolST", "");
            if (KolST != "") this.textBox39.Text = KolST;
            NasrExcel();
            SOZDlov("MAS#");
            Mas = HCtenSlov("MAS#", "1");
            if (Mas == "0.5") { this.radioButton1.Checked = true; }
            if (Mas == "1") { this.radioButton2.Checked = true; }
            if (Mas == "1.25") { this.radioButton3.Checked = true; }
            if (Mas == "2.5") { this.radioButton4.Checked = true; }
        }//Загрузка окна
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.checkBox1.Checked == true) { ObnTr(); this.dataGridView8.Rows.Clear(); ZagrDannNODisSLOV(); }
            this.Hide();
            string[] spVin;
            string T1 = "";
            string T2 = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument()) 
            { 
            List<TPodk> TochkiVzv = new List<TPodk>();
            List<Kab> VKsMaxINDotk = new List<Kab>();
            List<Kab> VKsMaxINDkud = new List<Kab>();
            List<Kab> VKdopol = VK.FindAll(x=> x.spVin!=null & x.spVin != "");
                foreach (Kab TKab in VKdopol) 
                {
                    spVin = TKab.spVin.Split(',');
                    T1 = TKab.INDOtk;
                    VK.Remove(TKab);
                    foreach (string tVin in spVin) 
                    {
                        if (tVin != "" & tVin != TKab.INDOtk & tVin != TKab.INDKud) 
                        {
                            T2 = tVin;
                            Kab NKab =new Kab();
                            NKab.NIND(TKab.IND);
                            NKab.NDIam(TKab.DIam);
                            NKab.NINDOtk(T1);
                            NKab.NINDKud(T2);
                            NKab.NspVin(TKab.spVin);
                            NKab.NSist(TKab.Sist);
                            NKab.NMassa(TKab.Massa);
                            NKab.NBlokNOD(TKab.BlokNOD);
                            VK.Add(NKab);
                            T1 = T2;
                        }
                    }
                    T2 = TKab.INDKud;
                    Kab NKabp = new Kab();
                    NKabp.NIND(TKab.IND);
                    NKabp.NDIam(TKab.DIam);
                    NKabp.NINDOtk(T1);
                    NKabp.NINDKud(T2);
                    NKabp.NspVin(TKab.spVin);
                    NKabp.NSist(TKab.Sist);
                    NKabp.NBlokNOD(TKab.BlokNOD);
                    NKabp.NMassa(TKab.Massa);
                    VK.Add(NKabp);
                }
            SgrupVK.Clear();
            int Vstr;
            SpVistTP(ref Tochki);
            SpVistTP_BL(ref Tochki);
            foreach (TPodk TPodk in TochkiRazv) 
            {
                Vstr = 0;
                foreach (Kab iKab in VK) {if ((iKab.INDOtk == TPodk.IND | iKab.INDKud == TPodk.IND) & iKab.Sist == TPodk.Sist & iKab.BlokNOD == TPodk.BlNOD) { Vstr = Vstr + 1; }}
                TPodk.NVstr(Vstr);
                TochkiVzv.Add(TPodk);
            }
            TochkiVzv.Sort(delegate(TPodk x,TPodk y){return y.Vstr.CompareTo(x.Vstr);});
            foreach (TPodk TPodkk in TochkiVzv) 
            {
                ZaPis NZap = new ZaPis();
                PodvKab tPodvKab = new PodvKab();
                TPodk TPodkKud=new TPodk(); 
                NZap.NSist(TPodkk.Sist);
                NZap.NBlokNOD(TPodkk.BlNOD);
                VKsMaxINDotk = VK.FindAll(x => (x.INDOtk == TPodkk.IND) & x.Sist == TPodkk.Sist & x.BlokNOD == TPodkk.BlNOD);
                VKsMaxINDkud = VK.FindAll(x => (x.INDKud == TPodkk.IND) & x.Sist == TPodkk.Sist & x.BlokNOD == TPodkk.BlNOD);
                if ((VKsMaxINDotk.Count > 0 | VKsMaxINDkud.Count > 0) & Tochki.Exists(x => x.IND == TPodkk.IND))
                {
                    NZap.NINDotk(TPodkk);
                    List<PodvKab> SpPodvKAB = new List<PodvKab>();
                    foreach (Kab tKab in VKsMaxINDotk) 
                    {
                        if (Tochki.Exists(x => x.IND == tKab.INDKud)) 
                        { 
                            TPodkKud = Tochki.Find(x => x.IND == tKab.INDKud);
                            tPodvKab.NIND(tKab.IND);
                            tPodvKab.NINDobor(tKab.INDKud);
                            tPodvKab.NDIam(tKab.DIam);
                            tPodvKab.NKoordObor(TPodkKud.Koord);
                            tPodvKab.NMassa(tKab.Massa);
                            if (Vkrazv.Exists(x => x.Split('*')[0] == tKab.IND) == false)SpPodvKAB.Add(tPodvKab);
                        }
                        VK.Remove(tKab); 
                    }
                    foreach (Kab tKab in VKsMaxINDkud)
                    {
                        if (Tochki.Exists(x => x.IND == tKab.INDOtk))
                        {
                            TPodkKud = Tochki.Find(x => x.IND == tKab.INDOtk);
                            tPodvKab.NIND(tKab.IND);
                            tPodvKab.NINDobor(tKab.INDOtk);
                            tPodvKab.NDIam(tKab.DIam);
                            tPodvKab.NKoordObor(TPodkKud.Koord);
                            tPodvKab.NMassa(tKab.Massa);
                            if (Vkrazv.Exists(x => x.Split('*')[0] == tKab.IND) == false)SpPodvKAB.Add(tPodvKab);
                        }
                        VK.Remove(tKab);
                    }
                    if (SpPodvKAB.Count > 0)
                    {
                        NZap.NSpVixKab(SpPodvKAB);
                        SgrupVK.Add(NZap);
                    }
                }
            }
            }
            Database db = doc.Database;
            Editor ed = doc.Editor;
            double Kol = 1;
            using (DocumentLock docLock = doc.LockDocument())
            {
                foreach (ZaPis zap in SgrupVK) { if (zap.Sist != "") { RazvOPoSp(zap); ed.WriteMessage("Кабели из точки-" + zap.INDotk.IND + " проложены (обработано-" + Kol.ToString() + " из " + SgrupVK.Count.ToString() + " точек)"); Kol = Kol + 1; ITOGI.Add("Кабели из точки-" + zap.INDotk.IND + " проложены (обработано-" + Kol.ToString() + " из " + SgrupVK.Count.ToString() + " точек)"); } }
                SOZDlov("OTPK");
                ZapisSlovSistTR("OTPK", ITOGI);
                //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\МАРШРУТ\ITOGI.txt")) { foreach (string line in ITOGI) { file.WriteLine(line); } }
            }
            using (DocumentLock docLock = doc.LockDocument())
            {
                VK.Clear();
                Vkrazv.Clear();
                VKizNOD(ref VK);
                SozdSpKab(ref Vkrazv);
                ZapTablK(VK);
            }
            this.Show();  
        }//развести кабели по оборудованию
        private void button2_Click(object sender, EventArgs e)
        {
            if (this.dataGridView2.CurrentRow.Cells[0].Value != null)
            {
                String FindK = this.dataGridView2.CurrentRow.Cells[0].Value.ToString();
                List<ObjectId> SpOBJID = new List<ObjectId>();
                //Application.ShowAlertDialog(Find);
                Document doc = Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    Database db = doc.Database;
                    Editor ed = doc.Editor;
                    ed.WriteMessage("Найти- " + FindK);
                    SBORpl_FIND(ref SpOBJID, FindK);
                    ObjectId[] idarrayEmpty = SpOBJID.ToArray();
                    ed.SetImpliedSelection(idarrayEmpty);
                }
            }
        }//найти кабель
        private void button3_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.CurrentRow.Cells[0].Value == null) return;
            if (Tochki.Exists(x => x.IND == this.dataGridView1.CurrentRow.Cells[0].Value)) return;
            String IND = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
            this.Hide();
            int Zvet = 52;
            string Gr = "";
            string tGrK = "";
            int[] KolZV = { 20, 2, 110, 6, 230, 4 };
            string[] Grupp = { "(I)", "(II)", "(III)", "(IV)", "(V)", "Силовые" };
            //List<Kab> VKt = VK.FindAll(x => x.INDKud == IND | x.INDOtk == IND | x.spVin.Contains(IND + ",") | x.spVin == IND |  x.spVin.Split(',').Last() == IND);
            List<Kab> VKt = VK.FindAll(x => x.INDKud == IND | x.INDOtk == IND );
            foreach (Kab Tkab in VK) { if (Tkab.spVin != null) if (Tkab.spVin.Contains(IND)) VKt.Add(Tkab); }
                foreach (Kab tKab in VKt) 
                {
                    tGrK = "(IV)";
                    foreach (string tGr in Grupp) if (tKab.IND.Contains(tGr)) tGrK = tGr;
                    if (Gr.Contains(tGrK) == false) Gr = Gr + tGrK; 
                }
            int i = 0;
            foreach (string tGr in Grupp) { if (Gr == tGr)Zvet = KolZV[i]; i = i + 1; }
            Document doc =Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            PromptPointResult pPtRes;
            PromptPointResult pPtRes1;
            PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку " + IND);
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (tr)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    pPtRes = doc.Editor.GetPoint(pPtOpts);
                    Point3d Toch1 = pPtRes.Value;
                    KRUGiTEXT(Toch1, Zvet, IND , Gr,15,15);
                    tr.Commit();
                }
                Tochki.Clear();
                SpVistTP(ref Tochki);
                ZapTablT(TochkiSort);
                ZapTablK(VK);
            }
            this.Show();
        }//выставить точку подключения
        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            Point3d BP = new Point3d();
            Point3d PSKtl = new Point3d(0, 0, 0);
            Point3d MSKtl = new Point3d(0, 0, 0);
            string OSItl = "";
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            using (DocumentLock docLock = doc.LockDocument())
            {
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    Editor ed =
                        Application.DocumentManager.MdiActiveDocument.Editor;
                    try
                    {
                        PromptEntityResult ers = ed.GetEntity("Укажите примитив ");
                        BlockReference bref = tr.GetObject(ers.ObjectId, OpenMode.ForRead) as BlockReference;
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
                            if (prop.PropertyName == "Видимость1") { OSItl = prop.Value.ToString(); }
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
                            }
                        }
                    }
                    PSKtl = new Point3d(BP.X + x1, BP.Y+y1, 0);
                    MSKtl = new Point3d(xMir, yMir, zMir);
                    tr.Commit();
                    }
                    catch
                    {
                        tr.Abort();
                    }
                }
            ProezT(MSKtl, PSKtl, OSItl);
            SpVistTP(ref Tochki);
            ZapTablT(TochkiSort);
            }
            this.Show();
        }//Кнопка спроецировать точки
        private void button5_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.CurrentRow.Cells[0].Value != null)
            {
                String FindK = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
                List<ObjectId> SpOBJID = new List<ObjectId>();
                //Application.ShowAlertDialog(Find);
                Document doc = Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = doc.LockDocument())
                {
                    Database db = doc.Database;
                    Editor ed = doc.Editor;
                    ed.WriteMessage("Найти- " + FindK);
                    SBORpl_FIND(ref SpOBJID, FindK);
                    ObjectId[] idarrayEmpty = SpOBJID.ToArray();
                    ed.SetImpliedSelection(idarrayEmpty);
                }
            }
        }//найти точку подключения
        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {

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
                VK.Clear();
                Vkrazv.Clear();
                VKizNOD(ref VK);
                SozdSpKab(ref Vkrazv);
                ZapTablK(VK);
            }
            this.Show();
        }//удалить все кабели
        private void button7_Click(object sender, EventArgs e)
        {
            Form7 form2 = new Form7();
            form2.Show();
        }//отчет ПТК
        private void button8_Click(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                VK.Clear();
                Vkrazv.Clear();
                //SozdSpKab(ref Vkrazv);
                VKizNOD(ref VK);
                SozdSpKV(ref Vkrazv);
                foreach (string Strok in Vkrazv) 
                {
                    if (VK.FindIndex(x => Strok.Split('*')[0] == x.IND) == -1)
                    {
                        Kab nkab = new Kab();
                        nkab.NIND(Strok.Split('*')[0]);
                        VK.Add(nkab);
                    }
                }
                ZapTablK(VK);
            }
        }//Маршруты
        private void button9_Click(object sender, EventArgs e)
        {
            this.dataGridView8.Rows.Clear();
            string File = "";
            int i = -1;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                UdalKrug();
                double[] KolZV = { 20, 2, 110, 6, 230, 4 };
                string[] Grupp = { "I", "II", "III", "IV", "V", "Силовые" };
                foreach (double Zvet in KolZV)
                {
                    i = i + 1;
                    if (EstLiSist(Zvet))
                    {
                        List<Noda> spNod = new List<Noda>();
                        List<Noda> spNod_Vin = new List<Noda>();
                        List<Noda> spPer = new List<Noda>();
                        List<Noda> spNodFin = new List<Noda>();
                        List<string> spLINI = new List<string>();
                        SozdSpNODPer(ref spPer, "INSERT");
                        SozdSpNODPerDinBl(ref spPer, "INSERT");
                        SozdSpNODPer(ref spPer, "CIRCLE");
                        SozdSpNOD1(Zvet, ref spNod, ref spLINI);
                        SozdSpNOD_Vin(ref spNod_Vin);
                        DOP_spNOD(ref spNod, ref spLINI, 10);
                        DOP_spNOD_Vin(ref spNod,spNod_Vin, 10);
                        SvazLin(ref spNod, ref spLINI, ref spNodFin);
                        foreach (Noda TnodEt in spPer) { spNodFin.Add(TnodEt);}
                        Skon(10, ref spNodFin, ref spLINI, ref spPer, Zvet);
                        SOZDlov("Трассы" + Grupp[i] + "группы");
                        ZapisSlovSistTR("Трассы" + Grupp[i] + "группы", NodiVStrList(spNodFin));
                    }
                }
                //ZagrDannNOD();
                ZagrDannNODisSLOV();
            }
        }//обновить трассы
        private void button10_Click(object sender, EventArgs e)
        {
            this.Hide();
            string SUF = this.textBox37.Text;
            string Adr = this.label34.Text;
            string Name = Adr.Split('\\').Last().Split('.')[0];
            double Nvin = 1;
            if (this.dataGridView3.RowCount == 0 & this.dataGridView4.RowCount == 0) Nvin = 1;
            string NmaxV = "";
            string SUFt = "";
                for (int i = 0; i < this.dataGridView3.RowCount; i++)
                {   if (this.dataGridView3.Rows[i].Cells[0].Value != null) 
                    {
                    NmaxV = this.dataGridView3.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView3.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                    } 
                }
                for (int i = 0; i < this.dataGridView4.RowCount ; i++)
                {
                    if (this.dataGridView4.Rows[i].Cells[0].Value != null)
                    {
                    NmaxV = this.dataGridView4.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                    }
                }
            InsBlockRef(@Adr, Name, Nvin.ToString(), "Выноска", SUF,"");
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()) {SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS);}
            this.dataGridView3.Rows.Clear();
            int i1 = -1;
            foreach (VINOS TV in Vinoski) 
            {
                i1 = i1 + 1;
                PovtVinoski = Vinoski.FindAll(x => x.Nazv == TV.Nazv & x.SUF == TV.SUF);
                this.dataGridView3.Rows.Add(TV.Nazv, TV.PolnNazv, TV.Spravka, TV.SUF);
                if (PovtVinoski.Count > 1)
                {
                    this.dataGridView3.Rows[i1].Cells[0].Style.BackColor = Color.LightPink;
                    this.dataGridView3.Rows[i1].Cells[1].Style.BackColor = Color.LightPink;
                    this.dataGridView3.Rows[i1].Cells[2].Style.BackColor = Color.LightPink;
                }
            }
            this.Show();
        }//выставить выноску
        private void button11_Click(object sender, EventArgs e)
        {
            string SUF = this.textBox37.Text;
            string Adr = this.label35.Text;
            string Name = Adr.Split('\\').Last().Split('.')[0];
            this.Hide();
            double Nvin = 1;
            if (this.dataGridView3.RowCount == 0 & this.dataGridView4.RowCount == 0) Nvin = 1;
            string NmaxV = "";
            string SUFt = "";
            for (int i = 0; i < this.dataGridView3.RowCount; i++)
            {
                if (this.dataGridView3.Rows[i].Cells[0].Value != null)
                {
                    NmaxV = this.dataGridView3.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView3.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                }
            }
            for (int i = 0; i < this.dataGridView4.RowCount; i++)
            {
                if (this.dataGridView4.Rows[i].Cells[0].Value != null)
                {
                    NmaxV = this.dataGridView4.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                }
            }
            InsBlockRef(@Adr, Name, Nvin.ToString(), "Переход", SUF,"");
            InsBlockRef(@Adr, Name, Nvin.ToString(), "Переход", SUF,"");
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()) { SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS); }
            this.dataGridView4.Rows.Clear();
            //foreach (PEREXOD TV in PEREX) { this.dataGridView4.Rows.Add(TV.Nazv, TV.Nazv, TV.Dlin.ToString(), TV.Nazv); }
            int i1 = -1;
            foreach (PEREXOD TV in PEREX)
            {
                i1 = i1 + 1;
                PovtPEREX = PEREX.FindAll(x => x.Nazv == TV.Nazv & x.SUF == TV.SUF);
                this.dataGridView4.Rows.Add(TV.Nazv, TV.PolnNazv, TV.Dlin.ToString(), TV.SUF);
                if (PovtPEREX.Count > 1)
                {
                    this.dataGridView4.Rows[i1].Cells[0].Style.BackColor = Color.LightPink;
                    this.dataGridView4.Rows[i1].Cells[1].Style.BackColor = Color.LightPink;
                    this.dataGridView4.Rows[i1].Cells[2].Style.BackColor = Color.LightPink;
                }
            }
            this.Show();
        }//Выставить переход
        private void button12_Click(object sender, EventArgs e)
        {
            string Adr = this.label36.Text;
            string Name = Adr.Split('\\').Last().Split('.')[0];
            this.Hide();
            //Application.ShowAlertDialog(@Adr + " " + Name);
            InsBlockRef(@Adr, Name,"","Плоскость","","");
            //InsBlockRef(@"D:\C#\TRASER_NC_AC19\Блоки\Плоскость_пр.dwg", "Плоскость_пр", "");
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()) { SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS); }
            this.dataGridView5.Rows.Clear();
            foreach (PLOS TV in SpPLOS) { this.dataGridView5.Rows.Add(TV.Vid, TV.List, TV.Osi, TV.Msk.ToString()); }
            this.Show();
        }//Выставить плоскость
        private void button13_Click(object sender, EventArgs e)
        {
            this.Hide();
            List<TPodk> SpVin = new List<TPodk>();
            using (DocumentLock docLock = doc.LockDocument())
            {

                SpVistVinDB(ref SpVin);
                Sozd_TB(ref SpVin);
                Postr_TB(ref SpVin);
            }
            this.Show();
        }//Перечень выносок
        private void button14_Click(object sender, EventArgs e)
        {
            using (DocumentLock docLock = doc.LockDocument())
            {SSil_Na_Vid();}
        }//Ссылки на виды
        private void button15_Click(object sender, EventArgs e)
        {
            double A = 400;
            double B = 100;
            double a = 40;
            double b = 80;
            string stIND = "";
            string stDlin = "";
            string stStar = "";
            string stNov = "";
            string stNom = "";
            List<string> Vk_Vibr = new List<string>();
            int nom;

            double Dlin = 0;
            double DlinN = 0;

            int Schet;
            this.Hide();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // Start a transaction
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Request for objects to be selected in the drawing area
                    TypedValue[] acTypValAr = new TypedValue[2];
                    acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
                    acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Kabeli"), 1);
                    // создаем фильтр
                    SelectionFilter filter = new SelectionFilter(acTypValAr);
                    PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection(filter);
                    // If the prompt status is OK, objects were selected
                    if (acSSPrompt.Status == PromptStatus.OK)
                    {
                        SelectionSet acSSet = acSSPrompt.Value;
                        //Application.ShowAlertDialog(acSSet.Count.ToString());
                        // Step through the objects in the selection set
                        foreach (SelectedObject acSSObj in acSSet)
                        {
                            Polyline ln = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                            var KolV = ln.NumberOfVertices;
                            Dlin = ln.Length;
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Noda Nnoda = new Noda();
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { stIND = value.Value.ToString(); Vk_Vibr.Add(stIND); }
                                    Schet = Schet + 1;
                                }
                            }
                        }
                        // Save the new object to the database
                        acTrans.Commit();
                    }
                }
                // Dispose of the transaction
            Vk_Vibr.Sort(delegate(string x, string y) { return x.CompareTo(y); });
            Transaction tr = acCurDb.TransactionManager.StartTransaction();
            using (tr)
            {
                PromptPointResult pPtRes;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d Toch1 = pPtRes.Value;
                Point3d TochT = Toch1;
                foreach (string Ind in Vk_Vibr)
                {
                    if (Ind != "")
                    {
                        LWPOIYiTEXT(TochT, 0, Ind, A, B, a, b, 50, "Выноски");
                        TochT = new Point3d(TochT.X + A, TochT.Y, 0);
                        if (TochT.X == Toch1.X + 8 * A) TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                    }
                }
            tr.Commit();
            }
            }
            this.Show();
        }//отдельная выноска
        private void button16_Click(object sender, EventArgs e)
        {
            string NOM = this.textBox5.Text;
            string NAME = this.textBox6.Text;
            string ISP = this.textBox7.Text;
            string PROV = this.textBox8.Text;
            string VIP = this.textBox9.Text;
            string TKONTR = this.textBox10.Text;
            string NKONTR = this.textBox11.Text;
            string UTV = this.textBox12.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("DOC");
                ZapisSlov("DOC", NOM + "&" + NAME + "&" + ISP + "&" + PROV + "&" + VIP + "&" + TKONTR + "&" + NKONTR + "&" + UTV);
            }
        }//сохранить реквизиты
        private void button17_Click(object sender, EventArgs e)
        {
            double A = Convert.ToDouble(this.textBox14.Text);
            double B = Convert.ToDouble(this.textBox13.Text);
            double PlusD = Convert.ToDouble(this.textBox16.Text);
            //double Otstup = Convert.ToDouble(this.textBox17.Text);
            double a = 0;
            double b = 0;
            string stIND = "";
            string stDlin = "";
            string stStar = "";
            string stNov = "";
            string stNom = "";
            string strMassa = "0";
            double douDiam = 0;
            List<Kab> Vk_Vibr = new List<Kab>();
            List<Kab> Vk_Postr = new List<Kab>();

            List<Kab> Vk_I = new List<Kab>();
            List<Kab> Vk_II = new List<Kab>();
            List<Kab> Vk_III = new List<Kab>();
            List<Kab> Vk_IV = new List<Kab>();
            List<Kab> Vk_V = new List<Kab>();

            int nom;

            double Dlin = 0;
            double DlinN = 0;

            int Schet;
            this.Hide();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // Start a transaction
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Request for objects to be selected in the drawing area
                    TypedValue[] acTypValAr = new TypedValue[2];
                    acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
                    acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Kabeli"), 1);
                    // создаем фильтр
                    SelectionFilter filter = new SelectionFilter(acTypValAr);
                    PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection(filter);
                    // If the prompt status is OK, objects were selected
                    if (acSSPrompt.Status == PromptStatus.OK)
                    {
                        SelectionSet acSSet = acSSPrompt.Value;
                        //Application.ShowAlertDialog(acSSet.Count.ToString());
                        // Step through the objects in the selection set
                        foreach (SelectedObject acSSObj in acSSet)
                        {
                            douDiam = 0;
                            stIND = "";
                            strMassa = "0";
                            Polyline ln = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Polyline;
                            var KolV = ln.NumberOfVertices;
                            Dlin = ln.Length;
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Noda Nnoda = new Noda();
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { stIND = value.Value.ToString(); }
                                    if (Schet == 2) { douDiam = Convert.ToDouble(value.Value.ToString()); }
                                    if (Schet == 6) { strMassa = value.Value.ToString(); }
                                    Schet = Schet + 1;
                                }
                            }
                            if (stIND != "")
                            {
                                Kab Tkab = new Kab();
                                Tkab.NIND(stIND);
                                Tkab.NDIam(douDiam + PlusD);
                                Tkab.NDIamKor(douDiam);
                                Tkab.NMassa(strMassa);
                                Vk_Vibr.Add(Tkab);
                            }
                        }
                        // Save the new object to the database
                        acTrans.Commit();
                    }
                }
                // Dispose of the transaction
                Vk_Vibr.Sort(delegate(Kab x, Kab y) { return y.DIam.CompareTo(x.DIam); });
                List<DiamVsehc> Sehc = new List<DiamVsehc>();
                if (this.checkBox2.Checked == false)
                {
                    Transaction tr = acCurDb.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        PromptPointResult pPtRes;
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                        PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                        pPtRes = doc.Editor.GetPoint(pPtOpts);
                        Point3d Toch1 = pPtRes.Value;
                        Vk_IV = Vk_Vibr;
                        while (Vk_IV.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_IV, ref Vk_Postr, "");
                            foreach (Kab Ind in Vk_Postr) Vk_IV.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        tr.Commit();
                    }
                }
                else
                {
                    Transaction tr = acCurDb.TransactionManager.StartTransaction();
                    using (tr)
                    {
                        PromptPointResult pPtRes;
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                        PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                        pPtRes = doc.Editor.GetPoint(pPtOpts);
                        Point3d Toch1 = pPtRes.Value;
                        Vk_I = Vk_Vibr.FindAll(x => x.IND.Contains("(I)") == true);
                        Vk_II = Vk_Vibr.FindAll(x => x.IND.Contains("(II)") == true);
                        Vk_III = Vk_Vibr.FindAll(x => x.IND.Contains("(III)") == true);
                        Vk_V = Vk_Vibr.FindAll(x => x.IND.Contains("(V)") == true);
                        Vk_IV = Vk_Vibr.FindAll(x => x.IND.Contains("I") == false & x.IND.Contains("V") == false);
                        while (Vk_I.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_I, ref Vk_Postr, "I");
                            foreach (Kab Ind in Vk_Postr) Vk_I.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_II.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_II, ref Vk_Postr, "II");
                            foreach (Kab Ind in Vk_Postr) Vk_II.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_III.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_III, ref Vk_Postr, "III");
                            foreach (Kab Ind in Vk_Postr) Vk_III.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_V.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_V, ref Vk_Postr, "V");
                            foreach (Kab Ind in Vk_Postr) Vk_V.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_IV.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_IV, ref Vk_Postr, "IV");
                            foreach (Kab Ind in Vk_Postr) Vk_IV.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        tr.Commit();
                    }
                }
            }
            this.Show();  
        }//Сечение трассы
        private void button18_Click(object sender, EventArgs e)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                SpPLOS.Clear();
                SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS);
                List<string> Vkrazv = new List<string>();
                SozdSpKV_n_k(ref Vkrazv);
                string stStar = "";
                string[] stSpT ;
                string[] stT0;
                string[] stT1;
                Point3d pT0 = new Point3d();
                Point3d pT1 = new Point3d();
                string stIND = "";
                string Kud = "";
                string stDlin = "";
                double Dlin = 0;
                double Koef = Convert.ToDouble(this.textBox15.Text);
                for (int i = 0; i < this.dataGridView2.RowCount; i++)
                {
                    Dlin = 0;
                    stIND = this.dataGridView2[0, i].Value.ToString();
                    if (this.dataGridView2[2, i].Value != "")  Dlin = Convert.ToDouble(this.dataGridView2[2, i].Value.ToString());
                    if (Vkrazv.Exists(x => x.Split('*')[0] == stIND) == true)
                    {
                        stStar = Vkrazv.Find(x => x.Split('*')[0] == stIND);
                        stSpT = stStar.Split('*')[1].Split(':');
                            if (stSpT.Length == 2) 
                            {
                                stT0 = stSpT[0].Split();
                                stT1 = stSpT[1].Split();
                                pT0 = new Point3d(Convert.ToDouble(stT0[0]),Convert.ToDouble(stT0[1]),Convert.ToDouble(stT0[2]));
                                pT1 = new Point3d(Convert.ToDouble(stT1[0]), Convert.ToDouble(stT1[1]), Convert.ToDouble(stT1[2]));
                            }
                            if (stSpT.Length > 2)
                            {
                                stT0 = stSpT[0].Split();
                                stT1 = stSpT[stSpT.Length-1].Split();
                                pT0 = new Point3d(Convert.ToDouble(stT0[0]), Convert.ToDouble(stT0[1]), Convert.ToDouble(stT0[2]));
                                pT1 = new Point3d(Convert.ToDouble(stT1[0]), Convert.ToDouble(stT1[1]), Convert.ToDouble(stT1[2]));
                            }
                            double PoPram = pT0.DistanceTo(pT1);
                            Dlin = Convert.ToDouble(this.dataGridView2[2, i].Value.ToString());
                            if (this.dataGridView2[2, i].Value != "") Dlin = Convert.ToDouble(this.dataGridView2[2, i].Value.ToString());
                            this.dataGridView2[9, i].Value = (PoPram / 1000).ToString("00.00");
                            this.dataGridView2[10, i].Value = (Dlin / (PoPram / 1000)).ToString("00.00");
                            if (Dlin / (PoPram / 1000) > Koef)
                            { this.dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightPink; }
                    }
                }
            }
        }//проверка длинн
        private void button19_Click(object sender, EventArgs e)
        {
            double A = Convert.ToDouble(this.textBox14.Text);
            double B = Convert.ToDouble(this.textBox13.Text);
            double PlusD = Convert.ToDouble(this.textBox16.Text);
            double a = 0;
            double b = 0;
            string stIND = "";
            string[] IndDi;
            double douDiam = 0;
            List<Kab> Vk_Vibr = new List<Kab>();
            List<Kab> Vk_Postr = new List<Kab>();
            List<Kab> Vk_IV = new List<Kab>();

            List<Kab> Vk_I = new List<Kab>();
            List<Kab> Vk_II = new List<Kab>();
            List<Kab> Vk_III = new List<Kab>();
            List<Kab> Vk_V = new List<Kab>();
            this.Hide();


            string[] lines = System.IO.File.ReadAllLines(@"C:\МАРШРУТ\ДИАМЕТРЫ.txt", Encoding.GetEncoding("Windows-1251"));
            foreach (string Strok in lines)
            {
                if (Strok != "")
                {
                    IndDi = Strok.Split(':');
                    Kab Tkab = new Kab();
                    Tkab.NIND(IndDi[0]);
                    Tkab.NDIam(Convert.ToDouble(IndDi[1].Replace(",",".")) + PlusD);
                    Tkab.NDIamKor(Convert.ToDouble(IndDi[1].Replace(",", ".")));
                    Vk_Vibr.Add(Tkab);
                }
            }
            Vk_Vibr.Sort(delegate(Kab x, Kab y) { return y.DIam.CompareTo(x.DIam); });
            List<DiamVsehc> Sehc = new List<DiamVsehc>();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Transaction tr = acCurDb.TransactionManager.StartTransaction();
            if (this.checkBox2.Checked == false)
            {
                using (DocumentLock docLock = acDoc.LockDocument())
                {
                    using (tr)
                    {
                        PromptPointResult pPtRes;
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                        PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                        pPtRes = doc.Editor.GetPoint(pPtOpts);
                        Point3d Toch1 = pPtRes.Value;
                        Vk_IV = Vk_Vibr;
                        while (Vk_IV.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_IV, ref Vk_Postr, "");
                            foreach (Kab Ind in Vk_Postr) Vk_IV.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        tr.Commit();
                    }
                }
            }
            else
            {
                using (DocumentLock docLock = acDoc.LockDocument()) 
                {
                    using (tr)
                    {
                        PromptPointResult pPtRes;
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                        PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                        pPtRes = doc.Editor.GetPoint(pPtOpts);
                        Point3d Toch1 = pPtRes.Value;
                        Vk_I = Vk_Vibr.FindAll(x => x.IND.Contains("(I)") == true);
                        Vk_II = Vk_Vibr.FindAll(x => x.IND.Contains("(II)") == true);
                        Vk_III = Vk_Vibr.FindAll(x => x.IND.Contains("(III)") == true);
                        Vk_V = Vk_Vibr.FindAll(x => x.IND.Contains("(V)") == true);
                        Vk_IV = Vk_Vibr.FindAll(x => x.IND.Contains("I") == false & x.IND.Contains("V") == false);
                        while (Vk_I.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_I, ref Vk_Postr, "I");
                            foreach (Kab Ind in Vk_Postr) Vk_I.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_II.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_II, ref Vk_Postr, "II");
                            foreach (Kab Ind in Vk_Postr) Vk_II.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_III.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_III, ref Vk_Postr, "III");
                            foreach (Kab Ind in Vk_Postr) Vk_III.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_V.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_V, ref Vk_Postr, "V");
                            foreach (Kab Ind in Vk_Postr) Vk_V.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        Vk_Postr.Clear();
                        while (Vk_IV.Count > 0)
                        {
                            Puchok(Toch1, ref Vk_IV, ref Vk_Postr, "IV");
                            foreach (Kab Ind in Vk_Postr) Vk_IV.Remove(Ind);
                            Toch1 = new Point3d(Toch1.X + A + 10, Toch1.Y, Toch1.Z);
                        }
                        tr.Commit();
                    }
                }
            }
            this.Show();
        }//Сечение трассы из TXT
        private void button20_Click(object sender, EventArgs e)
        {
            //this.Hide();
            //double DelX = Convert.ToDouble(this.textBox18.Text);
            //double DelY = Convert.ToDouble(this.textBox19.Text);
            //double DelZ = Convert.ToDouble(this.textBox20.Text);
            //Point3d T1_1 = new Point3d();
            //Document acDoc = Application.DocumentManager.MdiActiveDocument;
            //Database acCurDb = acDoc.Database;
            //Editor ed = acDoc.Editor;
            //using (DocumentLock docLock = acDoc.LockDocument())
            //{
            //    Transaction tr = acCurDb.TransactionManager.StartTransaction();
            //    using (tr)
            //    {
            //        PromptPointResult pPtRes;
            //        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
            //        PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
            //        pPtRes = doc.Editor.GetPoint(pPtOpts);
            //        T1_1 = pPtRes.Value;
            //    }
            //    Zagib(T1_1, DelX, DelY, DelZ,10);
            //}
            //this.Show();
        }//проверка радиусов изгиба
        private void button21_Click(object sender, EventArgs e)
        {
            VK.Clear();
            TochkiRazv.Clear();
            TochkiSort.Clear();
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else return;
            ZagrExcel(File);
            this.button27.Enabled = true;
            this.button23.Enabled = true;
        }//загрузка из excel
        private void button22_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
            return;
            string StrokDop = "";
            this.Hide();
            List<string> VKtxt = new List<string>();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                string[] lines = System.IO.File.ReadAllLines(@File, Encoding.Default);
                foreach (string Strok in lines)
                {
                    StrokDop = Strok + "::::::::::";
                    VKtxt.Add(StrokDop);
                }
                SOZDlov("VK");
                ZapisSlovSistTR("VK", VKtxt);
                this.Show();
                //TochkiRazv.Clear();
                string XYZ1 = "";
                string XYZ2 = "";
                int kVist = 0;
                List<String> TochkiSTR = new List<String>();
                SozdSpKab(ref Vkrazv);
                SpVistTP(ref Tochki);
                SpVistTP_BL(ref Tochki);
                foreach (string Strok in VKtxt)
                {
                    StrokDop = Strok + "::::::::::";
                    string[] Kabel = StrokDop.Split(':');
                    Kab nkab = new Kab();
                    nkab.NIND(Kabel[0]);
                    nkab.NDIam(Convert.ToDouble(Kabel[1]));
                    nkab.NINDOtk(Kabel[4]);
                    nkab.NINDKud(Kabel[7]);
                    nkab.NSist(Kabel[10]);
                    nkab.NBlokNOD(Kabel[11]);
                    nkab.KoorOt = Kabel[12];
                    nkab.KoorKud = Kabel[13];
                    VK.Add(nkab);
                    if (TochkiSort.Exists(x => x.IND == nkab.INDOtk & x.Sist == nkab.Sist & x.BlNOD == nkab.BlokNOD) == false)
                    {
                        TPodk NNOD = new TPodk();
                        NNOD.NIND(nkab.INDOtk);
                        NNOD.NSist(nkab.Sist);
                        NNOD.NBlNOD(nkab.BlokNOD);
                        NNOD.NKoorMod(Kabel[12]);
                        if (Kabel.Length > 13) NNOD.NNaim(Kabel[14]);
                        if (Tochki.Exists(x => x.IND == nkab.INDOtk)) { NNOD.NKoor(Tochki.Find(x => x.IND == nkab.INDOtk).Koord); }
                        if (TochkiSort.Exists(x => x.IND == nkab.INDOtk) == false) { TochkiSort.Add(NNOD); }
                    }
                    TochkiSTR.Add(nkab.INDKud);
                    if (TochkiSort.Exists(x => x.IND == nkab.INDKud & x.Sist == nkab.Sist & x.BlNOD == nkab.BlokNOD) == false)
                    {
                        TPodk NNOD = new TPodk();
                        NNOD.NIND(nkab.INDKud);
                        NNOD.NSist(nkab.Sist);
                        NNOD.NBlNOD(nkab.BlokNOD);
                        NNOD.NKoorMod(Kabel[13]);
                        if (Kabel.Length > 13) NNOD.NNaim(Kabel[15]);
                        if (Tochki.Exists(x => x.IND == nkab.INDKud)) { NNOD.NKoor(Tochki.Find(x => x.IND == nkab.INDKud).Koord); }
                        if (TochkiSort.Exists(x => x.IND == nkab.INDKud) == false) { TochkiSort.Add(NNOD); }
                    }
                }
                foreach (string Strok in Vkrazv)
                {
                    if (VK.FindIndex(x => Strok.Split('*')[0] == x.IND) == -1)
                    {
                        Kab nkab = new Kab();
                        nkab.NIND(Strok.Split('*')[0]);
                        VK.Add(nkab);
                    }
                }
                ZapTablT(TochkiSort);
                ZapTablK(VK);
            }
        }//загрузка из txt
        private void button23_Click(object sender, EventArgs e)
        {
            string Ind = "-";
            string SpOtr = "";
            int i = 1;
            List<string> VkrazvPl = new List<string>();
            List<string> SpOtrL = new List<string>();
            VK.Clear();
            Vkrazv.Clear();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                VKizNOD(ref VK);
                SozdSpKV(ref Vkrazv);
                SozdSpKVKoor(ref VkrazvPl, SpPLOS);
            }
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook xlWB; //рабочая книга              
            Excel.Worksheet xlSht; //лист Excel   
            string File = this.label33.Text;
            xlWB = xlApp.Workbooks.Open(@File); //название файла Excel                                             
            xlSht = xlWB.Worksheets[this.textBox20.Text]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            while (Ind != null) 
            {
                i = i + 1;
                SpOtr = "";
                Ind = xlWB.Worksheets[this.textBox20.Text].range[this.textBox36.Text + i.ToString()].value;
                if (Vkrazv.Exists(x => x.Split('*')[0] == Ind) == true)
                { 
                    xlWB.Worksheets[this.textBox20.Text].range[this.textBox34.Text + i.ToString()].value = Vkrazv.Find(x => x.Split('*')[0] == Ind).Split('*')[1];
                    xlWB.Worksheets[this.textBox20.Text].range[this.textBox38.Text + i.ToString()].value = Vkrazv.Find(x => x.Split('*')[0] == Ind).Split('*')[2].TrimEnd('-');
                }
                if (VkrazvPl.Exists(x => x.Split(':')[0] == Ind) == true)
                {
                    SpOtrL = VkrazvPl.FindAll(x => x.Split(':')[0] == Ind);
                    foreach (string tt in SpOtrL) SpOtr = SpOtr + "%" + tt;
                    xlWB.Worksheets[this.textBox20.Text].range[this.textBox35.Text + i.ToString()].value = SpOtr;
                }
            } 
            xlWB.Save();
            xlWB.Close(); //закрываем книгу, изменения не сохраняем
            xlApp.Quit(); //закрываем Excel
        }//Сохранить в excel
        private void button27_Click(object sender, EventArgs e)
        {
            VK.Clear();
            TochkiRazv.Clear();
            TochkiSort.Clear();
            string File = this.label33.Text;
            ZagrExcel(File);
        }//обновить в excel
        private void button24_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
            return;
            this.label34.Text = File;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("ADRblVin");
                ZapisSlov("ADRblVin", File);
            }
        }//изменить расположение блока выноски
        private void button25_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
                return;
            this.label35.Text = File;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("ADRblPer");
                ZapisSlov("ADRblPer", File);
            }
        }//изменить расположение блока перехода
        private void button26_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
                return;
            this.label36.Text = File;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("ADRblPlos");
                ZapisSlov("ADRblPlos", File);
            }
        }//изменить расположение блока перехода
        private void button20_Click_1(object sender, EventArgs e)
        {
            this.dataGridView9.Rows.Clear();
            int i = -1;
            foreach (TPodk TP in Tochki)
            {
                i = i + 1;
                Point3d KoorMod = TP.Koord;
                Point3d Zentr = TP.Koord;
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
                        if (iPlos.Osi == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; Znow = Znow + delZ - TP.VisT; }
                        if (iPlos.Osi == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                        if (iPlos.Osi == "ZY") { Xnow = Xnow + delZ; Ynow = Ynow + delX; Znow = Znow + delY; }
                        if (iPlos.Osi == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                        if (iPlos.Osi == "YZ") { Xnow = Xnow + delZ; Ynow = Ynow - delX; Znow = Znow + delY; }
                        KoorMod = new Point3d(Xnow, Ynow, Znow);
                    }
                }
                this.dataGridView9.Rows.Add(TP.IND, KoorMod, TP.Shem, TP.Naim, TP.Pom);
                this.dataGridView9.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
            }
        }//показать точки по чертежу
        private void button28_Click(object sender, EventArgs e)
        {
            LoadDataPC();
        }//показать точки из базы
        private void button29_Click(object sender, EventArgs e)
        {
            List<string> Vkrazv = new List<string>();
            SozdSpKab(ref Vkrazv);
            foreach (TPodk TP in TochkiSort)
            { if (Tochki.Exists(x => x.IND == TP.IND)) { ADDZapOCP(TP.IND, "Ind,Pos,Shem", TP.IND + "','" + Tochki.Find(x => x.IND == TP.IND).KoordMod + "','" + TP.Shem, TP.IND); } }
            LoadDataPC();
        }//перенести в базу точки
        private void button30_Click(object sender, EventArgs e)
        {
            string SpPom = this.textBox18.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("spPOM");
                ZapisSlov("spPOM", SpPom);
            }
            List<String> VKisSlov = HCtenSlovNod("VK");
            obnTP(VKisSlov, SpPom);
            foreach (string Strok in Vkrazv)
            {
                if (VK.FindIndex(x => Strok.Split('*')[0] == x.IND) == -1)
                {
                    Kab nkab = new Kab();
                    nkab.NIND(Strok.Split('*')[0]);
                    VK.Add(nkab);
                }
            }
            ZapTablT(TochkiSort);
            ZapTablK(VK);
        }//сохранить список помещений
        private void button31_Click(object sender, EventArgs e)
        {
            string Proj = this.textBox19.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("PROG");
                ZapisSlov("PROG", Proj);
            }
        }//сохранить номер проекта
        private void button32_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
            return;
            this.label21.Text = File;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("ADRblLn");
                ZapisSlov("ADRblLn", File);
            }
        }//изменить расположение формата для листа основной части
        private void button33_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
                return;
            this.label22.Text = File;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("ADRblL1");
                ZapisSlov("ADRblL1", File);
            }
        }//изменить расположение формата для титульного листа
        private void button34_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
                return;
            this.label23.Text = File;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("ADRblLreg");
                ZapisSlov("ADRblLreg", File);
            }
        }//изменить расположение формата для листа регистрации
        private void button36_Click(object sender, EventArgs e)
        {
            string Nastr = this.textBox20.Text + ":" + this.textBox21.Text + ":" + this.textBox22.Text + ":" + this.textBox23.Text + ":" + this.textBox24.Text + ":" + this.textBox25.Text + ":" + this.textBox26.Text + ":" + this.textBox27.Text + ":" + this.textBox28.Text + ":" + this.textBox29.Text + ":" + this.textBox30.Text + ":" + this.textBox31.Text + ":" + this.textBox32.Text + ":" + this.textBox33.Text + ":" + this.textBox34.Text + ":" + this.textBox35.Text + ":" + this.textBox36.Text + ":" + this.textBox38.Text;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("NastExel");
                ZapisSlov("NastExel", Nastr);
            }
        }//сохранить изменения в связях с excel
        private void button37_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string File = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            { File = ofd.FileName; }
            else
                return;
            this.label53.Text = File;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("AdrBD");
                ZapisSlov("AdrBD", File);
            }
        }//изменить расположение базы точек подключения
        private void button38_Click(object sender, EventArgs e)
        {
            string UTV = this.textBox37.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("SUF");
                ZapisSlov("SUF",  UTV);
            }
        }//сохранить суфикс
        private void button39_Click(object sender, EventArgs e)
        {
            this.Hide();
            string SUF = "БезПодиси";
            string Adr = this.label34.Text;
            string Name = Adr.Split('\\').Last().Split('.')[0];
            double Nvin = 1;
            if (this.dataGridView3.RowCount == 0 & this.dataGridView4.RowCount == 0) Nvin = 1;
            string NmaxV = "";
            string SUFt = "";
            for (int i = 0; i < this.dataGridView3.RowCount; i++)
            {
                if (this.dataGridView3.Rows[i].Cells[0].Value != null)
                {
                    NmaxV = this.dataGridView3.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView3.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                }
            }
            for (int i = 0; i < this.dataGridView4.RowCount; i++)
            {
                if (this.dataGridView4.Rows[i].Cells[0].Value != null)
                {
                    NmaxV = this.dataGridView4.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                }
            }
            InsBlockRef(@Adr, Name, "", "Выноска", SUF, Nvin.ToString() + SUF);
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()) { SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS); }
            this.dataGridView3.Rows.Clear();
            int i1 = -1;
            foreach (VINOS TV in Vinoski)
            {
                i1 = i1 + 1;
                PovtVinoski = Vinoski.FindAll(x => x.Nazv == TV.Nazv & x.SUF == TV.SUF);
                this.dataGridView3.Rows.Add(TV.Nazv, TV.PolnNazv, TV.Spravka, TV.SUF);
                if (PovtVinoski.Count > 1)
                {
                    this.dataGridView3.Rows[i1].Cells[0].Style.BackColor = Color.LightPink;
                    this.dataGridView3.Rows[i1].Cells[1].Style.BackColor = Color.LightPink;
                    this.dataGridView3.Rows[i1].Cells[2].Style.BackColor = Color.LightPink;
                }
            }
            this.Show();
        }//выставить выноску без подписи
        private void button40_Click(object sender, EventArgs e)
        {
            string SUF = "БезПодиси";
            string Adr = this.label35.Text;
            string Name = Adr.Split('\\').Last().Split('.')[0];
            this.Hide();
            double Nvin = 1;
            if (this.dataGridView3.RowCount == 0 & this.dataGridView4.RowCount == 0) Nvin = 1;
            string NmaxV = "";
            string SUFt = "";
            for (int i = 0; i < this.dataGridView3.RowCount; i++)
            {
                if (this.dataGridView3.Rows[i].Cells[0].Value != null)
                {
                    NmaxV = this.dataGridView3.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView3.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                }
            }
            for (int i = 0; i < this.dataGridView4.RowCount; i++)
            {
                if (this.dataGridView4.Rows[i].Cells[0].Value != null)
                {
                    NmaxV = this.dataGridView4.Rows[i].Cells[0].Value.ToString();
                    SUFt = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                    if (Convert.ToDouble(NmaxV) >= Nvin & SUFt == SUF) Nvin = Convert.ToDouble(NmaxV) + 1;
                }
            }
            InsBlockRef(@Adr, Name, "", "Переход", SUF, Nvin.ToString() + SUF);
            InsBlockRef(@Adr, Name, "", "Переход", SUF, Nvin.ToString() + SUF);
            Vinoski.Clear();
            PEREX.Clear();
            SpPLOS.Clear();
            using (DocumentLock docLock = doc.LockDocument()){SBOR_OBOR(ref Vinoski, ref PEREX, ref SpPLOS);}
            this.dataGridView4.Rows.Clear();
            int i1 = -1;
            foreach (PEREXOD TV in PEREX)
            {
                i1 = i1 + 1;
                PovtPEREX = PEREX.FindAll(x => x.Nazv == TV.Nazv & x.SUF == TV.SUF);
                this.dataGridView4.Rows.Add(TV.Nazv, TV.PolnNazv, TV.Dlin.ToString(), TV.SUF);
                if (PovtPEREX.Count > 1)
                {
                    this.dataGridView4.Rows[i1].Cells[0].Style.BackColor = Color.LightPink;
                    this.dataGridView4.Rows[i1].Cells[1].Style.BackColor = Color.LightPink;
                    this.dataGridView4.Rows[i1].Cells[2].Style.BackColor = Color.LightPink;
                }
            }
            this.Show();
        }//выставить перехлж без подписи
        private void button41_Click(object sender, EventArgs e)
        {
            this.Hide();
            List<TPodk> SpVin = new List<TPodk>();
            using (DocumentLock docLock = doc.LockDocument())
            {
                HsistSloi();
                SpVistVinBP_DB(ref SpVin);
                Sozd_TB(ref SpVin);
                Postr_VIN(ref SpVin);
            }
            this.Show();
        }//выноски без подписи
        private void button42_Click(object sender, EventArgs e)
        {
            string KolST = this.textBox39.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                SOZDlov("KolST");
                ZapisSlov("KolST", KolST);
            }
        }//сохранить количество столбцов
        private void button35_Click(object sender, EventArgs e)
        {
            this.Hide();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {

                Database db = doc.Database;
                Editor ed = doc.Editor;
                ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
                ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
                TypedValue[] acTypValAr = new TypedValue[2];
                acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 0);
                acTypValAr.SetValue(new TypedValue(8, "KabeliError"), 1);
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
                VK.Clear();
                Vkrazv.Clear();
                VKizNOD(ref VK);
                SozdSpKab(ref Vkrazv);
                ZapTablK(VK);
            }
            this.Show();
        }//удалить кабели с ошибкой
        private void button43_Click(object sender, EventArgs e)
        {
            List<String> SlovSTR = new List<String>();
            ShtenTXT_SLov(ref SlovSTR);
            string Toch = "";
            string Compon = "";
            string Name = "";
            string Shem = this.textBox40.Text;
            int i = -1;
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(@"C:\МАРШРУТ\точкиП.txt"))
            {
                foreach (TPodk TP in Tochki)
                {
                    i = i + 1;
                    Compon = "";
                    Point3d KoorMod = TP.Koord;
                    Point3d Zentr = TP.Koord;
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
                            if (iPlos.Osi == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; Znow = Znow + delZ - TP.VisT; }
                            if (iPlos.Osi == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                            if (iPlos.Osi == "ZY") { Xnow = Xnow + delZ; Ynow = Ynow + delX; Znow = Znow + delY; }
                            if (iPlos.Osi == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                            if (iPlos.Osi == "YZ") { Xnow = Xnow + delZ; Ynow = Ynow - delX; Znow = Znow + delY; }
                            KoorMod = new Point3d(Xnow, Ynow, Znow);
                        }
                    }
                    if(SlovSTR.Exists(x => x.Contains(TP.Naim))) Compon = SlovSTR.Find(x => x.Contains(TP.Naim)).Split(':')[1];
                    if (TP.Shem != "") Shem = TP.Shem;
                    Name = TP.IND.Replace('К', 'K');
                    Toch = Name + ":" + KoorMod.X.ToString("0.") + " " + KoorMod.Y.ToString("0.") + " " + KoorMod.Z.ToString("0.") + ":" + Compon + ":" + TP.Pom + ":" + Shem + ":-:" + TP.Povorot.X.ToString("0.##") + " " + TP.Povorot.Y.ToString("0.##") + " " + TP.Povorot.Z.ToString("0.##");
                    if(Compon!="") file.WriteLine(Toch);
                }
            }
        }//список точек в txt

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            List<Kab> VKf = new List<Kab>();
            if (this.textBox3.Text != "")
            {
                VKf = VK.FindAll(x => x.IND.Contains(this.textBox3.Text));
                ZapTablK(VKf);
            }
            else
            {
                ZapTablK(VK);
            }
        }//Фильтр кабельного
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            List<TPodk> SPTpf = new List<TPodk>();
            if (this.textBox2.Text != "")
            {
                SPTpf = TochkiSort.FindAll(x => x.IND.Contains(this.textBox2.Text));
                ZapTablT(SPTpf);
            }
            else
            {
                ZapTablT(TochkiSort);
            }
        }//Фильтр точек подключения
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (this.dataGridView1.CurrentRow.Cells[0].Value == null) return;
            String IND = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
            string[] spVin;
            List<Kab> VKt = VK.FindAll(x => x.INDKud == IND | x.INDOtk == IND );
            foreach (Kab Tkab in VK) { if (Tkab.spVin != null) if (Tkab.spVin.Contains(IND)) VKt.Add(Tkab); }
            this.dataGridView6.Rows.Clear();
            foreach (Kab tKab in VKt) 
            {
                if (tKab.INDOtk == IND)
                    this.dataGridView6.Rows.Add(tKab.IND, tKab.DIam, tKab.INDKud, tKab.Shem); 
                else if (tKab.INDKud == IND)
                    this.dataGridView6.Rows.Add(tKab.IND, tKab.DIam, tKab.INDOtk, tKab.Shem); 
                else
                {
                    spVin = tKab.spVin.Split(',');
                    foreach (string tV in spVin) { if (IND == tV) this.dataGridView6.Rows.Add(tKab.IND, tKab.DIam, tKab.INDOtk + "," + tKab.INDKud, tKab.Shem); }
                }
           
            }
        }//выбор точки подключения
        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (this.dataGridView2.CurrentRow.Cells[11].Value == null) { this.label20.Text = ""; return; }
            String IND = this.dataGridView2.CurrentRow.Cells[11].Value.ToString();
            this.label20.Text = IND;
        }
        private void dataGridView8_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            this.dataGridView7.Rows.Clear();
            int Nom = 0;
            int i = 0;
            if (this.dataGridView8.CurrentRow.Cells[0].Value == null) { return; }
            String strSist = this.dataGridView8.CurrentRow.Cells[0].Value.ToString();
            if (strSist.Contains("Трассы I группы")) { Nom = 0; }
            if (strSist.Contains("Трассы II группы")) { Nom = 1; }
            if (strSist.Contains("Трассы III группы")) { Nom = 2; }
            if (strSist.Contains("Трассы IV группы")) { Nom = 3; }
            if (strSist.Contains("Трассы V группы")) { Nom = 4; }
            if (strSist.Contains("Трассы cиловые")) { Nom = 5; }
            List<Noda> spNod_Vib = MSpNOD[Nom];
            foreach (Noda line in spNod_Vib)
            {
                i = i + 1;
                this.dataGridView7.Rows.Add(i, line.Nom, line.Hoz, line.Param, line.Koord, line.SpSmNod, line.Koord3D, line.Vin);
            }
        }

        public void DGVsist() 
        {
            this.dataGridView8.Rows.Clear();
            string[] MtSist;
            List<string> Sist = new List<string>();
            Sist.Add("Трассы I группы :1:20:1");
            Sist.Add("Трассы II группы :2:2:5");
            Sist.Add("Трассы III группы :3:110:3");
            Sist.Add("Трассы IV группы :4:6:231");
            Sist.Add("Трассы V группы :5:230:4");
            Sist.Add("Трассы cиловые  :6:4:136");
            int i = -1;
            foreach (string TSist in Sist) 
            {
                i = i + 1;
                MtSist = TSist.Split(':');
                this.dataGridView8.Rows.Add(MtSist[0], MtSist[1],MtSist[2], MtSist[3], MSpNOD[i].Count);
                if(MSpNOD[i].Count>0) this.dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
            }
        }//заполнение гирдвьювера с системати трасс
        public void SpisZag(ref List<Zag> spZag,string tZag) 
        {
            Zag nZag = new Zag();
            nZag.NomZagal(tZag);
            spZag.Add(nZag);
        }//пополнение списка загаловков
        public void ShtenTXT_SLov(ref List<string> spNod)
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\МАРШРУТ\Словарь.txt", Encoding.UTF8);
            foreach (string Strok in lines) { spNod.Add(Strok); }
        }//чтение текстового файла с нодами
        public void NasrExcel() 
        {
            string Nastr = HCtenSlov("NastExel", "");
            if (Nastr != "") 
            {
                Nastr = Nastr + ":::::";
                string[] NastM = Nastr.Split(':');
                this.textBox20.Text = NastM[0];
                this.textBox21.Text = NastM[1];
                this.textBox22.Text = NastM[2];
                this.textBox23.Text = NastM[3];
                this.textBox24.Text = NastM[4];
                this.textBox25.Text = NastM[5];
                this.textBox26.Text = NastM[6];
                this.textBox27.Text = NastM[7];
                this.textBox28.Text = NastM[8];
                this.textBox29.Text = NastM[9];
                this.textBox30.Text = NastM[10];
                this.textBox31.Text = NastM[11];
                this.textBox32.Text = NastM[12];
                this.textBox33.Text = NastM[13];
                this.textBox34.Text = NastM[14];
                this.textBox35.Text = NastM[15];
                this.textBox36.Text = NastM[16];
                this.textBox38.Text = NastM[17];
            }
        }

        //Фунуции для проверки радиусов зегиба кабеленй
        public void Zagib(Point3d T1_1, double DelX, double DelY, double DelZ, double rad) 
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            List<ObjectId> ObjId = new List<ObjectId>();
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                Point3d T1_2 = new Point3d(T1_1.X + 100, T1_1.Y, T1_1.Z);
                Point3d T1_4 = new Point3d(T1_1.X + 50, T1_1.Y, T1_1.Z);
                Point3d T2_1 = new Point3d(T1_1.X + DelX, T1_1.Y + DelY, T1_1.Z + DelZ);
                Point3d T2_2 = new Point3d(T1_1.X + DelX, T1_1.Y + DelY, T1_1.Z + DelZ - 100);
                Point3d T2_5 = new Point3d(T1_1.X + DelX, T1_1.Y + DelY, T1_1.Z + DelZ - 50);
                Point3d T1_3 = new Point3d();
                Point3d T2_3 = new Point3d();
                Point3d T2_4 = new Point3d();
                Point3d T3_3 = new Point3d();
                double x = Find_X(T1_1, T1_2, T2_2, rad);
                T1_2 = new Point3d(T1_1.X + 100 + x, T1_1.Y, T1_1.Z);
                T2_2 = new Point3d(T1_1.X + DelX, T1_1.Y + DelY, T1_1.Z + DelZ - 100-x);
                Sopr(T1_1, T1_2, T2_2, rad, ref T1_3, ref T2_3);
                Sopr(T1_2, T2_2, T2_1, rad, ref T2_4, ref T3_3);
                double dist = T2_3.DistanceTo(T2_4);
                Vector3d Vekt2 = (T2_4 - T2_3)/ dist;
                Point3d T2_6 = T2_3 + Vekt2*(dist/2);
                SPLINE(T1_1, T1_3, T1_2, T2_3, T2_4, T2_2, T3_3, T2_1, 0, ref ObjId, T1_4, T2_5, T2_6);
                //KRUG(T1_1,2, ref ObjId);
                //Solid(ObjId[1], ObjId[0]);
            }
        }
        private double Find_X(Point3d T1_1, Point3d T1_2, Point3d T2_2, double rad)
        {
            double a = T1_2.DistanceTo(T2_2);
            double b = T1_1.DistanceTo(T1_2);
            double c = T1_1.DistanceTo(T2_2);
            double cosAlf = (Math.Pow(a, 2) + Math.Pow(b, 2) - Math.Pow(c, 2)) / (2 * a * b);
            double Alf_2 = Math.Acos(cosAlf) / 2;
            double x = rad / Math.Tan(Alf_2);
            return x;
        }//поиск длинны после прямого участка
        static public void Sopr(Point3d T1_1, Point3d T1_2, Point3d T2_2,double rad,ref Point3d T1_3, ref Point3d T2_3)
        {
            double a = T1_2.DistanceTo(T2_2);
            double b = T1_1.DistanceTo(T1_2);
            double c = T1_1.DistanceTo(T2_2);
            double cosAlf = (Math.Pow(a, 2)+ Math.Pow(b, 2)- Math.Pow(c, 2))/(2*a*b);
            double Alf_2 = Math.Acos(cosAlf) / 2;
            double x = rad / Math.Tan(Alf_2);
            Vector3d Vekt1_1 = (T1_2 - T1_1)/ b;
            T1_3 = T1_2 - Vekt1_1 * x;
            Vector3d Vekt2_1 = (T2_2 - T1_2) / a;
            T2_3 = T1_2 + Vekt2_1* x;
        }//поиск точек для сопряжения
        static public void Solid(ObjectId regId, ObjectId splId) 
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                try
                {
                    Entity sweepEnt = tr.GetObject(regId, OpenMode.ForRead) as Entity;
                    Curve pathEnt = tr.GetObject(splId, OpenMode.ForRead) as Curve;
                    if (sweepEnt == null || pathEnt == null)
                    {
                        ed.WriteMessage("\nProblem opening the selected entities.");
                        return;
                    }
                    // We use a builder object to create
                    // our SweepOptions
                    SweepOptionsBuilder sob = new SweepOptionsBuilder();
                    // Align the entity to sweep to the path
                    sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath;
                    // The base point is the start of the path
                    sob.BasePoint = pathEnt.StartPoint;
                    // The profile will rotate to follow the path
                    sob.Bank = true;
                    // Now generate the surface...
                    SweptSurface ss = new SweptSurface();
                    ss.CreateSweptSurface(sweepEnt, pathEnt, sob.ToSweepOptions());
                    // ... and add it to the modelspace
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    ms.AppendEntity(ss);
                    tr.AddNewlyCreatedDBObject(ss, true);
                    tr.Commit();
                }
                catch
                { }
            }
        }//построение 3Д солида выдавливанием


        private void Puchok(Point3d Toch1, ref List<Kab> Vk_Vibr, ref List<Kab> Vk_Postr, string Grup)
        {

            double PlusD = Convert.ToDouble(this.textBox16.Text);
            double Otstup = Convert.ToDouble(this.textBox17.Text);
            double A = Convert.ToDouble(this.textBox14.Text) ;
            double B = Convert.ToDouble(this.textBox13.Text) ;
            double a = 0;
            double b = 0;
            string stIND = "";
            string stDlin = "";
            string stStar = "";
            string stNov = "";
            string stNom = "";
            double douDiam = 0;
            int nom;
            int Schet;
            double Dlin = 0;
            double DlinN = 0;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            // Start a transaction
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                // Dispose of the transaction
                Vk_Vibr.Sort(delegate(Kab x, Kab y) { return y.DIam.CompareTo(x.DIam); });
                List<DiamVsehc> Sehc = new List<DiamVsehc>();
                Transaction tr = acCurDb.TransactionManager.StartTransaction();
                using (tr)
                {
                    int row = 1;
                    int ID = 0;
                    double X;
                    double Y;
                    double X1;
                    double Y1;
                    double Xr;
                    double Yd;
                    double Xl;
                    double Yu;
                    //LWPOIYiTEXT(Toch1, 3, this.textBox13.Text + "x" + this.textBox13.Text, A, B, a, b,5);
                    Xl = Toch1.X;
                    Xr = Xl + A;
                    Yu = Toch1.Y - Otstup;
                    Yd = Yu - B;
                    //Application.ShowAlertDialog(Xl.ToString()+":" + Xr.ToString() + ":" + Yu.ToString() + ":" + Yd.ToString());
                    X = Toch1.X + Vk_Vibr[0].DIam / 2;
                    Y = Toch1.Y - B + Vk_Vibr[0].DIam / 2;
                    Point3d TochT = new Point3d(X, Y, 0);
                    Point3d TochTOts = new Point3d(Toch1.X, Toch1.Y - Otstup, 0);
                    Point3d TochTPr = new Point3d(0, 0, 0);
                    double a1 = 0;
                    double c1 = 0;
                    double r1r3 = 0;
                    double r2r3 = 0;
                    double Delt1 = 0;
                    double Delt2 = 0;
                    string Y_N = "N";
                    string zahc = "N";
                    string Peres = "";
                    string l_r = "";
                    //составление списка с координатами
                    foreach (Kab Ind in Vk_Vibr)
                    {
                        DiamVsehc iInd = new DiamVsehc();
                        iInd.Nind(Ind.IND);
                        iInd.NID(ID);
                        iInd.NMassa(Convert.ToDouble(Ind.Massa));
                        iInd.Nrad(Ind.DIam / 2);
                        iInd.NradK(Ind.DIamKor / 2);
                        if (row > 1)
                        {
                            DiamVsehc Ind1 = Sehc[ID - 1];
                            Y_N = "N";
                            zahc = "N";
                            foreach (DiamVsehc Ind2 in Sehc)
                            {
                                //if (Ind2.ID != ID - 1) TochTPr = Find3T_2Kr_Kos(Ind2, Ind1, Ind.DIam / 2);
                                if (row % 2 == 0) l_r = "r"; else l_r = "l";
                                TochTPr = Find3T_2Kr_Kos(Ind2, Ind1, Ind.DIam / 2, l_r);
                                zahc = "N";
                                foreach (DiamVsehc IndProv in Sehc) { if ((TochTPr.DistanceTo(IndProv.Zentr) - ((Ind.DIam / 2) + IndProv.rad)) < -0.01) { zahc = "Y-" + IndProv.ID + ":" + (TochTPr.DistanceTo(IndProv.Zentr) - ((Ind.DIam / 2) + IndProv.rad)).ToString(); } }
                                if (TochTPr.X - (Ind.DIam / 2) < Xl & (row % 2 == 0) == false) zahc = "за рамкой";
                                a1 = TochTPr.DistanceTo(Ind1.Zentr);
                                c1 = TochTPr.DistanceTo(Ind2.Zentr);
                                r1r3 = Ind1.rad + Ind.DIam / 2;
                                r2r3 = Ind2.rad + Ind.DIam / 2;
                                Delt1 = Math.Abs(a1 - r1r3);
                                Delt2 = Math.Abs(c1 - r2r3);
                                if (Delt1 < 0.01 & Delt2 < 0.01 & zahc == "N") { TochT = TochTPr; }
                            }
                            if (TochT.X - Ind.DIam / 2 < Xl & l_r == "r")
                            {
                                X = Xl + Ind.DIam / 2;
                                Y = Sehc[ID - 1].Zentr.Y + Math.Pow((Math.Pow((Ind.DIam / 2 + Sehc[ID - 1].rad), 2) - Math.Pow((Sehc[ID - 1].Zentr.X - (Xl + Ind.DIam / 2)), 2)), 0.5);
                                TochT = new Point3d(X, Y, 0);
                                row = row + 1;
                            }
                            if (TochT.X + Ind.DIam / 2 > Xr & l_r == "l")
                            {
                                X = Xr - Ind.DIam / 2;
                                Y = Sehc[ID - 1].Zentr.Y + Math.Pow((Math.Pow((Ind.DIam / 2 + Sehc[ID - 1].rad), 2) - Math.Pow(((Xr - Sehc[ID - 1].Zentr.X) - Ind.DIam / 2), 2)), 0.5);
                                TochT = new Point3d(X, Y, 0);
                                row = row + 1;
                            }
                            if (TochT.Y + Ind.DIam / 2 > Yu) break;
                        }
                        if (row == 1 & ID > 0)
                        {
                            X = TochT.X + Math.Pow((Math.Pow((Ind.DIam / 2 + Sehc[ID - 1].rad), 2) - Math.Pow((Ind.DIam / 2 - Sehc[ID - 1].rad), 2)), 0.5);
                            Y = Toch1.Y - B + Ind.DIam / 2;
                            if (X + Ind.DIam / 2 < Xr)
                                TochT = new Point3d(X, Y, 0);
                            else
                            {
                                X = Xr - Ind.DIam / 2;
                                Y = TochT.Y + Math.Pow((Math.Pow((Ind.DIam / 2 + Sehc[ID - 1].rad), 2) - Math.Pow(((Xr - Sehc[ID - 1].Zentr.X) - Ind.DIam / 2), 2)), 0.5);
                                TochT = new Point3d(X, Y, 0);
                                row = row + 1;
                            }
                        }
                        iInd.NZentr(TochT);
                        //Application.ShowAlertDialog("приготовилась удалять");
                        Vk_Postr.Add(Ind);
                        Sehc.Add(iInd);
                        ID = ID + 1;
                    }
                    //построение окружностей 
                    double SumPl = 0;
                    double SumMass = 0;
                    foreach (DiamVsehc Ind in Sehc)
                    {
                        KRUGiTEXT(Ind.Zentr, 0, Ind.ind, "", Ind.radK, 0.8);
                        SumPl = SumPl + Math.Pow(Ind.radK, 2) * Math.PI;
                        SumMass = SumMass + Ind.Massa;
                    }
                    LWPOIYiTEXT(Toch1, 3, "Сечение -" + this.textBox13.Text + "x" + this.textBox14.Text + " Кабели-" + Grup + " группы. Сумарная площадь кабелей-" + (SumPl/100).ToString("#0.") + "кв.см" + " . Нагрузка-" + (SumMass * 10).ToString("#0.") + "н/м", A, B, a, b, 2, "Выноски");
                    LWPOIYiTEXT(TochTOts, 3, "", A, B - Otstup, a, b, 2, "Выноски");
                    tr.Commit();
                }
            }

        }//построение пучков кабелей
        private Point3d Find3T_2Kr_Kos(DiamVsehc iInd1, DiamVsehc iInd2, double r3, string l_r)
        {
            double Ugol2;
            Point3d XYZ1 = iInd1.Zentr;
            Point3d XYZ2 = iInd2.Zentr;
            double r1 = iInd1.rad;
            double r2 = iInd2.rad;
            double Dx = XYZ1.X - XYZ2.X;
            double Dy = XYZ1.Y - XYZ2.Y;
            double tan = Dy / Dx;
            double Ugol1 = Math.Atan(tan);
            if (Ugol1 < 0) Ugol1 = Ugol1 + Math.PI;
            if (l_r == "r")
                Ugol2 = Ugol1 + Math.PI / 2;
            else
                Ugol2 = Ugol1 + 3 * Math.PI / 2;
            double a = r3 + r2;
            double b = r3 + r1;
            double c = XYZ1.DistanceTo(XYZ2);
            double cosalf = (Math.Pow(b, 2) + Math.Pow(c, 2) - Math.Pow(a, 2)) / (2 * b * c);
            double rast1 = cosalf * b;
            double rast2 = b * Math.Sqrt((1 - Math.Pow(cosalf, 2)));
            Point3d XYZ4 = polar(XYZ1, Ugol1, rast1);
            Point3d XYZ3 = polar(XYZ4, Ugol2, rast2);
            return XYZ3;
        }//поиск координаты центра круга касающегося двух
        public Point3d polar(Point3d XYZ1, double ugol, double Rast)
        {
            double X = Rast * Math.Cos(ugol);
            double Y = Rast * Math.Sin(ugol);
            Point3d XYZ2 = new Point3d(XYZ1.X + X, XYZ1.Y + Y, XYZ1.Z);
            return XYZ2;
        }//поиск точки отстаящей от заданной на растояние и угол заданный 


        public void SpVistTP(ref List<TPodk> SpNOD) 
        {
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
                            TNOD.NKoor(ln.Center);
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) {TNOD.NIND(value.Value.ToString());}
                                    Schet = Schet + 1;
                                    if (Schet > 1) { break;}
                                }
                            }
                            TNOD.NKoorMod(FIND_XYZmir_2D(SpPLOS, ln.Center));
                            SpNOD.Add(TNOD);
                        }
                    }
                    tr.Commit();
                }
            }
        }//создание списка точек подключения выставленых в виде окружностей
        public void SpVistTP_BL(ref List<TPodk> SpNOD)
        {
            string Ind = "";
            string Pom = "";
            string Compon = "";
            string Shem = "";
            string Visot = "0";
            string Vid = "";
            double UgolX = 0;
            double UgolY = 0;
            double UgolZ = 0;
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[7];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 2);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 3);
            acTypValAr.SetValue(new TypedValue(8, "Насыщение"), 4);
            acTypValAr.SetValue(new TypedValue(8, "НасыщениеСкрытое"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    //Zoom(new Point3d(db.Limmin.X, db.Limmin.Y, 0),
                    //new Point3d(db.Limmax.X, db.Limmax.Y, 0),
                    //new Point3d(), 1);
                    Editor ed1 = Application.DocumentManager.MdiActiveDocument.Editor;
                    //TPodk TNOD = new TPodk();
                    //Application.ShowAlertDialog(acSSet.Count.ToString());
                    foreach (ObjectId sobj in acSSet.GetObjectIds())
                    {
                        Ind = "";
                        Pom = "";
                        Compon = "";
                        Shem = "";
                        Visot = "0";
                        Vid = "";
                        UgolX = 0;
                        UgolY = 0;
                        UgolZ = 0;
                        BlockReference ln = tr.GetObject(sobj, OpenMode.ForWrite) as BlockReference;
                        if (ln != null)
                        {
                            //BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                            UgolZ = ln.Rotation;
                            if (ln.IsDynamicBlock)
                            {
                                DynamicBlockReferencePropertyCollection props = ln.DynamicBlockReferencePropertyCollection;
                                foreach (DynamicBlockReferenceProperty prop in props)
                                {
                                    object[] values = prop.GetAllowedValues();
                                    if (prop.PropertyName == "Видимость1") { Vid = prop.Value.ToString(); }
                                    if (prop.PropertyName == "Угол1") { UgolZ = Convert.ToDouble(prop.Value.ToString()); }
                                }
                            }
                            ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                            if (buffer != null)
                            {
                                Schet = 0;
                                foreach (TypedValue value in buffer)
                                {
                                    if (Schet == 1) { Ind = (value.Value.ToString()); }
                                    Schet = Schet + 1;
                                    if (Schet > 1) { break; }
                                }
                            }
                            foreach (ObjectId idAtrRef in ln.AttributeCollection)
                            {
                                using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                {
                                    if (atrRef != null)
                                    {
                                        if (atrRef.Tag == "Примечание" | atrRef.Tag == "ПРИМЕЧАНИЕ") { if(atrRef.TextString!="") Ind = atrRef.TextString; }
                                        if (atrRef.Tag == "Индекс") { if (atrRef.TextString != "") Ind = atrRef.TextString; }
                                        if (atrRef.Tag == "НОМЕР_ВЫНОСКИ") { Ind = atrRef.TextString; }
                                        if (atrRef.Tag == "Помещение") { Pom = atrRef.TextString; }
                                        if (atrRef.Tag == "ОКП") { Compon = atrRef.TextString; }
                                        if (atrRef.Tag == "Схема") { Shem = atrRef.TextString; }
                                        if (atrRef.Tag == "Исполнение") { Compon = atrRef.TextString; }
                                        if (atrRef.Tag == "Высота_установки") { Visot = atrRef.TextString; }
                                    }
                                }
                            }
                            if (Ind != "")
                            {
                                //Application.ShowAlertDialog(Ind);
                                if (Compon.Contains("#")) { Compon = Compon.Split('#').Last(); }
                                TPodk TNOD = new TPodk();
                                TNOD.NIND(Ind);
                                TNOD.NPom(Pom);
                                TNOD.NNaim(Compon);
                                TNOD.NKoor(ln.Position);
                                TNOD.NShem(Shem);
                                if (Vid == "Вид сверху") UgolX = 1.57;
                                if (Vid == "Вид с боку") UgolY = 1.57;
                                if (Visot != "") TNOD.NVisT(Convert.ToDouble(Visot));
                                TNOD.NPovorot(new Point3d(UgolX, UgolY, UgolZ));
                                SpNOD.Add(TNOD);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            //Application.ShowAlertDialog(SpNOD.Count.ToString());
        }//создание списка точек подключения выставленых в виде блоков
        public void SpVistVin(ref List<TPodk> SpVIN)
        {
            int Schet = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[6];
            acTypValAr.SetValue(new TypedValue(0, "CIRCLE"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Переходы"), 2);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            acTypValAr.SetValue(new TypedValue(40, 6.0), 5);
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
                            TNOD.NKoor(ln.Center);
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
        public void SpVistVinDB(ref List<TPodk> SpVIN)
        {
            string Nazv_Vin = "", Sprav_vin = "";
            string Nazv_per = "", Sprav_per = "";
            string SUF = "";
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            Point3d BP = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[7];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Переходы"), 2);
            acTypValAr.SetValue(new TypedValue(8, "Выноски"), 3);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 4);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
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
                            SUF = "";
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
                                    if (prop.PropertyName == "Положение3 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                                    if (prop.PropertyName == "Положение3 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
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
                                        if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА1") { Nazv_per = atrRef.TextString; BP = new Point3d(BP.X+ x1, BP.Y+y1, BP.Z);  }
                                        if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА") { Nazv_per = atrRef.TextString; }
                                        if (atrRef.Tag == "СУФИКС") { SUF = atrRef.TextString; }
                                    }
                                }
                            }
                            if (Nazv_Vin != "")
                            {
                                TPodk TNOD = new TPodk();
                                if (SUF != "" & Nazv_Vin.Length > SUF.Length) Nazv_Vin = Nazv_Vin.Substring(0, Nazv_Vin.Length - SUF.Length);
                                TNOD.NIND(Nazv_Vin);
                                TNOD.NShem(Nazv_Vin + SUF);
                                TNOD.NKoor(BP);
                                TNOD.NSist(SUF);
                                SpVIN.Add(TNOD);
                            }
                            if (Nazv_per != "")
                            {
                                TPodk TNOD = new TPodk();
                                if (SUF != "" & Nazv_per.Length > SUF.Length) Nazv_per = Nazv_per.Substring(0, Nazv_per.Length - SUF.Length);
                                TNOD.NIND(Nazv_per);
                                TNOD.NShem(Nazv_per + SUF);
                                TNOD.NKoor(BP);
                                TNOD.NSist(SUF);
                                if (SpVIN.Exists(x => x.IND == Nazv_per)==false) SpVIN.Add(TNOD);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            //Application.ShowAlertDialog(SpVIN.Count.ToString());
        }//создание списка выносок динБлоками
        //public void SpVistVinDB(ref List<TPodk> SpVIN)
        //{
        //    string Nazv_Vin = "", Sprav_vin = "";
        //    string Nazv_per = "", Sprav_per = "";
        //    double x1 = 0;
        //    double y1 = 0;
        //    double x2 = 0;
        //    double y2 = 0;
        //    double xMir = 0;
        //    double yMir = 0;
        //    double zMir = 0;
        //    Point3d BP = new Point3d();
        //    int Schet = 0;
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;
        //    ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
        //    ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
        //    TypedValue[] acTypValAr = new TypedValue[7];
        //    acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
        //    acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
        //    acTypValAr.SetValue(new TypedValue(8, "Переходы"), 2);
        //    acTypValAr.SetValue(new TypedValue(8, "Выноски"), 3);
        //    acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 4);
        //    acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 5);
        //    acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
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
        //                //Circle ln = tr.GetObject(sobj, OpenMode.ForWrite) as Circle;
        //                //TPodk TNOD = new TPodk();
        //                BlockReference bref = tr.GetObject(sobj, OpenMode.ForRead) as BlockReference;
        //                if (bref != null)
        //                {
        //                    Nazv_Vin = "";
        //                    Sprav_vin = "";
        //                    //для переходов
        //                    Nazv_per = "";
        //                    Sprav_per = "";
        //                    //для плоскостей
        //                    y1 = 0;
        //                    x2 = 0;
        //                    y2 = 0;
        //                    BP = bref.Position;
        //                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
        //                    {
        //                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
        //                        {
        //                            if (atrRef != null)
        //                            {
        //                                if (atrRef.Tag == "НОМЕР_ВЫНОСКИ") { Nazv_Vin = atrRef.TextString; }
        //                                if (atrRef.Tag == "НОМЕР_ПРОХОДКИ") { Nazv_Vin = atrRef.TextString; }
        //                                if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА") { Nazv_Vin = atrRef.TextString; }
        //                                if (atrRef.Tag == "Справочная_информация") { Sprav_vin = atrRef.TextString; }
        //                                //переходы
        //                                //if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА1") { Nazv_per = atrRef.TextString; }
        //                            }
        //                        }
        //                    }
        //                    if (Nazv_Vin != "")
        //                    {
        //                        TPodk TNOD = new TPodk();
        //                        TNOD.NIND(Nazv_Vin);
        //                        TNOD.NKoor(BP);
        //                        SpVIN.Add(TNOD);
        //                    }
        //                    if (Nazv_per != "")
        //                    {
        //                        Point3d T1 = new Point3d(BP.X + x1, BP.Y + y1, 0);
        //                        Point3d T2 = new Point3d(BP.X + x2, BP.Y + y2, 0);
        //                        TPodk TNOD1 = new TPodk();
        //                        TNOD1.NIND(Nazv_per);
        //                        TNOD1.NKoor(T1);
        //                        SpVIN.Add(TNOD1);
        //                        TPodk TNOD2 = new TPodk();
        //                        TNOD2.NIND(Nazv_per);
        //                        TNOD2.NKoor(T2);
        //                        SpVIN.Add(TNOD2);
        //                    }
        //                }
        //            }
        //            tr.Commit();
        //        }
        //    }
        //}//создание списка выносок динБлоками


        public void SpVistVinBP_DB(ref List<TPodk> SpVIN)
        {
            string Nazv_Vin = "", Sprav_vin = "";
            string Nazv_per = "";
            string SUF = "";
            double IDpust = 0;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            Point3d BP = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
            TypedValue[] acTypValAr = new TypedValue[7];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Переходы"), 2);
            acTypValAr.SetValue(new TypedValue(8, "Выноски"), 3);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 4);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
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
                            Nazv_per = "";
                            Sprav_vin = "";
                            SUF = "";
                            //для переходов        
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
                                    if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                                    if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
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
                                        if (atrRef.Tag == "ID_Вын") { Nazv_Vin = atrRef.TextString; }
                                        //переходы
                                        if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА") { Nazv_per = atrRef.TextString; }
                                        if (atrRef.Tag == "СУФИКС") { SUF = atrRef.TextString; }
                                        if (atrRef.Tag == "ID_перех") { Nazv_per = atrRef.TextString; }
                                    }
                                }
                            }
                            if ( Nazv_Vin != "")
                            {
                                TPodk TNOD = new TPodk();
                                if (SUF != "" & Nazv_Vin.Length > SUF.Length) Nazv_Vin = Nazv_Vin.Substring(0, Nazv_Vin.Length - SUF.Length);
                                TNOD.NIND(Nazv_Vin);
                                TNOD.NShem(Nazv_Vin + SUF);
                                TNOD.NKoor(BP);
                                TNOD.NKoor1(new Point3d(BP.X + x1, BP.Y + y1, BP.Z));
                                TNOD.NSist(SUF);
                                SpVIN.Add(TNOD);
                            }
                            if ( Nazv_per != "")
                            {
                                TPodk TNOD = new TPodk();
                                if (SUF != "" & Nazv_per.Length > SUF.Length) Nazv_per = Nazv_per.Substring(0, Nazv_per.Length - SUF.Length);
                                //Application.ShowAlertDialog(Nazv_per + " " + SUF);
                                TNOD.NIND(Nazv_per);
                                TNOD.NShem(Nazv_per + SUF);
                                TNOD.NKoor(BP);
                                TNOD.NKoor1(new Point3d(BP.X + x1, BP.Y + y1, BP.Z));
                                TNOD.NSist(SUF);
                                SpVIN.Add(TNOD);
                            }
                            //if ((Nazv_per == "" & (x1 != 0 & y1 != 0)) | (Nazv_Vin == "" & (x1 != 0 & y1 != 0)))
                            //{
                            //    //Application.ShowAlertDialog(Nazv_per + " " + Nazv_Vin);
                            //    IDpust = IDpust + 1;
                            //    TPodk TNOD = new TPodk();
                            //    TNOD.NIND(IDpust.ToString());
                            //    TNOD.NShem(IDpust.ToString() + SUF);
                            //    TNOD.NKoor(BP);
                            //    TNOD.NKoor1(new Point3d(BP.X + x1, BP.Y + y1, BP.Z));
                            //    TNOD.NSist(SUF);
                            //    SpVIN.Add(TNOD);
                            //}
                        }
                    }
                    tr.Commit();
                }
            }
            //Application.ShowAlertDialog(SpVIN.Count.ToString());
        }//создание списка выносок динБлоками

        public void ZagrDannNOD() 
        {
            int i = -1;
            string strSist = "";
            string strAdrRis = "";
            string[] Grupp = { "I", "II", "III", "IV", "V", "Силовые" };
            foreach (string Zvet in Grupp)
            {
                List<Noda> spNodT = new List<Noda>();
                i = i + 1;
                spNod.Clear();
                strAdrRis = @"C:\МАРШРУТ\Трассы " + Zvet + " группы.txt";
                if (System.IO.File.Exists(strAdrRis))
                {ShtenTXT(ref spNodT, "Трассы " + Zvet + " группы");}
                MSpNOD[i] = spNodT;
            }
            //i = -1;
            //foreach (string Zvet in Grupp) 
            //{
            //    i = i + 1;
            //    this.listBox2.Items.Add("Трассы " + Zvet + " группы " + MSpNOD[i].Count + " узлов");
            //}
        }//загрузка узлов из файла
        public void ZagrDannNODisSLOV()
        {
            int i = -1;
            string strSist = "";
            string strAdrRis = "";
            string[] Grupp = {"I", "II", "III", "IV", "V", "Силовые" };
            foreach (string Zvet in Grupp)
            {
                List<Noda> spNodT = new List<Noda>();
                i = i + 1;
                spNod.Clear();
                ShtenSLOV_NOD(ref spNodT, "Трассы" + Zvet + "группы"); 
                MSpNOD[i] = spNodT;
            }
            DGVsist();
        }//загрузка узлов из словоря


        public void RazvOPoSp(ZaPis Zap) 
        {
            int NSist=0;
            string Rez = "Есть";
            List<DUGA> DUGI = new List<DUGA>();
            List<Noda> OPEN = new List<Noda>();
            List<Noda> CLOSE = new List<Noda>();
            List<Noda> SpNODT2 = new List<Noda>();
            Noda T1 = new Noda();         
            Noda TNOD = new Noda();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Transaction tr = db.TransactionManager.StartTransaction();
            double Zvet = 6;
            NSist=Convert.ToInt16(Zap.Sist)-1;
            spNod = MSpNOD[NSist];
            if (spNod.Count == 0)  Application.ShowAlertDialog("Не создана система трасс");
            if (Zap.BlokNOD != "")
            {
                string[] strBLOKNODm = Zap.BlokNOD.Split(',');
                foreach (string TVin in strBLOKNODm) UdalSvasi(ref spNod, TVin);
            }
            using (tr)
            {
                T1.NomNod("Далеко от трасс");
                BlizhNOD(Zap.INDotk.Koord, spNod, ref T1);
                string strSpNODT2 = "";
                foreach (PodvKab Kab in Zap.SpVixKab) 
                {
                Noda T2 = new Noda();
                T2.NomNod("Далеко от трасс");
                BlizhNOD(Kab.KoordObor, spNod, ref T2);
                strSpNODT2 = strSpNODT2 + "," + T2.Nom;
                if (T1.Nom != T2.Nom)
                    SpNODT2.Add(T2);
                else
                { 
                    T2.NomNod("Первая и последняя точки совподают");
                    SpNODT2.Add(T2);
                }
                }
                T1.NVes(0);
                T1.NomNOtk("-");
                TNOD = T1;
                string strTNOD = TNOD.Nom;
                CLOSE.Add(TNOD);
                    while ((Rez == "Есть")&(T1.Nom!="Далеко от трасс"))
                    {
                        spDUG(ref DUGI, ref TNOD, ref OPEN, ref CLOSE, Zvet);
                        if (DUGI.Count > 0)
                        {
                            Relaks(ref DUGI, ref TNOD, ref OPEN);
                            SlNOD(ref DUGI, ref TNOD, ref OPEN, ref CLOSE, ref strTNOD);
                            if (strTNOD == "") { Rez = "Нет"; }
                        }
                        else
                        {
                            SlNOD(ref DUGI, ref TNOD, ref OPEN, ref CLOSE, ref strTNOD);
                            if (strTNOD == "") { Rez = "Нет"; }
                        }
                        DUGI.Clear(); 
                    }
                    int shetKab = 0;
                    foreach (Noda TN in SpNODT2) 
                    {
                        Rezult = "+";
                        POSTR(ref CLOSE, T1, TN, Zap.INDotk.Koord, Zap.SpVixKab[shetKab].KoordObor, Zap.SpVixKab[shetKab].IND, Zap.SpVixKab[shetKab].DIam.ToString(), ref Rezult, Zap.Sist, Zap.SpVixKab[shetKab].Massa);
                        if (T1.Nom == "Далеко от трасс" ) { Rezult = Zap.INDotk.IND +"-далеко от трасс"; }
                        if (TN.Nom == "Далеко от трасс") { Rezult = Zap.SpVixKab[shetKab].INDobor + "-далеко от трасс"; }
                        if (TN.Nom == "Первая и последняя точки совподают") { Rezult = Zap.SpVixKab[shetKab].INDobor + "-Первая и последняя точки совподают"; }
                        ed.WriteMessage(Zap.SpVixKab[shetKab].IND + "-" + Rezult);
                        ITOGI.Add(Zap.SpVixKab[shetKab].IND + ":" + Rezult); 
                        shetKab = shetKab + 1;          
                    }
                    tr.Commit();  
            }
        }//разводка по списку
        public void BlizhNOD(Point3d Toch, List<Noda> spNod, ref Noda blNOD) //поиск ближайшего нода
        {
            double MinDist = Convert.ToDouble(this.textBox1.Text);
            foreach (Noda TNOD in spNod) { if (Toch.DistanceTo(TNOD.Koord) < MinDist) { blNOD = TNOD; MinDist = Toch.DistanceTo(TNOD.Koord); } }
        }
        public void spDUG(ref List<DUGA> DUGI, ref Noda TNOD, ref List<Noda> OPEN, ref List<Noda> CLOSE, double Zvet)
        {
            string[] SpSmNOD = TNOD.SpSmNod.Split(',');
            Point3d KoorTN = TNOD.Koord3D;
            foreach (string NTSmNoda in SpSmNOD)
            {
                string[] tNoda = NTSmNoda.Split('*');
                //if (OPEN.Exists(x => x.Nom == tNoda[0]) == false & CLOSE.Exists(x => x.Nom == tNoda[0]) == false & spNod.Exists(x => x.Nom == tNoda[0]))
                if (CLOSE.Exists(x => x.Nom == tNoda[0]) == false)
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
        public void SlNOD(ref List<DUGA> DUGI, ref Noda TNOD, ref List<Noda> OPEN, ref List<Noda> CLOSE, ref   string strTNOD)
        {
            double minVes = 9999999999999.0;
            strTNOD = "";
            foreach (Noda TVNod in OPEN)
            {
                if (TVNod.Ves <= minVes)
                {
                    minVes = TVNod.Ves;
                    TNOD = TVNod;
                    strTNOD = TNOD.Nom;
                }
            }
            CLOSE.Add(TNOD);
            OPEN.Remove(TNOD);
        }//переход на слкдующий нод
        public void POSTR(ref List<Noda> CLOSE, Noda T1, Noda T2, Point3d Toch1, Point3d Toch2, string strInd, string strDiam,ref string Rezult, string Sist, string Massa)
        {
            Point3dCollection TkoorPL = new Point3dCollection();
            Point3d pToh;
            TkoorPL.Add(Toch2);
            string strHOZ;
            double dDlinKab = 0;
            int iNom = 0;
            string NTTohc = T2.Nom;
            Noda TNod = T1;
            pToh = Toch2;
            if (CLOSE.Exists(x => x.Nom == NTTohc) == true)
            {   
                TNod = CLOSE.Find(x => x.Nom == NTTohc);
                TkoorPL.Add(TNod.Koord);
                NTTohc = TNod.Otkuda;
                strHOZ = TNod.Nom;
            }
            else
            {
                TkoorPL.Add(Toch1);
                Rezult = "разрыв";
                FPoly2d(TkoorPL, strInd, strDiam, 0, dDlinKab, iNom, Sist, "KabeliError", Massa);
                return;
            }
            while (NTTohc != T1.Nom)
            {
                TNod = CLOSE.Find(x => x.Nom == NTTohc);
                if (TNod.Nom.Contains("Переход") == true & strHOZ.Contains("Переход") == true)
                {
                    FPoly2d(TkoorPL, strInd, strDiam, TNod.DlinP, dDlinKab, iNom, Sist, "Kabeli", Massa);
                    iNom = iNom + 1;
                    dDlinKab = 0;
                    TkoorPL.Clear();
                    TkoorPL.Add(TNod.Koord);
                    NTTohc = TNod.Otkuda;
                    strHOZ = TNod.Nom;
                }
                else
                {
                    TkoorPL.Add(TNod.Koord);
                    NTTohc = TNod.Otkuda;
                    strHOZ = TNod.Nom;
                }
            }
            TNod = CLOSE.Find(x => x.Nom == NTTohc);
            TkoorPL.Add(TNod.Koord);
            pToh = TNod.Koord;
            TkoorPL.Add(Toch1);
            FPoly2d(TkoorPL, strInd, strDiam, 0, dDlinKab, iNom, Sist, "Kabeli", Massa);
        }//построение кабеля 
        public void FPoly2d(Point3dCollection TkoorPL, string strInd, string strDiam, double dDlin, double dDlinKab, int nom, string Sist, string Sloi, string Massa)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            int CollorInd = 0;
            if (Sist=="1") {CollorInd = 1;}
            if (Sist == "2") { CollorInd = 5;}
            if (Sist == "3") { CollorInd = 3; }
            if (Sist == "4") { CollorInd = 231; }
            if (Sist == "5") { CollorInd = 4; }
            if (Sist == "6") { CollorInd = 136; }
            // Append the point to the database
            using (tr1)
            {
                Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = CollorInd;
                poly.Layer = Sloi;
                int i = 0;
                foreach (Point3d pt in TkoorPL)
                {
                    poly.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
                    i = i + 1;
                }
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);
                poly.XData = new ResultBuffer(
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, strInd),
                new TypedValue(1000, strDiam),
                new TypedValue(1040, dDlin),
                new TypedValue(1040, dDlinKab),
                new TypedValue(1000, Convert.ToString(nom)),
                new TypedValue(1040, Massa)
                //new TypedValue(1040, Convert.ToDouble(strVISk)),
                );
                btr.Dispose();
                tr1.Commit();
            }
        }//построение полилинии по списку координат и запись расширеных данных
        
        
        public void SozdSpKab(ref List<string> Vkrazv)
        {
            //List<string> Vkrazv=new List<string>();
            string stIND = "";
            string stDlin = "";
            string stStar = "";
            string stNov = "";
            string stNom = "";

            int nom;

            double Dlin = 0;
            double DlinN = 0;

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
                        var KolV = ln.NumberOfVertices;
                        Dlin = ln.Length/1000;
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
                            if (Vkrazv.Exists(x => x.Split('*')[0] == stIND) == true)
                            {
                                nom = Vkrazv.FindIndex(x => x.Split('*')[0] == stIND);
                                stStar = Vkrazv.Find(x => x.Split('*')[0] == stIND);
                                DlinN = Convert.ToDouble(stStar.Split('*')[1]) + Dlin;
                                stNov = stIND + "*" + DlinN.ToString("#0.") + "*" + stStar.Split('*')[2] + "," + stNom;
                                Vkrazv[nom] = stNov;
                            }
                            else
                                Vkrazv.Add(stIND + "*" + Dlin.ToString("#0.") + "*" + stNom);
                        }
                    }
                    tr.Commit();
                }
            }
        }//создания списка построеных кабелей
        public void SozdSpKV(ref List<string> Vkrazv)
        {
            string stIND = "";
            string stStar = "";
            string stNov = "";
            string stNom = "";
            string SpVinKab = "";
            int nom;

            double Dlin = 0;
            double DlinN = 0;

            List<TPodk> SpVin = new List<TPodk>();
            SpVistVin(ref SpVin);
            SpVistVinDB(ref SpVin);

            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "POLYLINE"), 1);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 2);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 3);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Kabeli"), 4);
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
                        Curve ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Curve;
                        //var KolV = ln.NumberOfVertices;
                        Dlin = ln.GetDistanceAtParameter(ln.EndParam) / 1000;
                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stIND = value.Value.ToString(); }
                                if (Schet == 5) { stNom = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                            if (Vkrazv.Exists(x => x.Split('*')[0] == stIND) == true)
                            {
                                nom = Vkrazv.FindIndex(x => x.Split('*')[0] == stIND);
                                stStar = Vkrazv.Find(x => x.Split('*')[0] == stIND);
                                DlinN = Convert.ToDouble(stStar.Split('*')[1]) + Dlin;
                                SpVinKab = stStar.Split('*')[2];
                                fSpVin(ln, SpVin, ref SpVinKab);
                                stNov = stIND + "*" + DlinN.ToString("#0.") + "*" + SpVinKab ;
                                Vkrazv[nom] = stNov;
                            }
                            else
                            {
                                SpVinKab = "";
                                fSpVin(ln, SpVin, ref SpVinKab);
                                Vkrazv.Add(stIND + "*" + Dlin.ToString("#0.") + "*" + SpVinKab );
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создания списка кабелей с маршрутами 

        public void fSpVin(Curve ln, List<TPodk> LSpVin, ref string SpVin)
        {
            //string SpVin = "";
            double DistA_B;
            double DistA_C;
            double DistC_B;
            double DELT;
            var KolV = ln.EndParam;
            //for (double i = 0; i < KolV; i++)
            for (double i = KolV; i > 0; i--)
                {
                Point3d TOtr1 = ln.GetPointAtParameter(i);
                //Point3d TOtr2 = ln.GetPointAtParameter(i + 1);
                Point3d TOtr2 = ln.GetPointAtParameter(i - 1);
                DistA_B = TOtr1.DistanceTo(TOtr2);
                foreach (TPodk TV in LSpVin)
                {
                    DistA_C = TOtr1.DistanceTo(TV.Koord);
                    DistC_B = TV.Koord.DistanceTo(TOtr2);
                    DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                    if (DELT < 0.5) { if (SpVin.Contains(TV.IND + "-") == false & SpVin != TV.IND) SpVin = SpVin + TV.IND + "-"; }
                }
            }
            //SpVin = SpVin.TrimEnd('-');
        }//public string fSpVin(Polyline ln, List<TPodk> LSpVin) 
        public void SozdSpKV_n_k(ref List<string> Vkrazv)
        {
            string stIND = "";
            string stStar = "";
            string stNov = "";
            string stNom = "";
            string SpVinKab = "";

            int nom;

            double Dlin = 0;
            double DlinN = 0;

            List<TPodk> SpVin = new List<TPodk>();
            SpVistVinDB(ref SpVin);

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
                        Dlin = ln.Length / 1000;
                        ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                        if (buffer != null)
                        {
                            Schet = 0;
                            foreach (TypedValue value in buffer)
                            {
                                if (Schet == 1) { stIND = value.Value.ToString(); }
                                if (Schet == 5) { stNom = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                            if (Vkrazv.Exists(x => x.Split('*')[0] == stIND) == true)
                            {
                                nom = Vkrazv.FindIndex(x => x.Split('*')[0] == stIND);
                                stStar = Vkrazv.Find(x => x.Split('*')[0] == stIND);
                                SpVinKab = stStar.Split('*')[1];
                                Point3d Tn = ln.StartPoint;
                                Point3d Tk = ln.EndPoint;
                                stNov = stIND + "*" + SpVinKab + ":" + FIND_XYZmir_2D(SpPLOS, Tn) + ":" + FIND_XYZmir_2D(SpPLOS, Tk);
                                Vkrazv[nom] = stNov;
                            }
                            else
                            {
                                Point3d Tn = ln.StartPoint;
                                Point3d Tk = ln.EndPoint;
                                Vkrazv.Add(stIND + "*" + FIND_XYZmir_2D(SpPLOS, Tn) + ":" + FIND_XYZmir_2D(SpPLOS, Tk));
                            }
                        }
                    }
                    tr.Commit();
                }
            }
        }//создания списка кабелей с маршрутами 
        public void Sozd_TB(ref List<TPodk> SpVin)
        {
            string stIND = "";
            string stNom = "";
            string SpVinKab = "";

            int nom;
            int nom1;
            int Schet;

            double Dlin = 0;

            //SpVin.Sort(delegate(TPodk x, TPodk y) { return Convert.ToDouble(x.IND).CompareTo(Convert.ToDouble(y.IND)); });

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
                                if (Schet == 5) { stNom = value.Value.ToString(); }
                                Schet = Schet + 1;
                            }
                            SpVinKab = "";
                            fSpVin(ln, SpVin, ref SpVinKab);
                            //Application.ShowAlertDialog(stIND + " " + SpVinKab);
                            foreach (string MT in SpVinKab.Split('-'))
                            {
                                if (MT != "")
                                {
                                    nom = SpVin.FindIndex(x => x.Shem == MT);
                                    nom1 = SpVin.FindLastIndex(x => x.Shem == MT);
                                    TPodk stTPodk = SpVin[nom];
                                    TPodk stTPodk1 = SpVin[nom1];
                                    if (stTPodk.BlNOD != null & stTPodk1.BlNOD != null)
                                    { if (stTPodk.BlNOD.Contains("," + stIND) == false & stTPodk1.BlNOD.Contains("," + stIND) == false) { stTPodk.NBlNOD(stTPodk.BlNOD + "," + stIND); stTPodk1.NBlNOD(stTPodk1.BlNOD + "," + stIND); } }
                                    else
                                    { stTPodk.NBlNOD(stTPodk.BlNOD + "," + stIND); stTPodk1.NBlNOD(stTPodk1.BlNOD + "," + stIND); }
                                    SpVin[nom] = stTPodk;
                                    SpVin[nom1] = stTPodk1;
                                }
                                //if (MT != "")
                                //{
                                //    //Application.ShowAlertDialog(MT);
                                //    nom = SpVin.FindIndex(x => x.Shem == MT);
                                //    TPodk stTPodk = SpVin[nom];
                                //    if (stTPodk.BlNOD != null)
                                //    { if (stTPodk.BlNOD.Contains("," + stIND) == false) stTPodk.NBlNOD(stTPodk.BlNOD + "," + stIND); }
                                //    else
                                //        stTPodk.NBlNOD(stTPodk.BlNOD + "," + stIND);
                                //    SpVin[nom] = stTPodk;
                                //}
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            //Application.ShowAlertDialog(SpVin.Count.ToString());
        }//создания перечня выносок
        //public void Postr_TB(ref List<TPodk> SpVin)
        //{
        //    TpodkVKTXTfail(SpVin, "ТБ");
        //    List<TPodk> SpVinSUF;
        //    double A = 430;
        //    double B = 110;
        //    double a = 40;
        //    double b = 80;
        //    double List = 2;
        //    //string NONDOC = this.textBox5.Text;
        //    string NOM = this.textBox5.Text;
        //    string NAME = this.textBox6.Text;
        //    string ISP = this.textBox7.Text;
        //    string PROV = this.textBox8.Text;
        //    string VIP = this.textBox9.Text;
        //    string TKONTR = this.textBox10.Text;
        //    string NKONTR = this.textBox11.Text;
        //    string UTV = this.textBox12.Text;
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Transaction tr = db.TransactionManager.StartTransaction();
        //    List<string> spIndL = new List<string>();
        //    List<string> spSuf = new List<string>();
        //    foreach (TPodk TT in SpVin) { if (spSuf.Exists(x => x == TT.Sist) == false)spSuf.Add(TT.Sist); }
        //    SpVin.Sort(delegate(TPodk x, TPodk y) { return Convert.ToDouble(x.IND).CompareTo(Convert.ToDouble(y.IND)); });
        //    using (tr)
        //    {
        //        PromptPointResult pPtRes;
        //        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        //        PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
        //        pPtRes = doc.Editor.GetPoint(pPtOpts);
        //        Point3d Toch1 = pPtRes.Value;
        //        Point3d TochT = Toch1;
        //        string AdrLn = this.label21.Text;
        //        string NameLn = AdrLn.Split('\\').Last().Split('.')[0];
        //        string AdrL1 = this.label22.Text;
        //        string NameL1 = AdrL1.Split('\\').Last().Split('.')[0];
        //        string AdrLreg = this.label23.Text;
        //        string NameLreg = AdrLreg.Split('\\').Last().Split('.')[0];
        //        InsBlockRef_NI(@AdrLn, NameLn, new Point3d(Toch1.X - B, TochT.Y + B, 0));
        //        SetDynamicBlkProperty_TB(NOM, "", "", "", "", "", "", "", "", List.ToString());
        //        List = List + 1;
        //        foreach (string tSUF in spSuf)
        //        {
        //            SpVinSUF = SpVin.FindAll(x => x.Sist == tSUF);
        //            foreach (TPodk TTpodk in SpVinSUF)
        //            {
        //                if (TTpodk.BlNOD != null)
        //                {
        //                    TEXT(TochT, 0, TTpodk.Shem, A, B, a, b);
        //                    TochT = new Point3d(Toch1.X, TochT.Y - B / 2, 0);
        //                    string[] spInd = TTpodk.BlNOD.Split(',');
        //                    spIndL.Clear();
        //                    foreach (string Ind in spInd) { spIndL.Add(Ind); }
        //                    spIndL.Sort(delegate(string x, string y) { return x.CompareTo(y); });
        //                    foreach (string Ind in spIndL)
        //                    {
        //                        if (Ind != "")
        //                        {
        //                            LWPOIYiTEXT(TochT, 0, Ind, A, B, a, b, 50, "Выноски");
        //                            TochT = new Point3d(TochT.X + A, TochT.Y, 0);
        //                            if (TochT.X == Toch1.X + 8 * A) TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
        //                            if (TochT.Y < Toch1.Y - 44 * B)
        //                            {
        //                                TochT = new Point3d(Toch1.X, TochT.Y - 10 * B, 0);
        //                                Toch1 = TochT;
        //                                InsBlockRef_NI(@AdrLn, NameLn, new Point3d(Toch1.X - B, TochT.Y + B, 0));
        //                                SetDynamicBlkProperty_TB(NOM, "", "", "", "", "", "", "", "", List.ToString());
        //                                List = List + 1;
        //                            }
        //                        }
        //                    }
        //                    TochT = new Point3d(Toch1.X, TochT.Y - 4 * B, 0);
        //                    if (TochT.Y < Toch1.Y - 44 * B)
        //                    {
        //                        TochT = new Point3d(Toch1.X, TochT.Y - 10 * B, 0);
        //                        Toch1 = TochT;
        //                        InsBlockRef_NI(@AdrLn, NameLn, new Point3d(Toch1.X - B, TochT.Y + B, 0));
        //                        SetDynamicBlkProperty_TB(NOM, "", "", "", "", "", "", "", "", List.ToString());
        //                        List = List + 1;
        //                    }
        //                }
        //            }
        //        }
        //        //List = List + 1;
        //        TochT = new Point3d(Toch1.X + 14 * A, Toch1.Y + B, 0);
        //        InsBlockRef_NI(@AdrL1, NameL1, TochT);
        //        SetDynamicBlkProperty_TB(NOM, NAME, ISP, PROV, VIP, TKONTR, NKONTR, UTV, List.ToString(), List.ToString());
        //        TochT = new Point3d(TochT.X + 14 * A, Toch1.Y + B, 0);
        //        InsBlockRef_NI(@AdrLreg, NameLreg, TochT);
        //        SetDynamicBlkProperty_TB(NOM, NAME, ISP, PROV, VIP, TKONTR, NKONTR, UTV, List.ToString(), List.ToString());
        //        tr.Commit();
        //    }
        //}//построение пречня выносок

        //public void Sozd_TB(ref List<TPodk> SpVin)
        //{
        //    string stIND = "";
        //    string stDlin = "";
        //    string stStar = "";
        //    string stNov = "";
        //    string stNom = "";
        //    string SpVinKab = "";

        //    int nom;
        //    int Schet;

        //    double Dlin = 0;
        //    double DlinN = 0;

        //    SpVin.Sort(delegate(TPodk x, TPodk y) { return Convert.ToDouble(x.IND).CompareTo(Convert.ToDouble(y.IND)); });


        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;
        //    TypedValue[] acTypValAr = new TypedValue[2];
        //    acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
        //    acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Kabeli"), 1);
        //    // создаем фильтр
        //    SelectionFilter filter = new SelectionFilter(acTypValAr);
        //    PromptSelectionResult selRes = ed.SelectAll(filter);
        //    if (selRes.Status == PromptStatus.OK)
        //    {
        //        SelectionSet acSSet = selRes.Value;
        //        Transaction tr = db.TransactionManager.StartTransaction();
        //        using (tr)
        //        {
        //            foreach (SelectedObject sobj in acSSet)
        //            {
        //                Polyline ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as Polyline;
        //                //var KolV = ln.NumberOfVertices;
        //                Dlin = ln.Length;
        //                ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
        //                if (buffer != null)
        //                {
        //                    Noda Nnoda = new Noda();
        //                    Schet = 0;
        //                    foreach (TypedValue value in buffer)
        //                    {
        //                        if (Schet == 1) { stIND = value.Value.ToString(); }
        //                        if (Schet == 5) { stNom = value.Value.ToString(); }
        //                        Schet = Schet + 1;
        //                    }
        //                    SpVinKab = "";
        //                    fSpVin(ln, SpVin, ref SpVinKab);
        //                    foreach (string MT in SpVinKab.Split('-'))
        //                    {
        //                        if (MT != "")
        //                        {
        //                            nom = SpVin.FindIndex(x => x.IND == MT);
        //                            TPodk stTPodk = SpVin[nom];
        //                            if (stTPodk.BlNOD != null)
        //                            { if (stTPodk.BlNOD.Contains("," + stIND) == false) stTPodk.NBlNOD(stTPodk.BlNOD + "," + stIND); }
        //                            else
        //                                stTPodk.NBlNOD(stTPodk.BlNOD + "," + stIND);
        //                            SpVin[nom] = stTPodk;
        //                        }
        //                    }
        //                }
        //            }
        //            tr.Commit();
        //        }
        //    }
        //}//создания перечня выносок
        public void Postr_TB(ref List<TPodk> SpVin)
        {
            double A = 400;
            double B = 100;
            double a = 40;
            double b = 80;
            double List = 2;
            //string NONDOC = this.textBox5.Text;
            string NOM = this.textBox5.Text;
            string NAME = this.textBox6.Text;
            string ISP = this.textBox7.Text;
            string PROV = this.textBox8.Text;
            string VIP = this.textBox9.Text;
            string TKONTR = this.textBox10.Text;
            string NKONTR = this.textBox11.Text;
            string UTV = this.textBox12.Text;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            List<string> spIndL = new List<string>();
            using (tr)
            {
                PromptPointResult pPtRes;
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                PromptPointOptions pPtOpts = new PromptPointOptions("Укажи точку ");
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d Toch1 = pPtRes.Value;
                Point3d TochT = Toch1;
                string AdrLn = this.label21.Text;
                string AdrL1 = this.label22.Text;
                string AdrLreg = this.label23.Text;
                //Application.ShowAlertDialog(AdrLn + " " + AdrL1 + AdrLreg);
                string NameLn = AdrLn.Split('\\').Last().Split('.')[0];
                string NameL1 = AdrL1.Split('\\').Last().Split('.')[0];
                string NameLreg = AdrLreg.Split('\\').Last().Split('.')[0];
                //InsBlockRef_NI(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\А4С#.dwg", "А4С#", new Point3d(Toch1.X - B, TochT.Y + B, 0));
                InsBlockRef_NI(@AdrLn, NameLn, new Point3d(Toch1.X - B, TochT.Y + B, 0));
                SetDynamicBlkProperty_TB(NOM, "", "", "", "", "", "", "", "", List.ToString());
                List = List + 1;
                foreach (TPodk TTpodk in SpVin)
                {
                    if (TTpodk.BlNOD != null)
                    {
                        TEXT(TochT, 0, TTpodk.IND, A, B, a, b);
                        TochT = new Point3d(Toch1.X, TochT.Y - B / 2, 0);
                        string[] spInd = TTpodk.BlNOD.Split(',');
                        spIndL.Clear();
                        foreach (string Ind in spInd) { spIndL.Add(Ind); }
                        spIndL.Sort(delegate(string x, string y) { return x.CompareTo(y); });
                        foreach (string Ind in spIndL)
                        {
                            if (Ind != "")
                            {
                                //LWPOIYiTEXT(TochT, 0, Ind, A, B, a, b, 50);
                                LWPOIYiTEXT(TochT, 0, Ind, A, B, a, b, 50, "Выноски");
                                TochT = new Point3d(TochT.X + A, TochT.Y, 0);
                                if (TochT.X == Toch1.X + 8 * A) TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                                if (TochT.Y < Toch1.Y - 44 * B)
                                {
                                    TochT = new Point3d(Toch1.X, TochT.Y - 10 * B, 0);
                                    Toch1 = TochT;
                                    //InsBlockRef_NI(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\А4С#.dwg", "А4С#", new Point3d(Toch1.X - B, TochT.Y + B, 0));
                                    InsBlockRef_NI(@AdrLn, NameLn, new Point3d(Toch1.X - B, TochT.Y + B, 0));
                                    SetDynamicBlkProperty_TB(NOM, "", "", "", "", "", "", "", "", List.ToString());
                                    List = List + 1;
                                }
                            }
                        }
                        TochT = new Point3d(Toch1.X, TochT.Y - 4 * B, 0);
                        if (TochT.Y < Toch1.Y - 44 * B)
                        {
                            TochT = new Point3d(Toch1.X, TochT.Y - 10 * B, 0);
                            Toch1 = TochT;
                            //InsBlockRef_NI(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\А4С#.dwg", "А4С#", new Point3d(Toch1.X - B, TochT.Y + B, 0));
                            InsBlockRef_NI(@AdrLn, NameLn, new Point3d(Toch1.X - B, TochT.Y + B, 0));
                            SetDynamicBlkProperty_TB(NOM, "", "", "", "", "", "", "", "", List.ToString());
                            List = List + 1;
                        }
                    }
                }
                //List = List + 1;
                TochT = new Point3d(Toch1.X + 14 * A, Toch1.Y + B, 0);
                //InsBlockRef_NI(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\А4_L1_С#.dwg", "А4_L1_С#", TochT);
                InsBlockRef_NI(@AdrL1, NameL1, TochT);
                SetDynamicBlkProperty_TB(NOM, NAME, ISP, PROV, VIP, TKONTR, NKONTR, UTV, List.ToString(), List.ToString());
                TochT = new Point3d(TochT.X + 14 * A, Toch1.Y + B, 0);
                //InsBlockRef_NI(@"\\tbserv.vympel.local\home\41 отдел\Ерошенко\МАРШРУТ_NC\А4_REGI_С#.dwg", "А4_REGI_С#", TochT);
                InsBlockRef_NI(@AdrLreg, NameLreg, TochT);
                SetDynamicBlkProperty_TB(NOM, NAME, ISP, PROV, VIP, TKONTR, NKONTR, UTV, List.ToString(), List.ToString());
                tr.Commit();
            }
        }//построение пречня выносок

        public void Postr_VIN(ref List<TPodk> SpVin)
        {
            double dMas = Convert.ToDouble(Mas);
            double dKolSt = Convert.ToDouble(this.textBox39.Text);
            double A = 430 * dMas;
            double B = 110 * dMas;
            double a = 40 * dMas;
            double b = 80 * dMas;
            double X = 0;
            double Y = 0;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            List<string> spIndL = new List<string>();
            List<string> spSuf = new List<string>();
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Point3d Toch1;
                Point3d TochT;
                Point3d TochB;
                    foreach (TPodk TTpodk in SpVin)
                    {
                        if (TTpodk.BlNOD != null)
                        {
                            TochB=TTpodk.Koord;
                            Toch1 = TTpodk.Koord1;
                            TochT = Toch1;
                            if (TochB.Y < TochT.Y & TochB.X < TochT.X) Toch1 = new Point3d(TochT.X, TochT.Y + B, 0);
                            if (TochB.Y < TochT.Y & TochB.X > TochT.X) Toch1 = new Point3d(TochT.X-A, TochT.Y + B, 0);
                            if (TochB.Y > TochT.Y & TochB.X > TochT.X) Toch1 = new Point3d(TochT.X-A, TochT.Y, 0);
                            TochT = Toch1;
                            string[] spInd = TTpodk.BlNOD.Split(',');
                            spIndL.Clear();
                            foreach (string Ind in spInd) {spIndL.Add(Ind);}
                            spIndL.Sort(delegate(string x, string y) { return x.CompareTo(y); });
                            foreach (string Ind in spIndL)
                            {
                                if (Ind != "")
                                {
                                    LWPOIYiTEXT(TochT, 3, Ind, A, B, a, b, 50 * dMas, "ВыноскиНаПолеЧертежа");
                                    if (TochB.X < Toch1.X) X = TochT.X + A; else X = TochT.X - A;
                                    TochT = new Point3d(X, TochT.Y, 0);
                                    if (TochT.X == Toch1.X + dKolSt * A & TochB.X < Toch1.X & TochB.Y < Toch1.Y) TochT = new Point3d(Toch1.X, TochT.Y + B, 0);
                                    if (TochT.X == Toch1.X - dKolSt * A & TochB.X > Toch1.X & TochB.Y < Toch1.Y) TochT = new Point3d(Toch1.X, TochT.Y + B, 0);
                                    if (TochT.X == Toch1.X + dKolSt * A & TochB.X < Toch1.X & TochB.Y > Toch1.Y) TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                                    if (TochT.X == Toch1.X - dKolSt * A & TochB.X > Toch1.X & TochB.Y > Toch1.Y) TochT = new Point3d(Toch1.X, TochT.Y - B, 0);
                                }
                            }
                        }
                    }
                tr.Commit();
            }
        }//построение пречня выносок
        public void HsistSloi() 
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {

                Database db = doc.Database;
                Editor ed = doc.Editor;
                ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
                ObjectIdCollection acObjIdColl2 = new ObjectIdCollection();
                TypedValue[] acTypValAr = new TypedValue[1];
                acTypValAr.SetValue(new TypedValue(8, "ВыноскиНаПолеЧертежа"), 0);
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
                            Entity ln = tr.GetObject(sobj, OpenMode.ForWrite) as Entity;
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
        }//чистка слоя

        public void fSpVin(Polyline ln, List<TPodk> LSpVin, ref string SpVin) 
        {
            //string SpVin = "";
            double DistA_B;
            double DistA_C;
            double DistC_B;
            double DELT;
            var KolV = ln.NumberOfVertices;
            for (int i = KolV - 1; i > 0; i--)
            {
                Point3d TOtr1 = ln.GetPointAtParameter(i);
                Point3d TOtr2 = ln.GetPointAtParameter(i-1);
                DistA_B = TOtr1.DistanceTo(TOtr2);
                foreach (TPodk TV in LSpVin)
                {
                    DistA_C = TOtr1.DistanceTo(TV.Koord);
                    DistC_B = TV.Koord.DistanceTo(TOtr2);
                    DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                    if (DELT < 0.5) { if (SpVin.Contains(TV.IND + "-")==false) SpVin = SpVin + TV.Shem + "-"; }
                }
            }
        }//Список выносок на отрезке кабеля


        public void ShtenTXT(ref List<Noda> spNod, string File)
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\МАРШРУТ\" + File + ".txt", Encoding.UTF8);
            foreach (string Strok in lines)
            {
                Noda Nod = new Noda();
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
                spNod.Add(Nod);
            }

        }//чтение текстового файла с нодами
        public void ShtenSLOV_NOD(ref List<Noda> spNod, string File)
        {
            List<string> DOC = HCtenSlovNod(File);
            if (DOC.Count == 0) return;
            //string[] lines = DOC.Split(';');
            foreach (string Strok in DOC)
            {
                Noda Nod = new Noda();
                if (Strok != "")
                {
                    string[] NODm = Strok.Split(':');
                    Nod.NomNod(NODm[0]);
                    Nod.NomHoz(NODm[1]);
                    Nod.NParam(Convert.ToDouble(NODm[2]));
                    Nod.NomNOtk("-");
                    string[] stKoor = NODm[3].TrimStart('(').TrimEnd(')').Split(',');
                    double X = Convert.ToDouble(stKoor[0]);
                    double Y = Convert.ToDouble(stKoor[1]);
                    double Z = Convert.ToDouble(stKoor[2]);
                    Point3d Koor = new Point3d(X, Y, Z);
                    Nod.NKoor(Koor);
                    Nod.NVes(9999999999999.0);
                    Nod.NSpSmNod(NODm[4]);
                    if (NODm.Length > 4)
                    {
                        //Nod.NSpSmNod(NODm[4]);
                        string[] stKoorMir = NODm[5].TrimStart('(').TrimEnd(')').Split(',');
                        double XMir = Convert.ToDouble(stKoorMir[0]);
                        double YMir = Convert.ToDouble(stKoorMir[1]);
                        double ZMir = Convert.ToDouble(stKoorMir[2]);
                        Point3d KoorMir = new Point3d(XMir, YMir, ZMir);
                        Nod.NKoor3D(KoorMir);
                        Nod.NomVin(NODm[6]);
                        if (NODm[6] == "" & NODm[0].Contains("Переход")) Nod.NomVin(NODm[1]);
                    }
                    spNod.Add(Nod);
                }
            }

        }//чтение текстового файла с нодами
        public void VKizTXT(ref List<Kab> VK) 
        {
            string StrokDop = "";
            string[] lines = System.IO.File.ReadAllLines(@"C:\МАРШРУТ\Кабели.txt", Encoding.Default);
            foreach (string Strok in lines)
            {
                StrokDop = Strok + "::::::::";
                string[] Kabel = StrokDop.Split(':');
                Kab nkab = new Kab();
                nkab.NIND(Kabel[0]);
                nkab.NDIam(Convert.ToDouble(Kabel[1]));
                nkab.NINDOtk(Kabel[4]);
                nkab.NINDKud(Kabel[7]);
                nkab.NSist(Kabel[10]);
                nkab.NBlokNOD(Kabel[11]);
                nkab.KoorOt = Kabel[12];
                nkab.KoorKud = Kabel[13];
                VK.Add(nkab);
            }
        }//Чтение кабельного из текстового файла
        public void VKizNOD(ref List<Kab> VK)
        {
            string[] DiamM;
            string StrokDop = "";
            List<String> VKisSlov = HCtenSlovNod("VK");
            foreach (string Strok in VKisSlov)
            {
                StrokDop = Strok + "::::::::";
                string[] Kabel = StrokDop.Split(':');
                Kab nkab = new Kab();
                nkab.NIND(Kabel[0]);
                DiamM = Kabel[1].Split('_');
                nkab.NDIam(Convert.ToDouble(DiamM[0]));
                if (DiamM.Length == 2) nkab.NMassa(DiamM[1]);else nkab.NMassa("0");
                    nkab.NINDOtk(Kabel[4]);
                nkab.NINDKud(Kabel[7]);
                nkab.NspVin(Kabel[9]);
                nkab.NSist(Kabel[10]);
                nkab.NBlokNOD(Kabel[11]);
                nkab.KoorOt = Kabel[12];
                nkab.KoorKud = Kabel[13];
                VK.Add(nkab);
            }
        }//Чтение кабельного из текстового файла

        public void SgrVKTXTfail(List<ZaPis> SpNOD)
        {
            List<PodvKab> SpSmNod;
            string strKoor;
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\SortSPKab.txt"))
            {
                foreach (ZaPis line in SpNOD)
                {
                    strKoor = "";
                    SpSmNod = line.SpVixKab;
                    foreach (PodvKab line1 in SpSmNod)
                    {
                        strKoor = strKoor + Convert.ToString(line1.IND) + " " + Convert.ToString(line1.INDobor) + " " + Convert.ToString(line1.KoordObor.ToString()) + ",";
                    }
                    file.WriteLine(Convert.ToString(line.INDotk.IND) + ":" + Convert.ToString(line.Sist) + ":" + Convert.ToString(line.BlokNOD) + ":" + strKoor);
                }
            }
        }//функция записи в файл дуг

        static public void SBORpl_FIND(ref List<ObjectId> SpOBJID, string FIND_K)
        {
            string stKomp = "";
            int Schet;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[8];
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 0);
            acTypValAr.SetValue(new TypedValue(0, "LWPOLYLINE"), 1);
            acTypValAr.SetValue(new TypedValue(0, "CIRCLE"), 2);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 4);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 5);
            acTypValAr.SetValue(new TypedValue(8, "Kabeli"), 6);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 7);
            // создаем фильтр
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);
            if (acSSPrompt.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Нет деталей...");
                return;
            }
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    stKomp = "";
                    Entity bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as Entity;
                    ResultBuffer buffer = bref.GetXDataForApplication("LAUNCH01");
                    if (buffer != null)
                    {
                        Schet = 0;
                        foreach (TypedValue value in buffer)
                        {
                            if (Schet == 1) { stKomp = value.Value.ToString(); }
                            Schet = Schet + 1;
                        }
                    }
                    if (stKomp == FIND_K ) { SpOBJID.Add(acSSObj.ObjectId); }
                }
                ObjectId[] idarrayEmpty = SpOBJID.ToArray();
                ed.SetImpliedSelection(idarrayEmpty);
                Tx.Commit();
            }
        }//поиск полилиний по индексу
        public void KRUGiTEXT(Point3d Toch1, int Zvet, string IND, string Gr, double rad, double VisT)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Circle KrugPOZVn = new Circle();
                KrugPOZVn.SetDatabaseDefaults();
                KrugPOZVn.Center = Toch1;
                KrugPOZVn.Radius = rad;
                KrugPOZVn.Layer = "ТРАССЫскрытые";
                KrugPOZVn.ColorIndex = Zvet;
                KrugPOZVn.XData = new ResultBuffer(
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, IND),
                new TypedValue(1040, 0)
                );
                btr.AppendEntity(KrugPOZVn);
                tr1.AddNewlyCreatedDBObject(KrugPOZVn, true);


                DBText Poz = new DBText();
                Poz.SetDatabaseDefaults();
                Poz.Position = Toch1;
                Poz.Height = VisT;
                Poz.ColorIndex = Zvet;
                Poz.TextString = IND+ Gr;
                Poz.Layer = "ТРАССЫскрытые";
                btr.AppendEntity(Poz);
                btr.XData = new ResultBuffer(
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, IND),
                new TypedValue(1040, 999999999)
                );

                DBText Diam = new DBText();
                Diam.SetDatabaseDefaults();
                Diam.Position = new Point3d(Toch1.X, Toch1.Y- VisT -0.1, 0);
                Diam.Height = VisT;
                Diam.ColorIndex = Zvet;
                Diam.TextString = "D=" +(rad*2).ToString("#.00");
                Diam.Layer = "ТРАССЫскрытые";
                btr.AppendEntity(Diam);
                btr.XData = new ResultBuffer(
                new TypedValue(1001, "LAUNCH01"),
                new TypedValue(1000, IND),
                new TypedValue(1040, 999999999)
                );

                tr1.AddNewlyCreatedDBObject(Poz, true);
                tr1.Commit();
            }
        }//Отрисовка кругов и текста для точе подключения
        public void LWPOIYiTEXT(Point3d Toch1, int Zvet, string IND, double A, double B, double a, double b, double Htext,string Sloi)
        {
            Point3d Ttext = new Point3d(Toch1.X + a, Toch1.Y-b,0);
            Point3d T1 = Toch1;
            Point3d T2 = new Point3d(Toch1.X + A, Toch1.Y , 0);
            Point3d T3 = new Point3d(Toch1.X + A, Toch1.Y-B, 0);
            Point3d T4 = new Point3d(Toch1.X, Toch1.Y - B, 0);
            Point3dCollection TkoorPL = new Point3dCollection();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                TkoorPL = new Point3dCollection { T1, T2, T3, T4 };
                Polyline poly = new Polyline();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = 1;
                poly.Closed = true;
                poly.Layer = Sloi;
                int i = 0;
                foreach (Point3d pt in TkoorPL)
                {
                    poly.AddVertexAt(i, new Point2d(pt.X, pt.Y), 0, 0, 0);
                    i = i + 1;
                }
                btr.AppendEntity(poly);
                tr1.AddNewlyCreatedDBObject(poly, true);

                if (IND != "")
                {
                    DBText Poz = new DBText();
                    Poz.SetDatabaseDefaults();
                    //Poz.ColorIndex = 3;
                    Poz.Position = Ttext;
                    Poz.Height = Htext;
                    Poz.ColorIndex = Zvet;
                    Poz.TextString = IND;
                    Poz.WidthFactor = 0.6;
                    Poz.Layer = Sloi;
                    btr.AppendEntity(Poz);
                    tr1.AddNewlyCreatedDBObject(Poz, true);
                }
                tr1.Commit();
            }
        }//Отрисовка прямоугольника и текста
        public void TEXT(Point3d Toch1, int Zvet, string IND, double A, double B, double a, double b)
        {
            //Point3d Ttext = new Point3d(Toch1.X + a, Toch1.Y + b, 0);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                DBText Poz = new DBText();
                Poz.SetDatabaseDefaults();
                Poz.Position = Toch1;
                Poz.Height = 50;
                Poz.ColorIndex = Zvet;
                Poz.TextString = IND;
                Poz.Layer = "Выноски";
                btr.AppendEntity(Poz);
                tr1.AddNewlyCreatedDBObject(Poz, true);

                tr1.Commit();
            }
        }//Отрисовка прямоугольника и текста
        static public void LINE(Point3d Toch1, Point3d Toch2, int Zvet)
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
            //PromptSelectionResult acSSPrompt = ed.SelectLast();
            //SelectionSet acSSet = acSSPrompt.Value;
            //using (Transaction Tx = db.TransactionManager.StartTransaction())
            //{
            //    foreach (SelectedObject acSSObj in acSSet)
            //    {
            //        ObjId.Add(acSSObj.ObjectId);
            //        Tx.Commit();
            //    }
            //}
        }//построение линии
        public void LWPOIY_3D(Point3d T1, Point3d T2, Point3d T3, Point3d T4, double rad)
        {
            Point3dCollection TkoorPL = new Point3dCollection();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Polyline3d poly = new Polyline3d();
                poly.SetDatabaseDefaults();
                poly.ColorIndex = 0;
                int i = 0;
                PolylineVertex3d V1 =new PolylineVertex3d(T1);
                PolylineVertex3d V2 = new PolylineVertex3d(T2);
                PolylineVertex3d V3 = new PolylineVertex3d(T3);
                PolylineVertex3d V4 = new PolylineVertex3d(T4);
                poly.AppendVertex(V1);
                poly.AppendVertex(V2);
                poly.AppendVertex(V3);
                poly.AppendVertex(V4);
                tr1.AddNewlyCreatedDBObject(poly, true);
                tr1.Commit();
            }
        }//Отрисовка прямоугольника и текста
        public void SPLINE(Point3d T1, Point3d T2, Point3d T3, Point3d T4, Point3d T5, Point3d T6, Point3d T7, Point3d T8, double rad,ref List<ObjectId> ObjId, Point3d T9, Point3d T10, Point3d T11) 
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = doc.Editor;
           
            //ed.Command("1");
            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
            //ed.Command("splmethod");
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // Create a Point3d Collection
                Point3dCollection acPt3dColl = new Point3dCollection();
                //acPt3dColl.Add(T1);
                //acPt3dColl.Add(T9);
                acPt3dColl.Add(T2);
                acPt3dColl.Add(T3);
                acPt3dColl.Add(T4);
                //acPt3dColl.Add(T11);
                //acPt3dColl.Add(T5);
                //acPt3dColl.Add(T6);
                //acPt3dColl.Add(T7);
                //acPt3dColl.Add(T10);
                //acPt3dColl.Add(T8);
                // Set the start and end tangency
                //Vector3d acStartTan = new Vector3d(1, 0, 0);
                //Vector3d acEndTan = new Vector3d(0, 0, 1);
                Vector3d acStartTan = (T2-T1)/T1.DistanceTo(T2);
                Vector3d acEndTan = (T5 - T4) / T5.DistanceTo(T4);
                // Create a spline
                using (Spline acSpline = new Spline(acPt3dColl, acStartTan, acEndTan, 2, 0))
                //using (Spline acSpline = new Spline())
                {
                    //Spline acSpline = new Spline(acPt3dColl, acStartTan, acEndTan, 0, 0);
                    //acSpline.SetDatabaseDefaults();
                    // Set a control point
                    //acSpline.SetControlPointAt(0, T1);
                    //acSpline.SetControlPointAt(1, T9);
                    //acSpline.SetControlPointAt(0, T2);
                    //acSpline.SetControlPointAt(1, T3);
                    //acSpline.SetControlPointAt(2, T4);
                    //acSpline.SetControlPointAt(5, T11);
                    //acSpline.SetControlPointAt(2, T5);
                    //acSpline.SetControlPointAt(3, T6);
                    //acSpline.SetControlPointAt(4, T7);
                    //acSpline.SetControlPointAt(5, T10);
                    //acSpline.SetControlPointAt(4, T8);
                    // Add the new object to the block table record and the transaction
                    acBlkTblRec.AppendEntity(acSpline);
                    acTrans.AddNewlyCreatedDBObject(acSpline, true);
                }
                // Save the new objects to the database
                acTrans.Commit();
            }
            PromptSelectionResult acSSPrompt = ed.SelectLast();
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = acDoc.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    ObjId.Add(acSSObj.ObjectId);
                    Tx.Commit();
                }
            }
        }//отрисовка сплайна
        public void KRUG(Point3d Toch1, double rad, ref List<ObjectId> ObjId)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Transaction tr1 = db.TransactionManager.StartTransaction();
            using (tr1)
            {
                BlockTableRecord btr = (BlockTableRecord)tr1.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Circle KrugPOZVn = new Circle();
                KrugPOZVn.SetDatabaseDefaults();
                KrugPOZVn.Center = Toch1;
                KrugPOZVn.Radius = rad;
                KrugPOZVn.Layer = "ТРАССЫскрытые";
                btr.AppendEntity(KrugPOZVn);
                tr1.AddNewlyCreatedDBObject(KrugPOZVn, true);
                tr1.Commit();
            }
            PromptSelectionResult acSSPrompt = ed.SelectLast();
            SelectionSet acSSet = acSSPrompt.Value;
            using (Transaction Tx = doc.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    ObjId.Add(acSSObj.ObjectId);
                    Tx.Commit();
                }
            }
        }//Отрисовка кругов и текста для точе подключения

        private void ZapTablT(List<TPodk> SpTP) 
        {
            this.dataGridView1.Rows.Clear();
            List<String> TochkiSTR = new List<String>();
            string XYZ1 = "";
            string XYZ2 = "";
            int kVist = 0;
            int i = -1;
            foreach (TPodk TP in SpTP)
            {
                if (Tochki.Exists(x => x.IND == TP.IND) == false)
                {
                    i = i + 1;
                    this.dataGridView1.Rows.Add(TP.IND, TP.KoordMod, TP.Naim, TP.Pom, TP.Shem);
                    this.dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                }
                else
                {
                    i = i + 1;
                    this.dataGridView1.Rows.Add(TP.IND, TP.KoordMod, TP.Naim, TP.Pom, TP.Shem);
                    this.dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
                    kVist = kVist + 1;
                }
            }
            this.label4.Text = "Выставлено-" + kVist.ToString() + " из " + TochkiSort.Count.ToString();
        }//Заполнение таблицы точек подключения
        private void ZapTablK(List<Kab> VKt)
        {
            string Ohsib = "";
            string[] spVin;
            int KolVisTP = 0;
            int KolVisTM = 0;
            int KolNVisTM = 0;
            int i = -1;
            this.label57.Text = "Разведено " + Vkrazv.Count.ToString() + " из " + VKt.Count.ToString() + "(" + (VKt.Count - Vkrazv.Count).ToString() + ")";
            this.dataGridView2.Rows.Clear();
            foreach (Kab Strok in VKt)
            {
                Ohsib = "";
                KolVisTP = 0;
                KolVisTM = 0;
                KolNVisTM = 0;
                i = i + 1;
                if (Vkrazv.Exists(x => x.Split('*')[0] == Strok.IND) == true)
                {
                    this.dataGridView2.Rows.Add(Strok.IND, Strok.DIam, Vkrazv.Find(x => x.Split('*')[0] == Strok.IND).Split('*')[1], Strok.INDOtk, Strok.INDKud, Strok.Sist, Strok.BlokNOD, Strok.spVin, Vkrazv.Find(x => x.Split('*')[0] == Strok.IND).Split('*')[2].TrimEnd('-'), "", "","", Strok.Massa);
                    this.dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
                }
                else
                {   
                    if (Tochki.Exists(x => x.IND == Strok.INDOtk) == false) Ohsib = Ohsib + " Не выставлена точка-" + Strok.INDOtk; else KolVisTP = KolVisTP + 1;
                    if (Tochki.Exists(x => x.IND == Strok.INDKud) == false) Ohsib = Ohsib + " Не выставлена точка-" + Strok.INDKud; else KolVisTP = KolVisTP + 1;
                    if (Strok.Sist == "") { Ohsib = "Не указанна система трасс"; KolVisTP = 0; KolVisTM = 0; }
                    else if (MSpNOD[Convert.ToInt16(Strok.Sist) - 1].Count == 0) { Ohsib = "Не созданна система трасс-" + Strok.Sist; KolVisTP = 0; KolVisTM = 0; }
                        if (Strok.spVin != null) 
                        {
                        spVin = Strok.spVin.Split(',');
                        foreach (string Tp in spVin) 
                        { 
                            if (Tochki.Exists(x => x.IND == Tp) & Tp != Strok.INDOtk & Tp != Strok.INDKud) KolVisTM = KolVisTM + 1;
                            if (Tochki.Exists(x => x.IND == Tp) == false & Tp != Strok.INDOtk & Tp != Strok.INDKud & Tp != "") { KolNVisTM = KolNVisTM + 1; Ohsib = Ohsib + " Не выставлена точка-" + Tp; }
                        }
                        }
                    this.dataGridView2.Rows.Add(Strok.IND, Strok.DIam, "", Strok.INDOtk, Strok.INDKud, Strok.Sist, Strok.BlokNOD, Strok.spVin, "", "","", Ohsib, Strok.Massa);
                    this.dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    if(Ohsib != "" | KolNVisTM>0) this.dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    if (KolVisTP >0 & KolVisTM > 0) this.dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Gold;
                    if (Ohsib == "" & KolNVisTM == 0) this.dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                }
            }
        }//Заполнение таблицы кабельного
   
        public void ProezT(Point3d MSK, Point3d PSK, string OSI) 
        {
         foreach (TPodk TToch in TochkiSort) 
         {
             if (TToch.KoordMod != "Не найден" & TToch.KoordMod != "" & Tochki.Exists(x => x.IND == TToch.IND) == false)
             {
                 double Glubin = Convert.ToDouble(this.textBox4.Text);
                 double Delt = 0;
                 string[] XYZmir = TToch.KoordMod.Split(' ');
                 int Zvet = 52;
                 string Gr = "";
                 int[] KolZV = { 20, 2, 110, 6, 230, 4 };
                 string[] Grupp = { "(I)", "(II)", "(III)", "(IV)", "(V)", "Силовые" };
                 List<Kab> VKt = VK.FindAll(x => x.INDKud == TToch.IND | x.INDOtk == TToch.IND);
                 foreach (Kab tKab in VKt) { foreach (string tGr in Grupp) { if (tKab.IND.Contains(tGr) & Gr.Contains(tGr)==false)Gr = Gr + tGr; } }
                 if (Gr == "") Gr = "(IV)";
                 int i = 0;
                 foreach (string tGr in Grupp) { if (Gr == tGr)Zvet = KolZV[i]; i = i + 1; }

                 double delX = Convert.ToDouble(XYZmir[0]) - MSK.X;
                 double delY = Convert.ToDouble(XYZmir[1]) - MSK.Y;
                 double delZ = Convert.ToDouble(XYZmir[2]) - MSK.Z;
                 double Xnow = PSK.X;
                 double Ynow = PSK.Y;
                 double Znow = PSK.Z;
                 if (OSI == "XY") { Delt = Math.Abs(delZ); }
                 if (OSI == "ZX") { Delt = Math.Abs(delY); }
                 if (OSI == "ZY") { Delt = Math.Abs(delX); }
                 if (OSI == "XZ") { Delt = Math.Abs(delY); }
                 if (OSI == "YZ") { Delt = Math.Abs(delX); }
                 if (Delt < Glubin)
                 {
                     if (OSI == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; }
                     if (OSI == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; }
                     if (OSI == "ZY") { Xnow = Xnow + delY; Ynow = Ynow + delZ; }
                     if (OSI == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; }
                     if (OSI == "YZ") { Xnow = Xnow - delY; Ynow = Ynow + delZ; }
                     Point3d IskTP = new Point3d(Xnow, Ynow, 0);
                     KRUGiTEXT(IskTP, Zvet, TToch.IND, Gr,15,15);
                 }
             }
         }
        }//проецирование точек

        public void ObnTr() 
        {
            this.dataGridView8.Rows.Clear();
            string File = "";
            int i = -1;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                UdalKrug();
                double[] KolZV = { 20, 2, 110, 6, 230, 4 };
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
                        SozdSpNODPerDinBl(ref spPer, "INSERT");
                        SozdSpNODPer(ref spPer, "CIRCLE");
                        SozdSpNOD1(Zvet, ref spNod, ref spLINI);
                        SozdSpNOD_Vin(ref spNod);
                        DOP_spNOD(ref spNod, ref spLINI, 10);
                        SvazLin(ref spNod, ref spLINI, ref spNodFin);
                        foreach (Noda TnodEt in spPer) { spNodFin.Add(TnodEt); }
                        Skon(10, ref spNodFin, ref spLINI, ref spPer, Zvet);
                        //NodiVKTXTfail(spNodFin, "Трассы " + Grupp[i] + " группы");
                        SOZDlov("Трассы" + Grupp[i] + "группы");
                        ZapisSlovSistTR("Трассы" + Grupp[i] + "группы", NodiVStrList(spNodFin));
                        //this.listBox2.Items.Add("Трассы " + Grupp[i] + " группы " + spNodFin.Count.ToString() + " узлов");
                    }
                }
            }
        }
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
        public void KRUG(Point3d Toch1, int Zvet)
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
        public void SozdSpNODPer(ref List<Noda> spPer, string TIP)
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
            Point3d BazT = new Point3d();
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
                         strPalub = "-";
                         strBort = "-";
                         strShpan = "-";
                         strPom1 = "-";
                         strPom2 = "-";
                         strNazv = "-";
                         douDlin = 0;
                         strNNod = "-";
                         SpDoPNOD = "";
                        if (TIP == "INSERT")
                        {
                            BlockReference ln = tr.GetObject(sobj, OpenMode.ForWrite) as BlockReference;
                            if (ln != null)
                            {
                                i = i + 1;
                                BazT = ln.Position;
                                strNNod = ln.Handle.ToString();
                                foreach (ObjectId idAtrRef in ln.AttributeCollection)
                                {
                                    using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                                    {
                                        if (atrRef != null)
                                        {
                                            if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА") { strNazv = atrRef.TextString; }
                                            if (atrRef.Tag == "ДЛИНА_ПЕРЕХОДА") { douDlin = Convert.ToDouble(atrRef.TextString); }
                                            if (atrRef.Tag == "ID_перех") { if (strNazv=="") strNazv = atrRef.TextString; }
                                        }
                                    }
                                }
                                ResultBuffer buffer = ln.GetXDataForApplication("LAUNCH01");
                                if (buffer != null)
                                {
                                   
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
                                }
                                if (strNazv != "-") 
                                { 
                                    Noda Nnoda = new Noda();
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
        public void SozdSpNODPerDinBl(ref List<Noda> spPer, string TIP)
        {
            //Application.ShowAlertDialog("Зашла сюда");
            string strNNod = "-";
            string Nazv_per = "", Sprav_per = "";
            double Dlin = 0;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            string SpDoPNOD = "";
            Point3d BazT = new Point3d();
            int i = 0;
            double DDist = Convert.ToDouble(this.textBox1.Text);
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
            TypedValue[] acTypValAr1 = new TypedValue[6];
            acTypValAr1.SetValue(new TypedValue(0, TIP), 0);
            acTypValAr1.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr1.SetValue(new TypedValue(8, "Выноски"), 2);
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
                    //Application.ShowAlertDialog(acObjIdColl1.Count.ToString());
                    foreach (ObjectId sobj in acObjIdColl1)
                    {
                        Nazv_per = "";
                        Sprav_per = "";
                        //для плоскостей
                        Dlin = 0;
                        x1 = 0;
                        y1 = 0;
                        x2 = 0;
                        y2 = 0;
                        if (TIP == "INSERT")
                        {
                            BlockReference bref = tr.GetObject(sobj, OpenMode.ForWrite) as BlockReference;
                            if (bref != null)
                            {
                                i = i + 1;
                                BazT = bref.Position;
                                strNNod = bref.Handle.ToString();
                                if (bref.IsDynamicBlock)
                                {
                                    DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                                    foreach (DynamicBlockReferenceProperty prop in props)
                                    {
                                        object[] values = prop.GetAllowedValues();
                                        //if (prop.PropertyName == "Положение1 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                                        //if (prop.PropertyName == "Положение1 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
                                        //if (prop.PropertyName == "Положение4 X") { x2 = Convert.ToDouble(prop.Value.ToString()); }
                                        //if (prop.PropertyName == "Положение4 Y") { y2 = Convert.ToDouble(prop.Value.ToString()); }
                                        if (prop.PropertyName == "Положение3 X") { x1 = Convert.ToDouble(prop.Value.ToString()); }
                                        if (prop.PropertyName == "Положение3 Y") { y1 = Convert.ToDouble(prop.Value.ToString()); }
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
                                            if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА1") { Nazv_per = atrRef.TextString; }
                                            if (atrRef.Tag == "СПРАВКА_ПЕРЕХОДА1") { Sprav_per = atrRef.TextString; }
                                            if (atrRef.Tag == "ДЛИНА_ПЕРЕХОДА") { Dlin = Convert.ToDouble(atrRef.TextString); }
                                        }
                                    }
                                }
                                //Application.ShowAlertDialog(Nazv_per);
                                if (Nazv_per != "") 
                                {
                                    Noda Nnoda1 = new Noda();
                                    Point3d T1 = new Point3d(BazT.X + x1, BazT.Y + y1, 0);
                                    Point3d T2 = new Point3d(BazT.X + x2, BazT.Y + y2, 0);
                                    Nnoda1.NomNod("ПереходБлок" + i);
                                    Nnoda1.NomHoz(Nazv_per);
                                    Nnoda1.NParam(i);
                                    Nnoda1.NKoor(T1);
                                    Nnoda1.NDlinP(Dlin);
                                    spPer.Add(Nnoda1);
                                    Noda Nnoda2 = new Noda();
                                    i = i + 1;
                                    Nnoda2.NomNod("ПереходБлок" + i);
                                    Nnoda2.NomHoz(Nazv_per);
                                    Nnoda2.NParam(i);
                                    Nnoda2.NKoor(T2);
                                    Nnoda2.NDlinP(Dlin);
                                    int indTnodEt = spPer.FindIndex(x => x.Hoz == Nnoda2.Hoz);
                                    Noda sNod = spPer.Find(x => x.Hoz == Nnoda2.Hoz);
                                    SpDoPNOD = Nnoda2.Nom + "*" + Dlin;
                                    DoboV_lSM_NOD(ref spPer, SpDoPNOD, sNod, indTnodEt);
                                    SpDoPNOD = sNod.Nom + "*" + Dlin;
                                    Nnoda2.NSpSmNod(SpDoPNOD);
                                    spPer.Add(Nnoda2);
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
            CreateLayer("Kabeli");
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
        public void SozdSpNOD_Vin(ref List<Noda> spNod)
        {
            string smNOD1 = "";
            string smNOD2 = "";
            string spSmNod = "";
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[5];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "Выноски"), 2);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "ТРАССЫскрытые"), 3);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 4);
            // создаем фильтр
            SelectionFilter filter = new SelectionFilter(acTypValAr);
            PromptSelectionResult selRes = ed.SelectAll(filter);
            CreateLayer("ТРАССЫскрытые");
            CreateLayer("Kabeli");
            if (selRes.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = selRes.Value;
                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    foreach (SelectedObject sobj in acSSet)
                    {
                        smNOD1 = "";
                        BlockReference ln = tr.GetObject(sobj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        foreach (ObjectId idAtrRef in ln.AttributeCollection)
                        {
                            using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                            {
                                if (atrRef != null)
                                {
                                    if (atrRef.Tag == "НОМЕР_ВЫНОСКИ") { smNOD1 = atrRef.TextString; }
                                }
                            }
                        }
                        if (smNOD1 != "")
                        {
                            Noda TNOD = new Noda();
                            TNOD.NomNod("УзелВыноска-" + smNOD1);
                            TNOD.NomHoz(smNOD1);
                            TNOD.NKoor(ln.Position);
                            TNOD.NParam(1);
                            TNOD.NSpSmNod(spSmNod);
                            TNOD.NomVin(smNOD1);
                            spNod.Add(TNOD);
                        }
                    }
                    tr.Commit();
                }
            }
        }//создания списка нод выносок
        public void DOP_spNOD(ref List<Noda> spNod, ref List<string> spLINI, double RAD)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            int i1 = 0;
            double DistA_B = 0;
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
                        for (int i = 0; i < spNodSR.Count - 1; i++)
                        {
                            Noda TOtr1 = spNodSR[i];
                            Noda TOtr2 = spNodSR[i + 1];
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
                                if (spNNod.Exists(x => x.Nom == TOtr1.Hoz + "-" + (i + (DistA_C / DistA_B)).ToString()) == false) spNNod.Add(NNoda);
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
                                if (spNNod.Exists(x => x.Nom == TOtr1.Hoz + "-" + (i + (DistA_C / DistA_B)).ToString()) == false) spNNod.Add(NNoda);
                            }
                        }
                    }
                }
                i1 = i1 + 1;
                ed.WriteMessage("Обработано " + i1 + " из " + spLINI.Count + " линий \n");
            }
            //this.progressBar1.Value = 0;
            foreach (Noda Lin1 in spNNod) { spNod.Add(Lin1); }
        }//соеденение концов линий с отрезками других линий
        public void DOP_spNOD_Vin(ref List<Noda> spNod, List<Noda> spNod_vin, double RAD)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            int i1 = 0;
            double DistA_B = 0;
            double DistA_C = 0;
            double DistC_B = 0;
            double DELT = 0;
            List<Noda> spNNod = new List<Noda>();
            foreach (Noda TVin in spNod_vin)
            {
                for (int i = 0; i < spNod.Count - 1; i++)
                {
                    Noda TOtr1 = spNod[i];
                    Noda TOtr2 = spNod[i + 1];
                    DistA_B = TOtr1.Koord.DistanceTo(TOtr2.Koord);
                    DistA_C = TOtr1.Koord.DistanceTo(TVin.Koord);
                    DistC_B = TVin.Koord.DistanceTo(TOtr2.Koord);
                    DELT = Math.Abs((DistA_C + DistC_B) - DistA_B);
                    if ((DELT < 0.5) & (DistA_C > RAD) & (DistC_B > RAD) & TOtr1.Hoz== TOtr2.Hoz)
                    {
                        Noda NNoda = new Noda();
                        NNoda.NomNod(TOtr1.Hoz + "-" + (TOtr1.Param + (DistA_C / DistA_B)).ToString());
                        NNoda.NomHoz(TOtr1.Hoz);
                        NNoda.NKoor(TVin.Koord);
                        NNoda.NParam(TOtr1.Param + (DistA_C / DistA_B));
                        NNoda.NomVin(TVin.Vin);
                        if (spNNod.Exists(x => x.Nom == TOtr1.Hoz + "-" + (TOtr1.Param + (DistA_C / DistA_B)).ToString()) == false) spNNod.Add(NNoda);
                        break;
                    }
                }
                i1 = i1 + 1;
                ed.WriteMessage("Обработано " + i1 + " из " + spNod_vin.Count + " выноски \n");
            }
            foreach (Noda Lin1 in spNNod) { spNod.Add(Lin1); }
        }//соеденение выносок с отрезками линий



        public void SvazLin(ref List<Noda> spNod, ref List<string> spLINI, ref List<Noda> spNodFin)
        {
            string smNOD1 = "";
            string smNOD2 = "";
            string spSmNod = "";
            foreach (string Lin in spLINI)
            {
                List<Noda> spNodEtal = spNod.FindAll(x => x.Hoz == Lin);
                spNodEtal.Sort(delegate(Noda x, Noda y) { return x.Param.CompareTo(y.Param); });
                var KolV = spNodEtal.Count;
                for (int i = 0; i <= KolV - 1; i++)
                {
                    smNOD1 = "";
                    smNOD2 = "";
                    spSmNod = "";
                    if (i > 0) { smNOD1 = spNodEtal[i].Hoz + "-" + spNodEtal[i - 1].Param + "*" + spNodEtal[i].Koord.DistanceTo(spNodEtal[i - 1].Koord).ToString(); }
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
        public void Skon(double RAD, ref List<Noda> spNod, ref List<string> spLINI, ref List<Noda> spPer, double Zvet)
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
                Noda Nod2 = spNodEtal[spNodEtal.Count - 1];
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
                if (Nod2.Koord.DistanceTo(Nod1.Koord) <= RAD)
                {
                    SpDoPNOD = Nod1.Nom + "*" + Nod2.Koord.DistanceTo(Nod1.Koord).ToString();
                    DoboV_lSM_NOD(ref spNod, SpDoPNOD, Nod2, ind2NOD);
                    SpDoPNOD = Nod2.Nom + "*" + Nod2.Koord.DistanceTo(Nod1.Koord).ToString();
                    DoboV_lSM_NOD(ref spNod, SpDoPNOD, Nod1, ind1NOD);
                    KRUG(Nod2.Koord, Convert.ToInt16(Zvet));
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
                if (StarNODA1_1.SpSmNod.Contains(stSMNOD) == false)
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
        public void NodiVKTXTfail(List<Noda> spNod, string File)
        {
            Point3d Zentr = new Point3d();
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\" + File + ".txt"))
            {
                foreach (Noda line in spNod)
                {
                    Zentr = line.Koord;
                    Point3d KoorMod = Zentr;
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
                            KoorMod = new Point3d(Xnow, Ynow, Znow);
                        }
                    }
                    string stSpSMNod = "";
                    if (line.SpSmNod != null) { stSpSMNod = line.SpSmNod.TrimStart(','); }
                    file.WriteLine(line.Nom + ":" + line.Hoz + ":" + line.Param + ":" + line.Koord.X.ToString() + "," + line.Koord.Y.ToString() + "," + line.Koord.Z.ToString() + ":" + stSpSMNod + ":" + KoorMod.X.ToString() + "," + KoorMod.Y.ToString() + "," + KoorMod.Z.ToString() + ":" + line.Vin);
                }
            }
        }//функция записи в файл нод
        public string NodiVStr(List<Noda> spNod)
        {
            string Nodi = "";
            Point3d Zentr = new Point3d();
                foreach (Noda line in spNod)
                {
                    Zentr = line.Koord;
                    Point3d KoorMod = Zentr;
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
                            KoorMod = new Point3d(Xnow, Ynow, Znow);
                        }
                    }
                    string stSpSMNod = "";
                    if (line.SpSmNod != null) { stSpSMNod = line.SpSmNod.TrimStart(','); }
                Nodi= Nodi + (line.Nom + ":" + line.Hoz + ":" + line.Param + ":" + line.Koord.X.ToString() + "," + line.Koord.Y.ToString() + "," + line.Koord.Z.ToString() + ":" + stSpSMNod + ":" + KoorMod.X.ToString() + "," + KoorMod.Y.ToString() + "," + KoorMod.Z.ToString() + ":" + line.Vin) + ";";
            }
            return Nodi;
        }//функция записи нод в текстовую переменную
        public List<string> NodiVStrList(List<Noda> spNod)
        {
            List<string> Nodi =new List<string>();
            Point3d Zentr = new Point3d();
            foreach (Noda line in spNod)
            {
                Zentr = line.Koord;
                Point3d KoorMod = Zentr;
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
                        KoorMod = new Point3d(Xnow, Ynow, Znow);
                    }
                }
                string stSpSMNod = "";
                if (line.SpSmNod != null) { stSpSMNod = line.SpSmNod.TrimStart(','); }
                Nodi.Add(line.Nom + ":" + line.Hoz + ":" + line.Param + ":" + line.Koord.X.ToString() + "," + line.Koord.Y.ToString() + "," + line.Koord.Z.ToString() + ":" + stSpSMNod + ":" + KoorMod.X.ToString() + "," + KoorMod.Y.ToString() + "," + KoorMod.Z.ToString() + ":" + line.Vin);
            }
            return Nodi;
        }//функция записи нод в текстовую переменную
        public void TpodkVKTXTfail(List<TPodk> spNod, string File)
        {
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\МАРШРУТ\" + File + ".txt"))
            {
                foreach (TPodk line in spNod)
                {
                    string stSpSMNod = "";
                    file.WriteLine(line.IND + ":" + line.Koord.ToString() + ":" + line.Sist + ":" + line.BlNOD);
                }
            }
        }//функция записи в файл нод
        public void obnTP(List<String> VKisSlov, string SpPom) 
        {
            string[] DiamM;
            string[] SpPomM = SpPom.Split(',');
            string[] SpVin;
            TochkiSort.Clear();
            TochkiRazv.Clear();
            string StrokDop = "";
            //string[] lines = System.IO.File.ReadAllLines(@"C:\МАРШРУТ\Кабели.txt", Encoding.Default);
            //List<String> VKisSlov = HCtenSlovNod("VK");
            foreach (string Strok in VKisSlov)
            {
                StrokDop = Strok + "::::::::::";
                string[] Kabel = StrokDop.Split(':');
                Kab nkab = new Kab();
                nkab.NIND(Kabel[0]);
                DiamM = Kabel[1].Split('_');
                nkab.NDIam(Convert.ToDouble(DiamM[0]));
                if (DiamM.Length == 2) nkab.NMassa((DiamM[1]));else nkab.NMassa("0");
                nkab.NINDOtk(Kabel[4]);
                nkab.NPomOt(Kabel[5]);
                nkab.NPomKud(Kabel[6]);
                nkab.NINDKud(Kabel[7]);
                nkab.NShem(Kabel[8]);
                nkab.NspVin(Kabel[9]);
                nkab.NSist(Kabel[10]);
                nkab.NBlokNOD(Kabel[11]);
                nkab.KoorOt = Kabel[12];
                nkab.KoorKud = Kabel[13];
                if (Kabel[9] != "")
                {
                    SpVin = Kabel[9].Split(',');
                    foreach (string TVin in SpVin)
                    {
                        if (TVin != "" & TochkiSort.Exists(x => x.IND == TVin & x.Sist == nkab.Sist & x.BlNOD == nkab.BlokNOD) == false)
                        {
                            TPodk NNOD = new TPodk();
                            NNOD.NIND(TVin);
                            NNOD.NPom("Выноска");
                            NNOD.NSist(nkab.Sist);
                            NNOD.NBlNOD(nkab.BlokNOD);
                            NNOD.NShem(nkab.Shem);
                            if (Tochki.Exists(x => x.IND == TVin)) { NNOD.NKoor(Tochki.Find(x => x.IND == TVin).Koord); }
                            if (TochkiSort.Exists(x => x.IND == TVin) == false) { TochkiSort.Add(NNOD); }
                            if (TochkiRazv.Exists(x => x.IND == TVin & x.Sist == nkab.Sist) == false) { TochkiRazv.Add(NNOD); }
                        }
                    }
                }
                VK.Add(nkab);
                if (TochkiSort.Exists(x => x.IND == nkab.INDOtk & x.Sist == nkab.Sist & x.BlNOD == nkab.BlokNOD) == false)
                {
                    TPodk NNOD = new TPodk();
                    NNOD.NIND(nkab.INDOtk);
                    NNOD.NPom(nkab.PomOt);
                    NNOD.NSist(nkab.Sist);
                    NNOD.NBlNOD(nkab.BlokNOD);
                    NNOD.NShem(nkab.Shem);
                    NNOD.NKoorMod(Kabel[12]);
                    if (Kabel.Length > 13) NNOD.NNaim(Kabel[14]);
                    if (Tochki.Exists(x => x.IND == nkab.INDOtk)) { NNOD.NKoor(Tochki.Find(x => x.IND == nkab.INDOtk).Koord); }
                    if (TochkiSort.Exists(x => x.IND == nkab.INDOtk) == false)
                    {
                        if (SpPom == "") TochkiSort.Add(NNOD);
                        if (Array.Exists(SpPomM, x => x == nkab.PomOt)) TochkiSort.Add(NNOD);
                    }
                    if (TochkiRazv.Exists(x => x.IND == nkab.INDOtk & x.Sist == nkab.Sist) == false)
                    {
                        if (SpPom == "") TochkiRazv.Add(NNOD);
                        if (Array.Exists(SpPomM, x => x == nkab.PomOt)) TochkiRazv.Add(NNOD);
                    }
                }
                //TochkiSTR.Add(nkab.INDKud);
                if (TochkiSort.Exists(x => x.IND == nkab.INDKud & x.Sist == nkab.Sist & x.BlNOD == nkab.BlokNOD) == false)
                {
                    TPodk NNOD = new TPodk();
                    NNOD.NIND(nkab.INDKud);
                    NNOD.NPom(nkab.PomKud);
                    NNOD.NSist(nkab.Sist);
                    NNOD.NBlNOD(nkab.BlokNOD);
                    NNOD.NShem(nkab.Shem);
                    NNOD.NKoorMod(Kabel[13]);
                    if (Kabel.Length > 13) NNOD.NNaim(Kabel[15]);
                    if (Tochki.Exists(x => x.IND == nkab.INDKud)) { NNOD.NKoor(Tochki.Find(x => x.IND == nkab.INDKud).Koord); }
                    if (TochkiSort.Exists(x => x.IND == nkab.INDKud) == false)
                    {
                        if (SpPom == "") TochkiSort.Add(NNOD);
                        if (Array.Exists(SpPomM, x => x == nkab.PomKud)) TochkiSort.Add(NNOD);
                    }
                    if (TochkiRazv.Exists(x => x.IND == nkab.INDKud & x.Sist == nkab.Sist) == false)
                    {
                        if (SpPom == "") TochkiRazv.Add(NNOD);
                        if (Array.Exists(SpPomM, x => x == nkab.PomOt)) TochkiRazv.Add(NNOD);
                    }
                }
            }
        }//Обновить координаты точек подключения
       

        public void SSil_Na_Vid() 
        {
            List<PERprim> spPerex = new List<PERprim>();
            string Nazv_per = "", Sprav_per1 = "", Sprav_per2 = "";
            string[] Sprav_per1M, Sprav_per2M;
            ObjectId Prim1=new ObjectId();
            ObjectId Prim2=new ObjectId();
            Point3d BP = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[7];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Выноски"), 2);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 3);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 4);
            acTypValAr.SetValue(new TypedValue(8, "Плоскости"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
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
                    //для переходов
                    Nazv_per = "";
                    //для плоскостей
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as BlockReference;
                    BP = bref.Position;
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                     using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference){if (atrRef != null){ if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА") { Nazv_per = atrRef.TextString; }}}}
                    if (Nazv_per != "")
                    {
                        if (Sprav_per1 != "" & Sprav_per2 != "") 
                        {
                        PERprim Tper = new PERprim();
                        Tper.NNazv(Nazv_per);
                        Tper.NSprav1(Sprav_per1); 
                        Tper.NSprav2(Sprav_per2);
                        Tper.NPrim1(Prim1);
                        Tper.NPrim2(Prim2);
                        spPerex.Add(Tper);
                        Sprav_per1 = "";
                        Sprav_per2 = "";
                        Prim1 = new ObjectId();
                        Prim2 = new ObjectId();
                        }
                        if (Sprav_per1 != "" & Sprav_per2 == "") 
                        {
                            Prim2 = acSSObj.ObjectId;
                            Sprav_per2 = FIND_List(ref SpPLOS, BP);
                            Sprav_per1M = Sprav_per1.Split('(');
                            Sprav_per2M = Sprav_per2.Split('(');
                                if (Sprav_per1M[1] == Sprav_per2M[1]) 
                                {
                                Sprav_per1 = Sprav_per1M[0];
                                Sprav_per2 = Sprav_per2M[0];
                                }
                        }
                        if (Sprav_per1 == "" & Sprav_per2 == "") 
                        {
                            Prim1 = acSSObj.ObjectId;
                            Sprav_per1 = FIND_List(ref SpPLOS, BP);
                            
                        }
                    }
                }
                if (Sprav_per1 != "" & Sprav_per2 != "")
                {
                    PERprim Tper = new PERprim();
                    Tper.NNazv(Nazv_per);
                    Tper.NSprav1(Sprav_per1);
                    Tper.NSprav2(Sprav_per2);
                    Tper.NPrim1(Prim1);
                    Tper.NPrim2(Prim2);
                    spPerex.Add(Tper);
                    Sprav_per1 = "";
                    Sprav_per2 = "";
                    Prim1 = new ObjectId();
                    Prim2 = new ObjectId();
                }
                foreach (PERprim Tper in spPerex) 
                {
                    BlockReference bref1 = Tx.GetObject(Tper.Prim1, OpenMode.ForRead) as BlockReference;
                    foreach (ObjectId idAtrRef in bref1.AttributeCollection){using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference){if (atrRef != null) { if (atrRef.Tag == "СПРАВКА_ПЕРЕХОДА") { atrRef.TextString = Tper.Sprav2;}}}}
                    BlockReference bref2 = Tx.GetObject(Tper.Prim2, OpenMode.ForRead) as BlockReference;
                    foreach (ObjectId idAtrRef in bref2.AttributeCollection) { using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference) { if (atrRef != null) { if (atrRef.Tag == "СПРАВКА_ПЕРЕХОДА") { atrRef.TextString = Tper.Sprav1;}}}}
                }
                Tx.Commit();
            }
        }
        static public void SBOR_OBOR(ref List<VINOS> SpVIN, ref List<PEREXOD> SpPEREX, ref List<PLOS> SpPLOS)
        {
            string Nazv_Vin = "", Sprav_vin = "";
            string Nazv_per = "", Sprav_per = "";
            string Vid = "", Mash = "", Sprav_plos = "", List = "", Osi = "", IDvin = "", IDper = "";
            string SUF = "";
            double Dlin = 0;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xMir = 0;
            double yMir = 0;
            double zMir = 0;
            Point3d BP = new Point3d();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            TypedValue[] acTypValAr = new TypedValue[7];
            acTypValAr.SetValue(new TypedValue(0, "INSERT"), 0);
            acTypValAr.SetValue(new TypedValue(-4, "<or"), 1);
            acTypValAr.SetValue(new TypedValue(8, "Выноски"), 2);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫ"), 3);
            acTypValAr.SetValue(new TypedValue(8, "ТРАССЫскрытые"), 4);
            acTypValAr.SetValue(new TypedValue(8, "Плоскости"), 5);
            acTypValAr.SetValue(new TypedValue(-4, "or>"), 6);
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
                    IDvin = "";
                    //для переходов
                    Nazv_per = "";
                    Sprav_per = "";
                    IDper = "";
                    //для плоскостей
                    Vid = "";
                    Mash = "";
                    Sprav_plos = "";
                    List = "";
                    Osi = "";
                    SUF = "";
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
                            if (prop.PropertyName == "Видимость1") { Osi = prop.Value.ToString(); }
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
                                if (atrRef.Tag == "СУФИКС") { SUF = atrRef.TextString; }
                                if (atrRef.Tag == "ID_Вын") { IDvin = atrRef.TextString; }
                                //переходы
                                if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА1") { Nazv_per = atrRef.TextString; }
                                if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА") { Nazv_per = atrRef.TextString; }
                                if (atrRef.Tag == "СПРАВКА_ПЕРЕХОДА") { Sprav_per = atrRef.TextString; }
                                if (atrRef.Tag == "ДЛИНА_ПЕРЕХОДА") { Dlin = Convert.ToDouble(atrRef.TextString); }
                                if (atrRef.Tag == "ID_перех") { IDper = atrRef.TextString; }
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
                    if (Nazv_Vin != "")
                    {
                        VINOS tPOZ = new VINOS();
                        if (SUF != "" & Nazv_Vin.Length > SUF.Length) Nazv_Vin = Nazv_Vin.Substring(0, Nazv_Vin.Length - SUF.Length);
                        tPOZ.NPolnNazv(Nazv_Vin + SUF);
                        tPOZ.NNazv(Nazv_Vin);
                        tPOZ.NSUF(SUF);
                        tPOZ.NSpravka(Sprav_vin);
                        tPOZ.NKoor(BP);
                        SpVIN.Add(tPOZ);
                    }
                    if (IDvin != "")
                    {
                        VINOS tPOZ = new VINOS();
                        if (SUF != "" & IDvin.Length > SUF.Length) IDvin = IDvin.Substring(0, IDvin.Length - SUF.Length);
                        tPOZ.NPolnNazv(IDvin + SUF);
                        tPOZ.NNazv(IDvin);
                        tPOZ.NSUF(SUF);
                        tPOZ.NSpravka(Sprav_vin);
                        tPOZ.NKoor(BP);
                        SpVIN.Add(tPOZ);
                    }
                    if (Nazv_per != "")
                    {
                        PEREXOD tPOZ = new PEREXOD();
                        if (SUF != "" & Nazv_per.Length > SUF.Length) Nazv_per = Nazv_per.Substring(0, Nazv_per.Length - SUF.Length);
                        tPOZ.NPolnNazv(Nazv_per + SUF);
                        tPOZ.NNazv(Nazv_per);
                        tPOZ.NSpravka(Sprav_vin);
                        tPOZ.NSUF(SUF);
                        tPOZ.NDlin(Dlin);
                        if (SpPEREX.Exists(x => x.PolnNazv == (Nazv_per + SUF)) == false) SpPEREX.Add(tPOZ);
                    }
                    if (IDper != "")
                    {
                        PEREXOD tPOZ = new PEREXOD();
                        if (SUF != "" & IDper.Length > SUF.Length) IDper = IDper.Substring(0, IDper.Length - SUF.Length);
                        tPOZ.NPolnNazv(IDper + SUF);
                        tPOZ.NNazv(IDper);
                        tPOZ.NSpravka(Sprav_vin);
                        tPOZ.NSUF(SUF);
                        tPOZ.NDlin(Dlin);
                        if (SpPEREX.Exists(x => x.PolnNazv == (IDper + SUF)) == false) SpPEREX.Add(tPOZ);
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
        public void InsBlockRef(string BlockPath, string NAME, string Vin, string TipDBL, string SUF, string ID)
        {
            // Активный документ в редакторе AutoCAD
            Document doc = Application.DocumentManager.MdiActiveDocument;
            // База данных чертежа (в данном случае - активного документа)
            Database db = doc.Database;
            // Редактор базы данных чертежа
            // Запускаем транзакцию
            using (DocumentLock docLock = doc.LockDocument())
            {
                CreateLayer("Выноски");
                CreateLayer("ТРАССЫ");
                CreateLayer("ТРАССЫскрытые");
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
                SetDynamicBlkProperty(NAME, Vin, TipDBL, SUF,ID);
            }
        }//Выставмить дин блок
        public void InsBlockRef_NI(string BlockPath, string NAME, Point3d Tvst)
        {
            // Активный документ в редакторе AutoCAD
            Document doc = Application.DocumentManager.MdiActiveDocument;
            // База данных чертежа (в данном случае - активного документа)
            Database db = doc.Database;
            // Редактор базы данных чертежа
            // Запускаем транзакцию
            using (DocumentLock docLock = doc.LockDocument())
            {
                CreateLayer("Выноски");
                CreateLayer("ТРАССЫ");
                CreateLayer("ТРАССЫскрытые");
                CreateLayer("Плоскости");
                using (DbS.Transaction tr = db.TransactionManager.StartTransaction())
                {
                    //EdI.Editor ed = doc.Editor;
                    //EdI.PromptPointOptions pPtOpts;
                    //pPtOpts = new EdI.PromptPointOptions("\nУкажите точку вставки блока: ");
                    //// Выбор точки пользователем
                    //var pPtRes = doc.Editor.GetPoint(pPtOpts);
                    //if (pPtRes.Status != EdI.PromptStatus.OK)
                    //    return;
                    //var ptStart = pPtRes.Value;

                    DbS.BlockTable bt = tr.GetObject(db.BlockTableId, DbS.OpenMode.ForRead) as DbS.BlockTable;
                    DbS.BlockTableRecord model = tr.GetObject(bt[DbS.BlockTableRecord.ModelSpace], DbS.OpenMode.ForWrite) as DbS.BlockTableRecord;
                    // Создаем новую базу
                    using (DbS.Database db1 = new DbS.Database(false, false))
                    {
                        // Получаем базу чертежа-донора
                        db1.ReadDwgFile(BlockPath, System.IO.FileShare.Read, true, null);
                        // Получаем ID нового блока
                        DbS.ObjectId BlkId = db.Insert(NAME, db1, false);
                        DbS.BlockReference bref = new DbS.BlockReference(Tvst, BlkId);
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
            }
        }//Выставмить дин блок не интерактивный способ
        static public void SetDynamicBlkProperty(string NAME, string Vin, string TipBl, string SUF, string ID)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            PromptSelectionResult acSSPrompt = doc.Editor.SelectLast();
            SelectionSet acSSet = acSSPrompt.Value;
            string tekM = "20";
            using (Transaction Tx = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject acSSObj in acSSet)
                {
                    RegAppTable regTable = (RegAppTable)Tx.GetObject(db.RegAppTableId, OpenMode.ForWrite);
                    BlockReference bref = Tx.GetObject(acSSObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                    if (TipBl == "Плоскость")
                        bref.Layer = "Плоскости";
                    else if(TipBl == "Выноска")
                        bref.Layer = "Выноски";
                    else if (TipBl == "Переход")
                        bref.Layer = "ТРАССЫ";
                    if (bref.IsDynamicBlock)
                    {
                        DynamicBlockReferencePropertyCollection props = bref.DynamicBlockReferencePropertyCollection;
                        foreach (DynamicBlockReferenceProperty prop in props)
                        {
                            if (prop.PropertyName == "МАСШТАБ") { tekM = prop.Value.ToString(); prop.Value = Convert.ToDouble(tekM) * Convert.ToDouble(HCtenSlov("MAS#", "1")); }
                        }
                    }
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "НОМЕР_ВЫНОСКИ") { if (Vin != "") atrRef.TextString = Vin + SUF; else atrRef.TextString = ""; }
                                if (atrRef.Tag == "НАЗВАНИЕ_ПЕРЕХОДА") { if (Vin != "") atrRef.TextString = Vin + SUF; else atrRef.TextString = ""; }
                                if (atrRef.Tag == "СУФИКС") { atrRef.TextString = SUF; }
                                if (atrRef.Tag == "ID_Вын") { atrRef.TextString = ID; }
                                if (atrRef.Tag == "ID_перех") { atrRef.TextString = ID; }
                            }
                        }
                    }
                }
                Tx.Commit();
            }
        }
        static public void SetDynamicBlkProperty_TB(string NONDOC, string NAIMDOC, string RAZRAB, string PROV, string VIPUST, string NKONTR, string TKONTR, string UTVER, string LISTOV, string LIST)
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
                    bref.Layer = "Выноски";
                    foreach (ObjectId idAtrRef in bref.AttributeCollection)
                    {
                        using (var atrRef = idAtrRef.Open(OpenMode.ForWrite, false, true) as AttributeReference)
                        {
                            if (atrRef != null)
                            {
                                if (atrRef.Tag == "РАЗРАБОТАЛ") { atrRef.TextString = RAZRAB; }
                                if (atrRef.Tag == "ПРОВЕРИЛ") { atrRef.TextString = PROV; }
                                if (atrRef.Tag == "ВЫПУСТИЛ") { atrRef.TextString = VIPUST; }
                                if (atrRef.Tag == "НКОНТР") { atrRef.TextString = NKONTR; }
                                if (atrRef.Tag == "УТВЕРДИЛ") { atrRef.TextString = UTVER; }
                                if (atrRef.Tag == "НОМЕР") { atrRef.TextString = NONDOC; }
                                if (atrRef.Tag == "НАИМЕНОВАНИЕ") { atrRef.TextString = NAIMDOC; }
                                if (atrRef.Tag == "Листов") { atrRef.TextString = LISTOV; }
                                if (atrRef.Tag == "Лист") { atrRef.TextString = LIST; }
                                if (atrRef.Tag == "ЛИСТ") { atrRef.TextString = LIST; }
                            }
                        }
                    }
                }
                Tx.Commit();
            }
        }
        public string FIND_List(ref List<PLOS> SpPOZ, Point3d IT)
        {
            string ind_pom = "";
            foreach (PLOS line1 in SpPOZ)
            {
                Point3d MAX = line1.max;
                Point3d MIN = line1.min;
                if (IT.X < MAX.X & IT.X > MIN.X & IT.Y < MAX.Y & IT.Y > MIN.Y) {  ind_pom = line1.Vid + " (" + line1.List + ")";  }

            }
            return ind_pom;
        }//поиск области
        public string FIND_XYZmir_2D(List<PLOS> SpPOZ, Point3d IT)
        {
            string XYmir = "0 0 0";
            foreach (PLOS line1 in SpPOZ)
            {
                Point3d MAX = line1.max;
                Point3d MIN = line1.min;
                if (IT.X < MAX.X & IT.X > MIN.X & IT.Y < MAX.Y & IT.Y > MIN.Y)
                {
                    double delX = IT.X - line1.Psk.X;
                    double delXn = IT.X - MAX.X;
                    double delY = IT.Y - line1.Psk.Y;
                    double delZ = IT.Z - line1.Psk.Z;
                    double Xnow = line1.Msk.X;
                    double Xnown = line1.Msk.X;
                    double Ynow = line1.Msk.Y;
                    double Znow = line1.Msk.Z;
                    double ID = 0;
                    double KolPovt = 0;
                    if (line1.Osi == "XY") { Xnow = Xnow + delX; Ynow = Ynow + delY; Znow = Znow + delZ; }
                    if (line1.Osi == "ZX") { Xnow = Xnow + delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                    if (line1.Osi == "ZY") { Xnow = Xnow + delZ; Ynow = Ynow + delX; Znow = Znow + delY; }
                    if (line1.Osi == "XZ") { Xnow = Xnow - delX; Ynow = Ynow + delZ; Znow = Znow + delY; }
                    if (line1.Osi == "YZ") { Xnow = Xnow + delZ; Ynow = Ynow - delX; Znow = Znow + delY; }
                    Point3d KoorMod = new Point3d(Xnow, Ynow, Znow);
                    XYmir = Xnow.ToString() + " " + Ynow.ToString() + " " + Znow.ToString();
                }
            }
            return XYmir;
        }//поиск шпангоут борт и палуба

        static void SOZDlov(string Slov)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            DBDictionary NewDict = new DBDictionary();
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForWrite) as DBDictionary;
                if (!nod.Contains(Slov))
                {
                    nod.SetAt(Slov, NewDict);
                    tr.AddNewlyCreatedDBObject(NewDict, true);
                }
                tr.Commit();
            }
        }//создание словоря если его нет
        static void ZapisSlov(string Slov, string ZNACH)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForWrite) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForWrite);
                            ResultBuffer reZ = new ResultBuffer(new TypedValue(1000, ZNACH));
                            Xrecord xRec = new Xrecord();
                            xRec.Data = reZ;
                            PomD.SetAt(Slov, xRec);
                            tr.AddNewlyCreatedDBObject(xRec, true);
                        }
                    }

                }
                tr.Commit();
            }
        }//запись данных в словарь одна запись
        static void HistSlov(string Slov)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForWrite) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForWrite);
                            PomD.Erase(true);
                        }
                    }

                }
                tr.Commit();
            }
        }//запись данных в словарь одна запись

        static void ZapisSlovSistTR(string Slov, List<string> ZNACH)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                DBDictionary nod = tr.GetObject(HostApplicationServices.WorkingDatabase.NamedObjectsDictionaryId, OpenMode.ForWrite) as DBDictionary;
                if (nod.Contains(Slov))
                {
                    foreach (DBDictionaryEntry de in nod)
                    {
                        if (de.Key == Slov)
                        {
                            DBDictionary PomD = (DBDictionary)tr.GetObject(de.Value, OpenMode.ForWrite);
                            ResultBuffer reZ = new ResultBuffer();
                            foreach(string TNod in ZNACH) reZ.Add(new TypedValue(1000, TNod));
                            Xrecord xRec = new Xrecord();
                            xRec.Data = reZ;
                            PomD.SetAt(Slov, xRec);
                            tr.AddNewlyCreatedDBObject(xRec, true);
                        }
                    }

                }
                tr.Commit();
            }
        }//запись данных в словарь список данных
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
                                foreach (TypedValue valSl in rez)
                                {
                                    ZNACH = valSl.Value.ToString();
                                }
                            }
                        }
                    }

                }
            }
            return ZNACH;
        }//чтение словоря с одной записью
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
                                //Application.ShowAlertDialog(rez.Length.ToString());
                                foreach (TypedValue valSl in rez)
                                {
                                    ZNACH.Add(valSl.Value.ToString());
                                }
                            }
                        }
                    }

                }
            }
            return ZNACH;
        }//чтение словоря со списком
        public void UdalSvasi(ref List<Noda> spNod, string TVin)
        {
            string UkorSpSmNod1 = "";
            string UkorSpSmNod2 = "";
            Noda BlNod = spNod.Find(x => x.Vin == TVin);
            string[] SpSmNOD = BlNod.SpSmNod.Split(',');
            foreach (string NTSmNoda in SpSmNOD)
            {
                string[] tNoda = NTSmNoda.Split('*');
                Noda TSMNod = spNod.Find(x => x.Nom == tNoda[0]);
                UkorSpSmNod1 = UkoroSpNod(BlNod.SpSmNod, tNoda[0]);
                if(TSMNod.SpSmNod!=null) UkorSpSmNod2 = UkoroSpNod(TSMNod.SpSmNod, BlNod.Nom);
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

        public void ZagrExcel(string File) 
        {

            string StrokDop = "";
            List<string> VKtxt = new List<string>();
            List<string> VKEcxel = new List<string>();
            List<Zag> SpZag = new List<Zag>();
            List<Zag> Hapka = new List<Zag>();

            SpisZag(ref SpZag, this.textBox22.Text);
            SpisZag(ref SpZag, this.textBox23.Text);
            SpisZag(ref SpZag, this.textBox24.Text);
            SpisZag(ref SpZag, this.textBox25.Text);
            SpisZag(ref SpZag, this.textBox26.Text);
            SpisZag(ref SpZag, this.textBox27.Text);
            SpisZag(ref SpZag, this.textBox28.Text);
            SpisZag(ref SpZag, this.textBox29.Text);
            SpisZag(ref SpZag, this.textBox30.Text);
            SpisZag(ref SpZag, this.textBox31.Text);
            SpisZag(ref SpZag, this.textBox32.Text);
            SpisZag(ref SpZag, this.textBox33.Text);

            //считываем данные из Excel файла в двумерный массив
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook xlWB; //рабочая книга              
            Excel.Worksheet xlShtVK; //лист Excel  
            Excel.Worksheet xlShtTR; //лист Excel
            this.label33.Text = File;
            xlWB = xlApp.Workbooks.Open(@File); //название файла Excel                                             
            xlShtVK = xlWB.Worksheets[this.textBox20.Text]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            xlShtTR = xlWB.Worksheets[this.textBox21.Text]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            ZagrLisExcel(xlShtVK, ref VKtxt, SpZag, Hapka, ref  VKEcxel);
            ZagrLisExcel(xlShtTR, ref VKtxt, SpZag, Hapka, ref  VKEcxel);   
            xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            xlApp.Quit(); //закрываем Excel
            this.dataGridView2.Rows.Clear();
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = acDoc.LockDocument())
            {
                foreach (string Strok in VKEcxel)
                {
                    StrokDop = Strok + "::::::::::";
                    VKtxt.Add(StrokDop);
                }
                SOZDlov("VK");
                SOZDlov("Istok");
                ZapisSlovSistTR("VK", VKtxt);
                ZapisSlov("Istok", File);
                this.Show();
                int kVist = 0;
                List<String> TochkiSTR = new List<String>();
                SozdSpKab(ref Vkrazv);
                SpVistTP(ref Tochki);
                SpVistTP_BL(ref Tochki);
                string SpPom = HCtenSlov("spPOM", "");
                string Proj = HCtenSlov("PROG", "");
                obnTP(VKtxt, SpPom);
                foreach (string Strok in Vkrazv)
                {
                    if (VK.FindIndex(x => Strok.Split('*')[0] == x.IND) == -1)
                    {
                        Kab nkab = new Kab();
                        nkab.NIND(Strok.Split('*')[0]);
                        VK.Add(nkab);
                    }
                }
                ZapTablT(TochkiSort);
                ZapTablK(VK);
            }
        }
        public void ZagrLisExcel(Excel.Worksheet xlSht, ref List<string> VKtxt, List<Zag> SpZag, List<Zag> Hapka,ref List<string> VKEcxel)
        {
            string Zagal = "";
            string Ind = "";
            string Diam = "";
            string Otk = "";
            string Kud = "";
            string Blok = "";
            string Sist = "";
            string Vinos = "";
            string StrokDop = "";
            string XYZ1 = "";
            string XYZ2 = "";
            string Shem = "";
            string PomOtk = "";
            string PomKud = "";
            int Nom = 0;
            //xlSht = xlWB.Worksheets[4]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "B"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
            var arrData = (object[,])xlSht.Range["A1:AA" + iLastRow].Value; //берём данные с листа Excel
            //xlApp.Visible = true; //отображаем Excel     
            //xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            //xlApp.Quit(); //закрываем Excel
            //настройка DataGridView
            this.dataGridView2.Rows.Clear();
            int RowsCount = arrData.GetUpperBound(0);
            int ColumnsCount = arrData.GetUpperBound(1);
            //dataGridView2.RowCount = RowsCount; //кол-во строк в DGV
            //dataGridView2.ColumnCount = ColumnsCount; //кол-во столбцов в DGV
            //заполняем DataGridView данными из массива
            int i, j;
            Hapka.Clear();
            for (j = 1; j <= ColumnsCount; j++)
            {
                if (arrData[1, j] != null)
                {
                    Zagal = arrData[1, j].ToString();
                    foreach (Zag tZag in SpZag)
                    {
                        if (tZag.Zagal == Zagal)
                        {
                            tZag.NomGraf(j); Hapka.Add(tZag);
                        }
                    }
                }
            }
            for (i = 2; i <= RowsCount; i++)
            {
                Ind = "";
                Diam = "0";
                Otk = "";
                Kud = "";
                Blok = "";
                Sist = "";
                Vinos = "";
                XYZ1 = "";
                XYZ2 = "";
                Shem = "";
                PomOtk = "";
                PomKud = "";
                for (j = 1; j <= ColumnsCount; j++)
                {
                    if (arrData[i, j] != null)
                    {
                        if (arrData[i, j] != null)
                            Zagal = arrData[i, j].ToString();
                        else
                            Zagal = "";
                        foreach (Zag tZag in Hapka)
                        {
                            if (tZag.Zagal == this.textBox22.Text & tZag.Graf == j) Ind = Zagal;
                            if (tZag.Zagal == this.textBox23.Text & tZag.Graf == j) Diam = Zagal;
                            if (tZag.Zagal == this.textBox24.Text & tZag.Graf == j) Otk = Zagal.Trim(' ');
                            if (tZag.Zagal == this.textBox25.Text & tZag.Graf == j) PomOtk = Zagal;
                            if (tZag.Zagal == this.textBox26.Text & tZag.Graf == j) PomKud = Zagal;
                            if (tZag.Zagal == this.textBox27.Text & tZag.Graf == j) Kud = Zagal.Trim(' ');
                            if (tZag.Zagal == this.textBox28.Text & tZag.Graf == j) Sist = Zagal;
                            if (tZag.Zagal == this.textBox29.Text & tZag.Graf == j) Shem = Zagal;
                            if (tZag.Zagal == this.textBox30.Text & tZag.Graf == j) Blok = Zagal;
                            if (tZag.Zagal == this.textBox31.Text & tZag.Graf == j) Vinos = Zagal;
                            if (tZag.Zagal == this.textBox32.Text & tZag.Graf == j) XYZ1 = Zagal;
                            if (tZag.Zagal == this.textBox33.Text & tZag.Graf == j) XYZ2 = Zagal;
                        }
                    }
                }
                if (Diam.Replace(" ", "") == "") Diam = "0";
                if (Ind != "") VKEcxel.Add(Ind + ":" + Diam + ":::" + Otk.Split()[0] + ":" + PomOtk + ":" + PomKud + ":" + Kud.Split()[0] + ":" + Shem + ":" + Vinos + ":" + Sist + ":" + Blok + ":" + XYZ1 + ":" + XYZ2 + ":" + Otk.Substring(Otk.Split()[0].Length, Otk.Length - Otk.Split()[0].Length) + ":" + Kud.Substring(Kud.Split()[0].Length, Kud.Length - Kud.Split()[0].Length));
            }
        }

        public static void SetConn(string adr)
        {
            //connection = new SQLiteConnection("Data Source=//avserv/ProjectsMarine/Электросхемы/DBC/KAB.db;Version=3;New=false;Compress=true");
            connection = new SQLiteConnection("Data Source=" + adr + ";Version=3;New=false;Compress=true");
        }
        public void LoadDataPC()
        {
            string Proj = HCtenSlov("PROG", "");
            SetConn(this.label53.Text);
            connection.Open();
            command = connection.CreateCommand();
            string ComTXT = "SELECT Ind as 'Индекс',Pos  as 'Координата' ,Shem  as 'Схема' FROM POINT WHERE Shem LIKE '" + Proj + "%'";
            SQLiteDataAdapter adapter = new SQLiteDataAdapter(ComTXT, connection);
            DS1.Reset();
            adapter.Fill(DS1);
            DT1 = DS1.Tables[0];
            this.dataGridView10.DataSource = DT1;
            connection.Close();
        }
        public void ADDZapOCP(string sfi_code, string Pola, string Znahc, string Del)
        {
            string Prov = "SELECT count(*) FROM POINT WHERE Ind='" + sfi_code + "'";
            string Zap = "INSERT INTO POINT(" + Pola + ") VALUES('" + Znahc + "')";
            string DelZ = "DELETE FROM POINT WHERE Ind='" + Del + "'";
            ProvZap(Zap, Prov, DelZ,this.label53.Text);
        }
        public static void ProvZap(string ADD, string PROV, string Del, string adr)
        {
            SetConn(adr);
            connection.Open();
            command = connection.CreateCommand();
            command.CommandText = PROV;// 
            int cur = Convert.ToInt32(command.ExecuteScalar());
            if (cur == 0)
            {
                command.CommandText = ADD;
                command.ExecuteNonQuery();
            }
            else
            {
                command.CommandText = Del;
                command.ExecuteNonQuery();
                command.CommandText = ADD;
                command.ExecuteNonQuery();
            }
            connection.Close();
        }//добавление нового оборудования в базу
        static public void SozdSpKVKoor(ref List<string> Vkrazv, List<PLOS> SpPLOS)
        {
            string stIND = "";

            double Dlin = 0;
            double Param = 0;
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

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "0.5";
            ZapisSlov("MAS#", "0.5");
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "2.5";
            ZapisSlov("MAS#", "2.5");
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "1";
            ZapisSlov("MAS#", "1");
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            Mas = "1.25";
            ZapisSlov("MAS#", "1.25");
        }


    }
}
