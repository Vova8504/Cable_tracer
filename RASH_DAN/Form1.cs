using System;
using System.Windows.Forms;

namespace RASH_DAN
{
#if NANOCAD
    using Teigha.DatabaseServices;
    using HostMgd.ApplicationServices;
    using HostMgd.EditorInput;
#elif AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Runtime;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.ApplicationServices;
    using System.Reflection;
    using Autodesk.AutoCAD.EditorInput;
#endif

    public partial class Form1 : Form
    {
        ObjectId ID;
        string strNAME = "-";
        string strSVOI = "-";
        string strNeSVI = "-";
        string strDlinK = "-";
        string strHOZ = "-";
        string HANDL="-";
        string strVIS = "-";
        string strDlin = "-";
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            int Schet = 0;
            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                Editor ed =
                    Application.DocumentManager.MdiActiveDocument.Editor;
                try
                {
                    // Просим пользователя выбрать примитив
                    PromptEntityResult ers = ed.GetEntity("Укажите примитив ");
                    // Открываем выбранный примитив
                    Entity ent = (Entity)tr.GetObject(ers.ObjectId, OpenMode.ForWrite);
                    ID = ent.ObjectId;
                    HANDL = ent.Handle.ToString();
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
                        ////Application.ShowAlertDialog(strDan);
                    }
                    tr.Commit();
                }
                catch
                {
                    tr.Abort();
                }
                //Form4 form1 = new Form4();
                //form1.Show();
                this.label4.Text = HANDL;
                this.label9.Text = strDlinK;
                this.textBox1.Text = strSVOI;
                this.textBox2.Text = strNeSVI;
                this.textBox3.Text = strHOZ;
                this.textBox5.Text = strNAME;
                this.textBox4.Text = strVIS;
                this.textBox6.Text = strDlin;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();
            strSVOI = this.textBox1.Text;
            strNeSVI = this.textBox2.Text;
            strHOZ = this.textBox3.Text;
            strNAME=this.textBox5.Text;
            strVIS = this.textBox4.Text;
            strDlin = this.textBox6.Text;
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
