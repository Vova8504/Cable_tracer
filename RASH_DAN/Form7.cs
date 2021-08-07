using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace RASH_DAN
{
#if NANOCAD
    using Teigha.DatabaseServices;
    using HostMgd.ApplicationServices;
    using HostMgd.EditorInput;
#elif AUTOCAD
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.EditorInput;
#endif

    public partial class Form7 : Form
    {
        public Form7()
        {
            InitializeComponent();
        }
        private void Form7_Load(object sender, EventArgs e)
        {
            List<String> VKisSlov = HCtenSlovNod("OTPK");
            foreach (string Strok in VKisSlov)
            {
                string[] Kabel = Strok.Split(':');
                if (Kabel.Length > 1)
                    this.dataGridView1.Rows.Add(Kabel[0], Kabel[1]);
            }
        }
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
    }
}
