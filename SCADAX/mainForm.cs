using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using FastReport;
using FirebirdSql.Data.FirebirdClient;
using System.Text.RegularExpressions;
using Telerik.WinControls.UI;
using System.IO;
using Color = System.Drawing.Color;
using DataTable = System.Data.DataTable;

namespace SCADAX
{
    public partial class MainForm : RadForm
    {
        private List<CPU> CPUs;
        
        
        private FbConnection _fb;
        public MainForm()
        {
            InitializeComponent();
        }
        
        public DataTable SqlExecute(string sql)
        {
            if (_fb.State == ConnectionState.Closed) _fb.Open();
            FbTransaction trans = _fb.BeginTransaction();
            var cmd = new FbCommand(sql, _fb, trans);
            try
            {
                var result = new DataTable();
                new FbDataAdapter(cmd).Fill(result);
                trans.Commit();
                return result;
            }
            catch (Exception )
            {
                trans.Rollback();
                //Helpers.Messages.Error(ex.Message);
                return null;
            }
            finally
            {
                _fb.Close();
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            var fbCon = new FbConnectionStringBuilder
            {
                Charset = "WIN1251",
                UserID = "sysdba",
                Password = "masterkey",
                Database = txtPath.Text,
                ServerType = 0
            };
            _fb = new FbConnection(fbCon.ToString());
            _fb.Open();
            var fbInf = new FbDatabaseInfo(_fb);
            if (_fb.State == ConnectionState.Open)
            {
                lblStatus.Text = @"Соединение с базой " + txtPath.Text + " установлено";
                //tabControl1.Enabled = true;

                GetCPU();
                GetMnemokadri();
            }
            else
            {
                lblStatus.Text = @"Соединение отсутствует ";
                //tabControl1.Enabled = false;
            }
            
        }

        private void GetCPU()
        {
            DataTable dt = SqlExecute(@"SELECT a.ID, a.MARKA FROM CARDS a where a.OBJTYPEID = '1458' or a.OBJTYPEID = '2119' or a.OBJTYPEID = '1843'");
            rlCPUs.DataSource = dt;
            rlCPUs.DisplayMember = "MARKA";
        }

        private void GetMnemokadri()
        {
            DataTable dt = SqlExecute(@"SELECT a.ID, a.PID, a.NAME FROM GRPAGES a");
            rtVideo.DataSource = dt;
            rtVideo.DisplayMember = "NAME";
            rtVideo.ChildMember = "ID";
            rtVideo.ParentMember = "PID";
            rtVideo.ValueMember = "ID";
            rtVideo.ExpandAll();
            rtVideo.Nodes[0].Checked = true;

        }

        private void radMenuItem1_Click(object sender, EventArgs e)
        {
            GetMnemokadri();
        }

        DataTable dt13;
        private void btnBlocks_Click(object sender, EventArgs e)
        {
            progr1.Value1 = 0;
            progr1.Minimum = 0;
            progr1.Maximum = 100;
            bw2.RunWorkerAsync();
            
        }

        private void btnA4_Click(object sender, EventArgs e)
        {
            SqlExecute(@"UPDATE ISAGRPAGES SET PRINTPAGEA4 = 9 WHERE PRINTPAGEA4 <> 9");
            MessageBox.Show("Скрипт A4 применен!");
        }

        private void btnKvit_Click(object sender, EventArgs e)
        {
            foreach (RadTreeNode selnode in rtVideo.CheckedNodes)
            {
                DataTable dt = SqlExecute(@"SELECT a.CARDID, b.MARKA, b.NAME,c.NAME
                            FROM PAGECONTENTS a
                            JOIN CARDS b
                            on b.ID=a.CARDID
                            JOIN OBJTYPE c
                            on b.OBJTYPEID =c.ID
                            WHERE a.PAGEID = " + selnode.Value + " and a.CARDID<>0 and b.OBJTYPEID>0");

                rgv.DataSource = dt;
                SqlExecute(
                    @"INSERT INTO PAGECONTENTS ( GROBJTYPE, PAGEID, X, Y, WIDTH, HEIGHT, MSORT, PARAMS, DRAWTYPE, PENCOLOR, BRUSHCOLOR, PENPARAMS, CARDID, OBJMSID, GRNUM, GRADCOLOR, NAME, GROUPID, LAYERNUM) VALUES ( '14', '" +
                    selnode.Value + @"', '10', '30', '150', '25', '298', '[TEXT]=Квитировать
                        [FONTID]=65535
                        [USERFONT]=12;Tahoma;0;0;
                        [HINT]=
                        [DELTAST]=2
                        [PRESSCOLOR]=536870911', '0', '0', '-16777201', '1', NULL, NULL, '0', '0', '', '0', '1')
                        ");

                string n = SqlExecute("SELECT GEN_ID( GEN_PAGECONTENTS, 0 ) FROM RDB$DATABASE;").Rows[0][0].ToString();
                // ID добавленной кнопки
                SqlExecute(
                    @"INSERT INTO PCRECEPTOR ( TYPEID, PRIMID, PARAM_INT, PARAM_FLOAT, PARAM_ST, EVPARAMS, DPARAMS, MSORT, USERRIGHTS, EVKEYPARAMS) VALUES ( '3', '" +
                    n + @"', '-10', '0.000000', '', '1', '0', '3725', '0', '');");
                n = SqlExecute("SELECT GEN_ID( GEN_PCRECEPTOR, 0 ) FROM RDB$DATABASE;").Rows[0][0].ToString();
                // ID рецептора

                DataTable paramid = SqlExecute(@"SELECT distinct e.ID FROM PAGECONTENTS a
                                JOIN CARDS b
                                on b.ID=a.CARDID
                                JOIN OBJTYPE c
                                on b.OBJTYPEID =c.ID
                                join OBJTYPEPARAM d
                                on d.PID = c.ID
                                join CARDPARAMS e
                                on e.CARDID=b.ID and e.OBJTYPEPARAMID = d.ID
                                WHERE a.PAGEID = " + selnode.Value +
                                          " and a.CARDID<>0 and b.OBJTYPEID>0 and d.NAME = 'КОМ. КВИТИРОВАТЬ'");
                foreach (DataRow pr in paramid.Rows)
                {
                    string str =
                        @"INSERT INTO CHARTS ( PID, MODE, PARAMID, COLOR, LSIDE, NDPARAMID, CHARTSTYLE, WIDTH, STAIRS, MARKS, BITNUM, NDBITNUM, PARAMMODE) VALUES ( '" +
                        n + "', '12', '" + pr[0] + "', '0', '1', '-1', '0', '1', '0', '0', '-1', '-1', '0');";
                    SqlExecute(str);
                }
            }
            MessageBox.Show(@"Кнопки квитирования расставлены и привязаны!");
        }

        private void btnTehNonExist_Click(object sender, EventArgs e)
        {
            DataTable dt = SqlExecute(@"SELECT a.MARKA, a.NAME,   d.NAME
                                                FROM CARDS a
                                                left join PAGECONTENTS b on a.ID = b.CARDID 
                                                left join GRPAGES c on b.PAGEID = c.ID
                                                join OBJTYPE d on a.OBJTYPEID = d.ID
                                                where a.OBJTYPEID<>0 and c.NAME is NULL
                                                order by a.MARKA
                                                ");
            rgv.DataSource = dt;
            lblStatus.Text = @"Количество технологических объектов, не находящихся на видеокадрах: " + dt.Rows.Count;
        }

        private void btnIsaNonExist_Click(object sender, EventArgs e)
        {
            DataTable dt = SqlExecute(@"SELECT a.MARKA, a.NAME,  d.NAME
                                                FROM ISACARDS a
                                                left join ISAPAGECONTENTS b on a.ID = b.CARDID
                                                left join ISAOBJ c on b.PAGEID = c.ID
                                                join ISAOBJ d on a.TID = d.ID
                                                where b.PAGEID is NULL and d.NAME not like 'V%'
                                                ");
            rgv.DataSource = dt;
            lblStatus.Text = @"Количество ISA объектов, не используемых в программах: " + dt.Rows.Count;
        }

        private void btnEmptyObj_Click(object sender, EventArgs e)
        {
            DataTable dt = SqlExecute(@"SELECT b.MARKA, b.NAME, b.PLC_ID, c.MARKA
                                                FROM ISACARDS a
                                                right join CARDS b
                                                on a.CARDSID=b.ID
                                                join cards c 
                                                on b.PLC_ID = c.ID
                                                where a.CARDSID is NULL
                                                ");
            rgv.DataSource = dt;
            lblStatus.Text = @"Количество пустых тех. объектов: " + dt.Rows.Count;
        }

        private void btnMKO_Click(object sender, EventArgs e)
        {
            DataTable dt = SqlExecute(@"SELECT e.MARKA, b.MARKA, a.SFIELD, g.MARKA, c.MARKA
                                                FROM MKOVAR a
                                                left JOIN ISACARDS b
                                                on a.SCARDID = b.ID
                                                left JOIN ISACARDS c
                                                on a.RCARDID = c.ID

                                                left JOIN RESOURCES d
                                                on b.RESID = d.ID
                                                left JOIN CARDS e
                                                on d.CARDID = e.ID

                                                left JOIN RESOURCES f
                                                on c.RESID = f.ID
                                                left JOIN CARDS g
                                                on f.CARDID = g.ID
                                                ");
            rgv.DataSource = dt;

            lblStatus.Text = @"Количество 'пар' МКО: " + dt.Rows.Count;
        }

        private void btnRez_Click(object sender, EventArgs e)
        {
//            DataTable dt = SqlExecute(@"SELECT a.NAME
//                                                FROM IO_DEVICE_C a
//                                                ");
//            rgv.DataSource = dt;
//            string s = dt.Rows.Cast<DataRow>()
//                .Aggregate("", (current, dr) => current + (Convert.ToString(dr["NAME"]) + Environment.NewLine));
//            var mc = Regex.Matches(s, @"\[\d\d?\]\D*(\d\d?)");
//            int sum = mc.Cast<Match>().Sum(m => Convert.ToInt32(m.Groups[1].Value));
//            var bDqerryexcel = new BDexelquerry(textBox2.Text, txtPath.Text);
//            String stemp = "";
//            s = "";
//            string all = "";
//            ExtractAo(bDqerryexcel, ref stemp, ref s, chkCPUlist1.CheckedItems);
//            all += stemp;
//            stemp = "";
//            ExtractAi(bDqerryexcel, ref stemp, ref s, chkCPUlist1.CheckedItems);
//            all += stemp;
//            stemp = "";
//            ExtractDi(bDqerryexcel, ref stemp, ref s, chkCPUlist1.CheckedItems);
//            all += stemp;
//            stemp = "";
//            ExtractDo(bDqerryexcel, ref stemp, ref s, chkCPUlist1.CheckedItems);
//            all += stemp;
//            /*
//                        stemp = "";
//            */
//            //textBox1.Text = all;
//            mc = Regex.Matches(all, @"- (\w*) -");
//            int f = mc.Count;
//            mc = Regex.Matches(all, @"- (.*OBJ\d\d\d\d\d) -");
//            int fobj = mc.Count;
//            label2.Text = @"Количество каналов: " + sum + @", модулей: " + dt.Rows.Count + @", резерв: " +
//                          (sum - f + fobj);
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            
            openFileDialog1.FileName = @"C:\SQUAD\replase.txt";
            openFileDialog1.ShowDialog();

            int i = 0;
            if (openFileDialog1.FileName != "")
            {       
                var dtt = new DataTable();
                dtt.Columns.Add("Тип 1");
                dtt.Columns.Add("Маска 1");
                dtt.Columns.Add("Маска 2");
                foreach (string line in File.ReadLines(openFileDialog1.FileName))
                {
                    string[] l = line.Split('\t');
                    dtt.Rows.Add(l[0], l[1], l[2]);
                    i++;
                }
                rgv.DataSource = dtt;
            }
        }

        private void radButton2_Click(object sender, EventArgs e)
        {
            var dtrez = new DataTable();
            dtrez.Columns.Add("Старый KKS");
            dtrez.Columns.Add("Новый KKS");

            progr1.Maximum = rgv.Rows.Count;
            foreach (GridViewRowInfo r in rgv.Rows)
            {
                DataTable dt = SqlExecute(@"SELECT c.ID, c.MARKA
                        FROM ISACARDS c
                        where c.CARDSID in (SELECT a.ID
                        FROM CARDS a
                        join OBJTYPE b
                        on a.OBJTYPEID = b.ID
                        where b.NAME = '" + r.Cells[0].Value + "' )");


                //progr1.Value1++;
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        if (row[1].ToString().Contains(r.Cells[1].Value.ToString()))
                        {
                            string oldmarka = row[1].ToString();
                            string newmarka = row[1].ToString()
                                .Replace(r.Cells[1].Value.ToString(), r.Cells[2].Value.ToString());
                            string t = @"UPDATE ISACARDS  SET MARKA = '" + newmarka + "' WHERE ID = '" + row[0] + "'";
                            SqlExecute(t);
                            dtrez.Rows.Add(oldmarka, newmarka);
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            progr1.Maximum = rgv.Rows.Count;
            progr1.Value1 = 0;
            foreach (GridViewRowInfo r in rgv.Rows)
            {
                //progr1.Value1++;

                try
                {
                    string t = @"UPDATE ISAPOUSTTEXT 
                                    SET DATA = REPLACE(DATA,'" + r.Cells[1].Value + @"','" +
                               r.Cells[2].Value + @"')
                                    WHERE DATA like '%ALL%';
                                    ";
                    SqlExecute(t);
                }
                catch (Exception)
                {
                }
            }
            rgv.DataSource = dtrez;
        }

        private void btnExpHierarh_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            foreach (RadTreeNode cn in rtVideo.CheckedNodes)
            {
                dt.Merge(SqlExecute(@"SELECT a.ID, a.PID, a.NAME FROM GRPAGES a WHERE a.ID ='" + cn.Value + "'"));
            }

            var dr = new DataTableReader(dt);


            repHierarh.Load(@"C:\Users\artemiev.TECONPC\Documents\Visual Studio 2013\Projects\SQADAX1.1\SCADAX\hierarh.frx");

            var dataSet2 = new DataSet();
            dataSet2.Tables.Add("Employees");
            dataSet2.Tables["Employees"].Load(dr);


            // create report instance
            var report = new Report();

            // load the existing report
            report.Load(@"C:\Users\artemiev.TECONPC\Documents\Visual Studio 2013\Projects\SQADAX1.1\SCADAX\hierarh.frx");

            // register the dataset
            report.RegisterData(dataSet2);

            // run the report
            report.Show();

            // free resources used by report
            report.Dispose();
        }

        private void mainForm_Load(object sender, EventArgs e)
        {

            
        }

        private void btnTableCon1_Click(object sender, EventArgs e)
        {
            progr1.Value1 = 0;
            progr1.Minimum = 0;
            progr1.Maximum = 100;
            bw3.RunWorkerAsync();
        }

        private MatchCollection _mc;
        private DataTable ExtractDi2(IEnumerable<ListViewDataItem> c)
        {
            var dtres = new DataTable();
            dtres.Columns.Add("Шкаф");
            dtres.Columns.Add("KKS");
            dtres.Columns.Add("Наименование");
            dtres.Columns.Add("Тип сигнала");
            dtres.Columns.Add("Модуль");
            dtres.Columns.Add("Канал");
            
            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                FROM ISAPOUSTTEXT a
                where a.RESOURCE in (
                    SELECT b.ID
                    FROM RESOURCES b
                    where b.CARDID in (
                                SELECT c.ID
                                FROM CARDS c
                                where c.MARKA = '" + item.Text + @"'
                )) and a.DATA like '%DI_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }


                _mc = Regex.Matches(s, @"(.*):=.*_(\d)(\d\d)._(\d\d).*");
                foreach (Match m in _mc)
                {
                    //----- Заполнение модели
                    int modnum = (Convert.ToInt32(m.Groups[3].Value) + Convert.ToInt32(m.Groups[2].Value) * 16);
                    int signum = Convert.ToInt32(m.Groups[4].Value);
                    string kks = m.Groups[1].Value;
                    string marka = GetNameFromKKs(kks);

                    dtres.Rows.Add(item.Text, kks, marka, "DI", modnum, signum);//, eqm[0]);
                }

                _mc = Regex.Matches(s, @"(.*)_DDR.*(\d\d\d)\w\._(\d\d).*(\d\d\d)\w\._(\d\d).*");
                foreach (Match m in _mc)
                {
                    int modnum = Convert.ToInt32(m.Groups[2].Value);
                    int signum = Convert.ToInt32(m.Groups[3].Value);
                    string kks = m.Groups[1].Value;
                    string marka = GetNameFromKKs(kks);
                    //var eqm = eq.Inptype(kks, "DI").Split('#');
                    dtres.Rows.Add(item.Text, kks, marka, "DI", modnum, signum);//,  eqm[0]);
                    modnum = Convert.ToInt32(m.Groups[4].Value);
                    dtres.Rows.Add(item.Text, kks, marka, "DI", modnum, signum);//, eqm[0]);
                }
                progr1.Value1++;
            }

            return dtres;
        }

        private DataTable ExtractAi2(IEnumerable<ListViewDataItem> c)
        {
            var dtres = new DataTable();
            dtres.Columns.Add("Шкаф");
            dtres.Columns.Add("KKS");
            dtres.Columns.Add("Наименование");
            dtres.Columns.Add("Тип сигнала");
            dtres.Columns.Add("Модуль");
            dtres.Columns.Add("Канал");
            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                FROM ISAPOUSTTEXT a
                where a.RESOURCE in (
                    SELECT b.ID
                    FROM RESOURCES b
                    where b.CARDID in (
                                SELECT c.ID
                                FROM CARDS c
                                where c.MARKA = '" + item.Text + @"'
                )) and a.DATA like '%AI_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }


                _mc = Regex.Matches(s, @"(.*)\(_IO_IU(.*)_(.*).Value?.*");
                foreach (Match m in _mc)
                {
                    int modnum = (Convert.ToInt32(m.Groups[2].Value) + 1);
                    int signum = Convert.ToInt32(m.Groups[3].Value);
                    string kks = m.Groups[1].Value;
                    string marka = GetNameFromKKs(kks);
                    dtres.Rows.Add(item.Text, kks, marka, "AI", modnum, signum);
                }
                progr1.Value1++;
            }
            return dtres;
        }

        private DataTable ExtractDo2(IEnumerable<ListViewDataItem> c)
        {
            var dtres = new DataTable();
            dtres.Columns.Add("Шкаф");
            dtres.Columns.Add("KKS");
            dtres.Columns.Add("Наименование");
            dtres.Columns.Add("Тип сигнала");
            dtres.Columns.Add("Модуль");
            dtres.Columns.Add("Канал");
            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                                                FROM ISAPOUSTTEXT a
                                                where a.RESOURCE in (
                                                    SELECT b.ID
                                                    FROM RESOURCES b
                                                    where b.CARDID in (
                                                                SELECT c.ID
                                                                FROM CARDS c
                                                                where c.MARKA = '" + item.Text + @"'
                                                )) and a.DATA like '%DO_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }
                //progr1.Value++;


                _mc = Regex.Matches(s, @"DO\d\d_(\d)(\d\d)\( 0, (.*)");
                foreach (Match m in _mc)
                {
                    MatchCollection mc2 = Regex.Matches(m.Groups[3].Value, @"\b(\w*)\b");
                    int i = 0;
                    foreach (Match m2 in mc2)
                    {
                        if ((m2.Groups[1].Value != "") && (m2.Groups[1].Value != "FALSE"))
                        {
                            int modnum = (Convert.ToInt32(m.Groups[2].Value) + Convert.ToInt32(m.Groups[1].Value) * 16);
                            int signum = i;
                            string kks = m2.Groups[1].Value;
                            string marka = GetNameFromKKs(kks);
                            dtres.Rows.Add(item.Text, kks, marka, "DO", modnum, signum);
                            i++;
                        }
                    }
                }
                progr1.Value1++;
            }
            return dtres;
        }

        private DataTable ExtractAo2(IEnumerable<ListViewDataItem> c)
        {
            var dtres = new DataTable();
            dtres.Columns.Add("Шкаф");
            dtres.Columns.Add("KKS");
            dtres.Columns.Add("Наименование");
            dtres.Columns.Add("Тип сигнала");
            dtres.Columns.Add("Модуль");
            dtres.Columns.Add("Канал");
            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                            FROM ISAPOUSTTEXT a
                            where a.RESOURCE in (
                                SELECT b.ID
                                FROM RESOURCES b
                                where b.CARDID in (
                                            SELECT c.ID
                                            FROM CARDS c
                                            where c.MARKA = '" + item.Text + @"'
                            )) and a.DATA like '%AO_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }


                _mc = Regex.Matches(s, @"_IO_QU(\d\d?)_(\d).*\((.*).OUT");
                foreach (Match m in _mc)
                {
                    int modnum = (Convert.ToInt32(m.Groups[1].Value) + 1);
                    int signum = Convert.ToInt32(m.Groups[2].Value);
                    string kks = m.Groups[3].Value;
                    string marka = GetNameFromKKs(kks);
                    dtres.Rows.Add(item.Text, kks, marka, "AO", modnum, signum);
                }
                progr1.Value1++;
            }
            return dtres;
        }
        
        
        private string GetNameFromKKs(string kks)
        {
            DataTable r = SqlExecute(@"SELECT a.NAME
                    FROM ISACARDS a
                    Where a.MARKA ='" + kks + "'");
            if (r.Rows.Count != 0)
            {
                return r.Rows[0][0].ToString();
            }
            return "";
        }

        private DataTable dt12;
        private void bw1_DoWork(object sender, DoWorkEventArgs e)
        {
            DataTable dt = SqlExecute(@"SELECT a.MARKA, a.NAME, a.INITIALVALUE
                                            FROM ISACARDS a
                                            where a.TID in (
                                            SELECT b.ID
                                            FROM ISAOBJ b
                                            where b.NAME in( 'AD3_v1' ))");
            string s = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                s += Convert.ToString(dr["MARKA"]) + "," + (Convert.ToString(dr["INITIALVALUE"])) + Environment.NewLine;
            }
            var mc = Regex.Matches(s, @"(\d?\d)\((.*?)\)");

            foreach (Match m in mc)
            {
                string temp = "";
                for (int i = 1; i <= Convert.ToInt32(m.Groups[1].Value); i++)
                {
                    if (i == Convert.ToInt32(m.Groups[1].Value))
                    {
                        temp += m.Groups[2].Value;
                    }
                    else
                    {
                        temp += m.Groups[2].Value + ",";
                    }
                }
                s = s.Replace(m.Groups[1].Value + "(" + m.Groups[2].Value + ")", temp);
            }
            string[] sm = s.Split('\n');


            rgv.Columns.Clear();


            dt12 = new DataTable();

            dt12.Columns.Add("KKS");
            dt12.Columns.Add("НШ");
            dt12.Columns.Add("НА");
            dt12.Columns.Add("НПА");
            dt12.Columns.Add("НП");
            dt12.Columns.Add("ВП");
            dt12.Columns.Add("ВПА");
            dt12.Columns.Add("ВА");
            dt12.Columns.Add("ВШ");
            dt12.Columns.Add("Упаковка");
            int k = 0;
            foreach (string sms in sm)
            {
                //progr1.Value1++;
                k++;
                bw1.ReportProgress(Convert.ToInt32(k*100/sm.Count()));
                string[] smsm = sms.Split(',');
                try
                {
                    if (smsm[1] != "")
                    {
                        dt12.Rows.Add(smsm[0], smsm[5], smsm[6], smsm[7], smsm[8], smsm[9], smsm[10], smsm[11], smsm[12],
                            Convert.ToString(Convert.ToInt32(smsm[19]), 2).PadLeft(8, '0'));
                    }
                    else
                    {
                        dt12.Rows.Add(smsm[0]);
                    }
                }
                catch
                {
                }
            }


        }

        private void btmSigs_Click(object sender, EventArgs e)
        {

            progr1.Value1 = 0;
            progr1.Minimum = 0;
            progr1.Maximum = 100;
            bw1.RunWorkerAsync();            
        }

        private void bw1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progr1.Value1 = e.ProgressPercentage;
        }

        private void bw1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            rgv.DataSource = dt12;

            //MessageBox.Show(Convert.ToByte(  radGridView1.Rows[1].Cells[9].Value.ToString()   ).ToString());

            foreach (var t in rgv.Rows)
            {
                for (int n = 2; n < 9; n++)
                {
                    string t1 = t.Cells[9].Value.ToString();
                    if (t1[9 - n] == '1')
                    {
                        t.Cells[n].Style.BackColor = Color.LightSalmon;
                        t.Cells[n].Style.CustomizeFill = true;
                    }
                }
            }
        }

        private void bw2_DoWork(object sender, DoWorkEventArgs e)
        {
            DataTable dt = SqlExecute(@"SELECT a.DATA, a.RESOURCE
                                                                FROM ISAPOUSTTEXT a
                                                                where a.DATA like '%_FCOM%'");
            string s = "";
            for (int ii = 0; ii < dt.Rows.Count; ii++)
            {
                DataRow dr = dt.Rows[ii];
                s += (Convert.ToString(dr["DATA"]));
            }

            var mc = Regex.Matches(s, @"(\w*)_FCOM\((.*)\.Wvid\)");
            dt13 = new DataTable();
            dt13.Columns.Add("KKS");
            dt13.Columns.Add("ТЗ закр");
            dt13.Columns.Add("ТЗ откр");
            dt13.Columns.Add("ЛТЗ закр");
            dt13.Columns.Add("ЛТЗ откр");
            dt13.Columns.Add("Запр закр");
            dt13.Columns.Add("Запр откр");
            dt13.Columns.Add("Авт закр");
            dt13.Columns.Add("Авт откр");
            dt13.Columns.Add("Лог закр");
            dt13.Columns.Add("Лог откр");
            dt13.Columns.Add("Опер закр");
            dt13.Columns.Add("Опер откр");
            dt13.Columns.Add("Ком МЩУ");

            var k= 0;
            foreach (Match m in mc)
            {
                k++;
                bw2.ReportProgress(Convert.ToInt32(k*100/mc.Count));
                string temp = Regex.Replace(m.Groups[2].Value, @"(\.C.*?),", ",");
                var mc2 = Regex.Matches(temp, @"(.*?),");
                var mass = new string[16];
                for (int i = 0; i < mc2.Count; i++)
                {
                    if ((mc2[i].Groups[1].Value != "") && (mc2[i].Groups[1].Value != " ")) mass[i] = "+";
                }

                dt13.Rows.Add(m.Groups[1].Value, mass[0], mass[1], mass[2], mass[3], mass[4], mass[5], mass[6], mass[7],
                    mass[8], mass[9], mass[10], mass[11], mass[12]);
            }


            
        }

        private void bw2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progr1.Value1 = e.ProgressPercentage;
        }

        private void bw2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            rgv.DataSource = dt13;
        }

        DataTable dt14;
        private void bw3_DoWork(object sender, DoWorkEventArgs e)
        {
            //eq = new BDexelquerry(@"C:\SQUAD\Пермь_ТМО_БД_Малыхов_Слитая.xlsx", txtPath.Text);
            var v = 0;
            if (checkDI.Checked) v++;
            if (checkAI.Checked) v++;
            if (checkDO.Checked) v++;
            if (checkAO.Checked) v++;
            
            int k = 0;
            dt14 = new DataTable();
            if (checkDI.Checked) {
                dt14.Merge(ExtractDi2(rlCPUs.CheckedItems));
                k++;
                bw3.ReportProgress(Convert.ToInt32(k *100/v));
            }
            if (checkAI.Checked) {dt14.Merge(ExtractAi2(rlCPUs.CheckedItems));
                k++;
                bw3.ReportProgress(Convert.ToInt32(k * 100 / v));
            }
            if (checkDO.Checked) {dt14.Merge(ExtractDo2(rlCPUs.CheckedItems));
                k++;
                bw3.ReportProgress(Convert.ToInt32(k * 100 / v));
            }
            if (checkAO.Checked) {dt14.Merge(ExtractAo2(rlCPUs.CheckedItems));
                k++;
                bw3.ReportProgress(Convert.ToInt32(k * 100 / v));
            }
            
        }

        private void bw3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            rgv.DataSource = dt14;
        }

        private void bw3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progr1.Value1 = e.ProgressPercentage;
        }

        private void btnTableCon2_Click(object sender, EventArgs e)
        {
            CPUs = new List<CPU>();
            foreach (ListViewDataItem cpu in rlCPUs.CheckedItems)
            {
                CPUs.Add(new CPU(cpu.Text));
            }
            var bq = new BDexelquerry(@"D:\Тимур\БД Киров.xlsx", txtPath.Text);
            if (checkDI.Checked) ExtractDi1(bq, rlCPUs.CheckedItems);
            if (checkAI.Checked) ExtractAi1(bq, rlCPUs.CheckedItems);
            if (checkDO.Checked) ExtractDo1(bq, rlCPUs.CheckedItems);
            if (checkAO.Checked) ExtractAo1(bq, rlCPUs.CheckedItems);

           // CPUs.se
            
            
        }

        private void ExtractDi1(BDexelquerry bDqerryexcel, IEnumerable<ListViewDataItem> c)
        {

            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                FROM ISAPOUSTTEXT a
                where a.RESOURCE in (
                    SELECT b.ID
                    FROM RESOURCES b
                    where b.CARDID in (
                                SELECT c.ID
                                FROM CARDS c
                                where c.MARKA = '" + item.Text + @"'
                )) and a.DATA like '%DI_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }


                _mc = Regex.Matches(s, @"(.*):=.*_(\d)(\d\d)._(\d\d).*");
                foreach (Match m in _mc)
                {
                    //----- Заполнение модели
                    int modnum = (Convert.ToInt32(m.Groups[3].Value) + Convert.ToInt32(m.Groups[2].Value) * 16);
                    int signum = Convert.ToInt32(m.Groups[4].Value);
                    string kks = m.Groups[1].Value;
                    string marka = GetNameFromKKs(kks);
                    var fstr = bDqerryexcel.Inptype(kks, "DI");
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Type = "DI";
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Signals[signum] = new Signal(kks, marka,
                       fstr.KMS, fstr.Connect,fstr.Cabel,fstr.KKS, fstr.ZhilkiList, fstr.ZlkforkeyList);
                }

                _mc = Regex.Matches(s, @"(.*)_DDR.*(\d\d\d)\w\._(\d\d).*(\d\d\d)\w\._(\d\d).*");
                foreach (Match m in _mc)
                {
                    int modnum = Convert.ToInt32(m.Groups[2].Value);
                    int signum = Convert.ToInt32(m.Groups[3].Value);
                    string kks = m.Groups[1].Value;
                    string marka = GetNameFromKKs(kks);
                    var fstr = bDqerryexcel.Inptype(kks, "DI");
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Type = "DI";
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Signals[signum] = new Signal(kks, marka,
                       fstr.KMS, fstr.Connect, fstr.Cabel, fstr.KKS, fstr.ZhilkiList, fstr.ZlkforkeyList);
                }
            }


        }

        private void ExtractAi1(BDexelquerry bDqerryexcel, IEnumerable<ListViewDataItem> c)
        {
            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                FROM ISAPOUSTTEXT a
                where a.RESOURCE in (
                    SELECT b.ID
                    FROM RESOURCES b
                    where b.CARDID in (
                                SELECT c.ID
                                FROM CARDS c
                                where c.MARKA = '" + item.Text + @"'
                )) and a.DATA like '%AI_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }


                _mc = Regex.Matches(s, @"(.*)\(_IO_IU(.*)_(.*).Value?.*");
                foreach (Match m in _mc)
                {
                    int modnum = (Convert.ToInt32(m.Groups[2].Value) + 1);
                    int signum = Convert.ToInt32(m.Groups[3].Value);
                    string kks = m.Groups[1].Value;
                    string marka = GetNameFromKKs(kks);
                    var fstr = bDqerryexcel.Inptype(kks, "AI");
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Type = "AI";
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Signals[signum] = new Signal(kks, marka,
                       fstr.KMS, fstr.Connect, fstr.Cabel, fstr.KKS, fstr.ZhilkiList, fstr.ZlkforkeyList);
                }
            }

        }

        private void ExtractDo1(BDexelquerry bDqerryexcel, IEnumerable<ListViewDataItem> c)
        {
            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                                                FROM ISAPOUSTTEXT a
                                                where a.RESOURCE in (
                                                    SELECT b.ID
                                                    FROM RESOURCES b
                                                    where b.CARDID in (
                                                                SELECT c.ID
                                                                FROM CARDS c
                                                                where c.MARKA = '" + item.Text + @"'
                                                )) and a.DATA like '%DO_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }
                //progr1.Value++;


                _mc = Regex.Matches(s, @"DO\d\d_(\d)(\d\d)\( 0, (.*)");
                foreach (Match m in _mc)
                {
                    MatchCollection mc2 = Regex.Matches(m.Groups[3].Value, @"\b(\w*)\b");
                    int i1 = 0;
                    foreach (Match m2 in mc2)
                    {
                        if ((m2.Groups[1].Value != "") && (m2.Groups[1].Value != "FALSE") && (m2.Groups[1].Value.Count() > 6))
                        {
                            int modnum = (Convert.ToInt32(m.Groups[2].Value) + Convert.ToInt32(m.Groups[1].Value) * 16);
                            int signum = i1;
                            string kks = m2.Groups[1].Value;
                            string marka = GetNameFromKKs(kks);
                            var fstr = bDqerryexcel.Inptype(kks, "DO");
                            
                            CPUs.Find(a => a.Name == item.Text).modules[modnum].Type = "DO";
                            CPUs.Find(a => a.Name == item.Text).modules[modnum].Signals[signum] = new Signal(kks, marka,
                                    fstr.KMS, fstr.Connect, fstr.Cabel, fstr.KKS, fstr.ZhilkiList, fstr.ZlkforkeyList);
                            i1++;
                        }
                    }
                }
            }

        }

        private void ExtractAo1(BDexelquerry bDqerryexcel, IEnumerable<ListViewDataItem> c)
        {
            foreach (ListViewDataItem item in c)
            {
                DataTable dt = SqlExecute(@"SELECT a.ISAOBJID, a.DATA
                            FROM ISAPOUSTTEXT a
                            where a.RESOURCE in (
                                SELECT b.ID
                                FROM RESOURCES b
                                where b.CARDID in (
                                            SELECT c.ID
                                            FROM CARDS c
                                            where c.MARKA = '" + item.Text + @"'
                            )) and a.DATA like '%AO_ALL%'");
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];

                    if (dr.IsNull("ISAOBJID") == false)
                    {
                        s += (Convert.ToString(dr["DATA"]));
                    }
                }


                _mc = Regex.Matches(s, @"_IO_QU(\d\d?)_(\d).*\((.*).OUT");
                foreach (Match m in _mc)
                {
                    int modnum = (Convert.ToInt32(m.Groups[1].Value) + 1);
                    int signum = Convert.ToInt32(m.Groups[2].Value);
                    string kks = m.Groups[3].Value;
                    string marka = GetNameFromKKs(kks);
                    var fstr = bDqerryexcel.Inptype(kks, "AO");
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Type = "AO";
                    CPUs.Find(a => a.Name == item.Text).modules[modnum].Signals[signum] = new Signal(kks, marka,
                       fstr.KMS, fstr.Connect, fstr.Cabel, fstr.KKS, fstr.ZhilkiList, fstr.ZlkforkeyList);
                }
            }

        }

        private BDexelquerry forcon;
        private void radButton7_Click(object sender, EventArgs e)
        {
            _toexcel = new ExcelTableConn(CPUs);
            forcon = new BDexelquerry(@"D:\Тимур\БД Киров.xlsx", txtPath.Text);
            rlCPUs.DataSource = forcon.ConnectionsTable;
            rlCPUs.DisplayMember= "Connections";
            //rlCPUs.DataMember = "Connections";
        }

        private void radButton8_Click(object sender, EventArgs e)
        {
            var t = new DataTable();
            foreach (var a in rlCPUs.CheckedItems)
            {
                t.Merge(forcon.GetzlkTable(a.Text, 0));

            }


            rgv.DataSource = t;

            //rgv.DataSource = forcon.GetzlkTable(rlCPUs.CheckedItems[0].Text, 2);
        }

        private ExcelTableConn _toexcel;
        int i = 1;
        private void saveuser_Click(object sender, EventArgs e)
        {
            
            if (i < 2)
            {
                //toexcel = new ExcelTableConn(CPUs);
                _toexcel.Uservaluezlk(rgv, rlCPUs.CheckedItems.First().Text);
                i++;
            }
            else
            {
                _toexcel.Uservaluezlk(rgv, rlCPUs.CheckedItems.First().Text);
            }
            _toexcel.Usersvaluesave();
            //string aa = rgv.Rows[1].Cells[0].Value.ToString();
            //Console.Write(aa);
        }

        private void Printbtn_Click(object sender, EventArgs e)
        {
            //CPUs = null;
            //var fs = new FileStream(@"C:\file.s", FileMode.Open, FileAccess.Read, FileShare.Read);
            //var bf = new BinaryFormatter();
            //CPUs = (List<CPU>)bf.Deserialize(fs);
            //fs.Close();

            var xl = new XLWorkbook(XLEventTracking.Disabled);
            foreach (var cpu in CPUs)
            {
                var tempsheet = xl.AddWorksheet(cpu.Name);
                tempsheet.Row(1).Cell(1).Value = cpu.Name;
                tempsheet.Row(1).Cell(1).Style.Font.Bold = true;
                var i = 0;
                for (int indmod = 0; indmod < cpu.modules.Count; indmod++)
                {
                    var module = cpu.modules[indmod];
                    if (module.Type == "") continue;
                    for (int index = 0; index < module.Signals.Count; index++)
                    {
                        var signal = module.Signals[index];
                        if (string.IsNullOrEmpty(signal.Key)) continue;
                        //tempsheet.LastRowUsed().RowBelow().Cell(1).Value = signal.KKS;
                        //tempsheet.LastRowUsed().RowBelow().Cell(3).Value = index;

                        if (signal.KMS != null) {if (signal.KMS.Contains("$")) signal.KMS = signal.KMS.Split('$')[0];}
                        // ищем первый попавшийся KMS
                        var indextemp = index;
                        while (string.IsNullOrEmpty(module.Signals[indextemp].KMS) && indextemp<49)
                        {
                            indextemp++;
                        }

                        if (string.IsNullOrEmpty(module.Signals[indextemp].KMS))
                        {
                            indextemp = index;
                            while (string.IsNullOrEmpty(module.Signals[indextemp].KMS) && indextemp > 1)
                            {
                                indextemp--;
                            }
                        }

                        var file =
                            @"C:\Users\malyhov\Google Диск\Рабочая\SQUAD\SQADAX1.2\SQADAX1.2\Templates\" +
                            module.Signals[indextemp].KMS + ".xlsx";

                        if (!File.Exists(file)) continue;


                        var XL2 = new XLWorkbook(file);
                        var sheetxl2 = XL2.Worksheet(1);

                        //Определяем размерность KMS
                        var rangekms = sheetxl2.Range(sheetxl2.FirstRowUsed().Cell(1).Address,
                            sheetxl2.LastRowUsed().Cell(13).Address);
                        
                        var kmschanels = rangekms.Cells().Count(a => a.Value.ToString().Contains("chanel"));
                        
                        //Заполняем левую часть таблицы для KMS
                        rangekms.Cells()
                                .First(a => a.Value.ToString().Contains("{#module}"))
                                .Value = rangekms
                                            .Cells()
                                            .First(a => a.Value.ToString().Contains("{#module}"))
                                            .Value.ToString().Replace("{#module}",indmod.ToString());
                        rangekms
                                    .Cells()
                                    .First(a => a.Value.ToString().Contains("{#kms}"))
                                    .Value = rangekms
                                            .Cells()
                                            .First(a => a.Value.ToString().Contains("{#kms}"))
                                            .Value.ToString().Replace("{#kms}", ((int)index / kmschanels+1).ToString());
                        rangekms
                                    .Cells()
                                    .First(a => a.Value.ToString().Contains("{krejt}"))
                                    .Value = rangekms
                                            .Cells()
                                            .First(a => a.Value.ToString().Contains("{krejt}"))
                                            .Value.ToString().Replace("{krejt}", ((int) indmod/16+1).ToString());
                        try
                        {
                            rangekms
                                .Cells()
                                .First(a => a.Value.ToString().Contains("XP{4+#kms}"))
                                .Value = rangekms
                                    .Cells()
                                    .First(a => a.Value.ToString().Contains("XP{4+#kms}"))
                                    .Value.ToString().Replace("{4+#kms}", (4 + (int)index / kmschanels + 1).ToString());
                        }
                        catch {}
                        //sheetxl2.Range(1, 1, 20, 20)
                        //    .Cells()
                        //    .First(a => a.Value.ToString().Contains("{4+#kms}"))
                        //    .Value = "XP"+((int)index / kmschanels + 1).ToString();

                        //По размерности KMS отрисовываем сигналы в соответствии с размерностью
                        index--;
                        for (int j = 0; j < kmschanels; j++)
                        {   index++;
                        rangekms.Cells()
                                .First(a => a.Value.ToString().Contains("#chanel" + (j + 1)))
                                .Value = index;

                        rangekms.Cells()
                                .First(a => a.Value.ToString().Contains("#kks" + (j + 1)))
                                .Value = module.Signals[index].KKS;
                        rangekms.Cells()
                                .First(a => a.Value.ToString().Contains("#name" + (j + 1)))
                                .Value = module.Signals[index].Marka;

                            string[] cabels = module.Signals[index].Cabel.Split('#');
                            rangekms.Cells()
                                .First(a => a.Value.ToString().Contains("#kkscabel1." + (j + 1)))
                                .Value = cabels[0];
                            if (cabels.Count() > 1)
                                rangekms.Cells()
                                    .First(a => a.Value.ToString().Contains("#typecabel1." + (j + 1)))
                                    .Value = cabels[1];
                            if (cabels.Count() > 3)
                                rangekms.Cells()
                                    .First(a => a.Value.ToString().Contains("#kkscabel2." + (j + 1)))
                                    .Value = cabels[3];
                            if (cabels.Count() > 4)
                                rangekms.Cells()
                                    .First(a => a.Value.ToString().Contains("#typecabel2." + (j + 1)))
                                    .Value = cabels[4];


                            try
                            {
                             string[] zhilki = null;
                             zhilki = _toexcel.userdictionary[module.Signals[index].Key].Split(',');

                                for (int index1 = 0; index1 < zhilki.Length; index1++)
                                {
                                    zhilki[index1] = zhilki[index1].Replace("{KKS}", module.Signals[index].RealKKS);
                                }

                                if (zhilki.Count()>=1)
                                    rangekms
                                    .Cells()
                                    .First(a => a.Value.ToString().Contains("#zhilka" + (j + 1) + ".1"))
                                    .Value = zhilki[0];
                                if (zhilki.Count() >= 2)
                                    rangekms
                                        .Cells()
                                        .First(a => a.Value.ToString().Contains("#zhilka" + (j + 1) + ".2"))
                                        .Value = zhilki[1];
                                if (zhilki.Count() >= 3)
                                    rangekms
                                        .Cells()
                                        .First(a => a.Value.ToString().Contains("#zhilka" + (j + 1) + ".3"))
                                        .Value = zhilki[2];
                                if (zhilki.Count() >= 4)
                                    rangekms
                                        .Cells()
                                        .First(a => a.Value.ToString().Contains("#zhilka" + (j + 1)+".4"))
                                        .Value = zhilki[3];


                            }
                            catch (Exception)
                            {
                                
                            }
                            
                        }
                        //Очищаем от переменных шаблона
                        rangekms.Cells().Where(a => a.Value.ToString().Contains("#")).ForEach(a => a.Value = "");
                        rangekms.Cells().Where(a => a.Value.ToString().Contains("err")).ForEach(a => a.Value = "");
                        // Копируем таблицу в результат
                        var rngk = tempsheet.Range(1 + i*35, 1, 35 + i*35, 20);
                        rangekms.CopyTo(rngk);
                        XL2.Dispose();
                        i++;
                    }
                }
            }
            xl.SaveAs(@"C:\1.xlsx");
            Process.Start(@"C:\1.xlsx");
        }

        private void autoprav_Click(object sender, EventArgs e)
        {
            _toexcel = new ExcelTableConn(CPUs);
            var it = rlCPUs.Items;
            foreach (var item in it)
            {
                rgv.DataSource = forcon.GetzlkTable(item.Text, 0);
                _toexcel.Uservaluezlk(rgv, item.Text);
            }
            //toexcel.Test();
            i++;
        }

        private void exporttoXL_Click(object sender, EventArgs e)
        {
            var xl = new XLWorkbook(XLEventTracking.Disabled);
            var shet = xl.AddWorksheet("вывод");
            int i = 1;

            foreach (var row in rgv.Rows)
            {
                for (int j=0; j < 6;j++)
                {
                    shet.Row(i).Cell(j+1).Value = row.Cells[j].Value;
                }
                i++;
            }

            xl.SaveAs(@"C:\2.xlsx");
            Process.Start(@"C:\2.xlsx");
        }

        private void openuser_Click(object sender, EventArgs e)
        {
            _toexcel.Usersvalueopen();
        }

        private void rlCPUs_SelectedItemChanged(object sender, EventArgs e)
        {

        }
        private void KMSs_Click(object sender, EventArgs e)
        {
            var kmsdictionary = new Dictionary<string, int>
            {
                {"AI16", 8},
                {"AIG16-TCC4A",4},
                {"AIG16-TCC4PW",4},
                {"AOC4-TCC4A",4},
                {"AOC8",8},
                {"DI32-TCC8L",8},
                {"DI32-TCC9A",8},
                {"DI32-TCC220AC",8},
                {"DI32-TCC220DC",8},
                {"DI48-24",16},
                {"DO16r-220",6},
                {"DO24r",8},
                {"DO32-TCB04",8},
                {"DO32-TCB08H",8},
                {"DO32-TCB08RT",8},
                {"LIG16-TCC4LT",4},
                {"LIG16-TCC4LTS",4},
                {"TCB08HP",8}
            };

            var xl = new XLWorkbook(XLEventTracking.Disabled);
           
            foreach (var cpu in CPUs)
            {
                var shet = xl.AddWorksheet(cpu.Name);
                shet.Row(1).Cell(1).Value = cpu.Name;
                int strind = 1;
                int indsigcpu = 0;
                int indmod = 0;
                int aa = cpu.modules.Count(module => module.Type != "");
                foreach (var module in cpu.modules)
                {
                    indmod++;
                    if (string.IsNullOrEmpty(module.Type)) continue;
                    int indkms = 0;
                    int indsigmod = 0;
                    int reset = 0;
                    int strind2 = strind;
                    shet.Row(strind).Cell(2).Value = module.Type;
                    foreach (var signal in module.Signals.Where(signal => signal.KKS != ""))
                    {
                        indsigcpu++;
                        indsigmod++;
                        if (reset != 0)
                        {
                            reset--;
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(signal.KMS)) continue;
                        reset = kmsdictionary[signal.KMS.Split('$')[0]];
                        shet.Row(strind).Cell(3).Value = signal.KMS;
                        shet.Row(strind).Cell(4).Value = indmod;
                        indkms++;
                        shet.Row(strind).Cell(5).Value = indkms;
                        shet.Row(strind).Cell(6).Value = ((int) indmod/16+1).ToString();
                        strind++;

                    }
                    shet.Row(strind2).Cell(7).Value = indsigmod;
                }
                shet.Row(1).Cell(8).Value = indsigcpu;
            }
            xl.SaveAs(@"C:\KMS.xlsx");
            Process.Start(@"C:\KMS.xlsx");
        }

        private void loadkms_Click(object sender, EventArgs e)
        {
            var xlopen = new XLWorkbook(@"C:\KMS.xlsx",XLEventTracking.Disabled);
            var xlsave = new XLWorkbook(XLEventTracking.Disabled);
            var xlsaveshet = xlsave.AddWorksheet("вывод");
            int indkms = 0;
            foreach (var cell in xlopen.Worksheets.SelectMany(shet => shet.Column(3).Cells(a => (string) a.Value != "")))
            {
                string file;
                if (cell.GetString().Contains("#"))
                    file = @"C:\Users\malyhov\Google Диск\Рабочая\SQUAD\SQADAX1.2\SQADAX1.2\Templates\" +
                           cell.GetString().Trim('#') + ".xlsx";
                else
                    file = @"C:\Users\malyhov\Google Диск\Рабочая\SQUAD\SQADAX1.2\SQADAX1.2\Templates\" +
                           cell.Value + ".xlsx";

                if (!File.Exists(file))
                {
                    cell.Value = cell.Value + "_Не найден шаблон KMS";
                    continue;
                }

            var KMSopen = new XLWorkbook(file,XLEventTracking.Disabled);
                var KMSsh = KMSopen.Worksheet(1);
                var rangefrom = KMSsh.Range(1, 1, 40, 13);

                if (cell.GetString().Contains("#")) //если добавленный кмс то заменяем ккс на резерв
                {
                    rangefrom.Cells().Where(a => a.GetString().Contains("#name")).ForEach(b => b.Value = "Резерв");
                    rangefrom.Cells().Where(a => a.GetString().Contains("#kks")).ForEach(b => b.Value = "");
                    rangefrom.Cells().Where(a => a.GetString().Contains("#typecabel")).ForEach(b => b.Value = "");
                    rangefrom.Cells().Where(a => a.GetString().Contains("#zhilka")).ForEach(b => b.Value = "");
                }
                var krejtcell = rangefrom.Cells().First(a => a.GetString().Contains("{krejt}"));

                krejtcell.Value = krejtcell.GetString().Replace("{krejt}", cell.CellRight(3).GetString())
                                                       .Replace("{#module}", cell.CellRight().GetString())
                                                       .Replace("{#kms}", cell.CellRight(2).GetString());


                rangefrom.Cells().First(a => a.GetString().Contains("#kms")).Value = "XP" + (4+cell.CellRight(2).GetString().CastTo<int>());

                

                var rangeto = xlsaveshet.Range(indkms*40 + 1, 1, indkms*40 + 41, 13);
                rangefrom.CopyTo(rangeto);
                indkms++;

            }//Дополнительно указывать добавленные кмс, если их нет в модели.


            xlsave.SaveAs(@"C:\" + xlopen.Worksheet(1).Name+@".xlsx");
            xlopen.SaveAs(@"C:\KMS.xlsx");
        }

        private void shablonfull_Click(object sender, EventArgs e)
        {
            var xlsave = new XLWorkbook(@"C:\10CRB51GH003.xlsx", XLEventTracking.Disabled);

            foreach (var cpu in CPUs)
            {
                
            }
        }
    }
}
