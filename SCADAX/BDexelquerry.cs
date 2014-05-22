using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DataTable = System.Data.DataTable;



namespace SCADAX
{
    public class BDexelquerry
    {
        public DataTable ConnectionsTable;
        private DataTable qresult;
        private string addres;
        private IXLWorksheet _sheet;
        XLWorkbook XL;
        public List<Signalinexcel> signals;
        private Dictionary<string, List<string>> connectDictionary;

        public class Signalinexcel
        {
            public string KKS { get; set; }
            public string Name { get; set; }
            public string T_Inp_Type_M { get; set; }
            public string T_Inp2_Type_M { get; set; }
            public string T_Out_Type_M { get; set; }
            public string T_Out2_Type_M { get; set; }
            public string KMS { get; set; }
            public string Cabel { get; set; }
            public string Connect { get; set; }
            public List<string> ZlkforkeyList; //для составления ключа
            private List<string> _zlk;
            public List<string> ZhilkiList //жилки должны быть через запятую.
            {

                get { return _zlk; } 
                set
                {
                    ZlkforkeyList = new List<string>();
                    value = _zlk;
                    value.ForEach(a => ZlkforkeyList.Add(a.Contains(KKS) ? @"{KKS}"+ a.Trim().Substring(KKS.Count()) : a));
                    
                }
            }

            public void Agregreturn(string intype)
            {
                var agreg = new List<string>();
                if (intype == "DI" || (intype == "AI"))
                {
                    agreg.Add(T_Inp_Type_M);
                    agreg.Add(T_Inp2_Type_M);
                }
                else
                {
                    agreg.Add(T_Out_Type_M);
                    agreg.Add(T_Out2_Type_M);
                }

                KMS = agreg.Aggregate("", (cur, ls) => cur.TrimEnd('$') + "$" + ls).Trim('$');
                
            }
            public Signalinexcel(string _KKS, string _Name, string _T_Inp_Type_M, string _T_Inp2_Type_M, string _T_Out_Type_M, string _T_Out2_Type_M,
                string _cabel, string _connection, string _zlkstr)
            {
                KKS = _KKS;
                Name = _Name;
                T_Inp_Type_M = _T_Inp_Type_M;
                T_Inp2_Type_M = _T_Inp2_Type_M;
                T_Out_Type_M = _T_Out_Type_M;
                T_Out2_Type_M = _T_Out2_Type_M;
                Cabel = _cabel;
                Connect = _connection;
                _zlk = new List<string>();
                _zlkstr.Split(',').ToList().ForEach(a =>_zlk.Add(a.Trim()));
                ZhilkiList = new List<string>();
            }
        }

        public BDexelquerry(string addresexel, string addresBd)
        {
            // TODO: Complete member initialization
            addres = addresexel;
            var BDconn = new BDconnect(addresBd);

            qresult = BDconn.GetBD(@"SELECT a.MARKA, b.MARKA, b.NAME, b.OBJSIGN, a.TID, a.CARDSID
                                                 FROM ISACARDS a
                                                 left join CARDS b
                                                 on a.CARDSID=b.ID
                                                 where a.CARDSID<>0");
            _sheet = openbook();
            getexceList();
            Analizeconnect();
            signals.Add(new Signalinexcel("err", "err", "err", "err", "err", "err", "err", "err", "err"));
        }      
        
        public IXLWorksheet openbook() //открывает БДэксель и возвращает страницу БД
        {         
                XL = new XLWorkbook(addres);
                _sheet = XL.Worksheet("БД");
                return _sheet;            
                      
       }

        public Signalinexcel Inptype(string KKSchan, string type)
        {
            try
            {
           
            string KKS = KKSchan.Trim(' ', '_');
            if (KKS.Contains('_'))
                KKS = KKS.Substring(0, KKS.IndexOf('_'));
            else
                KKS =
                    qresult.AsEnumerable()
                        .Where(dr => dr.Field<string>(0) == KKSchan.Trim())
                        .Select(dr => dr.Field<string>(1))
                        .Single();
            var sig = signals.Find(a => a.KKS == KKS);
            sig.Agregreturn(type);
            return sig;
            }

        
            catch {
                try
                {
                    var sig = signals.Find(a => a.KKS.Contains(KKSchan.Trim(' ', '_')));
                    sig.Agregreturn(type);
                    return sig;
                }
                catch 
                {
                    return signals.Find(a => a.KKS.Contains("err"));
                }
            }

        }

        public void CloseExel()
        {
            
        }

        public void Analizeconnect() //возвращает словарь datatable с одним столбиком Connections
        {
            
            string ret = "";
            connectDictionary = new Dictionary<string, List<string>>
            {
                {signals[0].Connect, new List<string>()}
            };

            signals.ForEach(a =>   //добавление остальных connections
            {
                if (connectDictionary.All(d => a.Connect != d.Key))
                {
                    connectDictionary.Add(a.Connect, new List<string>());
                }
            });

            foreach (string conn in connectDictionary.Keys)
            {
                string connectnumber = conn;
                connectDictionary[connectnumber].Add( signals.Find(a => a.Connect == connectnumber).ZlkforkeyList.Aggregate("",(a,b) => a + b +",").Trim(','));

                signals.FindAll(a => a.Connect == connectnumber).ForEach(k =>   //добавление остальных наборов жилок с предварительным сравнением с остальными экземплярами.
                {
                    string zlk = k.ZlkforkeyList.Aggregate("", (c, b) => c + b + ",").Trim(',');
                    if (connectDictionary[connectnumber].All(d => zlk != d))
                    {
                        connectDictionary[connectnumber].Add(zlk);
                    }
                });
                string connectzlktype = connectDictionary[connectnumber].Aggregate("", (current, lis) => current + lis + Environment.NewLine);
                ret += connectnumber + Environment.NewLine + connectzlktype;
            }

            ConnectionsTable = new DataTable();
            ConnectionsTable.Columns.Add("Connections");
            foreach (var con in connectDictionary.Keys)
            {
                ConnectionsTable.Rows.Add(con);
            }
           
        } //Проверяет переданный список конешна на разные наборы жилок.

        public DataTable GetzlkTable(string con, int ignorsign) //возвращает таблицу с заполненными полями вариантов жилок и суффиксами для переданного конешна.
        {
            var zlktable = new DataTable();
            zlktable.Columns.Add("Наборы жилок");
            connectDictionary[con].ForEach(a => zlktable.Rows.Add(a));

            var kkslocal = new List<string>();
            signals.FindAll(a => a.Connect == con).ForEach(a => kkslocal.Add(a.KKS.Substring(ignorsign)));

            int cardsid = 0;
            int i = 0;
            try
            {
                
                while (!qresult.AsEnumerable().Where(a => a.Field<string>(1).Trim('_') == kkslocal[i]).Select(a => a.Field<int>(5)).Any())
                {
                    i++;
                }
                cardsid = qresult.AsEnumerable().Where(a => a.Field<string>(1).Trim('_') == kkslocal[i]).Select(a => a.Field<int>(5)).First();
                List<string> listforcon = qresult.AsEnumerable()
                    .Where(a => a.Field<int>(5) == cardsid & (a.Field<int>(4) == -9 || a.Field<int>(4) == 1339 || a.Field<int>(4) == 1347)) //сделать отбор изаобъекта по ккс
                    .Select(a => a.Field<string>(0))
                    .ToList();
                List<string> listforcon2 = new List<string>();

                if (listforcon.Count > 1)
                {
                    listforcon2 = listforcon.Select(s => s.Trim('_').Substring(s.LastIndexOf('_'))).ToList();

                    foreach (var hvost in listforcon2)
                    {
                        zlktable.Columns.Add(hvost);
                        int zlkcount = 0;
                        for (i = 0; i < zlktable.Rows.Count; i++)
                        {
                            zlkcount = zlktable.Rows[i][0].ToString().Split(',').Count();
                            if (zlktable.Columns.Count - 2 < zlkcount)
                            {
                                zlktable.Rows[i][hvost] =
                                    zlktable.Rows[i][0].ToString().Split(',')[zlktable.Columns.Count - 2] + "," +
                                    zlktable.Rows[i][0].ToString().Split(',').Last();
                            }
                            else
                                zlktable.Rows[i][hvost] =
                                    "," + zlktable.Rows[i][0].ToString().Split(',').Last();
                        }
                    }
                }
                else
                {
                    zlktable.Columns.Add("#");
                    for (i = 0; i < zlktable.Rows.Count; i++)
                        zlktable.Rows[i][1] =
                            zlktable.Rows[i][0].ToString().Split(',')[zlktable.Columns.Count - 2] + "," +
                            zlktable.Rows[i][0].ToString().Split(',').Last();
                }
            }
            catch
            {
                zlktable.Columns[0].Caption = "Не найдены технологические объекты,невозможно определить состав Connections";
                zlktable.Columns.Add("#");
                for (i = 0; i < zlktable.Rows.Count; i++)
                    zlktable.Rows[i][1] =
                        zlktable.Rows[i][0];
            }
            i = 0;
            return zlktable;
        }

        public void getexceList() //заполняет модель сигнала и создает лист сигналов.
        {
            int KKScolumn = _sheet.FirstRowUsed().Cells(a => a.GetString().Contains("KKS")).First().WorksheetColumn().ColumnNumber();
            int typecolumn1 = _sheet.FirstRowUsed().Cells(a => a.GetString().Contains("T_Inp_Type_M")).First().WorksheetColumn().ColumnNumber() - KKScolumn;
            int typecolumn2 = _sheet.FirstRowUsed().Cells(a => a.GetString().Contains("T_Inp2_Type_M")).First().WorksheetColumn().ColumnNumber() - KKScolumn;
            int typecolumn3 = _sheet.FirstRowUsed().Cells(a => a.GetString().Contains("T_Out_Type_M")).First().WorksheetColumn().ColumnNumber() - KKScolumn;
            int typecolumn4 = _sheet.FirstRowUsed().Cells(a => a.GetString().Contains("T_Out2_Type_M")).First().WorksheetColumn().ColumnNumber() - KKScolumn;
            int cabpos = _sheet.FirstRowUsed().Cells(a => a.GetString().Contains("KKS кабеля")).First().WorksheetColumn().ColumnNumber() - KKScolumn;
            int conpos = _sheet.FirstRowUsed().Cells(a => a.GetString().Contains("Connection")).First().WorksheetColumn().ColumnNumber() - KKScolumn;

            signals = new List<Signalinexcel>();
            var rngs = _sheet.Range(_sheet.Column(KKScolumn).FirstCellUsed().CellBelow(), _sheet.Column(KKScolumn).LastCellUsed());
            rngs.Cells().ForEach(a =>
            {
                if (a.GetString() != "")
                {
                    var r1 = _sheet.Range(a.CellRight(cabpos).Address, a.CellRight(cabpos + 3).Address);//берет только первый кабель, исправить.
                    var r2 = _sheet.Range(a.CellRight(cabpos + 5).Address, a.CellRight(cabpos + 8).Address);
                    string zlk =
                        (a.CellRight(cabpos + 4).GetString().Trim(',') + "," + a.CellRight(cabpos + 9).GetString()).Trim(',');
                    string cab = r1.Cells().Aggregate("", (str, c) => str.Trim('#') + "#" + c.GetString())
                        + r2.Cells().Aggregate("", (str, c) => str.Trim('#') + "#" + c.GetString());

                    signals.Add(new Signalinexcel(a.GetString().Trim(),
                        a.CellRight(1).GetString().Trim(),
                        a.CellRight(typecolumn1).GetString().Trim(),
                        a.CellRight(typecolumn2).GetString().Trim(),
                        a.CellRight(typecolumn3).GetString().Trim(),
                        a.CellRight(typecolumn4).GetString().Trim(), cab.Trim(),
                        a.CellRight(conpos).GetString().Trim(), zlk));
                }
            });
        }

        




    }
}
