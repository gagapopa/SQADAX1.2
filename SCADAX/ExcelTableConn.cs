using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using  ClosedXML.Excel;
using Telerik.WinControls.UI;
//using Excel = Microsoft.Office.Interop.Excel;

namespace SCADAX
{
    class ExcelTableConn
    {
        List<CPU> CPUs;
        XLWorkbook XL;
        public Dictionary<string, string> userdictionary; //conect + "#" + zlkkey + "#" + hvostkey, userzlk

        public ExcelTableConn(List<CPU> cpUs)
        {
            CPUs = cpUs;
            openexcel();
             //createsheets(CPUs);
             userdictionary = new Dictionary<string, string>();
        }

        public void createsheets()
        {
            
            foreach (var cpu in CPUs)
            {

                var sheet = XL.Worksheets.Add(cpu.Name);
                // подготовка Шаблона

                var r1 = sheet.Range(1,5,2,5);
                r1.Cell(1,1).Value = "Номер разъема";
                r1.Merge();

                var r2 = sheet.Range(1, 6, 2, 6);
                r2.Cell(1, 1).Value = "№ входа";
                r2.Merge();

                var r3 = sheet.Range(1, 7, 2, 7);
                r3.Cell(1, 1).Value = "Обозначение KKS / Номер позиции";
                r3.Merge();


                var r4 = sheet.Range(1, 8, 2, 8);
                r4.Cell(1, 1).Value = "Наименование сигнала";
                r4.Merge();
                
                var r5 = sheet.Range(1, 9, 2, 9);
                r5.Cell(1, 1).Value = "KKS кабеля 1";
                r5.Merge();
                
                
                var r6 = sheet.Range(1, 10, 2, 10);
                r6.Cell(1, 1).Value = "Тип, жильность и сечение кабеля 1";
                r6.Merge();
                

                var r7 = sheet.Range(1, 11, 2, 11);
                r7.Cell(1, 1).Value = "KKS кабеля 2";
                r7.Merge();
                

                var r8 = sheet.Range(1, 12, 2, 12);
                r8.Cell(1, 1).Value = "Тип, жильность и сечение кабеля 2";
                r8.Merge();
                

                var r9 = sheet.Range(1, 13, 2, 13);
                r9.Cell(1, 1).Value = "Марка цепи";
                r9.Merge();

                sheet.Column(1).LastCellUsed();
                sheet.Columns(5, 13).AdjustToContents();
                ///------------------
                int j2 = 4;
                int j3 = 0;
                foreach (var m in cpu.modules)
                {
                    int j4 = 0;
                    int mergetype = 4;// сделать заисимость от типа КМС
                    foreach (var s in m.Signals)
                    {

                        //try
                        //{

                            if (s.KKS != "")
                            {
                                sheet.Column(1).AdjustToContents();
                                sheet.Column(6).LastCellUsed().CellBelow(mergetype).Value = j4.ToString();
                                sheet.Range(sheet.Column(6).LastCellUsed().Address, sheet.Column(6).LastCellUsed().CellBelow(mergetype-1).Address).Merge();

                                sheet.Column(7).LastCellUsed().CellBelow(mergetype).Value = s.KKS;
                                sheet.Range(sheet.Column(7).LastCellUsed().Address, sheet.Column(7).LastCellUsed().CellBelow(mergetype-1).Address).Merge();

                                sheet.Column(8).LastCellUsed().CellBelow(mergetype).Value = s.Marka;
                                sheet.Range(sheet.Column(8).LastCellUsed().Address, sheet.Column(8).LastCellUsed().CellBelow(mergetype-1).Address).Merge();

                                sheet.Column(9).LastCellUsed().CellBelow(mergetype).Value = s.Cabel;
                                sheet.Range(sheet.Column(9).LastCellUsed().Address, sheet.Column(9).LastCellUsed().CellBelow(mergetype-1).Address).Merge();

                                //sheet.Column(9).LastCellUsed().CellRight().Value = s.Key;
                               
                                
                                try
                                {
                                    var zlkcell = sheet.Column(9).LastCellUsed().CellRight();
                                    if (userdictionary[s.Key].Split(',')[0].Contains("{KKS}"))
                                    {
                                        int k = userdictionary[s.Key].Split(',')[0].IndexOf('}') + 1;
                                        zlkcell.Value = s.RealKKS + (userdictionary[s.Key].Split(',')[0].Substring(k));
                                    }
                                    else
                                        zlkcell.Value = (userdictionary[s.Key].Split(',')[0]);

                                    for (int i = 1; i < userdictionary[s.Key].Split(',').Count(); i++)
                                    {
                                        if (userdictionary[s.Key].Split(',')[i].Contains("{KKS}"))
                                        {
                                            int k = userdictionary[s.Key].Split(',')[i].IndexOf('}') + 1;
                                            zlkcell.CellBelow(i).Value = s.RealKKS + (userdictionary[s.Key].Split(',')[i].Substring(k));
                                        }
                                        else
                                            zlkcell.CellBelow(i).Value = (userdictionary[s.Key].Split(',')[i]);
                                    }
                                }
                                catch {}
                                sheet.Columns(6,10).AdjustToContents();
                            j2++;
                            }
                        //}
                        //catch { }
                        j4++;
                        
                    }
                j3++;
                }

                sheet.Columns(5, 13).AdjustToContents();
            }

            XL.SaveAs(@"C:\SQUAD\Template3.xlsx");
            

        }

        void openexcel()
        {
            try
            {
                XL = new XLWorkbook();
            }
            catch
            {
                MessageBox.Show("Не найдена БД Excel");
            }
            
            
        }

        public void Uservaluezlk(RadGridView userstable,string conect)
        {
            
            int datacolumn = userstable.Columns.Count;
            int datarow = userstable.Rows.Count;
            for (int i = 0; i < datarow; i++)
            {
                
                for (int j = 1; j < datacolumn; j++)
                {
                    string zlkkey = userstable.Rows[i].Cells[0].Value.ToString();
                    string hvostkey = userstable.Columns[j].HeaderText;
                    string userzlk = userstable.Rows[i].Cells[j].Value.ToString();
                    userdictionary.Add(conect + "#" + zlkkey + "#" + hvostkey, userzlk);
                }
            }
            
            int p = 0;
        }

        public void Usersvaluesave()
        {
            var xlvaluesave = new XLWorkbook();
            var sheetforvalue = xlvaluesave.Worksheets.Add("Словарь");
            sheetforvalue.FirstRow().Cell(1).Value = "Ключ";
            sheetforvalue.FirstRow().Cell(2).Value = "Значение";
            foreach (var item in userdictionary)
            {
                var lastrow = sheetforvalue.LastRowUsed();
                lastrow.Cell(1).CellBelow().Value = item.Key;
                lastrow.Cell(2).CellBelow().Value = item.Value;
            }
            xlvaluesave.SaveAs(@"C:\SQUAD\Uservalue.xlsx");
        }

        public void Usersvalueopen()
        {
            var xlvaluesave = new XLWorkbook(@"C:\SQUAD\Uservalue.xlsx");
            var sheetforvalue = xlvaluesave.Worksheet("Словарь");

            var r1 = sheetforvalue.Rows(sheetforvalue.FirstRowUsed().RowBelow().RowNumber(), sheetforvalue.LastRowUsed().RowNumber());
            foreach (var row in r1)
            {
                userdictionary.Add(row.Cell(1).GetString(), row.Cell(2).GetString());
            }
            int i = 0;
        }

        public void Test()
        {
            var i = CPUs.First().modules[16].Signals.FindAll(a =>  a.KKS != "");
            
        }

    }
}
