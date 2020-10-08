using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Translato.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Translato
{
    class ExcelReader
    {
        delegate int leven(string str1, string str2);
        string path = "";
        Application excel;
        public Workbook wb;
        public Worksheet ws;
        Cells cells = new Cells();
        public _Excel.Range range1;

        public ExcelReader()
        {



        }
        public ExcelReader(string path, int Sheet)
        {

        }
        public static int Leven(string string1, string string2)
        {
            if (string1 == null) throw new ArgumentNullException("string1");
            if (string2 == null) throw new ArgumentNullException("string2");
            int diff;
            int[,] m = new int[string1.Length + 1, string2.Length + 1];

            for (int i = 0; i <= string1.Length; i++) { m[i, 0] = i; }
            for (int j = 0; j <= string2.Length; j++) { m[0, j] = j; }

            for (int i = 1; i <= string1.Length; i++)
            {
                for (int j = 1; j <= string2.Length; j++)
                {
                    diff = (string1[i - 1] == string2[j - 1]) ? 0 : 1;

                    m[i, j] = Math.Min(Math.Min(m[i - 1, j] + 1,
                                             m[i, j - 1] + 1),
                                             m[i - 1, j - 1] + diff);
                }
            }
            return m[string1.Length, string2.Length];
        }

        public string ReadCell(string NameCell)
        {
            byte[] mas = cells.CellToInt(NameCell);
            byte i = mas[0];
            byte j = mas[1];
            if (ws.Cells[i, j].Value != null)
                return ws.Cells[i, j].Value2;
            else
                return "";
        }
        public void CreateNewFile(string path)
        {
            this.wb = excel.Workbooks.Add();
        }
        public void WriteToCell(int i,int j, string s)
        {
            ws.Cells[i, j].Value2 = s;
        }
        public void WriteToCell(IRow curRow, int Cellindex, string Value,HSSFCellStyle style)
        {
            ICell Cell = curRow.CreateCell(Cellindex);
            Cell.SetCellValue(Value);
            Cell.CellStyle = style;
        }
        public void CreateNewSheet()
        {
            Worksheet temptsheet = wb.Worksheets.Add(After:ws);
            ws = temptsheet;
        }
        public int Read(_Excel.Worksheet sheet,out object[,] Dato,out int lengthCol) 
        {
            _Excel.Range range;
            _Excel.Range last = sheet.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastCol = last.Column;
            int lastRow = last.Row;
            range = (_Excel.Range)sheet.Range["A1", "L200"];
            Dato = (object[,])range.Value;
            //  sheet.get_Range("K1").get_Resize(dataArr.GetUpperBound(0), dataArr.GetUpperBound(1)).Value = dataArr;
            lengthCol = lastCol;
            return lastRow;
        }
        public void ExcelToList(out List<string> articolo, out List<string> articolo_it,
                                out List<string> um, out List<string> colore, out List<string> colore_name,
                                out List<string> quantita, out List<string> prezzo, out List<string> aspetto,
                                out List<string> colli, out List<string> peso, out List<string> zh, int lengthCol, int lengthRow, object[,] Data,
                                out string numFactura, out string dateFactura,out string Pereviznik, out string Nomera, out string CMR,out string Kordon)
        {
            articolo = new List<string>();
            articolo_it = new List<string>();
            um = new List<string>();
            colore = new List<string>();
            colore_name = new List<string>();
            quantita = new List<string>();
            prezzo = new List<string>();
            aspetto = new List<string>();
            colli = new List<string>();
            peso = new List<string>();
            zh = new List<string>();
            numFactura = "";
            dateFactura = "";
            int count = 1;
            int count1 = 1;
            int count2 = 1;
            int count3 = 1;
            int count4 = 1;
            int count5 = 1;
            int count6 = 1;
            int count7 = 1;
            int count8 = 1;
            int count9 = 1;
            int count10 = 1;
            int count11 = 0;
            for (int j = 1; j < lengthCol + 1; j++)
            {
                for (int i = 1; i < lengthRow + 1; i++)
                {
                    if ((j == 1) && (i > 10) && (Data[i - count, j] != null))
                    {
                        if ((Data[i - count, j].ToString().Trim() == "ARTICOLO")&&count11==0)
                        {
                            count++;
                            
                                if (Data[i, j] != null && !(Data[i, j].ToString().Contains("VETTORE")))
                                {
                                    articolo.Add(Data[i, j].ToString());
                                }
                                else
                                {
                                    articolo.Add("empty");
                                }
                            if ((count>=3)&&articolo[count - 3] == "empty"&& (articolo[count - 2]=="empty"))
                            {
                                articolo.RemoveAt(count-2);
                                articolo.RemoveAt(count - 3);
                                count11++;
                                count-=2;
                            }
                            
                        }
                    }
                    if ((j == 2) && (i > 10) && (Data[i - count1, j] != null))
                    {
                        if ((Data[i - count1, j].ToString().Trim() == "DESCRIZIONE ARTICOLO")&&count11==1)
                        {
                            count1++;
                            if (Data[i, j] != null)
                                articolo_it.Add(Data[i, j].ToString());
                            else
                            {
                                articolo_it.Add("empty");
                            }
                            if ((count1>=3)&&articolo_it[count1 - 3] == "empty"&&(articolo_it[count1 - 2] == "empty"))
                            {
                                articolo_it.RemoveAt(count1 - 2);
                                articolo_it.RemoveAt(count1 - 3);
                                count11++;
                                count1-=2;
                            }
                        }
                    }
                    if ((j == 4) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count2, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count2, j].ToString().Trim() == "UM"))
                        {
                            count2++;
                            if (Data[i, j].ToString() != "empty")
                                um.Add(Data[i, j].ToString());
                            else
                            {
                                um.Add("empty");
                            }
                        }
                    }
                    if ((j == 5) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count3, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count3, j].ToString().Trim() == "COLORE"))
                        {
                            count3++;
                            if (Data[i, j].ToString() != "empty")
                                colore.Add(Data[i, j].ToString());
                            else
                            {
                                colore.Add("empty");
                            }
                        }
                    }
                    if ((j == 6) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count4, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count4, j].ToString().Trim() == "DESCRIZIONE COLORE"))
                        {
                            count4++;
                            if (Data[i, j].ToString() != "empty")
                                colore_name.Add(Data[i, j].ToString());
                            else
                            {
                                colore_name.Add("empty");
                            }
                        }
                    }
                    if ((j == 7) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count5, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count5, j].ToString().Trim() == "QUANTITA'"))
                        {
                            count5++;
                            if (Data[i, j].ToString() != "empty")
                                quantita.Add(Data[i, j].ToString().Replace(",", "."));
                            else
                            {
                                quantita.Add("empty");
                            }
                        }
                    }
                    if ((j == 8) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count6, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count6, j].ToString().Trim() == "PREZZO"))
                        {
                            count6++;
                            if (Data[i, j].ToString() != "empty")
                                prezzo.Add(Data[i, j].ToString().Replace(",", "."));
                            else
                            {
                                prezzo.Add("empty");
                            }
                        }
                    }
                    if ((j == 9) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count7, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count7, j].ToString().Trim() == "ASPETTO ESTERIORE"))
                        {
                            count7++;
                            if (Data[i, j].ToString() != "empty")
                                aspetto.Add(Data[i, j].ToString());
                            else
                            {
                                aspetto.Add("empty");
                            }
                        }
                    }
                    if ((j == 10) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count8, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count8, j].ToString().Trim() == "COLLI"))
                        {
                            count8++;
                            if (Data[i, j].ToString() != "empty")
                                colli.Add(Data[i, j].ToString());
                            else
                            {
                                colli.Add("empty");
                            }
                        }
                    }
                    if ((j == 11) && (i > 10) && (i < (10 + count + 1)) && (Data[i - count9, j] != null))
                    {
                        if (Data[i, j] == null)
                        {
                            Data[i, j] = "empty";
                        }
                        if ((Data[i - count9, j].ToString().Trim() == "PESO NETTO"))
                        {
                            count9++;
                            if (Data[i, j].ToString() != "empty")
                                peso.Add(Data[i, j].ToString().Replace(",", "."));
                            else
                            {
                                peso.Add("empty");
                            }
                        }
                    }

                }
            }
            /*Fix if not empty*/
            int counter = 0;
            for(int i = articolo.Count + 12; i < articolo.Count + 25; i++)
            {
                if (Data[i, 1] != "empty" && Data[i, 1] != null&&Data[i,1].ToString().Length>3)
                {
                    counter = i;
                    break;
                }
            }
            String[] splitPereviznik = Data[counter, 1].ToString().Split(new char[] { ':' });
            String[] splitNomera = Data[counter+1, 1].ToString().Split(new char[] { ':' });
            String[] splitCMR = Data[counter+3, 1].ToString().Split(new char[] { ':' });
            String[] splitKordon = Data[counter+4, 1].ToString().Split(new char[] { ':' });
            Pereviznik = splitPereviznik.Last().Trim(new char[] { ' ' });
            Nomera = splitNomera.Last().Trim(new char[] { ' ' });
            CMR = splitCMR.Last().Trim(new char[] { ' ' });
            Kordon = splitKordon.Last().Trim(new char[] { ' ' });
            String[] splitFactura = Data[7, 1].ToString().Split(new char[] { ' ' });
            String[] splitDate = Data[8, 1].ToString().Split(new char[] { ' ' });
            numFactura = splitFactura.Last();
            dateFactura = splitDate.Last();
        }
        public string ArticoloTranslate(string str1,string articolo)
        {
            if (str1 == "empty")
            {
                return " ";
            }
            transContext db = new transContext();
            String[] wordo = str1.Split(new char[] { ' ', '.', ',', '`', '"', '\\', '/', '(', ')', '+', '*', '?' });
            double length = 0;
            double temp = 0;
            string symbols = "";
            string strTrans = "";
            //array for chars
            string[] arr;
            //list for symbols (example 12, 4220, 58)
            List<string> symbol;
            //class for work
            var war = (from s in db.Trans
                       where s.It == str1
                       select s);
            //if we have full match
            if (war != null && war.Count() == 1)
            {
                foreach (var item in war)
                {
                    strTrans = item.Ua;
                }
                return strTrans;
            }
            //else comparing
            else
            {
                str1 = DeleteSymbols(out arr, out symbol, str1);
                string[] str1Symbols = arr;
                //take matches by first input word
                if (articolo.Length == 3)
                {
                    war = from s in db.Trans
                          where s.Group == "GRUPPI" && s.It.Contains(wordo[0])
                          select s;
                }
                else
                {
                    war = from s in db.Trans
                          where s.Group == "MATERIALI" && s.It.Contains(wordo[0])
                          select s;
                }
                if (war.Count() == 0)
                {
                    return "Нема перекладу в БД";
                }
                //results of comparing input string and string from db
                List<Result> results = new List<Result>();
                foreach (var item in war)
                {
                    string str2 = DeleteSymbols(out arr, out symbol, item.It);
                    results.Add(new Result() { Id = Leven(str1, str2), Name_it = item.It });
                }
                //taking the best match
                var sortedArray = (from s in results
                                   orderby s.Id
                                   select s).FirstOrDefault();
                //please, take a look. Not the best solution!!!
                DeleteSymbols(out arr, out symbol, sortedArray.Name_it);
                string NameUa = "";
                if (articolo.Length == 3)
                {
                    NameUa = (from s in db.Trans
                                     where s.Group=="GRUPPI" && s.It.Contains(sortedArray.Name_it)
                                     select s.Ua).FirstOrDefault();
                }
                else
                {
                    NameUa = (from s in db.Trans
                                     where s.Group == "MATERIALI"&& s.It.Contains(sortedArray.Name_it)
                                     select s.Ua).FirstOrDefault();
                }

                int j = 0;
                foreach (var item in arr)
                {
                    if (item != null&&(j<str1Symbols.Length) && str1Symbols[j] != null)
                    {
                        strTrans = NameUa.Replace(item, str1Symbols[j]);
                        j++;
                    }
                }
                return strTrans;
            }
        }
        public async Task<string> TrasnslateAllAsync(string str1, string articolo)
        {
            string str= await Task.Run(() => ArticoloTranslate(str1, articolo));
            return str;
        }
        public string ColorTranslate(string str1, string code)
        {
            if (str1 == "empty")
            {
                return " ";
            }
            transContext db = new transContext();
            double length = 0;
            double temp = 0;
            string symbols = "";
            string strTrans = "";
            //array for chars
            List<string> symbol;
            string[] arr;
            var war = from s in db.Trans
                      where s.Group == "COLORI"
                      select s;
            if (code != "empty")
            {
                var war1 = (from s in war
                       where s.Code == code && s.It == str1
                       select s).FirstOrDefault() ;
                if (war1 != null)
                {
                    strTrans = war1.Ua;
                    return code+" "+strTrans;
                }
                else
                {
                    String[] wordo = str1.Split(new char[] { ' ', '.', ',', '`', '"', '\\', '/', '(', ')', '+', '*', '?' });
                    str1 = DeleteSymbols(out arr, out symbol, str1);
                    string[] str1Symbols = arr;
                    //take matches by first input word
                    war = from s in war
                          where s.It.Contains(wordo[0])
                          select s;
                    if (war.Count() == 0)
                    {
                        return "Нема перекладу в БД";
                    }
                    //results of comparing input string and string from db
                    List<Result> results = new List<Result>();
                    foreach (var item in war)
                    {
                        string str2 = DeleteSymbols(out arr, out symbol, item.It);
                        results.Add(new Result() { Id = Leven(str1, str2), Name_it = item.It });
                    }
                    //taking the best match
                    var sortedArray = (from s in results
                                       orderby s.Id
                                       select s).FirstOrDefault();
                    //please, take a look. Not the best solution!!!
                    DeleteSymbols(out arr, out symbol, sortedArray.Name_it);
                    string NameUa = (from s in war
                                     where s.Group=="COLORI" && s.It.Contains(sortedArray.Name_it)
                                     select s.Ua).FirstOrDefault();
                    return code+" "+NameUa;
                }
            }
            else
            {
                var war1 = (from s in war
                      where s.Group=="COLORI" && s.It == str1
                      select s).FirstOrDefault();
                if (war1 != null)
                {
                    strTrans = war1.Ua;
                    return strTrans;
                }
                else
                {
                    String[] wordo = str1.Split(new char[] { ' ', '.', ',', '`', '"', '\\', '/', '(', ')', '+', '*', '?' });
                    str1 = DeleteSymbols(out arr, out symbol, str1);
                    string[] str1Symbols = arr;
                    //take matches by first input word
                    war = from s in war
                          where s.It.Contains(wordo[0])
                          select s;
                    if (war.Count() == 0)
                    {
                        return "Нема перекладу в БД";
                    }
                    //results of comparing input string and string from db
                    List<Result> results = new List<Result>();
                    foreach (var item in war)
                    {
                        string str2 = DeleteSymbols(out arr, out symbol, item.It);
                        results.Add(new Result() { Id = Leven(str1, str2), Name_it = item.It });
                    }
                    //taking the best match
                    var sortedArray = (from s in results
                                       orderby s.Id
                                       select s).FirstOrDefault();
                    //please, take a look. Not the best solution!!!
                    DeleteSymbols(out arr, out symbol, sortedArray.Name_it);
                    string NameUa = (from s in war
                                     where s.Group=="COLORI" && s.It.Contains(sortedArray.Name_it)
                                     select s.Ua).FirstOrDefault();
                    return NameUa;
                }
            }


        }
        public string UmTranslate(string str1)
        {
            if (str1 == "empty")
            {
                return " ";
            }
            transContext db = new transContext();
            var trans = (from s in db.Trans
                         where s.Group == "MISURE" && s.Code == str1
                         select s.Ua).FirstOrDefault();
            return trans;
        }
        public string AspettoTranslate(string str1)
        {
            if (str1 == "empty")
            {
                return " ";
            }
            transContext db = new transContext();
            var trans = (from s in db.Trans
                         where s.Group == "CONFEZIONE" && s.Code == str1
                         select s.Ua).FirstOrDefault();
            return trans;
        }
        public void SetColumnWidth(double width,string range)
        {
            ws.Columns[1].ColumnWidth = width;
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }
        public void Close()
        {
            wb.Close();
            excel.Quit();
        }
        public string DeleteSymbols(out string[] symbolArr, out List<string> charArr, string str)
        {
            String[] words = str.Split(new char[] { ' ', '.', ',', '`', '"', '\\', '/', '(', ')', '+', '*', '?' });
            symbolArr = new string[words.Length];
            charArr = new List<string>();
            string strTest = "";
            int i = 0;

            foreach (var item in words)
            {
                int number;
                bool isNum = int.TryParse(item, out number);

                if (isNum)
                {
                    symbolArr[i] = item;
                    i++;
                }
                else
                {
                    strTest += item;
                    charArr.Add(item);
                }
            }
            return strTest;
        }
        public void ColWidthPage1(ISheet Sheet)
        {
            Sheet.SetColumnWidth(0, 10 * 256);
            Sheet.SetColumnWidth(1, 53 * 256);
            Sheet.SetColumnWidth(2, 6 * 256);
            Sheet.SetColumnWidth(3, 43 * 256);
            Sheet.SetColumnWidth(4, 10 * 256);
            Sheet.SetColumnWidth(5, 8 * 256);
            Sheet.SetColumnWidth(6, 10 * 256);
            Sheet.SetColumnWidth(7, 17 * 256);
            Sheet.SetColumnWidth(8, 8 * 256);
            Sheet.SetColumnWidth(9, 10 * 256);
        }
        public void ColWidthPage2(ISheet Sheet)
        {
            Sheet.SetColumnWidth(0, 5 * 256);
            Sheet.SetColumnWidth(1, 11 * 256);
            Sheet.SetColumnWidth(2, 9 * 256);
            Sheet.SetColumnWidth(3, 55 * 256);
            Sheet.SetColumnWidth(4, 9 * 256);
            Sheet.SetColumnWidth(5, 11 * 256);
            Sheet.SetColumnWidth(6, 11 * 256);
            Sheet.SetColumnWidth(7, 13 * 256);
            Sheet.SetColumnWidth(8, 15 * 256);
            Sheet.SetColumnWidth(9, 11 * 256);
            Sheet.SetColumnWidth(10, 7 * 256);
            Sheet.SetColumnWidth(11, 7 * 256);
            Sheet.SetColumnWidth(12, 11 * 256);
            Sheet.SetColumnWidth(13, 9 * 256);
        }

    }
}
