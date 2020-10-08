using System;
using System.Windows;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Translato.Models;
using System.Threading;
using _Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;

namespace Translato
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            transContext db = new transContext();
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "12.xls");
            List<string> articolo;
            List<string> articolo_it;
            List<string> um;
            List<string> colore;
            List<string> colore_name;
            List<string> quantita;
            List<string> prezzo;
            List<string> aspetto;
            List<string> colli;
            List<string> peso;
            List<string> zh;
            List<Page1> allData=new List<Page1>();
            List<Page1> translation = new List<Page1>();
            string numFactura;
            string dateFactura;
            string Pereviznik;
            string Nomera;
            string CMR;
            string Kordon;
            IWorkbook wb;
            Stopwatch swf = new Stopwatch();
            swf.Start();
            using(FileStream file= new FileStream(path, FileMode.Open, FileAccess.Read,FileShare.ReadWrite))
            {
                if (path.Contains("xlsx"))
                {
                    wb = new XSSFWorkbook(file);
                }
                else
                {
                    wb = new HSSFWorkbook(file);
                }
            }
            ISheet ws = wb.GetSheetAt(0);
            object[,] Data=new object[ws.LastRowNum+2,15];
            for (int i = 0; i<=ws.LastRowNum; i++)
            {
                var currentRow = ws.GetRow(i);
                if (currentRow != null)
                {
                    for (int j = 0; j < 14; j++)
                    {
                        if (currentRow.GetCell(j) != null&& currentRow.GetCell(j).ToString()!="")
                        {
                            Data[i+1, j+1] = ws.GetRow(i).GetCell(j).ToString();
                        }
                        else
                        {
                            Data[i+1, j+1] = null;
                        }
                    }
                }
            }

            int lengthRow=ws.LastRowNum;
            int lengthCol=13;
            ExcelReader excelReader = new ExcelReader();
            excelReader.ExcelToList(out articolo, out articolo_it, out um, out colore,
                      out colore_name, out quantita, out prezzo, out aspetto,
                       out colli, out peso, out zh, lengthCol, lengthRow, Data, out numFactura, out dateFactura,out Pereviznik,out Nomera,out CMR,out Kordon);
            swf.Stop();
            label1.Text = swf.ElapsedMilliseconds.ToString();
            for (int i = 0; i < articolo.Count; i++)
            {
                allData.Add(new Page1
                {
                    articolo = articolo[i],
                    articolo_it = articolo_it[i],
                    aspetto = aspetto[i],
                    colli = colli[i],
                    colore = colore[i],
                    colore_name = colore_name[i],
                    peso = peso[i],
                    prezzo = prezzo[i],
                    quantita = quantita[i],
                    um = um[i]
                });
                translation.Add(new Page1
                {
                    //articolo = articolo[i],
                    articolo_it = excelReader.ArticoloTranslate(allData[i].articolo_it, allData[i].articolo),
                    aspetto = excelReader.AspettoTranslate(aspetto[i]),
                   // colli = colli[i],
                   // colore = colore[i],
                    colore_name = excelReader.ColorTranslate(allData[i].colore_name, allData[i].colore),
                   // peso = peso[i],
                   // prezzo = prezzo[i],
                   // quantita = quantita[i],
                    um = excelReader.UmTranslate(um[i])
                });
            }
            Kordon = (from s in db.Trans
                      where s.Group == "FRONTIERA" && s.It == Kordon
                      select s.Ua).FirstOrDefault();
            Pereviznik = (from s in db.Trans
                          where s.Group == "VETTORE" && s.It == Pereviznik
                          select s.Ua).FirstOrDefault();
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFCellStyle allCells = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFCellStyle justStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFCellStyle borderBottom = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFCellStyle borderTop= (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFCellStyle header = (HSSFCellStyle)workbook.CreateCellStyle();

            HSSFFont myFont = (HSSFFont)workbook.CreateFont();
            myFont.FontHeightInPoints = 10;
            myFont.FontName = "Arial Narrow";

            allCells.SetFont(myFont);
            allCells.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            borderBottom.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            borderBottom.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;

            justStyle.SetFont(myFont);

            borderTop.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            borderTop.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

            header.SetFont(myFont);
            header.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            header.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            header.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            header.FillForegroundColor = IndexedColors.Yellow.Index;
            header.FillPattern = FillPattern.SolidForeground;

            ISheet Sheet = workbook.CreateSheet("Report");
            excelReader.ColWidthPage1(Sheet);
            for (int i = 11; i< articolo.Count+11; i++)
            {
                IRow FirstRow = Sheet.CreateRow(i);
                if (i == 11)
                {
                    IRow row = Sheet.CreateRow(10);
                    IRow row1 = Sheet.CreateRow(9);
                    excelReader.WriteToCell(row, 0, "", borderTop);
                    excelReader.WriteToCell(row, 1, "", borderTop);
                    excelReader.WriteToCell(row, 2, "", borderTop);
                    excelReader.WriteToCell(row, 3, "", borderTop);
                    excelReader.WriteToCell(row, 4, "", borderTop);
                    excelReader.WriteToCell(row, 5, "", borderTop);
                    excelReader.WriteToCell(row, 6, "", borderTop);
                    excelReader.WriteToCell(row, 7, "", borderTop);
                    excelReader.WriteToCell(row, 8, "", borderTop);
                    excelReader.WriteToCell(row, 9, "", borderTop );

                    excelReader.WriteToCell(row1, 0, "АРТИКУЛ", header);
                    excelReader.WriteToCell(row1, 1, "ОПИС АРТИКУЛУ", header);
                    excelReader.WriteToCell(row1, 2, "О.В.", header);
                    excelReader.WriteToCell(row1, 3, "КОЛІР", header);
                    excelReader.WriteToCell(row1, 4, "КІЛЬКІСТЬ", header);
                    excelReader.WriteToCell(row1, 5, "ІНША О.В.", header);
                    excelReader.WriteToCell(row1, 6, "ЦІНА", header);
                    excelReader.WriteToCell(row1, 7, "ЗОВНІШНІЙ ВИГЛЯД", header);
                    excelReader.WriteToCell(row1, 8, "МІСЦЬ", header);
                    excelReader.WriteToCell(row1, 9, "ВАГА НЕТТО", header);
                    row1 = Sheet.CreateRow(0);
                    excelReader.WriteToCell(row1, 0, "СП ТОВ \"АРКОБАЛЕНО\"", justStyle);
                    row1 = Sheet.CreateRow(1);
                    excelReader.WriteToCell(row1, 0, "66-016 м. Червеньськ, вул.Квятова, 5; тел.: (0-68)3219100, 3219101-106", justStyle);
                    row1 = Sheet.CreateRow(2);
                    excelReader.WriteToCell(row1, 0, "Банківські реквізити:", justStyle);
                    row1 = Sheet.CreateRow(3);
                    excelReader.WriteToCell(row1, 0, "ПДВ ЄС: PL 9290100021 (ІД. НОМЕР)", justStyle);
                    row1 = Sheet.CreateRow(4);
                    excelReader.WriteToCell(row1, 0, "№ ПДВ: 929-010-00-21", justStyle);
                    row1 = Sheet.CreateRow(6);
                    excelReader.WriteToCell(row1, 0, $"Пакувальний лист для фактури: {numFactura} від {dateFactura}", justStyle);
                    row1 = Sheet.CreateRow(7);
                    excelReader.WriteToCell(row1, 0, $"ЕКСПОРТ ВІД : {dateFactura}", justStyle);
                    row1 = Sheet.CreateRow(articolo.Count+16);
                    excelReader.WriteToCell(row1, 0, $"ПЕРЕВІЗНИК : {Pereviznik}", justStyle);
                    row1 = Sheet.CreateRow(articolo.Count + 17);
                    excelReader.WriteToCell(row1, 0, $"НОМЕР АВТОМОБІЛЮ : {Nomera}", justStyle);
                    row1 = Sheet.CreateRow(articolo.Count + 18);
                    excelReader.WriteToCell(row1, 0, $"ЦМР НОМЕР : {CMR}", justStyle);
                    excelReader.WriteToCell(row1, 4, $"ПЕЧАТКА СП ТОВ 'АРКОБАЛЕНО'", justStyle);
                    row1 = Sheet.CreateRow(articolo.Count + 19);
                    excelReader.WriteToCell(row1, 0, $"МІСЦЕ ПЕРЕТИНУ КОРДОНУ: {Kordon} ", justStyle);
                    excelReader.WriteToCell(row1, 4, $"Підпис уповноваженої особи", justStyle);
                }
                if (articolo[i - 11] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 0, allData[i-11].articolo,allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 0, "", allCells);
                }
                excelReader.WriteToCell(FirstRow, 1, translation[i-11].articolo_it, allCells);
                excelReader.WriteToCell(FirstRow, 2, translation[i - 11].um, allCells);
                excelReader.WriteToCell(FirstRow, 3, translation[i-11].colore_name, allCells);
                if (quantita[i - 11] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 4, allData[i-11].quantita, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 4, "", allCells);
                }
                excelReader.WriteToCell(FirstRow, 5, "", allCells);
                if (prezzo[i - 11] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 6, allData[i-11].prezzo, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 6, "", allCells);
                }
                excelReader.WriteToCell(FirstRow, 7, translation[i-11].aspetto, allCells);
                if (colli[i - 11] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 8, allData[i-11].colli, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 8, "", allCells);
                }
                if (peso[i - 11] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 9, allData[i - 11].peso, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 9 , "", allCells);
                }
                if (i == articolo.Count + 10)
                {
                    FirstRow = Sheet.CreateRow(i + 1);
                    excelReader.WriteToCell(FirstRow, 0, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 1, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 2, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 3, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 4, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 5, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 6, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 7, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 8, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 9, "", borderBottom);
                }
            }
            ISheet Facture = workbook.CreateSheet("Factura");
            excelReader.ColWidthPage2(Facture);
            for (int i = 15; i < articolo.Count + 15; i++)
            {
                IRow FirstRow = Facture.CreateRow(i);
                if (i == 15)
                {
                    IRow row = Facture.CreateRow(14);
                    IRow row1 = Facture.CreateRow(13);
                    excelReader.WriteToCell(row, 0, "", borderTop);
                    excelReader.WriteToCell(row, 1, "", borderTop);
                    excelReader.WriteToCell(row, 2, "", borderTop);
                    excelReader.WriteToCell(row, 3, "", borderTop);
                    excelReader.WriteToCell(row, 4, "", borderTop);
                    excelReader.WriteToCell(row, 5, "", borderTop);
                    excelReader.WriteToCell(row, 6, "", borderTop);
                    excelReader.WriteToCell(row, 7, "", borderTop);
                    excelReader.WriteToCell(row, 8, "", borderTop);
                    excelReader.WriteToCell(row, 9, "", borderTop);
                    excelReader.WriteToCell(row, 10, "", borderTop);
                    excelReader.WriteToCell(row, 11, "", borderTop);
                    excelReader.WriteToCell(row, 12, "", borderTop);
                    excelReader.WriteToCell(row, 15, "", borderTop);

                    excelReader.WriteToCell(row1, 0, "№ п/п", header);
                    excelReader.WriteToCell(row1, 1, "Показник матеріалу", header);
                    excelReader.WriteToCell(row1, 2, "Симв.", header);
                    excelReader.WriteToCell(row1, 3, "Назва матеріалу", header);
                    excelReader.WriteToCell(row1, 4, "Од.вим.", header);
                    excelReader.WriteToCell(row1, 5, "Кількість", header);
                    excelReader.WriteToCell(row1, 6, "Ціна, євро", header);
                    excelReader.WriteToCell(row1, 7, "Націнка+знижка-", header);
                    excelReader.WriteToCell(row1, 8, "Ціна реальна, євро", header);
                    excelReader.WriteToCell(row1, 9, "Вартість без ПДВ", header);
                    excelReader.WriteToCell(row1, 10, "Став. ПДВ", header);
                    excelReader.WriteToCell(row1, 11, "Квота ПДВ", header);
                    excelReader.WriteToCell(row1, 12, "вартість з ПДВ", header);
                    excelReader.WriteToCell(row1, 13, "Вартість з ПДВ", header);


                    row1 = Facture.CreateRow(0);
                    excelReader.WriteToCell(row1, 0, "СП ТОВ \"АРКОБАЛЕНО\"", justStyle);
                    row1 = Facture.CreateRow(1);
                    excelReader.WriteToCell(row1, 0, "66-016 м. Червеньськ, вул.Квятова, 5; тел.: (0-68)3219100, 3219101-106", justStyle);
                    row1 = Facture.CreateRow(2);
                    excelReader.WriteToCell(row1, 0, "Банківські реквізити:", justStyle);
                    row1 = Facture.CreateRow(3);
                    excelReader.WriteToCell(row1, 0, "ПДВ ЄС: PL 9290100021 (ІД. НОМЕР)", justStyle);
                    row1 = Facture.CreateRow(4);
                    excelReader.WriteToCell(row1, 0, "№ ПДВ: 929-010-00-21", justStyle);
                    row1 = Facture.CreateRow(6);
                    excelReader.WriteToCell(row1, 0, $"Пакувальний лист для фактури: {numFactura} від {dateFactura}", justStyle);
                    row1 = Facture.CreateRow(7);
                    excelReader.WriteToCell(row1, 0, $"ЕКСПОРТ ВІД : {dateFactura}", justStyle);
                    row1 = Facture.CreateRow(articolo.Count + 16);
                    excelReader.WriteToCell(row1, 0, $"ПЕРЕВІЗНИК : {Pereviznik}", justStyle);
                    row1 = Facture.CreateRow(articolo.Count + 17);
                    excelReader.WriteToCell(row1, 0, $"НОМЕР АВТОМОБІЛЮ : {Nomera}", justStyle);
                    row1 = Facture.CreateRow(articolo.Count + 18);
                    excelReader.WriteToCell(row1, 0, $"ЦМР НОМЕР : {CMR}", justStyle);
                    excelReader.WriteToCell(row1, 4, $"ПЕЧАТКА СП ТОВ 'АРКОБАЛЕНО'", justStyle);
                    row1 = Facture.CreateRow(articolo.Count + 19);
                    excelReader.WriteToCell(row1, 0, $"МІСЦЕ ПЕРЕТИНУ КОРДОНУ: {Kordon} ", justStyle);
                    excelReader.WriteToCell(row1, 4, $"Підпис уповноваженої особи", justStyle);
                }
                if (articolo[i - 15] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 0, allData[i - 15].articolo, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 0, "", allCells);
                }
                excelReader.WriteToCell(FirstRow, 1, translation[i-15].articolo_it, allCells);
                excelReader.WriteToCell(FirstRow, 2, translation[i - 15].um, allCells);
                excelReader.WriteToCell(FirstRow, 3, translation[i - 15].colore_name, allCells);
                if (quantita[i - 15] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 4, allData[i - 15].quantita, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 4, "", allCells);
                }
                excelReader.WriteToCell(FirstRow, 5, "", allCells);
                if (prezzo[i - 15] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 6, allData[i - 15].prezzo, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 6, "", allCells);
                }
                excelReader.WriteToCell(FirstRow, 7, translation[i-15].aspetto, allCells);
                if (colli[i - 15] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 8, allData[i - 15].colli, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 8, "", allCells);
                }
                if (peso[i - 15] != "empty")
                {
                    excelReader.WriteToCell(FirstRow, 9, allData[i - 15].peso, allCells);
                }
                else
                {
                    excelReader.WriteToCell(FirstRow, 9, "", allCells);
                }
                if (i == articolo.Count + 10)
                {
                    FirstRow = Facture.CreateRow(i + 1);
                    excelReader.WriteToCell(FirstRow, 0, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 1, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 2, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 3, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 4, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 5, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 6, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 7, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 8, "", borderBottom);
                    excelReader.WriteToCell(FirstRow, 9, "", borderBottom);
                }
            }



            using (var fileData = new FileStream(@"C:\Users\1\Desktop\123\new.xlsReportName.xls", FileMode.Create))
            {
                workbook.Write(fileData);
            }
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            label1.Text = ts.ToString();
        }
    }
}

