using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace DomPecheyKP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Dictionary<string, double> list;
        Dictionary<string, double> listIC;
        Dictionary<string, double> listIW;
        Dictionary<string, double> listRD;

        double sumCE = 0;
        double sumIC = 0;
        double sumIW = 0;
        double sumRD = 0;
        double sumND = 0;

        string nameOfSheetIC = "Изоляц и расход материалы";
        string nameOfSheetIW = "Монтажные работы";
        string nameOfSheetRD = "Такелажные работы";
        
        Excel.Application ObjWorkExcel;
        Excel.Worksheet ObjWorkSheet;
        Excel.Workbook ObjWorkBook;

        private PdfPCell getMidHeader(string str, iTextSharp.text.Font font)
        {
            PdfPCell cell = new PdfPCell(new Phrase(new Phrase(str, font)));
            cell.Colspan = 5;
            cell.BorderWidthLeft = 0;
            cell.BorderWidthRight = 0;
            cell.PaddingBottom = 10;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            return cell;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            int countColumns = 5;
            string nameOfNewFile = @"newFile.pdf";
            //List<float> heightsRows = new List<float>();
            PdfPTable fTable = new PdfPTable(countColumns);
            int[] firstTablecellwidth = { 10, 51, 17, 11, 11 };
            using (var reader = new PdfReader(@"oldFile.pdf"))
            {
                //Запись в файл для подсчета высоты строк
                using (var fileStream = new FileStream(nameOfNewFile, FileMode.Create, FileAccess.Write))
                {
                    var document = new Document(reader.GetPageSizeWithRotation(1));
                    var writer = PdfWriter.GetInstance(document, fileStream);

                    document.Open();
                    document.NewPage();
                    var importedPage = writer.GetImportedPage(reader, 1);
                    var contentByte = writer.DirectContent;
                    contentByte.AddTemplate(importedPage, 0, 0);

                    System.Text.EncodingProvider ppp = System.Text.CodePagesEncodingProvider.Instance;
                    Encoding.RegisterProvider(ppp);

                    var fontHeaderName = "Sitka Text Bold";
                    if (!FontFactory.IsRegistered(fontHeaderName))
                    {
                        var fontHeaderPath = Environment.GetEnvironmentVariable("SitkaText-Bold.ttf");
                        FontFactory.Register("SitkaText-Bold.ttf");
                    }
                    iTextSharp.text.Font fontHeader = FontFactory.GetFont(fontHeaderName, BaseFont.IDENTITY_H);
                    fontHeader.Size = 23;
                    var fontName = "Sitka Banner";
                    if (!FontFactory.IsRegistered(fontName))
                    {
                        var fontPath = Environment.GetEnvironmentVariable("Sitka-Banner.ttf");
                        FontFactory.Register("Sitka-Banner.ttf");
                    }
                    iTextSharp.text.Font font = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                    font.Color = iTextSharp.text.BaseColor.BLACK;
                    font.Size = 25;

                    fTable.WidthPercentage = 80;
                    fTable.SetWidths(firstTablecellwidth);
                    //Добавим в таблицу общий заголовок
                    PdfPCell cell;
                    cell = new PdfPCell(new Phrase(new Phrase("Номер", fontHeader)));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Наименование", fontHeader)));
                    cell.BorderWidthTop = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Количество", fontHeader)));
                    cell.BorderWidthTop = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Цена за шт.", fontHeader)));
                    cell.BorderWidthTop = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Сумма", fontHeader)));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthRight = 0;
                    fTable.AddCell(cell);
                    /////////////////////////////////////////////////////////

                    var a = from RadioButton r in ProductType.Controls where r.Checked == true select r.Text;

                    fTable.AddCell(getMidHeader("1. " + a.First(), font));

                    ///////////////////////////////////////////////////////////////////////////
                    ///
                    cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[0].Value.ToString(), font)));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthLeft = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[1].Value.ToString(), font)));
                    cell.BorderWidthTop = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[2].Value.ToString(), font)));
                    cell.BorderWidthTop = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[3].Value.ToString(), font)));
                    cell.BorderWidthTop = 0;
                    fTable.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[4].Value.ToString(), font)));
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthRight = 0;
                    fTable.AddCell(cell);

                    a = from RadioButton r in Manufacturer.Controls where r.Checked == true select r.Text;
                    string str = "2. " + a.First();
                    if (OwnD.Checked)
                    {
                        str += " D" + OwnValue.Value.ToString() + " мм";
                    }
                    else
                    {
                        a = from RadioButton r in Diameter.Controls.OfType<RadioButton>() where (r.Checked) == true select r.Text;

                        str += " D" + a.First() + " мм";
                    }
                    a = from RadioButton r in MetalThickness.Controls where r.Checked == true select r.Text;
                    str += " (" + a.First() + " мм)";

                    fTable.AddCell(getMidHeader(str, font));

                    for (int j = 0; j < ChimneyElements.RowCount - 1; j++)
                    {
                        for (int k = 0; k < countColumns; k++)
                        {

                            cell = new PdfPCell(new Phrase(new Phrase(ChimneyElements.Rows[j].Cells[k].Value.ToString(), font)));
                            if (k == 0)
                                cell.BorderWidthLeft = 0;
                            else if (k == countColumns - 1)
                                cell.BorderWidthRight = 0;
                            cell.PaddingBottom = 10;
                            fTable.AddCell(cell);
                        }
                    }


                    fTable.AddCell(getMidHeader("3. Изоляционные и расходные материалы", font));



                    for (int j = 0; j < InsulationСonsumables.RowCount - 1; j++)
                    {
                        for (int k = 0; k < countColumns; k++)
                        {

                            cell = new PdfPCell(new Phrase(new Phrase(InsulationСonsumables.Rows[j].Cells[k].Value.ToString(), font)));
                            if (k == 0)
                                cell.BorderWidthLeft = 0;
                            else if (k == countColumns - 1)
                                cell.BorderWidthRight = 0;
                            cell.PaddingBottom = 10;
                            fTable.AddCell(cell);
                        }
                    }



                    fTable.AddCell(getMidHeader("4. Монтажные работы,выезд на замер", font));

                    for (int j = 0; j < InstallationWork.RowCount - 1; j++)
                    {
                        for (int k = 0; k < countColumns; k++)
                        {

                            cell = new PdfPCell(new Phrase(new Phrase(InstallationWork.Rows[j].Cells[k].Value.ToString(), font)));
                            if (k == 0)
                                cell.BorderWidthLeft = 0;
                            else if (k == countColumns - 1)
                                cell.BorderWidthRight = 0;
                            cell.PaddingBottom = 10;
                            fTable.AddCell(cell);
                        }
                    }


                    fTable.AddCell(getMidHeader("5. Такелажные работы и доставка", font));

                    for (int j = 0; j < RiggingDelivery.RowCount - 1; j++)
                    {
                        for (int k = 0; k < countColumns; k++)
                        {

                            cell = new PdfPCell(new Phrase(new Phrase(RiggingDelivery.Rows[j].Cells[k].Value.ToString()+ " ₽", font)));
                            if (k == 0)
                                cell.BorderWidthLeft = 0;
                            else if (k == countColumns - 1)
                                cell.BorderWidthRight = 0;
                            cell.PaddingBottom = 10;
                            fTable.AddCell(cell);
                        }
                    }


                    document.Add(fTable);

                    document.Close();
                    writer.Close();
                }


                using (var fileStream = new FileStream(nameOfNewFile, FileMode.Create, FileAccess.Write))
                {
                    var document = new Document(reader.GetPageSizeWithRotation(1));
                    var writer = PdfWriter.GetInstance(document, fileStream);

                    document.Open();

                    for (var i = 1; i <= reader.NumberOfPages; i++)
                    {
                        document.NewPage();
                        var importedPage = writer.GetImportedPage(reader, i);
                        var contentByte = writer.DirectContent;
                        contentByte.AddTemplate(importedPage, 0, 0);


                        if (i == 2)
                        {

                            System.Text.EncodingProvider ppp = System.Text.CodePagesEncodingProvider.Instance;
                            Encoding.RegisterProvider(ppp);
                            var fontName = "Sitka Text Italic";
                            if (!FontFactory.IsRegistered(fontName))
                            {
                                var fontPath = Environment.GetEnvironmentVariable("SitkaText.ttf");
                                FontFactory.Register("SitkaText.ttf");
                            }
                            iTextSharp.text.Font font1 = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                            font1.Color = iTextSharp.text.BaseColor.WHITE;


                            contentByte.BeginText();
                            contentByte.SetColorFill(BaseColor.WHITE);
                            contentByte.SetFontAndSize(font1.BaseFont, 70);
                            var multiLineString = "Александр,\nДобрый день!".Split('\n');
                            int y = 550;
                            foreach (var line in multiLineString)
                            {
                                contentByte.ShowTextAligned(PdfContentByte.ALIGN_CENTER, line, 650, y, 0);
                                y -= 80;
                            }
                            contentByte.EndText();
                        }

                        if (i == 11)
                        {

                            System.Text.EncodingProvider ppp = System.Text.CodePagesEncodingProvider.Instance;
                            Encoding.RegisterProvider(ppp);
                            var fontName = "Sitka Text Bold";
                            if (!FontFactory.IsRegistered(fontName))
                            {
                                var fontPath = Environment.GetEnvironmentVariable("SitkaText-Bold.ttf");
                                FontFactory.Register("SitkaText-Bold.ttf");
                            }
                            iTextSharp.text.Font font1 = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                            font1.Color = iTextSharp.text.BaseColor.WHITE;


                            contentByte.BeginText();
                            contentByte.SetColorFill(BaseColor.WHITE);
                            contentByte.SetFontAndSize(font1.BaseFont, 36);
                            contentByte.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Николай Макаркин", 640, 93, 0);
                            contentByte.EndText();
                        }

                        if (i == 8)//8
                        {
                            int maxHeightTable = 330;
                            Paragraph p;

                            PdfPTable sTable = new PdfPTable(countColumns);
                            sTable.WidthPercentage = 80;
                            sTable.SetWidths(firstTablecellwidth);

                            float sumHeight = 0;
                            for (int j = 0; j < fTable.Rows.Count; j++)
                            {
                                if ((sumHeight + fTable.Rows[j].MaxHeights) > maxHeightTable)
                                {
                                    p = new Paragraph("     ");
                                    p.SpacingAfter = 100;
                                    document.Add(p);
                                    sTable.SpacingBefore = 100;
                                    document.Add(sTable);
                                    document.NewPage();
                                    contentByte = writer.DirectContent;
                                    contentByte.AddTemplate(importedPage, 0, 0);
                                    sumHeight = 0;
                                    sTable = new PdfPTable(countColumns);
                                    sTable.SetWidths(firstTablecellwidth);
                                    sTable.Rows.Add(fTable.GetRow(0));
                                    sumHeight += fTable.Rows[0].MaxHeights;
                                }
                                sTable.Rows.Add(fTable.GetRow(j));
                                sumHeight += fTable.Rows[j].MaxHeights;
                            }


                            p = new Paragraph("     ");
                            p.SpacingAfter = 100;
                            document.Add(p);
                            sTable.SpacingBefore = 100;
                            document.Add(sTable);
                        }

                    }
                    document.Close();
                    writer.Close();
                }

            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            ObjWorkExcel = new Excel.Application();
            ObjWorkBook = ObjWorkExcel.Workbooks.Open(Environment.CurrentDirectory + @"\ДанныеКП.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            loadElD(2);
            NameOfKiln.Rows.Add("1", "", "1", "0", "0");

            loadIС();
            loadIW(1);
            loadRD();
            //удалить
            foreach (Object checkedItem in NewChimneyElements.Items)
            {
                ChimneyElements.Rows.Add("1", checkedItem.ToString(), "1", list[checkedItem.ToString()], list[checkedItem.ToString()]);
            }

            foreach (Object checkedItem in NewInstallationWork.Items)
            {
                InstallationWork.Rows.Add("1", checkedItem.ToString(), "1", listIW[checkedItem.ToString()], listIW[checkedItem.ToString()]);
            }

            foreach (Object checkedItem in NewInsulationСonsumables.Items)
            {
                InsulationСonsumables.Rows.Add("1", checkedItem.ToString(), "1", listIC[checkedItem.ToString()], listIC[checkedItem.ToString()]);
            }
            foreach (Object checkedItem in NewRiggingDelivery.Items)
            {
                RiggingDelivery.Rows.Add("1", checkedItem.ToString(), "1", listRD[checkedItem.ToString()], listRD[checkedItem.ToString()]);
            }

        }

        private void loadIС()
        {

            listIC = new Dictionary<string, double>();
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[nameOfSheetIC]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку          
            for (int i = 2; i <= (int)lastCell.Row; i++) // по всем строкам
            {
                listIC.Add(ObjWorkSheet.Cells[i, 1].Text.ToString(), Convert.ToDouble(ObjWorkSheet.Cells[i, 2].Text.ToString()));//считываем текст в строку
                NewInsulationСonsumables.Items.Add(ObjWorkSheet.Cells[i, 1].Text.ToString());
            }
            
            NewInsulationСonsumables.Sorted = true;

        }

        private void loadElD(int nColumn)
        {

            NewChimneyElements.Items.Clear();
            ChimneyElements.Rows.Clear();

            list = new Dictionary<string, double>();

            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку          
            for (int i = 3; i <= (int)lastCell.Row; i++) // по всем строкам
            {
                try
                {
                    list.Add(ObjWorkSheet.Cells[i, 1].Text.ToString(), Convert.ToDouble(ObjWorkSheet.Cells[i, nColumn].Text.ToString()));//считываем текст в строку

                }
                catch (Exception ex)
                {
                    list.Add(ObjWorkSheet.Cells[i, 1].Text.ToString(), Convert.ToDouble("0"));//считываем текст в строку

                }
                NewChimneyElements.Items.Add(ObjWorkSheet.Cells[i, 1].Text.ToString());
            }

        }

        private void calculateResults()
        {
            double sumC = Convert.ToDouble(NameOfKiln.Rows[0].Cells[4].Value.ToString());
            double sum1 = sumCE + sumIC + sumC;
            SumChimneyManufacturerAndInsulation.Text = (sum1).ToString() + " Руб.";
            double sum2 = sumRD + sumIW;
            SumRiggingAndInstall.Text = (sum2).ToString() + " Руб.";
            sumND = sum1 + sum2;
            SumNotDiscount.Text = (sumND).ToString() + " Руб.";
            AllSum.Text = (sumND - Convert.ToDouble(numericUpDown1.Value)).ToString() + " Руб.";

        }

        private void calculateChimneyElementsSum()
        {
            sumCE = 0;
            foreach (DataGridViewRow row in ChimneyElements.Rows)
                sumCE += Convert.ToDouble(row.Cells[4].Value);
            SumChimneyElements.Text = sumCE.ToString() + " Руб.";
            calculateResults();
        }

        

        private void loadIW(int nColumn)
        {

            listIW = new Dictionary<string, double>();
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[nameOfSheetIW]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку          
            for (int i = 2; i <= (int)lastCell.Row; i++) // по всем строкам
            {
                listIW.Add(ObjWorkSheet.Cells[i, nColumn].Text.ToString(), Convert.ToDouble(ObjWorkSheet.Cells[i, nColumn+1].Text.ToString()));//считываем текст в строку
                NewInstallationWork.Items.Add(ObjWorkSheet.Cells[i, 1].Text.ToString());
            }
            NewInstallationWork.Sorted = true;
        }

        private void loadRD()
        {

            listRD = new Dictionary<string, double>();
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[nameOfSheetRD]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку          
            for (int i = 2; i <= (int)lastCell.Row; i++) // по всем строкам
            {
                listRD.Add(ObjWorkSheet.Cells[i, 1].Text.ToString(), Convert.ToDouble(ObjWorkSheet.Cells[i, 2].Text.ToString()));//считываем текст в строку
                NewRiggingDelivery.Items.Add(ObjWorkSheet.Cells[i, 1].Text.ToString());
            }
            NewRiggingDelivery.Sorted = true;
        }


        private void addChimneyElements_Click(object sender, EventArgs e)
        {
            int n;
            if (ChimneyElements.RowCount == 0)
                n = 1;
            else
                n = ChimneyElements.RowCount + 1;
            foreach (Object checkedItem in NewChimneyElements.CheckedItems)
            {
                ChimneyElements.Rows.Add(n++, checkedItem.ToString(), "1", list[checkedItem.ToString()], list[checkedItem.ToString()]);
            }
            while (NewChimneyElements.CheckedItems.Count != 0)
            {
                NewChimneyElements.Items.Remove(NewChimneyElements.CheckedItems[0]);
            }
            calculateChimneyElementsSum();
        }

        private void deleteChimneyElements_Click(object sender, EventArgs e)
        {

            while (ChimneyElements.SelectedRows.Count != 0)
            {
                if (list.ContainsKey(ChimneyElements.SelectedRows[0].Cells[1].Value.ToString()))
                    NewChimneyElements.Items.Add(ChimneyElements.SelectedRows[0].Cells[1].Value);
                ChimneyElements.Rows.Remove(ChimneyElements.SelectedRows[0]);
            }

            for (int i = 0; i < ChimneyElements.RowCount - 1; i++)
            {
                ChimneyElements.Rows[i].Cells[0].Value = i + 1;
            }
            calculateChimneyElementsSum();
            checkAllChimneyElement();

        }

        private void checkAllChimneyElement()
        {
            NewChimneyElements.Items.Clear();
            foreach (string str in list.Keys)
            {
                bool f = false;
                for (int i = 0; i < ChimneyElements.RowCount - 1; i++)
                {
                    if (str == ChimneyElements.Rows[i].Cells[1].Value.ToString())
                        f = true;
                }
                if (!f)
                    NewChimneyElements.Items.Add(str);
            }
        }

        private void ChimneyElements_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

            for (int i = 0; i < ChimneyElements.RowCount - 1; i++)
            {
                ChimneyElements.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void ChimneyElements_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int nRow = e.RowIndex;
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {

                ChimneyElements.Rows[nRow].Cells[4].Value = Convert.ToDouble(ChimneyElements.Rows[nRow].Cells[2].Value) * Convert.ToDouble(ChimneyElements.Rows[nRow].Cells[3].Value);
                calculateChimneyElementsSum();
            }
            checkAllChimneyElement();
        }

        private void d_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(Convert.ToInt32(radioButton.Tag));
        }

        private void d115_CheckedChanged(object sender, EventArgs e)
        {

            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(2);


        }

        private void d120_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(3);
        }

        private void d130_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(4);
        }

        private void d150_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(5);
        }

        private void d180_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(6);
        }

        private void d200_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(7);
        }

        private void d250_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(8);
        }

        private void d300_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(9);
        }

        private void OwnD_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
                loadElD(0);
        }


        private DataGridView currentDataGridView;

        private void EditingControl_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar))
            {
                Control editingControl = (Control)sender;

                if (currentDataGridView.CurrentCell.ColumnIndex == 2)
                {
                    if (!Regex.IsMatch(editingControl.Text + e.KeyChar, "^[0-9]{0,4}$"))
                        e.Handled = true;
                }
                else if (currentDataGridView.CurrentCell.ColumnIndex == 3)
                {
                    if (!Regex.IsMatch(editingControl.Text + e.KeyChar, @"^[0-9,]{0,10}$"))
                        e.Handled = true;
                }

            }
        }

        private void NameOfKiln_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            currentDataGridView = NameOfKiln;
            NameOfKiln.EditingControl.KeyPress -= EditingControl_KeyPress;
            NameOfKiln.EditingControl.KeyPress += EditingControl_KeyPress;
        }

        private void NameOfKiln_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int nRow = e.RowIndex;
            var isValid = Regex.IsMatch(NameOfKiln.Rows[nRow].Cells[e.ColumnIndex].Value.ToString(), @"^[0-9]*[,]?[0-9]+$");
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {
                if (!isValid)
                {
                    MessageBox.Show("Ошибка при вводе числа. Введите значение заново.");
                    NameOfKiln.Rows[nRow].Cells[e.ColumnIndex].Value = "0";
                }
                NameOfKiln.Rows[nRow].Cells[4].Value = Convert.ToDouble(NameOfKiln.Rows[nRow].Cells[2].Value) * Convert.ToDouble(NameOfKiln.Rows[nRow].Cells[3].Value);
            }

        }

        private void ChimneyElements_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            currentDataGridView = ChimneyElements;
            ChimneyElements.EditingControl.KeyPress -= EditingControl_KeyPress;
            ChimneyElements.EditingControl.KeyPress += EditingControl_KeyPress;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
        }



        //////////////////////////////////////////////////////////////

        private void deleteInsulationСonsumables_Click(object sender, EventArgs e)
        {
            while (InsulationСonsumables.SelectedRows.Count != 0)
            {
                if (listIC.ContainsKey(InsulationСonsumables.SelectedRows[0].Cells[1].Value.ToString()))
                    NewInsulationСonsumables.Items.Add(InsulationСonsumables.SelectedRows[0].Cells[1].Value);
                InsulationСonsumables.Rows.Remove(InsulationСonsumables.SelectedRows[0]);
            }

            for (int i = 0; i < InsulationСonsumables.RowCount - 1; i++)
            {
                InsulationСonsumables.Rows[i].Cells[0].Value = i + 1;
            }
            calculateInsulationСonsumablesSum();

            checkAllInsulationСonsumables();
        }
        private void checkAllInsulationСonsumables()
        {
            NewInsulationСonsumables.Items.Clear();
            foreach (string str in listIC.Keys)
            {
                bool f = false;
                for (int i = 0; i < InsulationСonsumables.RowCount - 1; i++)
                {
                    if (str == InsulationСonsumables.Rows[i].Cells[1].Value.ToString())
                        f = true;
                }
                if (!f)
                    NewInsulationСonsumables.Items.Add(str);
            }
        }

        private void calculateInsulationСonsumablesSum()
        {
            sumIC = 0;
            foreach (DataGridViewRow row in InsulationСonsumables.Rows)
                sumIC += Convert.ToDouble(row.Cells[4].Value);
            SumInsulationСonsumables.Text = sumIC.ToString() + " Руб.";
            calculateResults();
        }
        

        private void addInsulationСonsumables_Click(object sender, EventArgs e)
        {
            int n;
            if (InsulationСonsumables.RowCount == 0)
                n = 1;
            else
                n = InsulationСonsumables.RowCount + 1;
            foreach (Object checkedItem in NewInsulationСonsumables.CheckedItems)
            {
                InsulationСonsumables.Rows.Add(n++, checkedItem.ToString(), "1", listIC[checkedItem.ToString()], listIC[checkedItem.ToString()]);
            }
            while (NewInsulationСonsumables.CheckedItems.Count != 0)
            {
                NewInsulationСonsumables.Items.Remove(NewInsulationСonsumables.CheckedItems[0]);
            }
            calculateInsulationСonsumablesSum();
        }

        private void InsulationСonsumables_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < InsulationСonsumables.RowCount - 1; i++)
            {
                InsulationСonsumables.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void InsulationСonsumables_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int nRow = e.RowIndex;
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {

                InsulationСonsumables.Rows[nRow].Cells[4].Value = Convert.ToDouble(InsulationСonsumables.Rows[nRow].Cells[2].Value) * Convert.ToDouble(InsulationСonsumables.Rows[nRow].Cells[3].Value);
                calculateInsulationСonsumablesSum();
            }
            checkAllInsulationСonsumables();
        }


        //////////////////////////////////////////////////////////////

        private void deleteInstallationWork_Click(object sender, EventArgs e)
        {
            while (InstallationWork.SelectedRows.Count != 0)
            {
                if (listIW.ContainsKey(InstallationWork.SelectedRows[0].Cells[1].Value.ToString()))
                    NewInstallationWork.Items.Add(InstallationWork.SelectedRows[0].Cells[1].Value);
                InstallationWork.Rows.Remove(InstallationWork.SelectedRows[0]);
            }

            for (int i = 0; i < InstallationWork.RowCount - 1; i++)
            {
                InstallationWork.Rows[i].Cells[0].Value = i + 1;
            }
            calculateInstallationWorkSum();

            checkAllInstallationWork();
        }
        private void checkAllInstallationWork()
        {
            NewInstallationWork.Items.Clear();
            foreach (string str in listIW.Keys)
            {
                bool f = false;
                for (int i = 0; i < InstallationWork.RowCount - 1; i++)
                {
                    if (str == InstallationWork.Rows[i].Cells[1].Value.ToString())
                        f = true;
                }
                if (!f)
                    NewInstallationWork.Items.Add(str);
            }
        }

        private void calculateInstallationWorkSum()
        {
            sumIW = 0;
            foreach (DataGridViewRow row in InstallationWork.Rows)
                sumIW += Convert.ToDouble(row.Cells[4].Value);
            SumInstallationWork.Text = sumIW.ToString() + " Руб.";
            calculateResults();
        }


        private void addInstallationWork_Click(object sender, EventArgs e)
        {
            int n;
            if (InstallationWork.RowCount == 0)
                n = 1;
            else
                n = InstallationWork.RowCount + 1;
            foreach (Object checkedItem in NewInstallationWork.CheckedItems)
            {
                InstallationWork.Rows.Add(n++, checkedItem.ToString(), "1", listIW[checkedItem.ToString()], listIW[checkedItem.ToString()]);
            }
            while (NewInstallationWork.CheckedItems.Count != 0)
            {
                NewInstallationWork.Items.Remove(NewInstallationWork.CheckedItems[0]);
            }
            calculateInstallationWorkSum();
        }

        private void InstallationWork_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < InstallationWork.RowCount - 1; i++)
            {
                InstallationWork.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void InstallationWork_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int nRow = e.RowIndex;
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {

                InstallationWork.Rows[nRow].Cells[4].Value = Convert.ToDouble(InstallationWork.Rows[nRow].Cells[2].Value) * Convert.ToDouble(InstallationWork.Rows[nRow].Cells[3].Value);
                calculateInstallationWorkSum();
            }
            checkAllInstallationWork();
        }


        //////////////////////////////////////////////////////////////

        private void deleteRiggingDelivery_Click(object sender, EventArgs e)
        {
            while (RiggingDelivery.SelectedRows.Count != 0)
            {
                if (listRD.ContainsKey(RiggingDelivery.SelectedRows[0].Cells[1].Value.ToString()))
                    NewRiggingDelivery.Items.Add(RiggingDelivery.SelectedRows[0].Cells[1].Value);
                RiggingDelivery.Rows.Remove(RiggingDelivery.SelectedRows[0]);
            }

            for (int i = 0; i < RiggingDelivery.RowCount - 1; i++)
            {
                RiggingDelivery.Rows[i].Cells[0].Value = i + 1;
            }
            calculateRiggingDeliverySum();

            checkAllRiggingDelivery();
        }
        private void checkAllRiggingDelivery()
        {
            NewRiggingDelivery.Items.Clear();
            foreach (string str in listRD.Keys)
            {
                bool f = false;
                for (int i = 0; i < RiggingDelivery.RowCount - 1; i++)
                {
                    if (str == RiggingDelivery.Rows[i].Cells[1].Value.ToString())
                        f = true;
                }
                if (!f)
                    NewRiggingDelivery.Items.Add(str);
            }
        }

        private void calculateRiggingDeliverySum()
        {
            sumRD = 0;
            foreach (DataGridViewRow row in RiggingDelivery.Rows)
                sumRD += Convert.ToDouble(row.Cells[4].Value);
            SumRiggingDelivery.Text = sumRD.ToString() + " Руб.";
            calculateResults();
        }


        private void addRiggingDelivery_Click(object sender, EventArgs e)
        {
            int n;
            if (RiggingDelivery.RowCount == 0)
                n = 1;
            else
                n = RiggingDelivery.RowCount + 1;
            foreach (Object checkedItem in NewRiggingDelivery.CheckedItems)
            {
                RiggingDelivery.Rows.Add(n++, checkedItem.ToString(), "1", listRD[checkedItem.ToString()], listRD[checkedItem.ToString()]);
            }
            while (NewRiggingDelivery.CheckedItems.Count != 0)
            {
                NewRiggingDelivery.Items.Remove(NewRiggingDelivery.CheckedItems[0]);
            }
            calculateRiggingDeliverySum();
        }

        private void RiggingDelivery_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < RiggingDelivery.RowCount - 1; i++)
            {
                RiggingDelivery.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void RiggingDelivery_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int nRow = e.RowIndex;
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {

                RiggingDelivery.Rows[nRow].Cells[4].Value = Convert.ToDouble(RiggingDelivery.Rows[nRow].Cells[2].Value) * Convert.ToDouble(RiggingDelivery.Rows[nRow].Cells[3].Value);
                calculateRiggingDeliverySum();
            }
            checkAllRiggingDelivery();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            calculateResults();
            if (sumND != 0)
            {
                decimal percentD = Convert.ToDecimal(Convert.ToDouble(numericUpDown1.Value) / sumND) * 100;
                if (percentD > 100)
                    percentD = 100;
                else if (percentD < 0)
                    percentD = 0;
                percentDiscount.Value = percentD;
            }
        }

        private void percentDiscount_ValueChanged(object sender, EventArgs e)
        {
            if (sumND != 0)
            {
                numericUpDown1.Value = Convert.ToDecimal(Math.Round(sumND * (Convert.ToDouble(percentDiscount.Value) / 100), 0));
            }
        }
    }
}
