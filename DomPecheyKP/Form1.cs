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
using System.Threading;
using System.Resources;
using System.Xml;

namespace DomPecheyKP
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            ChimneyElements.ForeColor = NameOfKiln.ForeColor = InsulationСonsumables.ForeColor = addInstallationWork.ForeColor = RiggingDelivery.ForeColor = Color.FromArgb(104, 51, 5);
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
        double resultSum1 = 0;
        double resultSum2 = 0;
        double sumFinal = 0;
        string nameOfSheetIC = "Изоляц и расход материалы";
        string nameOfSheetIW = "Монтажные работы и выезд";
        string nameOfSheetRD = "Такелажные работы и доставка";
        string managerName = "";
        Excel.Application ObjWorkExcel;
        Excel.Worksheet ObjWorkSheet;
        Excel.Workbook ObjWorkBook;
        Alert formAlert;
        DataGridView currentDataGridView;
        BaseColor borderColor = BaseColor.BLACK;

        private PdfPCell getMidHeader(string str, iTextSharp.text.Font font)
        {
            PdfPCell cell = new PdfPCell(new Phrase(new Phrase(str, font)));
            cell.BorderColor = borderColor;
            cell.Colspan = 5;
            cell.BorderWidthLeft = 0;
            cell.BorderWidthRight = 0;
            cell.PaddingBottom = 10;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            return cell;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            string nameOfNewFile;
            saveFileDialog.Filter = "Pdf files|*.pdf";
            
            var typeProduct = from RadioButton r in ProductType.Controls where r.Checked == true select r.Tag;
            var VIP = from RadioButton r in isVIP.Controls where r.Checked == true select r.Tag;
            var typeProductstring = from RadioButton r in ProductType.Controls where r.Checked == true select r.Text;
            var VIPstring = from RadioButton r in isVIP.Controls where r.Checked == true select r.Text;
            saveFileDialog.FileName = "Комерческое предложение " + VIPstring.First().ToString() + " " + typeProductstring.First().ToString()  + " " + ClientName.Text + " " + DateTime.Now.ToShortDateString() + ".pdf";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                nameOfNewFile = saveFileDialog.FileName;

                int numberOfFile = 3;
                string fileName = "";
                int pageClientName = 0;
                int pageTable = 0;
                int pageManagerName = 0;
                int firstMaxHeightTable = 0;
                int secondMaxHeightTable = 0;
                int firstTableMarginTop = 0;
                int secongTableMarginTop = 0;
                string fontClientNameName = "";
                string fontClientNamePath = "";
                string fontManagerNameName = "";
                string fontManagerNamePath = "";
                string fontHeaderTableName = "";
                string fontHeaderTablePath = "";
                string fontTableName = "";
                string fontTablePath = "";
                BaseColor mainColor = new BaseColor(55, 55, 55);

                if (typeProduct.First().ToString() != "4")
                {
                    if (VIP.First().ToString() == "0")
                    {
                        fileName = @"template/Fireplace.pdf";
                        pageClientName = 0;
                        pageTable = 8;
                        pageManagerName = 0;
                        firstMaxHeightTable = 460;
                        secondMaxHeightTable = 520;
                        firstTableMarginTop = 0;
                        secongTableMarginTop = 0;
                        fontClientNameName = "Sitka Text Italic";
                        fontClientNamePath = "Fonts/SitkaText.ttf";
                        fontManagerNameName = "Sitka Text Bold";
                        fontManagerNamePath = "Fonts/SitkaText-Bold.ttf";
                        fontHeaderTableName = "Roboto Bold";
                        fontHeaderTablePath = "Fonts/Roboto-Bold.ttf";
                        fontTableName = "Roboto";
                        fontTablePath = "Fonts/Roboto-Regular.ttf";
                    }
                    else
                    {
                        fileName = @"template/FireplaceVIP.pdf";
                        pageClientName = 2;
                        pageTable = 5;
                        pageManagerName = 9;
                        firstMaxHeightTable = 420;
                        secondMaxHeightTable = 520;
                        firstTableMarginTop = 140;
                        secongTableMarginTop = 0;
                        fontClientNameName = "Sitka Text Italic";
                        fontClientNamePath = "Fonts/SitkaText.ttf";
                        fontManagerNameName = "Sitka Text Bold";
                        fontManagerNamePath = "Fonts/SitkaText-Bold.ttf";
                        fontHeaderTableName = "Roboto Bold";
                        fontHeaderTablePath = "Fonts/Roboto-Bold.ttf";
                        fontTableName = "Roboto";
                        fontTablePath = "Fonts/Roboto-Regular.ttf";
                    }
                }

                else
                {
                    if (VIP.First().ToString() == "0")
                    {
                        fileName = @"template/Bath.pdf";
                        pageClientName = 0;
                        pageTable = 4;
                        pageManagerName = 0;
                        firstMaxHeightTable = 310;
                        secondMaxHeightTable = 520;
                        firstTableMarginTop = 250;
                        secongTableMarginTop = 0;
                        fontClientNameName = "Sitka Text Italic";
                        fontClientNamePath = "Fonts/SitkaText.ttf";
                        fontManagerNameName = "Sitka Text Bold";
                        fontManagerNamePath = "Fonts/SitkaText-Bold.ttf";
                        fontHeaderTableName = "Roboto Bold";
                        fontHeaderTablePath = "Fonts/Roboto-Bold.ttf";
                        fontTableName = "Roboto";
                        fontTablePath = "Fonts/Roboto-Regular.ttf";
                        mainColor = new BaseColor(100, 52, 13);
                        borderColor = new BaseColor(100, 52, 13);
                    }
                    else
                    {
                        fileName = @"template/BathVIP.pdf";
                        pageClientName = 2;
                        pageTable = 7;
                        pageManagerName = 11;
                        firstMaxHeightTable = 460;
                        secondMaxHeightTable = 520;
                        firstTableMarginTop = 100;
                        secongTableMarginTop = 0;
                        fontClientNameName = "Sitka Text Italic";
                        fontClientNamePath = "Fonts/SitkaText.ttf";
                        fontManagerNameName = "Sitka Text Bold";
                        fontManagerNamePath = "Fonts/SitkaText-Bold.ttf";
                        fontHeaderTableName = "Roboto Bold";
                        fontHeaderTablePath = "Fonts/Roboto-Bold.ttf";
                        fontTableName = "Roboto";
                        fontTablePath = "Fonts/Roboto-Regular.ttf";
                    }
                }

    

                int countColumns = 5;
                PdfPTable fTable = new PdfPTable(countColumns);
                int[] firstTablecellwidth = { 12, 53, 19, 12, 12 };
                using (var reader = new PdfReader(fileName))
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


                        var fontName = fontHeaderTableName;
                        if (!FontFactory.IsRegistered(fontName))
                        {
                            var fontPath = Environment.GetEnvironmentVariable(fontHeaderTablePath);
                            FontFactory.Register(fontHeaderTablePath);
                        }
                        iTextSharp.text.Font fontHeaderTable = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                        //BaseColor col = new BaseColor(55, 55, 55);
                        fontHeaderTable.Color = mainColor;// iTextSharp.text.BaseColor.BLACK;
                        fontHeaderTable.Size = 28;

                        fontName = fontTableName;
                        if (!FontFactory.IsRegistered(fontName))
                        {
                            var fontPath = Environment.GetEnvironmentVariable(fontTablePath);
                            FontFactory.Register(fontTablePath);
                        }
                        iTextSharp.text.Font fontTable = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                        fontTable.Color = mainColor;// iTextSharp.text.BaseColor.BLACK;
                        fontTable.Size = 25;


                        fTable.WidthPercentage = 80;
                        fTable.SetWidths(firstTablecellwidth);
                        //Добавим в таблицу общий заголовок
                        int paddingLeft = 5;
                        int paddingTop = 13;
                        if (typeProduct.First().ToString() != "4" && VIP.First().ToString() == "0") // печи не VIP
                        {
                            fTable.DefaultCell.BorderColor = mainColor;
                            fTable.DefaultCell.BorderColorBottom = mainColor;
                        }

                        PdfPCell cell;
                        cell = new PdfPCell(new Phrase(new Phrase("Номер", fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        cell.BorderWidthLeft = 0;
                        cell.PaddingLeft = paddingLeft;
                        cell.PaddingTop = paddingTop;
                        //cell.BorderColor = mainColor;
                        cell.VerticalAlignment = Element.ALIGN_CENTER;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Наименование", fontHeaderTable)));
                        cell.BorderColor = borderColor;

                        cell.BorderWidthTop = 0;
                        cell.PaddingLeft = paddingLeft;
                        cell.PaddingTop = paddingTop;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Количество", fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        cell.PaddingLeft = paddingLeft;
                        cell.PaddingTop = paddingTop;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Цена за шт.", fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        cell.PaddingLeft = paddingLeft;
                        cell.PaddingBottom = 5;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Сумма", fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        cell.BorderWidthRight = 0;
                        cell.PaddingLeft = paddingLeft;
                        cell.PaddingTop = paddingTop;
                        fTable.AddCell(cell);
                        /////////////////////////////////////////////////////////

                        var a = from RadioButton r in ProductType.Controls where r.Checked == true select r.Text;

                        fTable.AddCell(getMidHeader("1. " + a.First(), fontTable));


                        cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[0].Value.ToString(), fontTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        cell.BorderWidthLeft = 0;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[1].Value.ToString(), fontTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[2].Value.ToString(), fontTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[3].Value.ToString() + " ₽", fontTable)));
                        cell.BorderColor = borderColor;
                        cell.BorderWidthTop = 0;
                        fTable.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase(NameOfKiln.Rows[0].Cells[4].Value.ToString() + " ₽", fontTable)));
                        cell.BorderColor = borderColor;
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
                        str += " (" + a.First() + ")";

                        fTable.AddCell(getMidHeader(str, fontTable));


                        if (ChimneyElements.RowCount == 1)
                        {
                            cell = new PdfPCell(new Phrase(new Phrase("1", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.BorderWidthLeft = 0;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("Отсутствует", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);

                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.BorderWidthRight = 0;
                            fTable.AddCell(cell);
                        }
                        else
                            for (int j = 0; j < ChimneyElements.RowCount - 1; j++)
                        {
                            for (int k = 0; k < countColumns; k++)
                            {

                                string val = ChimneyElements.Rows[j].Cells[k].Value.ToString();
                                if (k > 2)
                                    val += " ₽";
                                cell = new PdfPCell(new Phrase(new Phrase(val, fontTable)));
                                cell.BorderColor = borderColor;
                                if (k == 0)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.BorderWidthLeft = 0;
                                }
                                else if (k == countColumns - 1)
                                    cell.BorderWidthRight = 0;
                                else if (k == 2)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                }
                                cell.PaddingBottom = 10;
                                cell.PaddingTop = 0;
                                cell.VerticalAlignment = Element.ALIGN_TOP;
                                fTable.AddCell(cell);
                            }
                        }


                        fTable.AddCell(getMidHeader("3. Изоляционные и расходные материалы", fontTable));


                        if(InsulationСonsumables.RowCount== 1)
                        {
                            cell = new PdfPCell(new Phrase(new Phrase("1", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.BorderWidthLeft = 0;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("Отсутствует", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);

                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.BorderWidthRight = 0;
                            fTable.AddCell(cell);
                        }
                        else
                        for (int j = 0; j < InsulationСonsumables.RowCount - 1; j++)
                        {
                            for (int k = 0; k < countColumns; k++)
                            {
                                string val = InsulationСonsumables.Rows[j].Cells[k].Value.ToString();
                                if (k > 2)
                                    val += " ₽";
                                cell = new PdfPCell(new Phrase(new Phrase(val, fontTable)));
                                cell.BorderColor = borderColor;
                                if (k == 0)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.BorderWidthLeft = 0;
                                }
                                else if (k == countColumns - 1)
                                    cell.BorderWidthRight = 0;
                                else if (k == 2)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                }
                                cell.PaddingBottom = 10;
                                fTable.AddCell(cell);
                            }
                        }



                        fTable.AddCell(getMidHeader("4. Монтажные работы, выезд на замер", fontTable));
                        if (InstallationWork.RowCount == 1)
                        {
                            cell = new PdfPCell(new Phrase(new Phrase("1", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.BorderWidthLeft = 0;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("Отсутствует", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);

                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.BorderWidthRight = 0;
                            fTable.AddCell(cell);
                        }
                        else
                            for (int j = 0; j < InstallationWork.RowCount - 1; j++)
                        {
                            for (int k = 0; k < countColumns; k++)
                            {
                                string val = InstallationWork.Rows[j].Cells[k].Value.ToString();
                                if (k > 2)
                                    val += " ₽";
                                cell = new PdfPCell(new Phrase(new Phrase(val, fontTable)));
                                cell.BorderColor = borderColor;
                                if (k == 0)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.BorderWidthLeft = 0;
                                }
                                else if (k == countColumns - 1)
                                    cell.BorderWidthRight = 0;
                                else if (k == 2)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                }
                                cell.PaddingBottom = 10;
                                fTable.AddCell(cell);
                            }
                        }


                        fTable.AddCell(getMidHeader("5. Такелажные работы и доставка", fontTable));

                        if (RiggingDelivery.RowCount == 1)
                        {
                            cell = new PdfPCell(new Phrase(new Phrase("1", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.BorderWidthLeft = 0;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("Отсутствует", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);

                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            fTable.AddCell(cell);
                            cell = new PdfPCell(new Phrase(new Phrase("", fontTable)));
                            cell.BorderColor = borderColor;
                            cell.BorderWidthRight = 0;
                            fTable.AddCell(cell);
                        }
                        else
                            for (int j = 0; j < RiggingDelivery.RowCount - 1; j++)
                        {
                            for (int k = 0; k < countColumns; k++)
                            {
                                string val = RiggingDelivery.Rows[j].Cells[k].Value.ToString();
                                if (k > 2)
                                    val += " ₽";
                                cell = new PdfPCell(new Phrase(new Phrase(val, fontTable)));
                                cell.BorderColor = borderColor;
                                if (k == 0)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.BorderWidthLeft = 0;
                                }
                                else if (k == countColumns - 1)
                                    cell.BorderWidthRight = 0;
                                else if (k == 2)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                }
                                cell.PaddingBottom = 10;
                                fTable.AddCell(cell);
                            }
                        }


                        cell = new PdfPCell(new Phrase(new Phrase("Скидка:", fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.Colspan = 2;
                        cell.BorderWidthLeft = 0;
                        cell.BorderWidthRight = 0;
                        cell.PaddingBottom = 10;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        fTable.AddCell(cell);

                        cell = new PdfPCell(new Phrase(new Phrase(sumDiscont.Value.ToString(), fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.Colspan = 3;
                        cell.BorderWidthLeft = 0;
                        cell.BorderWidthRight = 0;
                        cell.BorderWidthRight = 0;
                        cell.PaddingBottom = 10;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        fTable.AddCell(cell);

                        cell = new PdfPCell(new Phrase(new Phrase("Итого:", fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.Colspan = 2;
                        cell.BorderWidthLeft = 0;
                        cell.BorderWidthRight = 0;
                        cell.PaddingBottom = 10;
                        cell.HorizontalAlignment = Element.ALIGN_LEFT;
                        fTable.AddCell(cell);

                        cell = new PdfPCell(new Phrase(new Phrase(sumFinal.ToString(), fontHeaderTable)));
                        cell.BorderColor = borderColor;
                        cell.Colspan = 3;
                        cell.BorderWidthLeft = 0;
                        cell.BorderWidthRight = 0;
                        cell.BorderWidthRight = 0;
                        cell.PaddingBottom = 10;
                        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                        fTable.AddCell(cell);
                        document.Add(fTable);

                        document.Close();
                        writer.Close();
                    }


                    using (var fileStream = new FileStream(nameOfNewFile, FileMode.Create, FileAccess.Write))
                    {
                        var document = new Document(reader.GetPageSizeWithRotation(1));
                        var writer = PdfWriter.GetInstance(document, fileStream);
                        document.Open();

                        System.Text.EncodingProvider ppp = System.Text.CodePagesEncodingProvider.Instance;
                        Encoding.RegisterProvider(ppp);
                        var fontName = fontClientNameName;
                        if (!FontFactory.IsRegistered(fontName))
                        {
                            var fontPath = Environment.GetEnvironmentVariable(fontClientNamePath);
                            FontFactory.Register(fontClientNamePath);
                        }
                        iTextSharp.text.Font fontClientName = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                        fontClientName.Color = iTextSharp.text.BaseColor.WHITE;

                        fontName = fontManagerNameName;
                        if (!FontFactory.IsRegistered(fontName))
                        {
                            var fontPath = Environment.GetEnvironmentVariable(fontManagerNamePath);
                            FontFactory.Register(fontManagerNamePath);
                        }
                        iTextSharp.text.Font fontManagerName = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                        fontManagerName.Color = iTextSharp.text.BaseColor.WHITE;



                        for (var i = 1; i <= reader.NumberOfPages; i++)
                        {
                            document.NewPage();
                            var importedPage = writer.GetImportedPage(reader, i);
                            var contentByte = writer.DirectContent;
                            contentByte.AddTemplate(importedPage, 0, 0);



                            if (i == pageClientName)
                            {
                                contentByte.BeginText();
                                contentByte.SetColorFill(BaseColor.WHITE);
                                contentByte.SetFontAndSize(fontClientName.BaseFont, 70);
                                string name = ClientName.Text.ToString();
                                int y = 550;
                                contentByte.ShowTextAligned(PdfContentByte.ALIGN_CENTER, name + ",", 650, y, 0);
                                y -= 80;
                                contentByte.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Добрый день!", 650, y, 0);
                                contentByte.EndText();
                            }

                            if (i == pageTable)
                            {
                                int maxHeightTable = firstMaxHeightTable;
                                Paragraph p;

                                PdfPTable sTable = new PdfPTable(countColumns);
                                sTable.WidthPercentage = 80;
                                sTable.SetWidths(firstTablecellwidth);
                                sTable.SpacingBefore = firstTableMarginTop;
                                float sumHeight = 0;
                                for (int j = 0; j < fTable.Rows.Count; j++)
                                {
                                    int q = fTable.Rows[j].GetCells()[0].Colspan;
                                    bool isHeightMoreMaxHeight = (sumHeight + fTable.Rows[j].MaxHeights) > maxHeightTable;
                                    bool isHeaderLastRow = fTable.Rows[j].GetCells()[0].Colspan == 5 && ((sumHeight + fTable.Rows[j + 1].MaxHeights + fTable.Rows[j].MaxHeights) > maxHeightTable);
                                    bool isFooterNotNextRow = fTable.Rows[j].GetCells()[0].Colspan != 2;
                                    if ((isHeightMoreMaxHeight || isHeaderLastRow) && isFooterNotNextRow)
                                    {

                                        //new table
                                        p = new Paragraph("     ");
                                        p.SpacingAfter = 90;
                                        document.Add(p);
                                        document.Add(sTable);
                                        document.NewPage();
                                        if (i == pageTable)
                                        {
                                            i++;
                                            importedPage = writer.GetImportedPage(reader, i);
                                            sTable.SpacingBefore = secongTableMarginTop;
                                            maxHeightTable = secondMaxHeightTable;
                                        }
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
                                p.SpacingAfter = 90;
                                document.Add(p);
                                //sTable.SpacingBefore = secongTableMarginTop[numberOfFile];
                                document.Add(sTable);
                            }
                            if (i == pageManagerName)
                            {
                                contentByte.BeginText();
                                contentByte.SetColorFill(BaseColor.WHITE);
                                contentByte.SetFontAndSize(fontManagerName.BaseFont, 33);
                                contentByte.ShowTextAligned(PdfContentByte.ALIGN_CENTER, ManagerName.Text, 640, 93, 0);
                                contentByte.EndText();
                            }

                        }
                        if (numberOfFile ==3)
                        {
                            document.NewPage();
                            var importedPage = writer.GetImportedPage(reader, reader.NumberOfPages);
                            var contentByte = writer.DirectContent;
                            contentByte.AddTemplate(importedPage, 0, 0);
                        }
                        document.Close();
                        writer.Close();
                    }

                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            long size = 0;
            System.IO.FileInfo file;

            file = new System.IO.FileInfo(@"template/Bath.pdf");
            size += file.Length;
            file = new System.IO.FileInfo(@"template/BathVIP.pdf");
            size += file.Length;
            file = new System.IO.FileInfo(@"template/Fireplace.pdf");
            size += file.Length;
            file = new System.IO.FileInfo(@"template/FireplaceVIP.pdf");
            size += file.Length;

            if (size != 38417940)
                this.Close();

            string name = "";
            XmlDocument doc = new XmlDocument();
            doc.Load("manager.xml");
            foreach (XmlNode node in doc.DocumentElement)
            {
                name = node.Attributes[0].Value;                
            }
            ManagerName.Text = name;


            ObjWorkExcel = new Excel.Application();
            ObjWorkBook = ObjWorkExcel.Workbooks.Open(Environment.CurrentDirectory + @"\ДанныеКП.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            loadElD("Дымок 0,5", 2);
            NameOfKiln.Rows.Add("1", "", "1", "0", "0");

            loadIС();
            loadIW(1);
            loadRD();
            //удалить
            //foreach (Object checkedItem in NewChimneyElements.Items)
            //{
            //    ChimneyElements.Rows.Add("1", checkedItem.ToString(), "1", list[checkedItem.ToString()], list[checkedItem.ToString()]);
            //}

            //foreach (Object checkedItem in NewInsulationСonsumables.Items)
            //{
            //    InsulationСonsumables.Rows.Add("1", checkedItem.ToString(), "1", listIC[checkedItem.ToString()], listIC[checkedItem.ToString()]);
            //}

            //foreach (Object checkedItem in NewInstallationWork.Items)
            //{
            //    InstallationWork.Rows.Add("1", checkedItem.ToString(), "1", listIW[checkedItem.ToString()], listIW[checkedItem.ToString()]);
            //}


            //foreach (Object checkedItem in NewRiggingDelivery.Items)
            //{
            //    RiggingDelivery.Rows.Add("1", checkedItem.ToString(), "1", listRD[checkedItem.ToString()], listRD[checkedItem.ToString()]);
            //}

        }

        private void loadIС()
        {
            sumIC = 0;
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

        private void loadElD(string nameOfPage, int nColumn)
        {
            //sumCE = 0;
            NewChimneyElements.Items.Clear();
            //ChimneyElements.Rows.Clear();

            list = new Dictionary<string, double>();

            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[nameOfPage]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            for (int i = 3; i <= (int)lastCell.Row; i++) // по всем строкам
            {
                string name = ObjWorkSheet.Cells[i, 1].Text.ToString();
                bool nameNotNull = name != "";
                bool nameNotContainsInList = !list.ContainsKey(name);
                bool nameNotContainsInDataGrid = true;
                foreach (DataGridViewRow row in ChimneyElements.Rows)
                {
                    if (row.Cells[1].Value != null && name == row.Cells[1].Value.ToString())
                        nameNotContainsInDataGrid = false;
                }

                if (nameNotNull && nameNotContainsInList)
                {
                    try
                    {
                        list.Add(name, Convert.ToDouble(ObjWorkSheet.Cells[i, nColumn].Text.ToString()));//считываем текст в строку
                    }
                    catch (Exception ex)
                    {
                        list.Add(name, 0);//считываем текст в строку

                    }
                    if(nameNotContainsInDataGrid)
                    NewChimneyElements.Items.Add(name);
                }

            }

        }

        private void loadIW(int nColumn)
        {
            //sumIW = 0;
            NewInstallationWork.Items.Clear();
            //InstallationWork.Rows.Clear();
            listIW = new Dictionary<string, double>();
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[nameOfSheetIW]; //получить 1 лист
            var lastCell = ObjWorkSheet.Columns[nColumn].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку

            for (int i = 3; i <= (int)lastCell.Row; i++) // по всем строкам
            {
                string name = ObjWorkSheet.Cells[i, nColumn].Text.ToString();
                bool nameNotNull = name.Trim() != "";
                bool nameNotContainsInList = !listIW.ContainsKey(name);
                bool nameNotContainsInDataGrid = true;
                foreach (DataGridViewRow row in InstallationWork.Rows)
                {
                    if (row.Cells[1].Value != null && name == row.Cells[1].Value.ToString())
                        nameNotContainsInDataGrid = false;
                }


                if (nameNotNull && nameNotContainsInList)
                {
                    listIW.Add(name, Convert.ToDouble(ObjWorkSheet.Cells[i, nColumn + 1].Text.ToString()));//считываем текст в строку
                    if (nameNotContainsInDataGrid)
                    NewInstallationWork.Items.Add(name);
                }

            }
            NewInstallationWork.Sorted = true;
        }

        private void calculateResults()
        {
            double sumC = Convert.ToDouble(NameOfKiln.Rows[0].Cells[4].Value.ToString());
            resultSum1 = sumCE + sumIC + sumC;
            SumChimneyManufacturerAndInsulation.Text = (resultSum1).ToString() + " Руб.";
            resultSum2 = sumRD + sumIW;
            SumRiggingAndInstall.Text = (resultSum2).ToString() + " Руб.";
            sumND = resultSum1 + resultSum2;
            SumNotDiscount.Text = (sumND).ToString() + " Руб.";
            sumFinal = sumND - Convert.ToDouble(sumDiscont.Value);
            AllSum.Text = (sumFinal).ToString() + " Руб.";
        }

        private void calculateChimneyElementsSum()
        {
            sumCE = 0;
            foreach (DataGridViewRow row in ChimneyElements.Rows)
                sumCE += Convert.ToDouble(row.Cells[4].Value);
            SumChimneyElements.Text = sumCE.ToString() + " Руб.";
            calculateResults();
        }





        private void loadRD()
        {
            sumRD = 0;
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

            foreach (DataGridViewCell cell in ChimneyElements.SelectedCells)
            {
                if (ChimneyElements.Rows[cell.RowIndex].Cells[1].Value != null)
                {
                    if (list.ContainsKey(this.ChimneyElements.Rows[cell.RowIndex].Cells[1].Value.ToString()))
                        NewChimneyElements.Items.Add(ChimneyElements.Rows[cell.RowIndex].Cells[1].Value);
                    ChimneyElements.Rows.RemoveAt(cell.RowIndex);
                }
            }

            for (int i = 0; i < ChimneyElements.RowCount - 1; i++)
               ChimneyElements.Rows[i].Cells[0].Value = i + 1;

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

        private void currentDataGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridView data = (DataGridView)sender;

            for (int i = 0; i < data.RowCount - 1; i++)
                data.Rows[i].Cells[0].Value = i + 1;

        }


        private void d_CheckedChanged(object sender, EventArgs e)
        {
            var a = from RadioButton r in Manufacturer.Controls where r.Checked == true select r.Text;
            var b = from RadioButton r in MetalThickness.Controls where r.Checked == true select r.Tag;
            var c = from RadioButton r in Diameter.Controls.OfType<RadioButton>() where r.Checked == true select r.Tag;
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                string page = a.First().ToString() + " " + b.First().ToString();
                int col = Convert.ToInt32(c.First().ToString());
                loadElD(page, col);
                SumChimneyElements.Text = "0 Руб.";
                calculateResults();
            }

        }

         private void mt_CheckedChanged(object sender, EventArgs e)
        {
            var a = from RadioButton r in Manufacturer.Controls where r.Checked == true select r.Text;
            var b = from RadioButton r in MetalThickness.Controls where r.Checked == true select r.Tag;
            var c = from RadioButton r in Diameter.Controls.OfType<RadioButton>() where r.Checked == true select r.Tag;
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                string page = a.First().ToString() + " " + b.First().ToString();
                int col = Convert.ToInt32(c.First().ToString());
                loadElD(page,col);
                //SumChimneyElements.Text = "0 Руб.";
                //calculateResults();
            }

        }

        private void m_CheckedChanged(object sender, EventArgs e)
        {


            var a = from RadioButton r in Manufacturer.Controls where r.Checked == true select r.Text;
            if (a.First().ToString() == "Schiedel")
            {
                MetalThickness1.Text = "PM25";
                MetalThickness2.Text = "PM50";
                MetalThickness1.Tag = "PM25";
                MetalThickness2.Tag = "PM50";
                MetalThickness.Text = "Толщина изоляции";
            }
            else
            {
                MetalThickness1.Text = "0,5";
                MetalThickness2.Text = "0,8";
                MetalThickness1.Tag = "0,5";
                MetalThickness2.Tag = "0,8";
                MetalThickness.Text = "Толщина металла";
            }
            var b = from RadioButton r in MetalThickness.Controls where r.Checked == true select r.Tag;
            var c = from RadioButton r in Diameter.Controls.OfType<RadioButton>() where r.Checked == true select r.Tag;



            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                string page = a.First().ToString() + " " + b.First().ToString();
                int col = Convert.ToInt32(c.First().ToString());
                loadElD(page, col);
                SumChimneyElements.Text = "0 Руб.";
                calculateResults();
            }


        }

        private void iw_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = (RadioButton)sender;
            if (radioButton.Checked)
            {
                loadIW(Convert.ToInt32(radioButton.Tag));
                //SumInstallationWork.Text = "0 Руб.";
                //calculateResults();
            }
        }


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

        private void DataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridView data = (DataGridView)sender;
            currentDataGridView = data;
            currentDataGridView.EditingControl.KeyPress -= EditingControl_KeyPress;
            currentDataGridView.EditingControl.KeyPress += EditingControl_KeyPress;
        }

        private void currentDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int nRow = e.RowIndex;

            if ((e.ColumnIndex == 2 || e.ColumnIndex == 3) && currentDataGridView.Rows[e.RowIndex].Cells[1].Value != null)
            {
                var isValid = Regex.IsMatch(currentDataGridView.Rows[nRow].Cells[e.ColumnIndex].Value.ToString(), @"^[0-9]*[,]?[0-9]+$");
                if (!isValid)
                {
                    MessageBox.Show("Ошибка при вводе числа. Введите значение заново.");
                    currentDataGridView.Rows[nRow].Cells[e.ColumnIndex].Value = "0";
                }
                currentDataGridView.Rows[nRow].Cells[4].Value = Convert.ToDouble(currentDataGridView.Rows[nRow].Cells[2].Value) * Convert.ToDouble(currentDataGridView.Rows[nRow].Cells[3].Value);
            }
            switch (currentDataGridView.Name)
            {
                case "ChimneyElements":
                    calculateChimneyElementsSum();
                    break;
                case "InsulationСonsumables":
                    calculateInsulationСonsumablesSum();
                    break;
                case "InstallationWork":
                    calculateInstallationWorkSum();
                    break;
                case "RiggingDelivery":
                    calculateRiggingDeliverySum();
                    break;
            }
            calculateResults();

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            ObjWorkBook = null;
            ObjWorkExcel = null;
            ObjWorkSheet = null;
            GC.Collect();
        }


        private void deleteInsulationСonsumables_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in InsulationСonsumables.SelectedCells)
            {
                if (InsulationСonsumables.Rows[cell.RowIndex].Cells[1].Value != null)
                {
                    if (listIC.ContainsKey(this.InsulationСonsumables.Rows[cell.RowIndex].Cells[1].Value.ToString()))
                        NewInsulationСonsumables.Items.Add(InsulationСonsumables.Rows[cell.RowIndex].Cells[1].Value);
                    InsulationСonsumables.Rows.RemoveAt(cell.RowIndex);
                }

            }

            for (int i = 0; i < InsulationСonsumables.RowCount - 1; i++)
                InsulationСonsumables.Rows[i].Cells[0].Value = i + 1;
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

        private void calculateSum()
        {
            sumIC = 0;
            foreach (DataGridViewRow row in InsulationСonsumables.Rows)
                sumIC += Convert.ToDouble(row.Cells[4].Value);
            SumInsulationСonsumables.Text = sumIC.ToString() + " Руб.";
            calculateResults();
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


        //////////////////////////////////////////////////////////////

        private void deleteInstallationWork_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in InstallationWork.SelectedCells)
            {
                if (InstallationWork.Rows[cell.RowIndex].Cells[1].Value != null)
                {
                    if (listIW.ContainsKey(this.InstallationWork.Rows[cell.RowIndex].Cells[1].Value.ToString()))
                        NewInstallationWork.Items.Add(InstallationWork.Rows[cell.RowIndex].Cells[1].Value);
                    InstallationWork.Rows.RemoveAt(cell.RowIndex);
                }

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


        //////////////////////////////////////////////////////////////

        private void deleteRiggingDelivery_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in RiggingDelivery.SelectedCells)
            {
                if (RiggingDelivery.Rows[cell.RowIndex].Cells[1].Value != null)
                {
                    if (listRD.ContainsKey(this.RiggingDelivery.Rows[cell.RowIndex].Cells[1].Value.ToString()))
                        NewRiggingDelivery.Items.Add(RiggingDelivery.Rows[cell.RowIndex].Cells[1].Value);
                    RiggingDelivery.Rows.RemoveAt(cell.RowIndex);
                }

            }

            for (int i = 0; i < RiggingDelivery.RowCount - 1; i++)
                RiggingDelivery.Rows[i].Cells[0].Value = i + 1;
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


        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            calculateResults();
            if (sumND != 0)
            {
                decimal percentD = Convert.ToDecimal(Convert.ToDouble(sumDiscont.Value) / sumND) * 100;
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
                sumDiscont.Value = Convert.ToDecimal(Math.Round(sumND * (Convert.ToDouble(percentDiscount.Value) / 100), 0));
            }
        }

        BackgroundWorker worker;
        private void button2_Click(object sender, EventArgs e)
        {
            managerName = ManagerName.Text;
            
            saveFileDialog.Filter = "Excel files|*.xls;*.xlsx";
            var typeProductstring = from RadioButton r in ProductType.Controls where r.Checked == true select r.Text;
            var VIPstring = from RadioButton r in isVIP.Controls where r.Checked == true select r.Text;
            saveFileDialog.FileName = "Локальный сметный счет " + VIPstring.First().ToString() + " " + typeProductstring.First().ToString() + " " + ClientName.Text + " " + DateTime.Now.ToShortDateString() + ".xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                formAlert = new Alert();
                formAlert.Show();
                worker = new BackgroundWorker();
                worker.DoWork += new DoWorkEventHandler(exportExcel);
                worker.RunWorkerCompleted +=
                           new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
                worker.RunWorkerAsync();
            }
        }
        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            formAlert.Hide();
            formAlert.Close();
            MessageBox.Show("Экспорт завершен");
        }

        private void exportExcel(object sender, DoWorkEventArgs e)
        {
            //formAlert = new Alert();
            //formAlert.Refresh();
            //formAlert.Show();
            int nRow = 10;
            int s1 = 12, s2, f1, f2;
            string nameOfNewFile = Environment.CurrentDirectory + @"/template/2.xlsx";
            nameOfNewFile = saveFileDialog.FileName;
            Excel.Application excelApp;

            string fileTarget = nameOfNewFile;
            string fileTemplate = Environment.CurrentDirectory + @"/template/template1.xlsx";
            excelApp = new Excel.Application();
            Excel.Workbook wbTarget;
            Excel.Worksheet sh;

            //Create target workbook    
            wbTarget = excelApp.Workbooks.Open(fileTemplate);

            //Fill target workbook
            //Open the template sheet
            sh = wbTarget.Worksheets[1];
            sh.Name = "Коммерческое предложение";
            sh.Cells[1,2] = DateTime.Now.ToShortDateString();
            sh.Cells[4,2] = " Ваш менеджер: " + managerName;


            //Вывод первой таблицы
            var a = from RadioButton r in ProductType.Controls where r.Checked == true select r.Text;
            sh.Cells[9, 1] = "1. " + a.First();
            for (int j = 0; j < 5; j++)
                sh.Cells[nRow, j + 1] = NameOfKiln.Rows[0].Cells[j].Value.ToString();
            nRow++;
            //Вывод второй таблицы

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
            str += " (" + a.First() + ")";
            sh.Cells[nRow, 1] = str;
            nRow++;
            if (ChimneyElements.RowCount == 1)
            {
                sh.Cells[nRow, 1] = "1";
                sh.Cells[nRow, 2] = "Отсутствует";
                nRow++;
            }
            else
            for (int i = 0; i < ChimneyElements.RowCount - 1; i++)
            {
                if (i != 0)
                {
                    Excel.Range cellRange = (Excel.Range)sh.Cells[nRow, 1];
                    Excel.Range rowRange = cellRange.EntireRow;
                    rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                }
                for (int j = 0; j < 5; j++)
                {
                    sh.Cells[nRow, j + 1] = ChimneyElements.Rows[i].Cells[j].Value.ToString();
                }

                nRow++;
            }
            nRow++;
            //вывод третьей таблицы
            if (InsulationСonsumables.RowCount == 1)
            {
                sh.Cells[nRow, 1] = "1";
                sh.Cells[nRow, 2] = "Отсутствует";
                nRow++;
            }
            else
                for (int i = 0; i < InsulationСonsumables.RowCount - 1; i++)
            {
                if (i != 0)
                {
                    Excel.Range cellRange = (Excel.Range)sh.Cells[nRow, 1];
                    Excel.Range rowRange = cellRange.EntireRow;
                    rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                }
                for (int j = 0; j < 5; j++)
                    sh.Cells[nRow, j + 1] = InsulationСonsumables.Rows[i].Cells[j].Value.ToString();
                nRow++;
            }
            f1 = nRow;
            nRow++;
            sh.Cells[nRow++, 5] = resultSum1;
            double disc = Convert.ToDouble(sumDiscont.Value);
            sh.Cells[nRow++, 5] = disc;
            sh.Cells[nRow++, 5] = resultSum1 - disc;
            nRow += 2;
            s2 = nRow + 1;
            //вывод четвертой таблицы
            if (InstallationWork.RowCount==1)
            {
                sh.Cells[nRow, 1] = "1";
                sh.Cells[nRow, 2] = "Отсутствует";
                nRow++;
            }
            else
            for (int i = 0; i < InstallationWork.RowCount - 1; i++)
            {
                if (i != 0)
                {
                    Excel.Range cellRange = (Excel.Range)sh.Cells[nRow, 1];
                    Excel.Range rowRange = cellRange.EntireRow;
                    rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                }
                for (int j = 0; j < 5; j++)
                    sh.Cells[nRow, j + 1] = InstallationWork.Rows[i].Cells[j].Value.ToString();
                nRow++;
            }
            nRow++;
            //вывод пятой таблицы
            if (RiggingDelivery.RowCount == 1)
            {
                sh.Cells[nRow, 1] = "1";
                sh.Cells[nRow, 2] = "Отсутствует";
                nRow++;
            }
            else                    
            for (int i = 0; i < RiggingDelivery.RowCount - 1; i++)
            {
                if (i != 0)
                {
                    Excel.Range cellRange = (Excel.Range)sh.Cells[nRow, 1];
                    Excel.Range rowRange = cellRange.EntireRow;
                    rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                }
                for (int j = 0; j < 5; j++)
                    sh.Cells[nRow, j + 1] = RiggingDelivery.Rows[i].Cells[j].Value.ToString();
                nRow++;
            }
            f2 = nRow;
            nRow++;
            sh.Cells[nRow++, 5] = resultSum2;
            sh.Cells[nRow++, 5] = 0;
            sh.Cells[nRow++, 5] = resultSum2;
            nRow++;
            sh.Cells[nRow++, 5] = sumFinal;
            sh.Cells[nRow++, 5] = disc;
            sh.Cells[nRow++, 5] = sumFinal - disc;
                

            Excel.Range workSheet_range = sh.get_Range("A" + s1, "E" + f1);
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range = sh.get_Range("A" + s2, "E" + f2);
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();

            workSheet_range = sh.get_Range("A5", "E" + (nRow - 1));
            workSheet_range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexNone, Color.Black, Type.Missing);


            //Save file
            excelApp.DisplayAlerts = false;
            wbTarget.SaveAs(fileTarget, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            //Close and save target workbook
            wbTarget.Close(true);
            //Kill excelapp
            excelApp.Quit();
            excelApp = null;
            wbTarget = null;
            sh = null;
            GC.Collect();
        }

        private void ManagerName_TextChanged(object sender, EventArgs e)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode rootNode = xmlDoc.CreateElement("managers");
            xmlDoc.AppendChild(rootNode);

            XmlNode userNode = xmlDoc.CreateElement("managers");
            XmlAttribute attribute = xmlDoc.CreateAttribute("name");
            attribute.Value = ManagerName.Text;
            userNode.Attributes.Append(attribute);
            rootNode.AppendChild(userNode);

            xmlDoc.Save("manager.xml");

        }

        private void Form1_Click(object sender, EventArgs e)
        {
            this.Focus();

        }
    }
}
