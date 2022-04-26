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

namespace DomPecheyKP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            reader = new PdfReader("oldFile.pdf");
            sourceDocument = new Document(reader.GetPageSizeWithRotation(1));
            pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream("newFile.pdf", System.IO.FileMode.Create));

            sourceDocument.Open();

            for (int i = 1; i <= 5; i++)
            {
                importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                pdfCopyProvider.AddPage(importedPage);
            }
            sourceDocument.Close();
            reader.Close();*/

            using (var reader = new PdfReader(@"oldFile.pdf"))
            {
                using (var fileStream = new FileStream(@"newFile.pdf", FileMode.Create, FileAccess.Write))
                {
                    var document = new Document(reader.GetPageSizeWithRotation(1));
                    var writer = PdfWriter.GetInstance(document, fileStream);

                    document.Open();

                    for (var i = 1; i <= reader.NumberOfPages; i++)
                    {
                        document.NewPage();

                        var baseFont = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                        var importedPage = writer.GetImportedPage(reader, i);

                        var contentByte = writer.DirectContent;
                        contentByte.AddTemplate(importedPage, 0, 0);
                        contentByte.BeginText();
                        contentByte.SetFontAndSize(baseFont, 40);

                        var multiLineString = "Hello,\r\nWo                   rld!".Split('\n');

                        foreach (var line in multiLineString)
                        {
                            contentByte.ShowTextAligned(PdfContentByte.ALIGN_CENTER, line, 1000, 200, 0);
                        }

                        contentByte.EndText();
                        PdfPTable table = new PdfPTable(2);
                        table.WidthPercentage = 80;
                        int[] firstTablecellwidth = { 25, 75 };
                        table.SetWidths(firstTablecellwidth);
                        //Добавим в таблицу общий заголовок
                        PdfPCell cell = new PdfPCell(new Phrase("БД  таблица №", font));

                        cell.Colspan = 1;
                        cell.HorizontalAlignment = 1;
                        //Убираем границу первой ячейки, чтобы балы как заголовок
                        cell.Border = 0;
                        table.AddCell(cell);

                        //Сначала добавляем заголовки таблицы
                        for (int j = 0; j < 2; j++)
                        {
                            cell = new PdfPCell(new Phrase(new Phrase("132", font)));
                            //Фоновый цвет (необязательно, просто сделаем по красивее)
                            cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                            table.AddCell(cell);
                        }

                        //Добавляем все остальные ячейки
                        for (int j = 0; j < 1; j++)
                        {
                            for (int k = 0; k < 1; k++)
                            {
                                table.AddCell(new Phrase("111", font));
                            }
                        }
                        //Добавляем таблицу в документ
                        document.Add(table);

                    }

                    document.Close();
                    writer.Close();
                }
            }
        }
    }
}
