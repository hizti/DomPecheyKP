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
            int countColumns = 2;
            string nameOfNewFile = @"newFile.pdf";
            //List<float> heightsRows = new List<float>();
            PdfPTable fTable = new PdfPTable(countColumns);

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
                    var fontName = "Sitka Banner";
                    if (!FontFactory.IsRegistered(fontName))
                    {
                        var fontPath = Environment.GetEnvironmentVariable("Sitka-Banner.ttf");
                        FontFactory.Register("Sitka-Banner.ttf");
                    }
                    iTextSharp.text.Font font = FontFactory.GetFont(fontName, BaseFont.IDENTITY_H);
                    font.Color = iTextSharp.text.BaseColor.WHITE;
                    font.Size = 35;
                    
                    fTable.WidthPercentage = 80;
                    int[] firstTablecellwidth = { 25, 75 };
                    fTable.SetWidths(firstTablecellwidth);
                    //Добавим в таблицу общий заголовок
                    PdfPCell cell;

                    float sum = 0;
                    int nRow = 0;
                    string str = "рррррр ррррррррррррррррррррррррррррр";
                    for (int j = 0; j < 10; j++)
                    {
                        for (int k = 0; k < 2; k++)
                        {

                            cell = new PdfPCell(new Phrase(new Phrase(str, font)));
                            str += " 123";
                            fTable.AddCell(cell);
                        }
                        /*sum += fTable.GetRowHeight(j);
                        nRow++;
                        if (sum > 600)
                        {
                            Paragraph p1 = new Paragraph("     ");
                            p1.SpacingAfter = 100;
                            document.Add(p1);
                            fTable.SpacingBefore = 100;
                            document.Add(fTable);

                            document.NewPage();
                            importedPage = writer.GetImportedPage(reader, 1);
                            contentByte = writer.DirectContent;
                            contentByte.AddTemplate(importedPage, 0, 0);

                            fTable = new PdfPTable(2);
                            fTable.WidthPercentage = 80;
                            fTable.SetWidths(firstTablecellwidth);

                            fTable.DeleteRow(nRow);
                            sum = 0;
                            nRow = 0;
                            j--;
                        }*/

                    }
                    document.Add(fTable);
                    /*for (int i = 0; i < fTable.Rows.Count; i++)
                    {
                        heightsRows.Add(fTable.Rows[i].MaxHeights);
                    }*/
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

                        if (i == 1)
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

                        if (i == 10)
                        {
                            int maxHeightTable = 800;
                            Paragraph p;

                            PdfPTable sTable = new PdfPTable(countColumns);
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

        private void Diameter_Enter(object sender, EventArgs e)
        {

        }

        private void SumChimneyElements_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
