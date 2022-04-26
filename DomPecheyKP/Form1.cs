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

                        if (i == 1)
                        {
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
                            PdfPTable table = new PdfPTable(2);
                            table.WidthPercentage = 80;
                            int[] firstTablecellwidth = { 25, 75 };
                            table.SetWidths(firstTablecellwidth);
                            //Добавим в таблицу общий заголовок
                            PdfPCell cell;

                            //Сначала добавляем заголовки таблицы
                            

                            for (int j = 0; j < 10; j++)
                            {
                                for (int k = 0; k < 2; k++)
                                {
                                    cell = new PdfPCell(new Phrase(new Phrase("   ", font)));
                                    table.AddCell(cell);
                                }
                                
                            }
                            Paragraph p = new Paragraph("1233");
                            p.SpacingAfter = 100;
                            document.Add(p);
                            table.SpacingBefore=100;
                            document.Add(table);
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
