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
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
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

       

        private void Form1_Load(object sender, EventArgs e)
        {
            loadElD(2);
            NameOfKiln.Rows.Add("1","","1","0","0");
        }

        private void loadElD(int nColumn)
        {
            
            NewChimneyElements.Items.Clear();            
            ChimneyElements.Rows.Clear();

            list  = new Dictionary<string, double>();
            Excel.Application ObjWorkExcel = new Excel.Application(); 
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(Environment.CurrentDirectory + @"\ДанныеКП.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку          
            for (int i = 3; i < (int)lastCell.Row; i++) // по всем строкам
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
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit(); 
           
        }

        private void calculateChimneyElementsSum()
        {
            double sum = 0;
            foreach (DataGridViewRow row in ChimneyElements.Rows)
                sum += Convert.ToDouble(row.Cells[4].Value);
            SumChimneyElements.Text = sum.ToString() + " Руб.";
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
               ChimneyElements.Rows.Add(n++, checkedItem.ToString(), "1",list[checkedItem.ToString()], list[checkedItem.ToString()]);
            }
            while( NewChimneyElements.CheckedItems.Count!=0)
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

            for (int i = 0; i < ChimneyElements.RowCount-1; i++)
            {
                ChimneyElements.Rows[i].Cells[0].Value=i+1;                
            }
            calculateChimneyElementsSum();
        }

        private void ChimneyElements_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            
            for (int i = 0; i < ChimneyElements.RowCount-1; i++)
            {
                ChimneyElements.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void ChimneyElements_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int nRow = e.RowIndex;
            if(e.ColumnIndex==2 || e.ColumnIndex == 3)
            {

                ChimneyElements.Rows[nRow].Cells[4].Value = Convert.ToDouble(ChimneyElements.Rows[nRow].Cells[2].Value) * Convert.ToDouble(ChimneyElements.Rows[nRow].Cells[3].Value);
                calculateChimneyElementsSum();
            }
            
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
                    if (!Regex.IsMatch(editingControl.Text + e.KeyChar, @"^[0-9,.]{0,10}$"))
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
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
                NameOfKiln.Rows[nRow].Cells[4].Value = Convert.ToDouble(NameOfKiln.Rows[nRow].Cells[2].Value) * Convert.ToDouble(NameOfKiln.Rows[nRow].Cells[3].Value);
                
        }

        private void ChimneyElements_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            currentDataGridView = ChimneyElements;
            ChimneyElements.EditingControl.KeyPress -= EditingControl_KeyPress;
            ChimneyElements.EditingControl.KeyPress += EditingControl_KeyPress;
        }
    }
}
