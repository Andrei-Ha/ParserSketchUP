using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using NativeExcel;

namespace ParserSketchUP
{
    class Panel
    {
        public string Name { get; set; }
        public string Name2 { get; set; }
        public int Length { get; set; }
        public string Lkrom { get; set; }
        public int Width { get; set; }
        public string Wkrom { get; set; }
        public string Brutto { get; set; }
        public string Netto {
            get => Length.ToString() + "x" + Width.ToString();
        }
        public int quantity { get; set; }
        public Panel(string n,string n2, int l, string lk, int w, string wk, string b)
        {
            Name = n;
            Name2 = n2;
            Length = l;
            Lkrom = lk;
            Width = w;
            Wkrom = wk;
            Brutto = b;
            quantity = 1;
        }
    }
    public partial class Form1 : Form
    {
        private int i = 1;
        private FileInfo file = null;
        private string path_file = string.Empty;
        private List<Panel> details= new List<Panel>();
        private string path_repfolder = string.Empty;
        private string fileName = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                path_file = ofd.FileName;
                file = new FileInfo(path_file);
                label1.Text = path_file;
                DirectoryInfo dir_file = file.Directory;
                path_repfolder = dir_file + @"\reports";
                fileName = file.Name.Split('.')[0];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (TextFieldParser parser = new TextFieldParser(path_file))
            {
                List<string> lines = new List<string>();
                string temp_Name = string.Empty, temp_Description = string.Empty,line = string.Empty;
                while (!parser.EndOfData)
                {
                    line = parser.ReadLine();
                    string[] a = new string[6];
                    //if (temp_string2.Count<char>(p => p == '!') == 5)
                    if (line.IndexOf("fiberboard") != -1 || line.IndexOf("panel") != -1) {
                        temp_Name = line.Split(',')[0].Trim('"');
                        temp_Description = line.Split(',')[1].Trim('"');
                        a = temp_Description.Split('!');
                        if (line.IndexOf("fiberboard") != -1)
                            details.Add(new Panel(temp_Name, a[0], Convert.ToInt16(a[1]), "", Convert.ToInt16(a[2]), "", $"({a[1]}x{a[2]})"));
                        else
                            details.Add(new Panel(temp_Name, a[0], Convert.ToInt16(a[1]), a[2], Convert.ToInt16(a[3]), a[4], a[5]));
                    }
                    
                }
                details = details.OrderBy(e => e.Length).ThenBy(e => e.Width).ToList<Panel>();
                // create .xls
                IWorkbook workbook = NativeExcel.Factory.CreateWorkbook();
                IWorksheet w1 = workbook.Worksheets.Add();
                IWorksheet w2 = workbook.Worksheets.Add();
                IWorksheet w3 = workbook.Worksheets.Add();
                IWorksheet w4 = workbook.Worksheets.Add();
                w1.Name = "My_specification";
                w2.Name = "specification";
                w3.Name = "Cutting3";
                w4.Name = "Fiberboard";
                //header My_specification
                string[] h1 = new string[] { "№","название","длинна, мм.", "кромка", "ширина, мм.", "кромка", "с учетом кромки" };
                IRange rh1 = w1.Cells[2, 1, 2, h1.Length];
                SheetFormat(rh1, w1, h1);
                List<Panel> panels = details.Where(e=>e.Name2!="fiberboard").ToList<Panel>();
                int y = 1;
                foreach (Panel panel in panels)
                {
                    label_result.Text += panel.Netto + " - " + panel.Brutto + " q: " + panel.quantity + "\n";
                    w1.Cells[y + 2, 1].Value = y+".";
                    w1.Cells[y + 2, 2].Value = panel.Name + $"({panel.Name2})";
                    w1.Cells[y + 2, 3].Value = panel.Length;
                    w1.Cells[y + 2, 4].Value = panel.Lkrom == "" ? 0 : Convert.ToInt16(panel.Lkrom);
                    w1.Cells[y + 2, 5].Value = panel.Width;
                    w1.Cells[y + 2, 6].Value = panel.Wkrom == "" ? 0 : Convert.ToInt16(panel.Wkrom);
                    w1.Cells[y + 2, 7].Value = panel.Brutto;
                    y++;
                }
                w1.Cells.Columns[1, 7].Autofit(); ;
                //header specification
                string[] h2 = new string[] { "№", "длинна, мм.", "кромка", "ширина, мм.", "кромка", "с учетом кромки","количество" };
                IRange rh2 = w2.Cells[2, 1, 2, h2.Length];
                SheetFormat(rh2, w2, h2);
                List<Panel> panels2 = panels.GroupBy(e => e.Netto + e.Brutto).Select(p => { var i = p.First();i.quantity = p.Count(); return i; }).ToList<Panel>();
                y = 1;
                foreach (Panel panel2 in panels2)
                {
                    label_result2.Text += panel2.Netto + " - " + panel2.Brutto + " q: " + panel2.quantity +  "\n";
                    w2.Cells[y + 2, 1].Value = y + ".";
                    w2.Cells[y + 2, 2].Value = panel2.Length;
                    w2.Cells[y + 2, 3].Value = panel2.Lkrom == "" ? 0 : Convert.ToInt16(panel2.Lkrom);
                    w2.Cells[y + 2, 4].Value = panel2.Width;
                    w2.Cells[y + 2, 5].Value = panel2.Wkrom == "" ? 0 : Convert.ToInt16(panel2.Wkrom);
                    w2.Cells[y + 2, 6].Value = panel2.Brutto;
                    w2.Cells[y + 2, 7].Value = panel2.quantity;
                    y++;
                }
                //header Cutting3
                string[] h3 = new string[] { "const", "длинна, мм.", "ширина, мм.", "количество" };
                IRange rh3 = w3.Cells[2, 1, 2, h3.Length];
                SheetFormat(rh3, w3, h3);
                List<Panel> panels3 = panels.GroupBy(e => e.Netto).Select(p => { var i = p.First(); i.quantity = p.Count(); return i; }).ToList<Panel>();
                y = 1;
                foreach (Panel panel3 in panels3)
                {
                    w3.Cells[y + 2, 1].Value = 7;
                    w3.Cells[y + 2, 2].Value = panel3.Length;
                    w3.Cells[y + 2, 3].Value = panel3.Width;
                    w3.Cells[y + 2, 4].Value = panel3.quantity;
                    y++;
                }
                //header Fiberboard
                string[] h4 = new string[] { "№", "название", "длинна, мм.", "ширина, мм.","количество" };
                IRange rh4 = w4.Cells[2, 1, 2, h4.Length];
                SheetFormat(rh4, w4, h4);
                List<Panel> fiberboards = details.Where(e => e.Name2 == "fiberboard").GroupBy(o => o.Netto).Select(o => { var r = o.First(); r.quantity = o.Count(); return r; }).ToList<Panel>();
                y = 1;
                foreach (Panel panel in fiberboards)
                {
                    w4.Cells[y + 2, 1].Value = y + ".";
                    w4.Cells[y + 2, 2].Value = panel.Name + $"({panel.Name2})";
                    w4.Cells[y + 2, 3].Value = panel.Length;
                    w4.Cells[y + 2, 4].Value = panel.Width;
                    w4.Cells[y + 2, 5].Value = panel.quantity;
                    y++;
                }
                // creating folder Reports to save reports.xlx
                if (!Directory.Exists(path_repfolder))
                    Directory.CreateDirectory(path_repfolder);
                // save report                
                workbook.SaveAs(path_repfolder + @"\" + fileName + ".xls");
            }
        }
        private void SheetFormat(IRange range, IWorksheet worksheet, string[] h)
        {
            for (int i = 0; i < h.Length; i++)
                worksheet.Cells[2, i + 1].Value = h[i];
            range.Font.Bold = true;
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            range.Borders[XlBordersIndex.xlAround].ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            range.Borders[XlBordersIndex.xlInsideAll].ColorIndex = XlColorIndex.xlColorIndexAutomatic;
            worksheet.Cells.Columns[1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Cells.Columns[1, 7].Autofit();
        }
    }
}

