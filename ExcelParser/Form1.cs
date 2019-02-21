using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace ExcelParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string _fName = "";
        private string _savePath = "";
        private List<string> _texts;
        
        private void SelectSavePath(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                var result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    _savePath = fbd.SelectedPath;
                }
            }
        }

        private void SelectTextFile(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                var file = openFileDialog1.FileName;
                try
                {
                    if (openFileDialog1.SafeFileName != null) _fName = openFileDialog1.SafeFileName.Replace(".txt", "");
                    var textContent = File.ReadAllText(file);
                    if (!string.IsNullOrEmpty(textContent))
                    {
                        _texts = new List<string>();
                        var stringSeparators = new[] { "\r\n" };
                        _texts = textContent.Split(stringSeparators, StringSplitOptions.None).ToList();
                    }
                }
                catch (IOException)
                {
                }
            }
        }

        private void SelectExcelFile(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(_savePath))
            {
                MessageBox.Show(@"Save Path Error", @"Please select file save directory");
                return;
            }
            var result = openFileDialog2.ShowDialog();
            if (result == DialogResult.OK)
            {
                var file = openFileDialog2.FileName;
                try
                {
                    using (var fileStream = new FileStream(file, FileMode.Open))
                    {
                        var excel = new ExcelPackage(fileStream);
                        var workSheet = excel.Workbook.Worksheets[1];

                        var workSheetName = excel.Workbook.Worksheets[1].Name;

                        var start = workSheet.Dimension.Start;
                        var end = workSheet.Dimension.End;

                        var columns = new List<string>();

                        for (var col = start.Column; col <= end.Column; col++)
                        {
                            var cellValue = workSheet.Cells[1, col].Text;  
                            columns.Add(cellValue);
                        }
                        

                        var lists = new List<ItemCol>();

                        for (var i = 2; i < 7; i++)
                        {
                            lists.Add(new ItemCol
                            {
                                LABEL = workSheet.Cells[i, 1].Text,
                                PURPOSE = workSheet.Cells[i, 2].Text,
                                BARCODE_ID = workSheet.Cells[i, 3].Text,
                                BARCODE_ID2 = workSheet.Cells[i, 4].Text,
                            });
                        }
                          
                        var outL = new List<ItemCol>();

                        if (lists.Any() && _texts.Any())
                        {

                            _texts.ForEach(t =>
                            {
                                if (!string.IsNullOrEmpty(t))
                                {
                                    var li0 = new ItemCol
                                    {
                                        LABEL = lists[0].LABEL,
                                        PURPOSE = lists[0].PURPOSE,
                                        BARCODE_ID = lists[0].BARCODE_ID,
                                        BARCODE_ID2 = t
                                    };

                                    var li1 = new ItemCol
                                    {
                                        LABEL = lists[1].LABEL,
                                        PURPOSE = lists[1].PURPOSE,
                                        BARCODE_ID = t,
                                        BARCODE_ID2 = lists[1].BARCODE_ID2
                                    };

                                    var li2 = new ItemCol
                                    {
                                        LABEL = lists[2].LABEL,
                                        PURPOSE = lists[2].PURPOSE,
                                        BARCODE_ID = t,
                                        BARCODE_ID2 = lists[2].BARCODE_ID2
                                    };

                                    var li3 = new ItemCol
                                    {
                                        LABEL = lists[3].LABEL,
                                        PURPOSE = lists[3].PURPOSE,
                                        BARCODE_ID = t,
                                        BARCODE_ID2 = lists[3].BARCODE_ID2
                                    };

                                    var li4 = new ItemCol
                                    {
                                        LABEL = lists[4].LABEL,
                                        PURPOSE = lists[4].PURPOSE,
                                        BARCODE_ID = t,
                                        BARCODE_ID2 = lists[4].BARCODE_ID2
                                    };

                                    outL.Add(li0);
                                    outL.Add(li1);
                                    outL.Add(li2);
                                    outL.Add(li3);
                                    outL.Add(li4);
                                }
                            });
                        }

                        if (outL.Any())
                        {
                            if (!_savePath.EndsWith("\\"))
                            {
                                _savePath = _savePath + @"\";
                            }

                            var fi = new FileInfo(_savePath + _fName + ".xlsx");

                            using (var p = new ExcelPackage(fi))
                            {
                                var ws = p.Workbook.Worksheets.Add(workSheetName);
                               
                                ws.Cells[1, 1].Value = columns[0];
                                ws.Cells[1, 2].Value = columns[1];
                                ws.Cells[1, 3].Value = columns[2];
                                ws.Cells[1, 4].Value = columns[3];
                                ws.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                ws.Row(1).Style.Font.Bold = true;
                                ws.Row(1).Height = 11;
                                ws.Row(1).Style.Font.Size = 10;
                            
                                var recordIndex = 2;
                                outL.ForEach(f =>
                                {
                                    ws.Cells[recordIndex, 1].Value = f.LABEL;
                                    ws.Cells[recordIndex, 2].Value = f.PURPOSE;
                                    ws.Cells[recordIndex, 3].Value = f.BARCODE_ID;
                                    ws.Cells[recordIndex, 4].Value = f.BARCODE_ID2;
                                    ws.Row(recordIndex).Height = 11;
                                    ws.Row(recordIndex).Style.Font.Size = 10;
                                    ws.Row(recordIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    recordIndex++;
                                });

                                ws.Column(1).Width = 12;
                                ws.Column(2).Width = 18;
                                ws.Column(3).Width = 15;
                                ws.Column(4).Width = 15;
                                p.Save();
                            }
                        }

                    }
                }
                catch (Exception ex)
                {

                }
            }

        }

        public class ItemCol
        {
            public string LABEL { get; set; }
            public string PURPOSE { get; set; }
            public string BARCODE_ID { get; set; }
            public string BARCODE_ID2 { get; set; }

        }

    }
}
