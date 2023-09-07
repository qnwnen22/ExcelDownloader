using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDownloader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.testToolStripMenuItem.Enabled = false;
            var json2 = File.ReadAllText(@"D:\product.txt");
            var excelConvertor = new ExcelConvertor();
            excelConvertor.DownloadExcel(json2);
        }

        private void test2ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private Task DownloadExcel(string json)
        {
            var resultAsync = new Task(() =>
            {
                var excelConvertor = new ExcelConvertor();
                excelConvertor.DownloadExcel(json);
                excelConvertor = null;
                this.Invoke(new System.Action(() =>
                {
                    this.testToolStripMenuItem.Enabled = true;
                    MessageBox.Show("작업이 완료되었습니다.");
                }));
            });
            return resultAsync;
        }
    }


    public class ExcelConvertor
    {
        private class Header
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public string Value { get; set; }
            public List<Header> Headers { get; set; } = new List<Header>();
        }

        private int width = 0;
        private int height = 0;
        private int nextHeight = 0;

        private Excel.Application excelApp = null;
        private Excel.Workbook wb = null;
        private Excel.Worksheet ws = null;
        private List<Header> Headers { get; set; } = new List<Header>();

        private int GetMaxDepth(Header parent, JToken jToken, int depth = 1)
        {
            int result = depth;
            switch (jToken.Type)
            {
                case JTokenType.Object:
                    result++;
                    JObject jObject = JObject.Parse(jToken.ToString());
                    List<JProperty> properties = jObject.Properties().ToList();

                    for (int i = 0; i < properties.Count; i++)
                    {
                        JProperty property = properties[i];

                        if (property.Value.HasValues)
                        {
                            Header header = Add(property.Name, result, width);
                            var getDepth1 = GetMaxDepth(header, property.Value, result);
                            parent.Headers.Add(header);
                            width++;

                            if (result < getDepth1)
                            {
                                result = getDepth1;
                            }
                        }
                        else
                        {
                            var getDepth1 = GetMaxDepth(parent, property.Value, result);
                            Header header = Add(property.Name, getDepth1, width);
                            parent.Headers.Add(header);
                            width++;

                            if (result < getDepth1)
                            {
                                result = getDepth1;
                            }
                        }

                        if (properties.Count - 1 == i)
                        {
                            width--;
                        }

                    }
                    break;
                case JTokenType.Array:
                    int currentDepth = result;
                    JToken[] array = jToken.ToArray();
                    if (array != null && array.Length > 0)
                    {
                        JToken first = array.First();
                        var getDepth2 = GetMaxDepth(parent, first, currentDepth);
                        if (result < getDepth2)
                        {
                            result = getDepth2;
                        }
                    }
                    break;
                default:
                    break;
            }
            return result;
        }
        private Header Add(string value, int row, int column)
        {
            ws.Cells[row, column] = value;
            return new Header
            {
                Value = value,
                Row = row,
                Column = column
            };
        }
        private int Merge(Header header, int maxDepth, int column = 0)
        {
            Range range = null;
            if (header.Headers.Count == 0)
            {
                range = ws.Range[ws.Cells[header.Row, header.Column], ws.Cells[maxDepth, header.Column]];
            }
            else
            {
                var newColumn = (header.Column + header.Headers.Count - 1);
                if (column < newColumn)
                {
                    column = newColumn;
                }

                for (int i = 0; i < header.Headers.Count; i++)
                {
                    Header child = header.Headers[i];
                    var addColumn = Merge(child, maxDepth, newColumn);

                    if (newColumn < addColumn)
                    {
                        newColumn = addColumn;
                    }
                }
                range = ws.Range[ws.Cells[header.Row, header.Column], ws.Cells[header.Row, newColumn]];
            }
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.Merge();
            return column;
        }
        public void WriteHeader(JObject jObject)
        {
            List<JProperty> properties = jObject.Properties().ToList();
            int maxDepth = 0;

            for (int i = 0; i < properties.Count; i++)
            {
                JProperty property = properties[i];
                width++;

                JToken jToken = property.Value;
                Header header = Add(property.Name, 1, width);
                Headers.Add(header);

                var depth = 1;
                if (jToken.HasValues)
                {
                    depth = GetMaxDepth(header, jToken);
                    if (maxDepth < depth)
                    {
                        maxDepth = depth;
                    }
                }

            }

            foreach (Header item in Headers)
            {
                Merge(item, maxDepth);
            }

            height = maxDepth + 1;
        }

        public int WriteBody(JToken jToken, int row, int column)
        {
            int result = 0;

            switch (jToken.Type)
            {
                case JTokenType.Object:
                    JObject jobj = (JObject)jToken;
                    List<JProperty> jProperties = jobj.Properties().ToList();
                    int appandColumn = 0;
                    for (int i = 0; i < jProperties.Count; i++)
                    {
                        JProperty property = jProperties[i];
                        appandColumn = column + i;
                        appandColumn += WriteBody(property.Value, row, appandColumn);
                    }
                    result = appandColumn - column;
                    break;
                case JTokenType.Property:

                    break;
                case JTokenType.Array:
                    List<JToken> values = jToken.Children().ToList();
                    for (int i = 0; i < values.Count; i++)
                    {
                        int addRow = (height + i);
                        if (nextHeight < addRow)
                        {
                            nextHeight = addRow;
                        }
                        JToken value = values[i];
                        result = WriteBody(value, addRow, column);
                    }
                    break;
                default:
                    ws.Cells[row, column] = jToken.ToString(); ;
                    break;
            }
            return result;
        }

        public void DownloadExcel(string json)
        {
            excelApp = new Excel.Application();
            wb = excelApp.Workbooks.Add();
            ws = wb.Worksheets.get_Item(1) as Worksheet;

            var token = JToken.Parse(json);

            try
            {
                if (token is JArray)
                {
                    DownloadExcel(JArray.Parse(json));
                }
                else
                {
                    var jObject = JObject.Parse(json);
                    WriteHeader(jObject);
                    DownloadExcel(jObject);
                }

                string save = Directory.GetCurrentDirectory() + "\\test.xls";
                // 엑셀파일 저장
                wb.SaveAs(save, Excel.XlFileFormat.xlWorkbookNormal);
                wb.Close(true);
                excelApp.Quit();
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception.Message);
            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
                GC.Collect();
            }
        }
        private void DownloadExcel(JArray jArray)
        {
            Console.WriteLine("DownloadExcel Start");
            for (int i = 0; i < jArray.Count; i++)
            {
                Console.WriteLine(i);
                JToken jToken = jArray[i];
                var jObject = (JObject)jToken;
                if (i == 0)
                {
                    Console.WriteLine("WriteHeader Start");
                    WriteHeader(jObject);
                }
                
                DownloadExcel(jObject);
                height = nextHeight + 1;
            }
        }


        private void DownloadExcel(JObject jObject)
        {
            List<JToken> tokens = jObject.Values().ToList();
            int addValue = 1;
            for (int i = 0; i < tokens.Count; i++)
            {
                Console.WriteLine($"WriteBody [{i}] Start");
                int res = WriteBody(tokens[i], height, i + addValue);
                addValue += res;
            }
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
        }
    }

}
