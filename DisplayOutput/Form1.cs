using MaterialSkin;
using MaterialSkin.Controls;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace DisplayOutput
{
    public partial class Form1 : MaterialForm
    {


        public Form1()
        {
            InitializeComponent();

            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);
        }


        private void materialLabel1_Click(object sender, EventArgs e)
        {

        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            {
                string GetData()
                // generic class
                {

                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook excelWB = excelApp.Workbooks.Open(@"FlightSearchResults.xlsx");
                    Excel._Worksheet excelWS = excelWB.Sheets[1];
                    Excel.Range excelRange = excelWS.UsedRange;
                    //MessageBox.Show("Flight Date  Average Delay in Minutes\n");
                    int rowCount = excelRange.Rows.Count;
                    int columnCount = excelRange.Columns.Count;
                    string data = "";
                    for (int i = 2; i <= rowCount; i++)
                    // for loop to iterate through excel data and write it to the message box in the form
                    {

                        if (excelRange.Cells[i, 1] != null)
                        {
                            data = data + DateTime.FromOADate(Double.Parse(excelRange.Cells[i, 1].Value2.ToString())).ToString("MM/dd/yyyy\t ");

                            // this function converts the excel date storage format to a string that is logical for the user

                        }


                        if (excelRange.Cells[i, 2] != null)
                        {
                            data = data + System.Convert.ToInt32(float.Parse(excelRange.Cells[i, 2].Value2.ToString()));

                            // this function converts the function in excel that calculates the average from the large data set to a string readable by the messagebox function


                        }
                        data = data + "\n\n";

                    }

                    return data.ToString();

                    Marshal.ReleaseComObject(excelWS);
                    Marshal.ReleaseComObject(excelRange);
                    excelWB.Close();
                    Marshal.ReleaseComObject(excelWB);
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);

                }

                MessageBox.Show($"Flight Date  Average Delay in Minutes\n\n{GetData()}\t\t");



            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}

