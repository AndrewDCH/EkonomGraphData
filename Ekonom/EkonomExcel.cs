using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Ekonom
{
    class EkonomExcel
    {
        private Excel.Application excel;

        private Excel.Workbook workbook;

        private Excel.Worksheet sheet;

        private List<double> data = new List<double>();

        private int row = 2;

        private int counter;

        public EkonomExcel()
        {
            ExcelManager();
            UserInterface();
        }

        private void GetResults()
        {
            counter=0;
            double sum = 0;
            Console.Write("price = ");
            double price = Convert.ToDouble(Console.ReadLine());
            double t = 0.03;
            Console.Write("power = ");
            double power = Convert.ToDouble(Console.ReadLine());
            double time = 1000;
            double i = 0.05;
            Console.Write("n = ");
            int n = Convert.ToInt32(Console.ReadLine());
            while (i <= 0.35)
            {
                for (int j = 1; j < n; j++)
                {
                    if (j <= 10)
                    {
                        sum += (0 + power * time * t) / Math.Pow(1 + i, j);
                    }
                    else
                    {
                        sum += (price + power * time * t) / Math.Pow(1 + i, j);
                    }
                }
                sum += price;
                data.Add(sum);
                i += 0.01;
                sum = 0;
                counter++;
                
            }
            ExcelFiller(counter);
            row++;
            data.Clear();
        }

        private void ExcelManager()
        {
            excel = new Excel.Application();
            excel.Visible = true;
            excel.SheetsInNewWorkbook = 1;
            workbook = excel.Workbooks.Add(Type.Missing);
            excel.DisplayAlerts = false;
            sheet = (Excel.Worksheet)excel.Worksheets[1];
            sheet.Name = "Data";
        }

        private void ExcelFiller(int n)
        {
            for(int j = 1; j < n; j++)
            {
                sheet.Cells[row, j] = data[j - 1].ToString().Substring(0,5);
                Console.WriteLine(data[j - 1].ToString());
            }
        }

        private void MakeSpaceBetweenData()
        {
            row++;
        }

        private void Cancel()
        {
            row--;
        }

        private void UserInterface()
        {
            int ind = 0;
            while(true)
            {
                if (ind == 0)
                {
                    GetResults();
                    Console.Write("Make your choice (0=continue, 1=save and exit(D:\\ - save directory)) 2=make space between data");
                    ind = Convert.ToInt32(Console.ReadLine());
                }
                if(ind == 1)
                {
                    MakeSpaceBetweenData();
                    Console.Write("Make your choice (0=continue,1=make space between data 2=cancel 3=save and exit(D:\\ - save directory)) ");
                    ind = Convert.ToInt32(Console.ReadLine());
                }
                if(ind == 2)
                {
                    Cancel();
                    Console.Write("Make your choice (0=continue,1=make space between data 2=cancel 3=save and exit(D:\\ - save directory)) ");
                    ind = Convert.ToInt32(Console.ReadLine());
                }
                if(ind == 3)
                {
                    break;
                }
            }
            excel.Application.ActiveWorkbook.SaveAs(@"D:\Data.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

    }
}
