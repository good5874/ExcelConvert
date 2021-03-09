using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConvert
{
    class Program
    {
        static void Main(string[] args)
        {
            Start();
        }

        public async static void Start()
        {
            List<List<string>> listFirs = await FilOpen("file1.xlsx");
            List<List<string>> listSecond = await FilOpen("file2.xlsx");

            foreach (var itemFirst in listFirs)
            {
                await Search(itemFirst, listSecond);
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("list1");
                var currentRow = 0;
                foreach (var item in listSecond)
                {
                    currentRow++;
                    for (int i = 0; i < item.Count; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = item[i];
                    }
                }
                workbook.SaveAs($"Result.xlsx");
            }
        }

        public static async Task Search(List<string> itemFirst, List<List<string>> listSecond)
        {
            var a = itemFirst[2].Split('.');
            foreach (var itemSecond in listSecond)
            {
                var b = itemSecond[2].Split('.');
                if (itemFirst[0] == itemSecond[0] && a[0] == b[0] && a[1] == b[1])
                {
                    listSecond.Remove(itemSecond);
                    break;
                }
            }
        }


        public static async Task<List<List<string>>> FilOpen(string file)
        {
            Application ObjWorkExcelFirst = new Application();
            Workbook ObjWorkBookFirst = ObjWorkExcelFirst.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + $@"\{file}", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Worksheet ObjWorkSheetFirst = (Worksheet)ObjWorkBookFirst.Sheets[1]; //получить 1 лист
            var lastCellFirs = ObjWorkSheetFirst.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);//1 ячейку
            List<List<string>> listFirs = new List<List<string>>();

            for (int i = 1; i <= (int)lastCellFirs.Row; i++)
            {
                List<string> temp = new List<string>();
                for (int j = 1; j <= (int)lastCellFirs.Column; j++)
                {

                    temp.Add(ObjWorkSheetFirst.Cells[i, j].Text.ToString());
                }
                listFirs.Add(temp);
            }
            ObjWorkBookFirst.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcelFirst.Quit(); // выйти из экселя
            GC.Collect();
            return listFirs;
        }
    }
}
