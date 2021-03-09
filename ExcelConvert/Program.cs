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
            Application ObjWorkExcelFirst = new Application(); //открыть эксельC:\Users\user\source\repos\ExcelConvert\ExcelConvert\bin\Debug
            Workbook ObjWorkBookFirst = ObjWorkExcelFirst.Workbooks.Open(System.IO.Directory.GetCurrentDirectory()+@"\file1.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
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



            Application ObjWorkExcelSecond = new Application(); //открыть эксель
            Workbook ObjWorkBookSecond = ObjWorkExcelSecond.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\file2.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Worksheet ObjWorkSheetSecond = (Worksheet)ObjWorkBookSecond.Sheets[1]; //получить 1 лист
            var lastCellSecond = ObjWorkSheetSecond.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);//1 ячейку

            List<List<string>> listSecond = new List<List<string>>();
            for (int i = 1; i <= (int)lastCellSecond.Row; i++)
            {
                List<string> temp = new List<string>();
                for (int j = 1; j <= (int)lastCellSecond.Column; j++)
                {

                    temp.Add(ObjWorkSheetSecond.Cells[i, j].Text.ToString());
                }
                listSecond.Add(temp);
            }
            ObjWorkBookSecond.Close(true, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcelSecond.Quit(); // выйти из экселя

            GC.Collect(); // убрать за собой


            foreach (var itemFirst in listFirs)
            {
                foreach(var itemSecond in listSecond)
                {
                    if (itemFirst[0] == itemSecond[0] && itemFirst[2] == itemSecond[2])
                    {
                        listSecond.Remove(itemSecond);
                        break;
                    }
                }
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("list1");
                var currentRow = 0;
                foreach (var item in listSecond)
                {
                    currentRow++;
                    for(int i = 0; i< item.Count;i++)
                    {
                        worksheet.Cell(currentRow, i+1).Value = item[i];
                    }
                }
                workbook.SaveAs($"Result.xlsx");
            }
        }
    }
}
