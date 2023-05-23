using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelDemo.Helper
{
    public static class ExcelFileReader
    {
        public static System.Data.DataTable read(Range excelRange)
        {
            DataRow row;
            System.Data.DataTable dt = new System.Data.DataTable();
            int rowCount = excelRange.Rows.Count; //get row count of excel data

            int colCount = excelRange.Columns.Count; // get column count of excel data

            //Get the first Column of excel file which is the Column Name

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                }
                break;
            }

            //Get Row Data of Excel

            int rowCounter; //This variable is used for row index number
            for (int i = 2; i <= rowCount; i++) //Loop for available row of excel data
            {
                row = dt.NewRow(); //assign new row to DataTable
                rowCounter = 0;
                for (int j = 1; j <= colCount; j++) //Loop for available column of excel data
                {
                    //check if cell is empty
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {
                        row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                    }
                    else
                    {
                        row[i] = "";
                    }
                    rowCounter++;
                }
                dt.Rows.Add(row); //add row to DataTable
            }

            return dt;
        }
    }
}
