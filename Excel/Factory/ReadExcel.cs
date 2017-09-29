using Excel.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Excel.Factory
{
    public class ReadExcel
    {
        public static List<Dish> WorkReadExcel(string p_Path)
        {
            string l_filePath = "";//檔案路徑
            string l_fileCollection = Directory.GetFiles(p_Path, "檔名" + ".xls").First();//取檔案
            List<Dish> DishList = new List<Dish>();

            FileInfo l_fileName = new FileInfo(l_fileCollection);
            string l_path = l_filePath + @"\" + l_fileName.Name;
            IWorkbook wk = oldOrNewExcel(l_path, l_fileName);
            IRow hr; //取得SHEET中的某個ROW
            ISheet l_sheet = wk.GetSheet("Department");//讀取Sheet1 工作表
            for (int x = 1; x <= l_sheet.LastRowNum; x++)//逐步跑每個sheet的內容轉成DataTable
            {
                if (l_sheet.GetRow(x) != null)
                {
                    hr = l_sheet.GetRow(x);
                    if (!string.IsNullOrEmpty(hr.GetCell(0).ToString()))
                    {
                        Dish l_dish = new Dish();
                        l_dish.Number = Convert.ToInt32(hr.GetCell(0) == null ? "" :  hr.GetCell(0).ToString());//編號
                        l_dish.Name = hr.GetCell(1) == null ? "" :  hr.GetCell(1).ToString();//名稱
                        l_dish.Category = hr.GetCell(2) == null ? "" :  hr.GetCell(2).ToString();//種類

                        DishList.Add(l_dish);
                    }
                }
            }
            return DishList;
        }

        public static IWorkbook oldOrNewExcel(string p_path, FileInfo p_fileName)
        {
            IWorkbook wk;
            using (FileStream fs = new FileStream(p_path, FileMode.Open, FileAccess.ReadWrite))//fs為stream 物件
            {
                if (p_fileName.Extension.Equals(".xls"))
                {
                    wk = new HSSFWorkbook(fs);//舊版EXCEL
                }
                else if (p_fileName.Extension.Equals(".xlsx"))
                {
                    wk = new XSSFWorkbook(fs);//新版EXCEL
                }
                else
                {
                    Exception ex = new Exception("副檔名錯誤");
                    throw ex;
                }
            }
            return wk;
        }
    }
}