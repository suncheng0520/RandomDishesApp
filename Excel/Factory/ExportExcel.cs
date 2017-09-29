using Aspose.Cells;
using Excel.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Factory
{
    public class ExportExcel
    {
        public void processExcel(string p_path, List<Dish> p_class, DateTime p_selectedDate, bool p_forLaborInspectionCheck)
        {
            if (!Directory.Exists(p_path))
            {
                //新增資料夾
                Directory.CreateDirectory(p_path);
            }

            Workbook l_excel = Ede.Uof.Utility.ExcelHelper.GetAsposeWorkbook();

            DataTable l_exportDT = ToDataTable(p_class);
            var l_formatDate = p_selectedDate.Year.ToString() + "年" + p_selectedDate.Month.ToString() + "月";
            //匯入欄位名稱
           // string[] l_NameAndTime = { "姓名：", p_class[0].NAME, "", "", "", l_formatDate };
            if (p_forLaborInspectionCheck)
            {
              //  l_NameAndTime = new string[] { "姓名：", p_class[0].NAME, "", "", "", l_formatDate, "出勤紀錄表" };
            }
           // l_excel.Worksheets[0].Cells.ImportArray(l_NameAndTime, 0, 0, false);
            string[] getFieldName = { "日期", "工作內容", "實際加班(起)", "實際加班(訖)", "應付時數", "加班免稅(1)", "加班免稅(1.34)", "加班免稅(1.67)", "加班免稅(2.00)", "加班免稅(2.67)", "加班扣稅(1.34)", "加班扣稅(1.67)", "加班扣稅(2.00)", "加班扣稅(2.67)", "備註" };
            if (p_forLaborInspectionCheck)
            {
                getFieldName = new string[] { "日期", "工作內容", "上班時間", "下班時間", "實際加班(起)", "實際加班(訖)", "應付時數", "加班免稅(1)", "加班免稅(1.34)", "加班免稅(1.67)", "加班免稅(2.00)", "加班免稅(2.67)", "加班扣稅(1.34)", "加班扣稅(1.67)", "加班扣稅(2.00)", "加班扣稅(2.67)", "備註" };
            }
            else
            {
                l_exportDT.Columns.Remove("WORKTIME_START");
                l_exportDT.Columns.Remove("WORKTIME_END");
            }

            l_excel.Worksheets[0].Cells.ImportArray(getFieldName, 1, 0, false);

            l_exportDT.Columns.Remove("NAME");
            l_exportDT.Columns.Remove("GUID");
            l_exportDT.Columns.Remove("num");


            //資料寫入EXCEL
            l_excel.Worksheets[0].Cells.ImportDataTable(l_exportDT, false, "A3");//查詢條件
            l_excel.Worksheets[0].Cells.StandardWidth = 13;

            foreach (Cell item in l_excel.Worksheets[0].Cells)
            {
                item.Style.HorizontalAlignment = TextAlignmentType.Right;
            }
            //l_excel.Worksheets[0].Cells["A1"].Style.HorizontalAlignment = TextAlignmentType.Right;
            //string l_fileName = "Export_" + p_class[0].NAME + DateTime.Now.ToString("yyyyMMdd");
            MemoryStream stream = new MemoryStream();
            l_excel.Save(stream, FileFormatType.Default);



            byte[] l_datas = stream.ToArray();

            //if (File.Exists(p_path + l_fileName + ".xls"))
            //{
            //    File.Delete(p_path + l_fileName + ".xls");
            //}

            //FileStream FS = File.Create(p_path + l_fileName + ".xls");
            //FS.Write(l_datas, 0, l_datas.Length);
            //FS.Close();

        }
        public static DataTable ToDataTable<T>(IList<T> p_data)
        {
            PropertyDescriptorCollection l_properties = TypeDescriptor.GetProperties(typeof(T));

            DataTable l_table = new DataTable();

            foreach (PropertyDescriptor prop in l_properties)
            {
                l_table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (T item in p_data)
            {
                DataRow l_row = l_table.NewRow();
                foreach (PropertyDescriptor prop in l_properties)
                    l_row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                l_table.Rows.Add(l_row);
            }
            return l_table;

        }
    }
}
