using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

//NPOI範例使用到的using
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelDemoProject
{
    public partial class ExportExcelPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void export_btn_Click(object sender, EventArgs e)
        {
            //1.開啟我們的範例檔案
            using (FileStream file = new FileStream(Server.MapPath(@"~/Templete/FixedDemoTemplete.xlsx"), FileMode.Open, FileAccess.Read))
            {
                //2.輸出檔名(我習慣加上日期時間)
                string fileName = "學生成績單資料_" + DateTime.Now.ToString();
                //3.開啟workbook檔案
                IWorkbook workbook = new XSSFWorkbook(file);
                //4.指定為第一個頁籤
                ISheet sheet = workbook.GetSheetAt(0);

                var lst = CreateExportData();

                //從第二列到第三列，往下移動lst.Count列
                sheet.ShiftRows(1, 2, lst.Count);

                for (int i = 0; i < lst.Count; i++)
                {
                    //5.依列逐格填入資料
                    var l = lst.ElementAt(i);
                    sheet.CreateRow(i + 1);
                    sheet.GetRow(i + 1).CreateCell(0).SetCellValue(l.StuName);
                    sheet.GetRow(i + 1).CreateCell(1).SetCellValue(l.Chinese);
                    sheet.GetRow(i + 1).CreateCell(2).SetCellValue(l.English);
                    sheet.GetRow(i + 1).CreateCell(3).SetCellValue(l.Math);
                }

                //資料總比數
                var dataCount = lst.Count;

                //科目總分計算行
                string[] subjectTotalColumns = new string[] { "B", "C", "D", "E" };
                if (dataCount > 0)
                {
                    //設定要填入公式的列數
                    IRow subjectTotalRow = sheet.GetRow(dataCount + 1);
                    IRow subjectAverageRow = sheet.GetRow(dataCount + 2);
                    for (int i = 0; i < subjectTotalColumns.Count(); i++)
                    {
                        subjectTotalRow.CreateCell(i + 1).CellFormula = string.Format("SUM({0}2:{0}{1})", subjectTotalColumns[i], dataCount + 1);
                        subjectAverageRow.CreateCell(i + 1).CellFormula = string.Format("AVERAGE({0}2:{0}{1})", subjectTotalColumns[i], dataCount + 1);
                    }

                    //個人總分計算               
                    string startPersonColunn = "B";
                    string endPersonColumn = "D";

                    for (int i = 0; i < dataCount; i++)
                    {
                        IRow personTotalRow = sheet.GetRow(i + 1);
                        personTotalRow.CreateCell(4).CellFormula = string.Format("SUM({0}{1}:{2}{1})", startPersonColunn, i + 2, endPersonColumn);
                    }
                    //由於最外層平均要多下面兩列一起計算
                    for (int i = 0; i < dataCount + 2; i++)
                    {
                        IRow personTotalRow = sheet.GetRow(i + 1);
                        personTotalRow.CreateCell(5).CellFormula = string.Format("AVERAGE({0}{1}:{2}{1})", startPersonColunn, i + 2, endPersonColumn);
                    }
                }
                //重新計算公式內的值
                sheet.ForceFormulaRecalculation = true;
                #region 輸出
                MemoryStream ms = new MemoryStream();
                workbook.Write(ms);

                //設定檔名, IE 要特殊處理
                if (HttpContext.Current.Request.Browser.Browser == "IE" || HttpContext.Current.Request.Browser.Browser == "InternetExplorer") fileName = HttpContext.Current.Server.UrlPathEncode(fileName);

                HttpContext.Current.Response.ClearHeaders();
                HttpContext.Current.Response.Clear();

                //此列針對xlsx錯誤做修正處理，若使用hssfworkbook則不需要此行
                HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                HttpContext.Current.Response.Cache.SetNoStore();
                HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + fileName + ".xlsx"));
                HttpContext.Current.Response.BinaryWrite(ms.ToArray());

                workbook = null;
                ms.Close();
                ms.Dispose();

                HttpContext.Current.Response.End();
                #endregion
            }
        }
        private List<GradeBook> CreateExportData()
        {
            var lst = new List<GradeBook> {
                new GradeBook() { StuName="鄭XX",Chinese=100,English=90,Math=95},
                new GradeBook() { StuName="陳XX",Chinese=60,English=90,Math=80},
                new GradeBook() { StuName="吳XX",Chinese=70,English=60,Math=75},
                new GradeBook() { StuName="李XX",Chinese=80,English=30,Math=85},
                new GradeBook() { StuName="林XX",Chinese=50,English=20,Math=65}
            };
            return lst;
        }

        public class GradeBook
        {
            public string StuName { get; set; }
            public double Chinese { get; set; }
            public double English { get; set; }
            public double Math { get; set; }
        }
    }
}