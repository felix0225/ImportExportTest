using Core.Utility;
using LinqToExcel;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ImportExportTest
{
    public partial class index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            UploadFile("LinqToExcel");
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            UploadFile("EPPlus");
        }

        private void UploadFile(string type)
        {
            FileUploadErr.Visible = false;

            const string uploadFail = "上傳失敗！檔名 : {0}";

            var uploadExcelPath = ConfigurationHelper.GetFilesUploadPath("xlsx");

            if (!FileUpload1.HasFile) return;
            var fileType = Path.GetExtension(FileUpload1.FileName).ToLower();

            try
            {
                var fileName = GuidHelper.Get32String();
                var xPatch = $"{uploadExcelPath}{fileName}{fileType}";
                FileUpload1.PostedFile.SaveAs(xPatch);

                if (type == "LinqToExcel")
                {
                    var excel = new ExcelQueryFactory(xPatch);
                    var worksheetCount = excel.GetWorksheetNames().Count();

                    //判斷有資料時才做寫入
                    if (worksheetCount > 0)
                    {
                        //取得excel內的資料
                        var importdatas = from c in excel.Worksheet(0)
                                          select c;

                        foreach (var importdata in importdatas)
                        {
                            var title = importdata["標題"].ToString().Trim();

                            Response.Write(title + "<br/>");
                        }
                    }
                }
                else
                {
                    //載入Excel檔案
                    var fileStream = new FileStream(xPatch, FileMode.Open, FileAccess.Read);
                    using (var ep = new ExcelPackage(fileStream))
                    {
                        //判斷有資料時才做寫入
                        foreach (var sheet in ep.Workbook.Worksheets)
                        {
                            if (sheet.Dimension == null) continue;

                            var startRowIndex = sheet.Dimension.Start.Row;  //起始列
                            var endRowIndex = sheet.Dimension.End.Row;      //結束列
                            var startColumn = sheet.Dimension.Start.Column; //開始欄
                            var endColumn = sheet.Dimension.End.Column;     //結束欄

                            //不含標題，資料開始行，一般是1，代表由第2行開始
                            startRowIndex += 1;

                            for (var currentRow = startRowIndex; currentRow <= endRowIndex; currentRow++)
                            {
                                //抓出當前的資料範圍
                                var range = sheet.Cells[currentRow, startColumn, currentRow, endColumn];

                                //全部儲存格是完全空白時則跳過
                                if (range.Any(c => !string.IsNullOrEmpty(c.Text)) == false)
                                    continue;
                                
                                var title = range[currentRow, 1].Text;

                                Response.Write(title + "<br/>");
                            }
                        }
                    }
                    fileStream.Close();
                }

                FilesHelper.DeleteFile(xPatch);
            }
            catch (Exception ex)
            {
                var uploadFileName = FileUpload1.PostedFile.FileName;
                FileUploadErr.Text = string.Format(uploadFail, uploadFileName) + @"，" + ex.Message;
                FileUploadErr.Visible = true;
            }
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            var exportdt = new DataTable();
            exportdt.Columns.Add("標題", typeof(string));

            var dr = exportdt.NewRow();
            dr["標題"] = "export";
            exportdt.Rows.Add(dr);

            FilesHelper.ExportDatasToExcel(exportdt, "匯出_" + DateTime.Now.ToString("yyyyMMddHHmm"));
        }
    }
}