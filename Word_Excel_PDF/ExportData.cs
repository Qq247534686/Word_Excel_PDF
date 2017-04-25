using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox;
using GemBox.Spreadsheet;
using GemBox.Document.MailMerging;
using GemBox.Document;

namespace Word_Excel_PDF
{
    public class ExportData
    {
        /// <summary>
        /// 导出Excel
        /// </summary>
        public static void exportExcel(DataTable dt)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = new ExcelFile();
            ExcelWorksheet ws = ef.Worksheets.Add("DataSheet");
            ws.InsertDataTable(dt, new InsertDataTableOptions(0, 0) { ColumnHeaders = true });
            //ef.Save(ws.Name+".xls");
            ef.Save("Writing.html");
            //ef.Save(this.Response, "Report." + name);

        }
      
    /// <summary>
    /// 导出Word
    /// </summary>
    public static void exportDocx()
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        DocumentModel document = new DocumentModel();
        document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "English: asd"),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Russian: "),
                    new Run(document, new string(new char[] { '\u0417', '\u0434', '\u0440', '\u0430', '\u0432', '\u0441', '\u0442', '\u0432', '\u0443', '\u0439', '\u0442', '\u0435' })),
                    new SpecialCharacter(document, SpecialCharacterType.LineBreak),
                    new Run(document, "Chinese: "),
                    new Run(document, new string(new char[] { '\u4f60', '\u597d' }))),
               new Paragraph(document, "In order to see Russian and Chinese characters you need to have appropriate fonts on your machine.")));
        document.Save("Writing.docx");

    }
    public static DataTable dt()
    {
        DataTable tblDatas = new DataTable("Datas");
        DataColumn dc = null;
        dc = tblDatas.Columns.Add("ID", Type.GetType("System.Int32"));
        dc.AutoIncrement = true;//自动增加
        dc.AutoIncrementSeed = 1;//起始为1
        dc.AutoIncrementStep = 1;//步长为1
        dc.AllowDBNull = false;//

        dc = tblDatas.Columns.Add("Product", Type.GetType("System.String"));
        dc = tblDatas.Columns.Add("Version", Type.GetType("System.String"));
        dc = tblDatas.Columns.Add("Description", Type.GetType("System.String"));

        DataRow newRow;
        newRow = tblDatas.NewRow();
        newRow["Product"] = "大话西游";
        newRow["Version"] = "2.0";
        newRow["Description"] = "我很喜欢";
        tblDatas.Rows.Add(newRow);

        newRow = tblDatas.NewRow();
        newRow["Product"] = "梦幻西游";
        newRow["Version"] = "3.0";
        newRow["Description"] = "比大话更幼稚";
        tblDatas.Rows.Add(newRow);
        return tblDatas;
    }
}
}
