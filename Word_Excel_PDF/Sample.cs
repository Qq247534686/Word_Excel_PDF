using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using GemBox.Spreadsheet;
using Word_Excel_PDF;

class Sample
{
    /*
      备注：行的下标为0，列队下标为0开始
            修改人：李嘉成
            修改日期：2017-04-25
    */
    public void exportExcel()
    {
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");//免费版密钥
        ExcelFile ef = new ExcelFile();//操作对象
        ExcelWorksheet ws = ef.Worksheets.Add("Writing");//工作簿名
        object[,] skyscrapers = new object[21, 7]//测试数据集
        {
            {"Rank", "Building", "City", "Metric", "Imperial", "Floors", "Built (Year)"},
            { 1, "Taipei 101", "Taipei", 509, 1671, 101, 2004},
            { 2, "Petronas Tower 1", "Kuala Lumpur", 452, 1483, 88, 1998},
            { 3, "Petronas Tower 2", "Kuala Lumpur", 452, 1483, 88, 1998},
            { 4, "Sears Tower", "Chicago", 442, 1450, 108, 1974},
            { 5, "Jin Mao Tower", "Shanghai", 421, 1380, 88, 1998},
            { 6, "2 International Finance Centre", "Hong Kong", 415, 1362, 88, 2003},
            { 7, "CITIC Plaza", "Guangzhou", 391, 1283, 80, 1997},
            { 8, "Shun Hing Square", "Shenzhen", 384, 1260, 69, 1996},
            { 9, "Empire State Building", "New York City", 381, 1250, 102, 1931},
            {10, "Central Plaza", "Hong Kong", 374, 1227, 78, 1992},
            {11, "Bank of China Tower", "Hong Kong", 367, 1205, 72, 1990},
            {12, "Emirates Office Tower", "Dubai", 355, 1163, 54, 2000},
            {13, "Tuntex Sky Tower", "Kaohsiung", 348, 1140, 85, 1997},
            {14, "Aon Center", "Chicago", 346, 1136, 83, 1973},
            {15, "The Center", "Hong Kong", 346, 1135, 73, 1998},
            {16, "John Hancock Center", "Chicago", 344, 1127, 100, 1969},
            {17, "Ryugyong Hotel", "Pyongyang", 330, 1083, 105, 1992},
            {18, "Burj Al Arab", "Dubai", 321, 1053, 60, 1999},
            {19, "Chrysler Building", "New York City", 319, 1046, 77, 1930},
            {20, "Bank of America Plaza", "Atlanta", 312, 1023, 55, 1992}
        };
        //标题,第0行第0列
        ws.Cells[0, 0].Value = "测试html,png,excel.pdf";
        //设定每一列的列宽
        ws.Columns[0].Width = 8 * 256;
        ws.Columns[1].Width = 30 * 256;
        ws.Columns[2].Width = 16 * 256;
        ws.Columns[3].Width = 9 * 256;
        ws.Columns[4].Width = 9 * 256;
        ws.Columns[5].Width = 9 * 256;
        ws.Columns[6].Width = 9 * 256;

        int i, j;
        // Write header data to Excel cells.
        for (j = 0; j < 7; j++)
        {
            ws.Cells[3, j].Value = skyscrapers[0, j];//创建表的列名，共7列，从第4行开始
            ws.Cells.GetSubrangeAbsolute(2, j, 3, j).Merged = true;//合并单元格，从第3行开始到第4行，从第j列到第j列
            ws.Cells.GetSubrangeAbsolute(2, j, 3, j).Merged = true;
            ws.Cells.GetSubrangeAbsolute(2, j, 3, j).Merged = true;
            ws.Cells.GetSubrangeAbsolute(2, j, 3, j).Merged = true;
            ws.Cells.GetSubrangeAbsolute(2, j, 3, j).Merged = true;
            ws.Cells.GetSubrangeAbsolute(2, j, 3, j).Merged = true;
            ws.Cells.GetSubrangeAbsolute(2, j, 3, j).Merged = true;
        }
        //列样式
        CellStyle tmpStyle = new CellStyle();
        tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;//内容水平居中
        tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;//内容垂直居中
        tmpStyle.FillPattern.SetSolid(Color.Chocolate);//背景颜色
        tmpStyle.Font.Weight = ExcelFont.BoldWeight;
        tmpStyle.Font.Color = Color.White;//字体颜色
        tmpStyle.WrapText = true;//文本类型
        tmpStyle.Borders.SetBorders(MultipleBorders.Right | MultipleBorders.Top, Color.Black, LineStyle.Thin);//设置边框线
        ws.Cells.GetSubrangeAbsolute(2, 0, 3, 6).Style = tmpStyle;//样式应用的范围，第3行到到第4行的1到7列

        tmpStyle = new CellStyle();//重写样式
        tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
        tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
        tmpStyle.Font.Weight = ExcelFont.BoldWeight;
        ////扩展列
        //CellRange mergedRange = ws.Cells.GetSubrangeAbsolute(4, 7, 13, 7);//从第4行扩展到底13行，从第七列扩展到第七列
        //mergedRange.Merged = true;
        //mergedRange.Value = "T o p   1 0";//填充值
        //tmpStyle.Rotation = -90;
        //tmpStyle.FillPattern.SetSolid(Color.Lime);//填充色
        //mergedRange.Style = tmpStyle;
        ////扩展列
        //mergedRange = ws.Cells.GetSubrangeAbsolute(4, 8, 23, 8);
        //mergedRange.Merged = true;
        //mergedRange.Value = "T o p   2 0";
        //tmpStyle.IsTextVertical = true;
        //tmpStyle.FillPattern.SetSolid(Color.Gold);
        //mergedRange.Style = tmpStyle;

        //mergedRange = ws.Cells.GetSubrangeAbsolute(14, 7, 23, 7);
        //mergedRange.Merged = true;
        //mergedRange.Style = tmpStyle;

        // Write and format sample data to Excel cells.
        for (i = 0; i < 20; i++)//20行
        {
            for (j = 0; j < 7; j++)//7列
            {
                ExcelCell cell = ws.Cells[i + 4, j];//第i+5行，第j列的对象;
                cell.Value = skyscrapers[i + 1, j];//单元格赋值
                if (i % 2 == 0)
                    cell.Style.FillPattern.SetSolid(Color.LightSkyBlue);
                else
                    cell.Style.FillPattern.SetSolid(Color.FromArgb(210, 210, 230));

                if (j == 3)
                    cell.Style.NumberFormat = "#\" m\"";//格式化字符

                if (j == 4)
                    cell.Style.NumberFormat = "#\" ft\"";

                if (j > 2)
                    cell.Style.Font.Name = "Courier New";

                cell.Style.Borders[IndividualBorder.Right].LineStyle = LineStyle.Thin;
                cell.Style.Borders[IndividualBorder.Bottom].LineStyle = LineStyle.Thin;
            }
        }

        //ws.Cells.GetSubrange("A5", "I24").Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Double);
        //ws.Cells.GetSubrange("A3", "G4").Style.Borders.SetBorders(MultipleBorders.Vertical | MultipleBorders.Top, Color.Black, LineStyle.Double);
        //ws.Cells.GetSubrange("A5", "H14").Style.Borders.SetBorders(MultipleBorders.Bottom | MultipleBorders.Right, Color.Black, LineStyle.Double);

        ws.Cells["A27"].Value = "备注:";
        ws.Cells["A28"].Value = "a)行的下标为0，列队下标为0开始";
        ws.Cells["A29"].Value = "b)官网地址：https://www.gemboxsoftware.com/";

        ws.PrintOptions.FitWorksheetWidthToPages = 1;

        ef.Save("123.pdf");
    }
    public void exportExcelTpl(int count, string path, exportType filetype, Mytable data)
    {
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");//免费版密钥
        ExcelFile ef = new ExcelFile();//操作对象
        ExcelWorksheet ws = ef.Worksheets.Add("Writing");//工作簿名

        //标题,第0行第0列
        //ws.Cells.GetSubrangeAbsolute(0, 5, 0, 6).Merged = true;
        //ws.Cells[0, 0].Value = "页眉";
        //设定每一列的列宽
        ws.Columns[0].Width = 8 * 256;
        ws.Columns[1].Width = 12 * 256;
        ws.Columns[2].Width = 9 * 256;
        ws.Columns[3].Width = 9 * 256;
        ws.Columns[4].Width = 9 * 256;
        ws.Columns[5].Width = 9 * 256;
        ws.Columns[6].Width = 9 * 256;
        int rows = 1;
        for (int i = 0; i < count; i++)
        {
            ws.Cells[rows, 0].Value = "厂区";
            ws.Cells[rows, 1].Value = data.MydataList[i].cq;
            ws.Cells[rows, 2].Value = "盘点地点";
            ws.Cells[rows, 3].Value = data.MydataList[i].pddd;
            ws.Cells[rows, 4].Value = "盘点时间";
            ws.Cells.GetSubrangeAbsolute(rows, 5, rows, 6).Merged = true;
            ws.Cells[rows, 5].Value = data.MydataList[i].pdsj;
            ws.Cells[rows+1, 0].Value = "序号";
            ws.Cells[rows+1, 1].Value = "名称";
            ws.Cells[rows+1, 2].Value = "单位";
            ws.Cells[rows+1, 3].Value = "盘点数";
            ws.Cells[rows+1, 4].Value = "账目数";
            ws.Cells[rows+1, 5].Value = "盘盈数";
            ws.Cells[rows+1, 6].Value = "盘亏数";
            ws.Cells[rows+2, 0].Value = "1";
            ws.Cells[rows+2, 1].Value = "RFID托盘";
            ws.Cells[rows+2, 2].Value = "个";
            ws.Cells[rows+2, 3].Value = data.MydataList[i].pds1;
            ws.Cells[rows+2, 4].Value = data.MydataList[i].zms1;
            ws.Cells[rows+2, 5].Value = data.MydataList[i].zms1;
            ws.Cells[rows+2, 6].Value = data.MydataList[i].pks1;
            ws.Cells[rows+3, 0].Value = "2";
            ws.Cells[rows+3, 1].Value = "RFID围板箱";
            ws.Cells[rows+3, 2].Value = "个";
            ws.Cells[rows+3, 3].Value =data.MydataList[i].pds2;
            ws.Cells[rows+3, 4].Value =data.MydataList[i].zms2;
            ws.Cells[rows+3, 5].Value =data.MydataList[i].zms2;
            ws.Cells[rows+3, 6].Value =data.MydataList[i].pks2;
            ws.Cells[rows+3, 3].Value =
            ws.Cells.GetSubrangeAbsolute(rows+4, 0, rows+4, 6).Merged = true;
            ws.Cells[rows + 4, 0].Value = "";
            rows +=5;
        }
       
    
        //列样式
        CellStyle tmpStyle = new CellStyle();
        tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;//内容水平居中
        tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;//内容垂直居中
        tmpStyle.Font.Weight = ExcelFont.BoldWeight;
        tmpStyle.Font.Color = Color.Black;//字体颜色
        tmpStyle.WrapText = true;//文本类型
        tmpStyle.Borders.SetBorders(MultipleBorders.All, Color.Black, LineStyle.Thin);//设置边框线
        for (int i = 1; i <rows ; i++)
        {
            for (int j = 0; j < 7; j++)
            {
                ws.Cells[i, j].Style = tmpStyle;
            }
        }
        CellStyle last = new CellStyle();
        last.Borders.SetBorders(MultipleBorders.Bottom | MultipleBorders.Right | MultipleBorders.Left, Color.Black, LineStyle.None);//设置边框线
        last.Borders.SetBorders(MultipleBorders.Top, Color.Black, LineStyle.Thin);
        ws.Cells[rows-1, 0].Style = last;
        ws.Cells[rows - 1, 0].Value = "";
        //CellStyle tmpStyle = new CellStyle();
        //tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;//内容水平居中
        //tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;//内容垂直居中
        //tmpStyle.Font.Weight = ExcelFont.BoldWeight;
        //tmpStyle.Font.Color = Color.White;//字体颜色
        //tmpStyle.WrapText = true;//文本类型
        //tmpStyle.Borders.SetBorders(MultipleBorders.Right | MultipleBorders.Top, Color.Black, LineStyle.Thin);//设置边框线
        //ws.Cells.GetSubrangeAbsolute(1, 0, 4, 6).Style = tmpStyle;//样式应用的范围，第3行到到第4行的1到7列

        //tmpStyle = new CellStyle();//重写样式
        //tmpStyle.HorizontalAlignment = HorizontalAlignmentStyle.Center;
        //tmpStyle.VerticalAlignment = VerticalAlignmentStyle.Center;
        //tmpStyle.Font.Weight = ExcelFont.BoldWeight;



        ws.PrintOptions.FitWorksheetWidthToPages = 1;

        ef.Save(path+"." + filetype.ToString());
    }
}
