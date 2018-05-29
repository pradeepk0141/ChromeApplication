using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using ClosedXML.Excel;
using System.IO;

public partial class ExcelDownload : System.Web.UI.Page
{
    DataAccess objDataAccess = new DataAccess();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {

        }
    }

    protected void btnTop10News_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_Top10News]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "Top 10 News", Page.Response);
    }

    protected void btnStoryROI_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_StoryROI]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "Story ROI", Page.Response);
    }

    protected void btnGenreROI_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_GenreROI]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "Genre ROI", Page.Response);
    }

    protected void btnProgROI_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_ProgROI]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "Prog ROI", Page.Response);
    }

    protected void btnLogisticsROI_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_LogisticsROI]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "Logistics ROI", Page.Response);
    }


    protected void btnTelecastFormatROI_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_TelecastFormatROI]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "Telecast Format ROI", Page.Response);
    }

    protected void btnAnchorROI_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_AnchorROI]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "Anchor ROI", Page.Response);
    }

    protected void btnExclusive_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_EXCLUSIVE]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "EXCLUSIVE", Page.Response);
    }

    protected void btnURSplitRoI_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_URSplitRoI]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel_URSplit(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "UR Split ROI", Page.Response);
    }
    protected void btnDoDStoryTrend_Click(object sender, EventArgs e)
    {
        string strSQL = "EXEC [proc_DoDStoryTrend]";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt;
        gvUploadFiles.DataBind();
        ExportToExcel_URSplit(dt, "Output_File_" + System.DateTime.Now.ToString("dd_MMM_yyyy") + ".xlsx", "UR Split ROI", Page.Response);
    }
    #region Export To Excel With Formatting   
    private string ExportToExcel_URSplit(DataTable table, string FileName, string WorksheetName, HttpResponse response)
    {
        DataTable dtChannel = new DataTable();
        dtChannel = objDataAccess.GetDataTable("SELECT DISTINCT Channel FROM tb_RawData_Master ORDER  BY Channel");
        var wb = new XLWorkbook();
        #region Create Sheet <Dynamic Name>
        var ws = wb.Worksheets.Add(WorksheetName);
        string DataCell1, Range = "B1";
        int verticalColumn = 0, horiZontalColumn = 1;
        verticalColumn = 1;
        #region  Channel Name Header (Main Header)
        for (int i = 0; i < dtChannel.Rows.Count; i++)
        {
            //Urban
            DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();//B1
            ws.Cell(DataCell1).Value = dtChannel.Rows[i][0].ToString() + " Urban";
            setAlignment(ws, DataCell1, true, true);
            Range = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString() + ":" + getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();//C3:F3
            ws.Range(Range).Merge();
            setBorderAndBg(ws, Range, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            verticalColumn += 4;
            //Rural
            DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();//B1
            ws.Cell(DataCell1).Value = dtChannel.Rows[i][0].ToString() + " Rural";
            setAlignment(ws, DataCell1, true, true);
            Range = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString() + ":" + getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();//C3:F3
            ws.Range(Range).Merge();
            setBorderAndBg(ws, Range, XLBorderStyleValues.Thin, XLColor.GreenYellow, true, true);
            verticalColumn += 4;
            if (i == dtChannel.Rows.Count - 1)
            {
                DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();//B1
                ws.Cell(DataCell1).Value = " Urban";
                setAlignment(ws, DataCell1, true, true);
                Range = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString() + ":" + getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();//C3:F3
                ws.Range(Range).Merge();
                setBorderAndBg(ws, Range, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                verticalColumn += 4;
                //Rural
                DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();//B1
                ws.Cell(DataCell1).Value = " Rural";
                setAlignment(ws, DataCell1, true, true);
                Range = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString() + ":" + getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();//C3:F3
                ws.Range(Range).Merge();
                setBorderAndBg(ws, Range, XLBorderStyleValues.Thin, XLColor.GreenYellow, true, true);
                verticalColumn += 4;
            }
        }
        #endregion
        #region  Channel Name Header (Secondary Header)
        horiZontalColumn = 2;
        verticalColumn = 0;
        for (int i = 0; i < table.Columns.Count; i++)
        {
            DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();
            ws.Cell(DataCell1).Value = getColumnHeaderName(table.Columns[i].ColumnName.ToString(), dtChannel, "UR");
            setAlignment(ws, DataCell1, true, true);
            setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            if (i == table.Columns.Count - 1)
            {
                DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Total Sum of HSM Urban NTVT";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Total Sum of Duration2";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "SOV";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Efficiency Index (EI)";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                verticalColumn += 4;
                DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Total Sum of HSM Rural NTVT";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Total Sum of Duration2";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "SOV";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Efficiency Index (EI)";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            }
            verticalColumn++;
        }
        #endregion
        #region  Put Table Data In Sheet
        horiZontalColumn = 3;
        for (int i = 0; i < table.Rows.Count; i++)
        {
            for (int j = 0; j < table.Columns.Count; j++)
            {
                verticalColumn = j;
                DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetValue(table.Rows[i][j].ToString());
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                ws.Range(DataCell1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                if (j == 0) { ws.Range(DataCell1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left); continue; }
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                if ((j - 1) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000"; }
                else
                if ((j - 2) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0"; }
                else
                if ((j - 3) % 4 == 0 || (j - 4) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%"; }
                if (j == table.Columns.Count - 1)
                {
                    DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 8, horiZontalColumn, dtChannel));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000";
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(2, 8, horiZontalColumn, dtChannel));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0";

                    DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1("=" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%";
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);

                    DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + ")/(" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + ")");
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%";

                    verticalColumn += 4;
                    DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(5, 8, horiZontalColumn, dtChannel));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000";
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(6, 8, horiZontalColumn, dtChannel));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0";

                    DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1("=" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%";
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);

                    DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + ")/(" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + ")");
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%";
                }
            }
            horiZontalColumn++;
        }
        #endregion
        #region  Summrize the Columns
        for (int i = 0; i < table.Columns.Count; i++)
        {
            verticalColumn = i;
            DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();
            setAlignment(ws, DataCell1, true, true);
            setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            if (i == 0) { ws.Cell(DataCell1).Value = "Grand Total"; continue; }
            else
            {
                ws.Cell(DataCell1).SetFormulaA1("=ROUND(SUM(" + getExcelColumnCode(i) + "3:" + getExcelColumnCode(i) + (table.Rows.Count + 2) + "),6)");
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                if ((i - 1) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000"; }
                else
                if ((i - 2) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0"; }
                else
                if ((i - 3) % 4 == 0 || (i - 4) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0%"; }
                if ((i - 4) % 4 == 0) { ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(i - 2) + (table.Rows.Count + 3).ToString() + "/" + getExcelColumnCode(i - 1) + (table.Rows.Count + 3) + ")/(" + getExcelColumnCode(i - 2) + (table.Rows.Count + 3).ToString() + "/" + getExcelColumnCode(i - 1) + (table.Rows.Count + 3) + ")"); }
            }
            if (i == table.Columns.Count - 1)
            {
                DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=SUM(" + getExcelColumnCode(verticalColumn + 1) + "3:" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 2) + ")");
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000";
                DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=SUM(" + getExcelColumnCode(verticalColumn + 2) + "3:" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 2) + ")");

                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0";

                DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3));
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000%";

                DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + ")/(" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + ")");
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000%";

                verticalColumn += 4;
                DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=SUM(" + getExcelColumnCode(verticalColumn + 1) + "3:" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 2) + ")");
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000";
                DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=SUM(" + getExcelColumnCode(verticalColumn + 2) + "3:" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 2) + ")");

                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0";

                DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3));
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000%";

                DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + ")/(" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + ")");
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000%";
            }
        }
        #endregion
        #endregion
        HttpResponse httpResponse = response;
        httpResponse.Clear();
        httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        httpResponse.AddHeader("content-disposition", "attachment;filename=" + WorksheetName.Replace(" ", "_") + ".xlsx");
        // Flush the workbook to the Response.OutputStream
        using (MemoryStream memoryStream = new MemoryStream())
        {
            wb.SaveAs(memoryStream);
            memoryStream.WriteTo(httpResponse.OutputStream);
            memoryStream.Close();
        }
        httpResponse.End();
        return "OK";
    }
    #endregion
    #region Export To Excel With Formatting   
    private string ExportToExcel(DataTable table, string FileName, string WorksheetName, HttpResponse response)
    {
        DataTable dtChannel = new DataTable();
        dtChannel = objDataAccess.GetDataTable("SELECT DISTINCT Channel FROM tb_RawData_Master ORDER  BY Channel");
        var wb = new XLWorkbook();
        #region Create Sheet <Dynamic Name>
        var ws = wb.Worksheets.Add(WorksheetName);
        string DataCell1, Range = "B1";
        int verticalColumn = 0, horiZontalColumn = 1;
        verticalColumn = 1;
        #region  Channel Name Header (Main Header)
        for (int i = 0; i < dtChannel.Rows.Count; i++)
        {
            DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();//B1
            ws.Cell(DataCell1).Value = dtChannel.Rows[i][0].ToString();
            setAlignment(ws, DataCell1, true, true);
            Range = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString() + ":" + getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();//C3:F3
            ws.Range(Range).Merge();
            setBorderAndBg(ws, Range, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            verticalColumn += 4;
        }
        #endregion
        #region  Channel Name Header (Secondary Header)
        horiZontalColumn = 2;
        verticalColumn = 0;
        for (int i = 0; i < table.Columns.Count; i++)
        {
            DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();
            ws.Cell(DataCell1).Value = getColumnHeaderName(table.Columns[i].ColumnName.ToString(), dtChannel);
            setAlignment(ws, DataCell1, true, true);
            setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            if (i == table.Columns.Count - 1)
            {
                DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Total Sum of HSM NTVT";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Total Sum of Duration2";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "SOV";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).Value = "Efficiency Index (EI)";
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            }
            verticalColumn++;
        }
        #endregion
        #region  Put Table Data In Sheet
        horiZontalColumn = 3;
        for (int i = 0; i < table.Rows.Count; i++)
        {
            for (int j = 0; j < table.Columns.Count; j++)
            {
                verticalColumn = j;
                DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();
                //ws.Cell(DataCell1).Value = 
                ws.Cell(DataCell1).SetValue(table.Rows[i][j].ToString());
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                ws.Range(DataCell1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                if (j == 0) { ws.Range(DataCell1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left); continue; }
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                if ((j - 1) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000"; }
                else
                if ((j - 2) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0"; }
                else
                if ((j - 3) % 4 == 0 || (j - 4) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%"; }
                if (j == table.Columns.Count - 1)
                {
                    DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000";
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(2, 4, horiZontalColumn, dtChannel));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0";

                    DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1("=" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3));
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%";
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);

                    DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                    ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + ")/(" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + ")");
                    ws.Cell(DataCell1).DataType = XLDataType.Number;
                    setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.White, false);
                    ws.Cell(DataCell1).Style.NumberFormat.Format = "0.00000%";
                }
            }
            horiZontalColumn++;
        }
        #endregion
        #region  Summrize the Columns
        for (int i = 0; i < table.Columns.Count; i++)
        {
            verticalColumn = i;
            DataCell1 = getExcelColumnCode(verticalColumn) + horiZontalColumn.ToString();
            setAlignment(ws, DataCell1, true, true);
            setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
            if (i == 0) { ws.Cell(DataCell1).Value = "Grand Total"; continue; }
            else
            {
                ws.Cell(DataCell1).SetFormulaA1("=ROUND(SUM(" + getExcelColumnCode(i) + "3:" + getExcelColumnCode(i) + (table.Rows.Count + 2) + "),6)");
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                if ((i - 1) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000"; }
                else
                if ((i - 2) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0"; }
                else
                if ((i - 3) % 4 == 0 || (i - 4) % 4 == 0) { ws.Cell(DataCell1).Style.NumberFormat.Format = "0%"; }
                if ((i - 4) % 4 == 0) { ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(i - 2) + (table.Rows.Count + 3).ToString() + "/" + getExcelColumnCode(i - 1) + (table.Rows.Count + 3) + ")/(" + getExcelColumnCode(i - 2) + (table.Rows.Count + 3).ToString() + "/" + getExcelColumnCode(i - 1) + (table.Rows.Count + 3) + ")"); }
            }
            if (i == table.Columns.Count - 1)
            {
                DataCell1 = getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=SUM(" + getExcelColumnCode(verticalColumn + 1) + "3:" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 2) + ")");
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000";
                DataCell1 = getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=SUM(" + getExcelColumnCode(verticalColumn + 2) + "3:" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 2) + ")");
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(2, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0";

                DataCell1 = getExcelColumnCode(verticalColumn + 3) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3));
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000%";

                DataCell1 = getExcelColumnCode(verticalColumn + 4) + horiZontalColumn.ToString();
                ws.Cell(DataCell1).SetFormulaA1("=(" + getExcelColumnCode(verticalColumn + 1) + horiZontalColumn.ToString() + "/" + getExcelColumnCode(verticalColumn + 2) + horiZontalColumn.ToString() + ")/(" + getExcelColumnCode(verticalColumn + 1) + (table.Rows.Count + 3) + "/" + getExcelColumnCode(verticalColumn + 2) + (table.Rows.Count + 3) + ")");
                //ws.Cell(DataCell1).SetFormulaA1(getHorizontalFormula(1, 4, horiZontalColumn, dtChannel));
                ws.Cell(DataCell1).DataType = XLDataType.Number;
                setAlignment(ws, DataCell1, true, true);
                setBorderAndBg(ws, DataCell1, XLBorderStyleValues.Thin, XLColor.LightSteelBlue, true, true);
                ws.Cell(DataCell1).Style.NumberFormat.Format = "0.0000%";
            }
        }
        #endregion
        #endregion
        HttpResponse httpResponse = response;
        httpResponse.Clear();
        httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        httpResponse.AddHeader("content-disposition", "attachment;filename=" + WorksheetName.Replace(" ", "_") + ".xlsx");
        // Flush the workbook to the Response.OutputStream
        using (MemoryStream memoryStream = new MemoryStream())
        {
            wb.SaveAs(memoryStream);
            memoryStream.WriteTo(httpResponse.OutputStream);
            memoryStream.Close();
        }
        httpResponse.End();
        return "OK";
    }
    #endregion
    #region Excel Formatting Methods    
    private void setBorderAndBg(IXLWorksheet ws, string Range, XLBorderStyleValues v)
    {
        try
        {
            ws.Range(Range).Style.Border.LeftBorder = v;
            ws.Range(Range).Style.Border.RightBorder = v;
            ws.Range(Range).Style.Border.TopBorder = v;
            ws.Range(Range).Style.Border.BottomBorder = v;
        }
        catch { }
    }
    private void setBorderAndBg(IXLWorksheet ws, string Range, XLBorderStyleValues v, XLColor bgcolor)
    {
        try
        {
            setBorderAndBg(ws, Range, v);
            ws.Range(Range).Style.Fill.BackgroundColor = bgcolor;
        }
        catch { }
    }
    private void setBorderAndBg(IXLWorksheet ws, string Range, XLBorderStyleValues v, XLColor bgcolor, bool fontBold)
    {
        try
        {
            setBorderAndBg(ws, Range, v, bgcolor);
            ws.Range(Range).Style.Font.Bold = fontBold;
        }
        catch { }
    }
    private void setBorderAndBg(IXLWorksheet ws, string Range, XLBorderStyleValues v, XLColor bgcolor, bool fontBold, bool wrapText)
    {
        try
        {
            setBorderAndBg(ws, Range, v, bgcolor, fontBold);
            ws.Range(Range).Style.Alignment.WrapText = wrapText;
        }
        catch { }
    }
    private void setBorderAndBg(IXLWorksheet ws, string Range, XLBorderStyleValues v, XLColor bgcolor, bool fontBold, bool wrapText, double fontSize)
    {
        try
        {
            setBorderAndBg(ws, Range, v, bgcolor, fontBold, wrapText);
            ws.Cell(Range).WorksheetColumn().Width = fontSize;
        }
        catch { }
    }
    private void setAlignment(IXLWorksheet ws, string Range, bool vertical, bool horizontal)
    {
        try
        {
            if (vertical == true) { ws.Range(Range).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top); }
            if (horizontal == true) { ws.Range(Range).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center); }
        }
        catch { }
    }
    private string getExcelColumnCode(int columnCount)
    {
        string columnCode = (((columnCount / 26) <= 0) ? "" : ((char)(64 + (columnCount / 26))).ToString()) + ((char)(65 + (columnCount % 26))).ToString();
        return columnCode;
    }
    private string getColumnHeaderName(string ColumnName, DataTable dtChannel)
    {
        string HeaderName = ColumnName;
        for (int i = 0; i < dtChannel.Rows.Count; i++)
        {
            HeaderName = HeaderName.Replace(dtChannel.Rows[i][0].ToString() + "_", "");
        }
        return HeaderName;
    }
    private string getColumnHeaderName(string ColumnName, DataTable dtChannel, string typeSheet)
    {
        string HeaderName = ColumnName;
        for (int i = 0; i < dtChannel.Rows.Count; i++)
        {
            HeaderName = HeaderName.Replace(dtChannel.Rows[i][0].ToString() + "_", "").Replace("Urban", "").Replace("Rural", "");
        }
        return HeaderName;
    }
    private string getHorizontalFormula(int startColumns, int step, int horiZontalColumn, DataTable columnTable)
    {
        string Formula = "=";
        for (int i = 0; i < columnTable.Rows.Count; i++)
        {
            Formula += getExcelColumnCode(i * step + startColumns).ToString() + horiZontalColumn.ToString();
            if (i == columnTable.Rows.Count - 1) { continue; }
            Formula += "+";
        }
        return Formula;
    }
    #endregion

    
}
