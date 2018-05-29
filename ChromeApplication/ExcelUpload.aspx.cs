using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class ExcelUpload : System.Web.UI.Page
{
    DataAccess objDataAccess = new DataAccess();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Session["userID"] = "1";
            //txtDate.Text = System.DateTime.Now.ToString("dd-MMM-yyyy");
            getUnSyncFiles();
        }
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        lblMessage.ForeColor = Color.Red; lblMessage.Text = "";
        if (txtDate.Text.Trim() == "")
        {
            lblMessage.Text = "Fill A Date For Excel File";
            return;
        }
        if (!fileUpload.HasFile)
        {
            lblMessage.Text = "Select a file first";
            return;
        }
        string strFilename, strMessage;
        strFilename = fileUpload.PostedFile.FileName.ToString();
        strFilename = System.DateTime.Now.ToString("yyyyMMddHHmmssfff") + strFilename.Substring(strFilename.IndexOf("."));
        uploadFile(fileName: strFilename, folderName: WebConfigurationManager.AppSettings["folderPath"].ToString(), outChar: out string outChar, outMessage: out strMessage);
        lblMessage.Text = strMessage;
        lblMessage.ForeColor = (outChar == "!") ? Color.Red : Color.Blue;
        if (outChar == "!") { return; }
        string strSQL = "EXEC [proc_Uploaded_Files] @action='insert', @upload_id=0, @Selected_date='" + txtDate.Text + "'"
            + ", @original_file_name = '" + fileUpload.PostedFile.FileName.ToString() + "', @system_file_name='" + strFilename + "'"
            + ", @upload_by='" + Session["userID"].ToString() + "'";
        DataTable dt = new DataTable();
        dt = objDataAccess.GetDataTable(strSQL);
        if (dt.Rows.Count > 0)
        {
            getUnSyncFiles();
        }
    }

    public void getUnSyncFiles()
    {
        string strSQL = "EXEC [proc_Uploaded_Files] @action='select', @upload_by='" + Session["userID"].ToString() + "',@upload_id=0";
        DataTable dt1 = new DataTable();
        dt1 = objDataAccess.GetDataTable(strSQL);
        gvUploadFiles.DataSource = dt1;
        gvUploadFiles.DataBind();
    }
    public void uploadFile(string fileName, string folderName, out string outChar, out string outMessage)
    {
        if (fileName == "")
        {
            outMessage = "Invalid filename supplied"; outChar = "!";
            return;
        }
        if (fileUpload.PostedFile.ContentLength == 0)
        {
            outMessage = "Invalid file content"; outChar = "!";
            return;
        }
        fileName = System.IO.Path.GetFileName(fileName);
        if (folderName == "")
        {
            outMessage = "Path not found"; outChar = "!";
            return;
        }
        try
        {
            if (fileUpload.PostedFile.ContentLength <= 204800000)
            {
                fileUpload.PostedFile.SaveAs(Server.MapPath(folderName) + "\\" + fileName);
                outMessage = "File uploaded successfully"; outChar = "";
                return;
            }
            else
            {
                outMessage = "Unable to upload,file exceeds maximum limit"; outChar = "!";
                return;
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            outMessage = ex.Message + "Permission to upload file denied"; outChar = "!";
            return;
        }
    }

    public void saveExcelData(string fileName, string Upload_ID)
    {
        fileName = Server.MapPath(WebConfigurationManager.AppSettings["folderPath"].ToString() + "/" + fileName); // @"F:\Chrome DM\Help Files\Raw Input File.xlsx";       
        var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
        var ds = new DataSet();
        using (var conn = new OleDbConnection(connectionString))
        {
            conn.Open();
            var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] WHERE F1 ";
                var adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(ds);
            }
        }
        #region Save Data to Excel
        int i = 0;
        RawData objRawData = new RawData();
        for (i = 1; i < ds.Tables[0].Rows.Count; i++)
        {
            if (ds.Tables[0].Rows[i]["F1"].ToString().Trim() == "")
                continue;
            objRawData.Upload_id = Upload_ID;
            objRawData.RecordID = "0";
            objRawData.Assign_Min = ds.Tables[0].Rows[i]["F1"].ToString();
            objRawData.Channel = ds.Tables[0].Rows[i]["F2"].ToString();
            objRawData.Story = ds.Tables[0].Rows[i]["F3"].ToString();
            objRawData.Sub_Story = ds.Tables[0].Rows[i]["F4"].ToString();
            objRawData.Story_Genre_1 = ds.Tables[0].Rows[i]["F5"].ToString();
            objRawData.Story_Genre_2 = ds.Tables[0].Rows[i]["F6"].ToString();
            objRawData.Pgm_Name = ds.Tables[0].Rows[i]["F7"].ToString();
            objRawData.Pgm_Start_Time = ds.Tables[0].Rows[i]["F8"].ToString();
            objRawData.Pgm_End_Time = ds.Tables[0].Rows[i]["F9"].ToString();
            objRawData.Clip_Start_Time = ds.Tables[0].Rows[i]["F10"].ToString();
            objRawData.Clip_End_Time = ds.Tables[0].Rows[i]["F11"].ToString();
            objRawData.Pgm_Date = ds.Tables[0].Rows[i]["F12"].ToString();
            objRawData.Week = ds.Tables[0].Rows[i]["F13"].ToString();
            objRawData.Week_Day = ds.Tables[0].Rows[i]["F14"].ToString();
            objRawData.Pgm_Hour = ds.Tables[0].Rows[i]["F15"].ToString();
            objRawData.Geography = ds.Tables[0].Rows[i]["F16"].ToString();
            objRawData.Duration = ds.Tables[0].Rows[i]["F17"].ToString();
            objRawData.Duration_Seconds = ds.Tables[0].Rows[i]["F18"].ToString();
            objRawData.Duration2 = ds.Tables[0].Rows[i]["F19"].ToString();
            objRawData.Personality = ds.Tables[0].Rows[i]["F20"].ToString();
            objRawData.Guest = ds.Tables[0].Rows[i]["F21"].ToString();
            objRawData.Anchor = ds.Tables[0].Rows[i]["F22"].ToString();
            objRawData.Reporter = ds.Tables[0].Rows[i]["F23"].ToString();
            objRawData.Logistics = ds.Tables[0].Rows[i]["F24"].ToString();
            objRawData.Telecast_Format = ds.Tables[0].Rows[i]["F25"].ToString();
            objRawData.Assist_Used = ds.Tables[0].Rows[i]["F26"].ToString();
            objRawData.Split = ds.Tables[0].Rows[i]["F27"].ToString();
            objRawData.Story_Format = ds.Tables[0].Rows[i]["F28"].ToString();
            objRawData.HSM = ds.Tables[0].Rows[i]["F29"].ToString();
            objRawData.HSM_Urban = ds.Tables[0].Rows[i]["F30"].ToString();
            objRawData.HSM_Rural = ds.Tables[0].Rows[i]["F31"].ToString();
            objRawData.HSM_NTVT = ds.Tables[0].Rows[i]["F32"].ToString();
            objRawData.HSM_Urban_NTVT = ds.Tables[0].Rows[i]["F33"].ToString();
            objRawData.HSM_Rural_NTVT = ds.Tables[0].Rows[i]["F34"].ToString();
            objRawData.Hour = ds.Tables[0].Rows[i]["F35"].ToString();
            objRawData.State = ds.Tables[0].Rows[i]["F36"].ToString();
            objRawData.Live_Coverage = ds.Tables[0].Rows[i]["F37"].ToString();
            objRawData.InsertRawData();
        }
        if(i==ds.Tables[0].Rows.Count)
        {
            string strSQL = "EXEC [proc_Uploaded_Files] @action='sync-completed', @upload_id ='" + Upload_ID + "'" +
                           ",@synchronize_status='2',@synchronize_date='" + System.DateTime.Now.ToString() + "',@synchronize_by='" + Session["userID"].ToString() + "'";
            DataTable dt = new DataTable();
            dt = objDataAccess.GetDataTable(strSQL);
            lblMessage.Text = (dt.Rows.Count > 0) ? dt.Rows[0]["Result"].ToString() : "";
            lblMessage.ForeColor = (dt.Rows.Count > 0) ? Color.Blue : Color.Red;
            getUnSyncFiles();
        }
        #endregion        
    }

    protected void gvUploadFiles_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        string strSQL = "";
        DataTable dt = new DataTable();
        if (e.CommandName == "DeleteNow")
        {
            strSQL = "EXEC [proc_Uploaded_Files] @action='delete', @upload_id ='" + e.CommandArgument.ToString() + "'";
            dt = objDataAccess.GetDataTable(strSQL);
            lblMessage.Text = (dt.Rows.Count > 0) ? dt.Rows[0]["Result"].ToString() : "";
            lblMessage.ForeColor = (dt.Rows.Count > 0) ? Color.Blue : Color.Red;
            getUnSyncFiles();
        }
        else
            if (e.CommandName == "SyncNow")
        {
            strSQL = "EXEC [proc_Uploaded_Files] @action='sync-start', @upload_id ='" + e.CommandArgument.ToString() + "'" +
                ",@synchronize_status='1',@synchronize_date='" + System.DateTime.Now.ToString() + "',@synchronize_by='" + Session["userID"].ToString() + "'";
            dt = objDataAccess.GetDataTable(strSQL);
            lblMessage.Text = (dt.Rows.Count > 0) ? dt.Rows[0]["Result"].ToString() : "";
            lblMessage.ForeColor = (dt.Rows.Count > 0) ? Color.Blue : Color.Red;
            getUnSyncFiles();
            saveExcelData(dt.Rows[0]["system_file_name"].ToString(), e.CommandArgument.ToString());
        }
    }
}