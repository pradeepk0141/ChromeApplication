using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Data.SqlClient;
using System.Data;

/// <summary>
/// Summary description for DataAccess
/// </summary>
public class DataAccess
{
    public DataAccess()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    private string getConnectionString()
    {
        string str = WebConfigurationManager.ConnectionStrings["Sqlonnection"].ConnectionString;
        return str;
    }
    public DataTable GetDataTable(string strSQL)
    {
        SqlConnection con = new SqlConnection(getConnectionString());
        try
        {
            DataTable dt = new DataTable();
            SqlCommand cmd = new SqlCommand(strSQL, con);
            if (con.State == ConnectionState.Closed)
                con.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            dt.Load(dr);
            return dt;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            con.Close();
        }
    }
    public int ExecuteQuery(string strSQL)
    {
        SqlConnection con = new SqlConnection(getConnectionString());
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmd = new SqlCommand(strSQL, con);
            if (con.State == ConnectionState.Closed)
                con.Open();
            int i = cmd.ExecuteNonQuery();
            return i;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            con.Close();
        }
    }
}