using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for RawData
/// </summary>
public class RawData
{
    private string recordID;
    private string upload_id;
    private string assign_Min;
    private string channel;
    private string story;
    private string sub_Story;
    private string story_Genre_1;
    private string story_Genre_2;
    private string pgm_Name;
    private string pgm_Start_Time;
    private string pgm_End_Time;
    private string clip_Start_Time;
    private string clip_End_Time;
    private string pgm_Date;
    private string week;
    private string week_Day;
    private string pgm_Hour;
    private string geography;
    private string duration;
    private string duration_Seconds;
    private string duration2;
    private string personality;
    private string guest;
    private string anchor;
    private string reporter;
    private string logistics;
    private string telecast_Format;
    private string assist_Used;
    private string split;
    private string story_Format;
    private string hSM;
    private string hSM_Urban;
    private string hSM_Rural;
    private string hSM_NTVT;
    private string hSM_Urban_NTVT;
    private string hSM_Rural_NTVT;
    private string hour;
    private string state;
    private string live_Coverage;
    DataAccess objDataAccess = new DataAccess();
    public RawData()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    public string RecordID { get => recordID; set => recordID = value; }
    public string Upload_id { get => upload_id; set => upload_id = value; }
    public string Assign_Min { get => assign_Min; set => assign_Min = value; }
    public string Channel { get => channel; set => channel = value; }
    public string Story { get => story; set => story = value; }
    public string Sub_Story { get => sub_Story; set => sub_Story = value; }
    public string Story_Genre_1 { get => story_Genre_1; set => story_Genre_1 = value; }
    public string Story_Genre_2 { get => story_Genre_2; set => story_Genre_2 = value; }
    public string Pgm_Name { get => pgm_Name; set => pgm_Name = value; }
    public string Pgm_Start_Time { get => pgm_Start_Time; set => pgm_Start_Time = value; }
    public string Pgm_End_Time { get => pgm_End_Time; set => pgm_End_Time = value; }
    public string Clip_Start_Time { get => clip_Start_Time; set => clip_Start_Time = value; }
    public string Clip_End_Time { get => clip_End_Time; set => clip_End_Time = value; }
    public string Pgm_Date { get => pgm_Date; set => pgm_Date = value; }
    public string Week { get => week; set => week = value; }
    public string Week_Day { get => week_Day; set => week_Day = value; }
    public string Pgm_Hour { get => pgm_Hour; set => pgm_Hour = value; }
    public string Geography { get => geography; set => geography = value; }
    public string Duration { get => duration; set => duration = value; }
    public string Duration_Seconds { get => duration_Seconds; set => duration_Seconds = value; }
    public string Duration2 { get => duration2; set => duration2 = value; }
    public string Personality { get => personality; set => personality = value; }
    public string Guest { get => guest; set => guest = value; }
    public string Anchor { get => anchor; set => anchor = value; }
    public string Reporter { get => reporter; set => reporter = value; }
    public string Logistics { get => logistics; set => logistics = value; }
    public string Telecast_Format { get => telecast_Format; set => telecast_Format = value; }
    public string Assist_Used { get => assist_Used; set => assist_Used = value; }
    public string Split { get => split; set => split = value; }
    public string Story_Format { get => story_Format; set => story_Format = value; }
    public string HSM { get => hSM; set => hSM = value; }
    public string HSM_Urban { get => hSM_Urban; set => hSM_Urban = value; }
    public string HSM_Rural { get => hSM_Rural; set => hSM_Rural = value; }
    public string HSM_NTVT { get => hSM_NTVT; set => hSM_NTVT = value; }
    public string HSM_Urban_NTVT { get => hSM_Urban_NTVT; set => hSM_Urban_NTVT = value; }
    public string HSM_Rural_NTVT { get => hSM_Rural_NTVT; set => hSM_Rural_NTVT = value; }
    public string Hour { get => hour; set => hour = value; }
    public string State { get => state; set => state = value; }
    public string Live_Coverage { get => live_Coverage; set => live_Coverage = value; }
    public int InsertRawData()
    {
        string strSQL = "EXEC [proc_tb_RawData] @action='insert', @RecordID='0' , @upload_id = '" + upload_id + "'" +
            ", @Assign_Min='" + Assign_Min + "',@Channel='" + Channel + "',@Story='" + Story + "',@Sub_Story='" + Sub_Story + "'" +
            ",@Story_Genre_1='" + Story_Genre_1 + "',@Story_Genre_2='" + Story_Genre_2 + "',@Pgm_Name='" + Pgm_Name + "'" +
            ",@Pgm_Start_Time='" + Pgm_Start_Time + "',@Pgm_End_Time='" + Pgm_End_Time + "',@Clip_Start_Time='" + Clip_Start_Time + "'" +
            ",@Clip_End_Time='" + Clip_End_Time + "',@Pgm_Date='" + Pgm_Date + "',@Week='" + Week + "',@Week_Day='" + Week_Day + "'" +
            ",@Pgm_Hour='" + Pgm_Hour + "',@Geography='" + Geography + "',@Duration='" + Duration + "'" +
            ",@Duration_Seconds='" + Duration_Seconds + "',@Duration2='" + Duration2 + "',@Personality='" + Personality + "'" +
            ",@Guest='" + Guest + "',@Anchor='" + Anchor + "',@Reporter='" + Reporter + "',@Logistics='" + Logistics + "'" +
            ",@Telecast_Format='" + Telecast_Format + "',@Assist_Used='" + Assist_Used + "',@Split='" + Split + "'" +
            ",@Story_Format='" + Story_Format + "',@HSM='" + HSM + "',@HSM_Urban='" + HSM_Urban + "',@HSM_Rural='" + HSM_Rural + "'" +
            ",@HSM_NTVT='" + HSM_NTVT + "',@HSM_Urban_NTVT='" + HSM_Urban_NTVT + "',@HSM_Rural_NTVT='" + HSM_Rural_NTVT + "'" +
            ",@Hour='" + Hour + "',@State='" + State + "',@Live_Coverage='" + Live_Coverage + "'";
        return objDataAccess.ExecuteQuery(strSQL);
    }
}