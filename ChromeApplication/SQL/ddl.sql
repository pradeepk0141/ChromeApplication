SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_RawData](
	[RecordID] [int] IDENTITY(1,1) NOT NULL,
	[upload_id] [int] NULL,
	[Assign_Min] [nvarchar](max) NULL,
	[Channel] [nvarchar](max) NULL,
	[Story] [nvarchar](max) NULL,
	[Sub_Story] [nvarchar](max) NULL,
	[Story_Genre_1] [nvarchar](max) NULL,
	[Story_Genre_2] [nvarchar](max) NULL,
	[Pgm_Name] [nvarchar](max) NULL,
	[Pgm_Start_Time] [nvarchar](max) NULL,
	[Pgm_End_Time] [nvarchar](max) NULL,
	[Clip_Start_Time] [nvarchar](max) NULL,
	[Clip_End_Time] [nvarchar](max) NULL,
	[Pgm_Date] [nvarchar](max) NULL,
	[Week] [nvarchar](max) NULL,
	[Week_Day] [nvarchar](max) NULL,
	[Pgm_Hour] [nvarchar](max) NULL,
	[Geography] [nvarchar](max) NULL,
	[Duration] [nvarchar](max) NULL,
	[Duration_Seconds] [nvarchar](max) NULL,
	[Duration2] [nvarchar](max) NULL,
	[Personality] [nvarchar](max) NULL,
	[Guest] [nvarchar](max) NULL,
	[Anchor] [nvarchar](max) NULL,
	[Reporter] [nvarchar](max) NULL,
	[Logistics] [nvarchar](max) NULL,
	[Telecast_Format] [nvarchar](max) NULL,
	[Assist_Used] [nvarchar](max) NULL,
	[Split] [nvarchar](max) NULL,
	[Story_Format] [nvarchar](max) NULL,
	[HSM] [nvarchar](max) NULL,
	[HSM_Urban] [nvarchar](max) NULL,
	[HSM_Rural] [nvarchar](max) NULL,
	[HSM_NTVT] [nvarchar](max) NULL,
	[HSM_Urban_NTVT] [nvarchar](max) NULL,
	[HSM_Rural_NTVT] [nvarchar](max) NULL,
	[Hour] [nvarchar](max) NULL,
	[State] [nvarchar](50) NULL,
	[Live_Coverage] [nvarchar](50) NULL,
 CONSTRAINT [PK_tb_RawData] PRIMARY KEY CLUSTERED 
(
	[RecordID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tb_RawData_Master]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_RawData_Master](
	[Data_ID] [int] IDENTITY(1,1) NOT NULL,
	[RecordID] [int] NOT NULL,
	[Assign_Min] [time](7) NULL,
	[Channel] [nvarchar](50) NULL,
	[Story] [nvarchar](max) NULL,
	[Sub_Story] [nvarchar](max) NULL,
	[Story_Genre_1] [nvarchar](max) NULL,
	[Story_Genre_2] [nvarchar](max) NULL,
	[Pgm_Name] [nvarchar](max) NULL,
	[Pgm_Start_Time] [time](7) NULL,
	[Pgm_End_Time] [time](7) NULL,
	[Clip_Start_Time] [time](7) NULL,
	[Clip_End_Time] [time](7) NULL,
	[Pgm_Date] [date] NULL,
	[Week] [int] NULL,
	[Week_Day] [nvarchar](50) NULL,
	[Pgm_Hour] [int] NULL,
	[Geography] [nvarchar](50) NULL,
	[Duration] [time](7) NULL,
	[Duration_Seconds] [int] NULL,
	[Duration2] [int] NULL,
	[Personality] [nvarchar](max) NULL,
	[Guest] [nvarchar](max) NULL,
	[Anchor] [nvarchar](max) NULL,
	[Reporter] [nvarchar](max) NULL,
	[Logistics] [nvarchar](50) NULL,
	[Telecast_Format] [nvarchar](50) NULL,
	[Assist_Used] [nvarchar](50) NULL,
	[Split] [tinyint] NULL,
	[Story_Format] [nvarchar](50) NULL,
	[HSM] [decimal](18, 6) NULL,
	[HSM_Urban] [decimal](18, 6) NULL,
	[HSM_Rural] [decimal](18, 6) NULL,
	[HSM_NTVT] [decimal](18, 6) NULL,
	[HSM_Urban_NTVT] [decimal](18, 6) NULL,
	[HSM_Rural_NTVT] [decimal](18, 6) NULL,
	[Hour] [int] NULL,
	[State] [nvarchar](50) NULL,
	[Live_Coverage] [nvarchar](50) NULL,
 CONSTRAINT [PK_tb_RawData_Master] PRIMARY KEY CLUSTERED 
(
	[Data_ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[tb_Uploaded_Files]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tb_Uploaded_Files](
	[upload_id] [int] IDENTITY(1,1) NOT NULL,
	[Selected_date] [date] NULL,
	[original_file_name] [nvarchar](max) NULL,
	[system_file_name] [nvarchar](50) NULL,
	[upload_date] [datetime] NULL CONSTRAINT [DF_tb_Uploaded_Files_upload_date]  DEFAULT (getdate()),
	[upload_by] [int] NULL,
	[synchronize_status] [tinyint] NULL CONSTRAINT [DF_tb_Uploaded_Files_synchronize_status]  DEFAULT ((0)),
	[synchronize_date] [datetime] NULL,
	[synchronize_by] [nchar](10) NULL,
 CONSTRAINT [PK_tb_Uploaded_Files] PRIMARY KEY CLUSTERED 
(
	[upload_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  StoredProcedure [dbo].[proc_AnchorROI]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO