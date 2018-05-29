CREATE PROC [dbo].[proc_AnchorROI]
AS
--SELECT Channel,Anchor,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Logistics ORDER BY Channel,Anchor DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6)
SELECT Channel,Anchor,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #AnchorData
FROM tb_RawData_Master GROUP BY Channel,Anchor 
--SELECT * FROM #AnchorData ORDER BY Channel,Anchor
SELECT DISTINCT Anchor  into #AnchorTable FROM #AnchorData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #AnchorData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #AnchorTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #AnchorTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #AnchorTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #AnchorTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #AnchorTable lt JOIN 
					(SELECT Channel,Anchor,Duration2,HSM_NTVT FROM #AnchorData) tbl1
					ON lt.[Anchor]=tbl1.Anchor and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #AnchorData WHERE Channel=@Channel GROUP BY Channel
		SET @sql = 'update #AnchorTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	

		print @SUM_HSM_NTVT
		print @SUM_Duration2
		print @sql
		EXEC(@sql)
		--SELECT SUM(Duration2),SUM(HSM_NTVT) FROM #AnchorData WHERE Channel=@Channel GROUP BY Channel
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels

SELECT * FROM #AnchorTable ORDER BY Anchor
DROP TABLE #AnchorTable


GO
/****** Object:  StoredProcedure [dbo].[proc_DoDStoryTrend]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[proc_DoDStoryTrend]
AS
--SELECT Channel,Story,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Story ORDER BY Channel,Story DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6),@OrderColumns NVARCHAR(1000)=''
SELECT Channel,Story,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #StoryData
FROM tb_RawData_Master GROUP BY Channel,Story 
--SELECT * FROM #StoryData ORDER BY Channel,Story
SELECT DISTINCT Story  into #StoryTable FROM #StoryData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #StoryData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #StoryTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	SET @OrderColumns=@OrderColumns+'['+@Channel+'_Duration2] +'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #StoryTable lt JOIN 
					(SELECT Channel,Story,Duration2,HSM_NTVT FROM #StoryData) tbl1
					ON lt.[Story]=tbl1.Story and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #StoryData WHERE Channel=@Channel GROUP BY Channel
		print @SUM_HSM_NTVT
		print @SUM_Duration2
		SET @sql = 'update #StoryTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		--print @sql
		EXEC(@sql)
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels
SET @OrderColumns=SUBSTRING(@OrderColumns,1,LEN(@OrderColumns)-1)
--print @OrderColumns
SET @sql = 'SELECT TOP 25 * FROM #StoryTable ORDER BY ('+@OrderColumns+') DESC'
print @sql
EXEC(@sql)
DROP TABLE #StoryTable






GO
/****** Object:  StoredProcedure [dbo].[proc_EXCLUSIVE]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[proc_EXCLUSIVE]
@ChannelName NVARCHAR(50)='AAJ TAK'
AS
--SELECT Channel,Story,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Story ORDER BY Channel,Story DESC
declare @Channel nvarchar(50),@sql NVARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6),@ExclusiveColumns NVARCHAR(1000)=''
SELECT Channel,Story,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #StoryData
FROM tb_RawData_Master GROUP BY Channel,Story 
--SELECT * FROM #StoryData ORDER BY Channel,Story
SELECT DISTINCT Story  into #StoryTable FROM #StoryData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #StoryData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #StoryTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
		--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #StoryTable lt JOIN 
					(SELECT Channel,Story,Duration2,HSM_NTVT FROM #StoryData) tbl1
					ON lt.[Story]=tbl1.Story and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		--SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #StoryData WHERE Channel=@Channel GROUP BY Channel
		--print @SUM_HSM_NTVT
		--print @SUM_Duration2
		--SET @sql = 'update #StoryTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		--,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		--print @sql
		--EXEC(@sql)
		IF(@Channel<>@ChannelName)
		BEGIN
			IF (@ExclusiveColumns<>'')
				SET @ExclusiveColumns=@ExclusiveColumns+' AND '
			SET @ExclusiveColumns=@ExclusiveColumns+'CAST(['+@Channel+'_Duration2] AS INT) =0'
		END	
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels
print @ExclusiveColumns
--Exclusive Coverage
			--SET @sql='SELECT @SUM_Duration2=SUM(['+@ChannelName+'_Duration2]),@SUM_HSM_NTVT=SUM(['+@ChannelName+'_HSM NTVT])
			SET @sql='SELECT @SUM_Duration2=SUM(['+@ChannelName+'_Duration2]) FROM #StoryTable WHERE '+@ExclusiveColumns
			--print @sql			
			exec sp_executesql @sql, N'@SUM_Duration2 decimal out', @SUM_Duration2 out
			SET @sql='SELECT @SUM_HSM_NTVT=SUM(['+@ChannelName+'_HSM NTVT]) FROM #StoryTable WHERE '+@ExclusiveColumns
			exec sp_executesql @sql, N'@SUM_HSM_NTVT decimal out', @SUM_HSM_NTVT out
			--select @SUM_Duration2
			--SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #StoryData WHERE Channel=@Channel GROUP BY Channel
			print @SUM_HSM_NTVT
			print @SUM_Duration2
			SET @sql = 'update #StoryTable set ['+@ChannelName+'_SOV] = ['+@ChannelName+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
			,['+@ChannelName+'_Efficiency Index (EI)] = (['+@ChannelName+'_HSM NTVT]/ISNULL(NULLIF(['+@ChannelName+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)
			WHERE '+@ExclusiveColumns	
			print @sql
			EXEC(@sql)
--End of Exclusive coverage
SET @sql = 'SELECT * FROM #StoryTable WHERE '+@ExclusiveColumns
print @sql
EXEC(@sql)
DROP TABLE #StoryTable





GO
/****** Object:  StoredProcedure [dbo].[proc_GenreROI]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[proc_GenreROI]
AS
--SELECT Channel,Story_Genre_1,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Logistics ORDER BY Channel,Story_Genre_1 DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6)
SELECT Channel,Story_Genre_1,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #Story_Genre_1Data
FROM tb_RawData_Master GROUP BY Channel,Story_Genre_1 
--SELECT * FROM #Story_Genre_1Data ORDER BY Channel,Story_Genre_1
SELECT DISTINCT Story_Genre_1  into #Story_Genre_1Table FROM #Story_Genre_1Data

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #Story_Genre_1Data ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #Story_Genre_1Table ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Story_Genre_1Table ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Story_Genre_1Table ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Story_Genre_1Table ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #Story_Genre_1Table lt JOIN 
					(SELECT Channel,Story_Genre_1,Duration2,HSM_NTVT FROM #Story_Genre_1Data) tbl1
					ON lt.[Story_Genre_1]=tbl1.Story_Genre_1 and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #Story_Genre_1Data WHERE Channel=@Channel GROUP BY Channel
		SET @sql = 'update #Story_Genre_1Table set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		print @sql
		EXEC(@sql)
		--SELECT SUM(Duration2),SUM(HSM_NTVT) FROM #Story_Genre_1Data WHERE Channel=@Channel GROUP BY Channel
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels

SELECT * FROM #Story_Genre_1Table ORDER BY Story_Genre_1
DROP TABLE #Story_Genre_1Table





GO
/****** Object:  StoredProcedure [dbo].[proc_LogisticsROI]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[proc_LogisticsROI]
AS
--SELECT Channel,Logistics,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Logistics ORDER BY Channel,Logistics DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6)
SELECT Channel,Logistics,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #LogisticsData
FROM tb_RawData_Master GROUP BY Channel,Logistics 
--SELECT * FROM #LogisticsData ORDER BY Channel,Logistics
SELECT DISTINCT Logistics  into #LogisticsTable FROM #LogisticsData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #LogisticsData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #LogisticsTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #LogisticsTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #LogisticsTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #LogisticsTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #LogisticsTable lt JOIN 
					(SELECT Channel,Logistics,Duration2,HSM_NTVT FROM #LogisticsData) tbl1
					ON lt.[Logistics]=tbl1.Logistics and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #LogisticsData WHERE Channel=@Channel GROUP BY Channel
		SET @sql = 'update #LogisticsTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		print @sql
		EXEC(@sql)
		--SELECT SUM(Duration2),SUM(HSM_NTVT) FROM #LogisticsData WHERE Channel=@Channel GROUP BY Channel
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels

SELECT * FROM #LogisticsTable ORDER BY Logistics
DROP TABLE #LogisticsTable




GO
/****** Object:  StoredProcedure [dbo].[proc_ProgROI]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROC [dbo].[proc_ProgROI]
AS
--SELECT Channel,Pgm_Name,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Logistics ORDER BY Channel,Pgm_Name DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6)
SELECT Channel,Pgm_Name,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #Pgm_NameData
FROM tb_RawData_Master GROUP BY Channel,Pgm_Name 
--SELECT * FROM #Pgm_NameData ORDER BY Channel,Pgm_Name
SELECT DISTINCT Pgm_Name  into #Pgm_NameTable FROM #Pgm_NameData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #Pgm_NameData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #Pgm_NameTable lt JOIN 
					(SELECT Channel,Pgm_Name,Duration2,HSM_NTVT FROM #Pgm_NameData) tbl1
					ON lt.[Pgm_Name]=tbl1.Pgm_Name and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #Pgm_NameData WHERE Channel=@Channel GROUP BY Channel
		SET @sql = 'update #Pgm_NameTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		print @sql
		EXEC(@sql)
		--SELECT SUM(Duration2),SUM(HSM_NTVT) FROM #Pgm_NameData WHERE Channel=@Channel GROUP BY Channel
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels

SELECT * FROM #Pgm_NameTable ORDER BY Pgm_Name
DROP TABLE #Pgm_NameTable




GO
/****** Object:  StoredProcedure [dbo].[proc_StoryROI]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[proc_StoryROI]
AS
--SELECT Channel,Story,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Story ORDER BY Channel,Story DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6),@OrderColumns NVARCHAR(1000)=''
SELECT Channel,Story,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #StoryData
FROM tb_RawData_Master GROUP BY Channel,Story 
--SELECT * FROM #StoryData ORDER BY Channel,Story
SELECT DISTINCT Story  into #StoryTable FROM #StoryData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #StoryData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #StoryTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	SET @OrderColumns=@OrderColumns+'['+@Channel+'_Duration2] +'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #StoryTable lt JOIN 
					(SELECT Channel,Story,Duration2,HSM_NTVT FROM #StoryData) tbl1
					ON lt.[Story]=tbl1.Story and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #StoryData WHERE Channel=@Channel GROUP BY Channel
		print @SUM_HSM_NTVT
		print @SUM_Duration2
		SET @sql = 'update #StoryTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		--print @sql
		EXEC(@sql)
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels
SET @OrderColumns=SUBSTRING(@OrderColumns,1,LEN(@OrderColumns)-1)
--print @OrderColumns
SET @sql = 'SELECT * FROM #StoryTable ORDER BY ('+@OrderColumns+') DESC'
print @sql
EXEC(@sql)
DROP TABLE #StoryTable




GO
/****** Object:  StoredProcedure [dbo].[proc_tb_RawData]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROC [dbo].[proc_tb_RawData]
@action NVARCHAR(50)
,@RecordID  int 
,@upload_id  int   
,@Assign_Min  NVARCHAR(MAX) 
,@Channel  NVARCHAR(MAX) 
,@Story  NVARCHAR(MAX) 
,@Sub_Story  NVARCHAR(MAX) 
,@Story_Genre_1  NVARCHAR(MAX) 
,@Story_Genre_2  NVARCHAR(MAX) 
,@Pgm_Name  NVARCHAR(MAX) 
,@Pgm_Start_Time  NVARCHAR(MAX) 
,@Pgm_End_Time  NVARCHAR(MAX) 
,@Clip_Start_Time  NVARCHAR(MAX) 
,@Clip_End_Time  NVARCHAR(MAX) 
,@Pgm_Date  NVARCHAR(MAX) 
,@Week  NVARCHAR(MAX) 
,@Week_Day  NVARCHAR(MAX) 
,@Pgm_Hour  NVARCHAR(MAX) 
,@Geography  NVARCHAR(MAX) 
,@Duration  NVARCHAR(MAX) 
,@Duration_Seconds  NVARCHAR(MAX) 
,@Duration2  NVARCHAR(MAX) 
,@Personality  NVARCHAR(MAX) 
,@Guest  NVARCHAR(MAX) 
,@Anchor  NVARCHAR(MAX) 
,@Reporter  NVARCHAR(MAX) 
,@Logistics  NVARCHAR(MAX) 
,@Telecast_Format  NVARCHAR(MAX) 
,@Assist_Used  NVARCHAR(MAX) 
,@Split  NVARCHAR(MAX) 
,@Story_Format  NVARCHAR(MAX) 
,@HSM  NVARCHAR(MAX) 
,@HSM_Urban  NVARCHAR(MAX) 
,@HSM_Rural  NVARCHAR(MAX) 
,@HSM_NTVT  NVARCHAR(MAX) 
,@HSM_Urban_NTVT  NVARCHAR(MAX) 
,@HSM_Rural_NTVT  NVARCHAR(MAX) 
,@Hour  NVARCHAR(MAX) 
,@State  NVARCHAR(MAX) 
,@Live_Coverage  NVARCHAR(MAX) 
AS

IF (@action='select')
BEGIN
	SELECT upload_id,CONVERT(NVARCHAR(50),Selected_date,106) AS Selected_date, original_file_name, system_file_name
	, CONVERT(NVARCHAR(50),upload_date,100) AS upload_date, upload_by
	, synchronize_status
	, synchronize_date, synchronize_by   FROM tb_Uploaded_Files --WHERE upload_by=case when @upload_by=0 then upload_by else @upload_by end
	ORDER BY upload_date DESC 
END
ELSE
IF (@action='insert')
BEGIN
	INSERT INTO tb_RawData([upload_id],[Assign_Min],[Channel],[Story],[Sub_Story],[Story_Genre_1],[Story_Genre_2],[Pgm_Name],[Pgm_Start_Time]
	,[Pgm_End_Time]	,[Clip_Start_Time],[Clip_End_Time],[Pgm_Date],[Week],[Week_Day],[Pgm_Hour],[Geography],[Duration],[Duration_Seconds]
	,[Duration2],[Personality],[Guest],[Anchor],[Reporter],[Logistics],[Telecast_Format],[Assist_Used],[Split],[Story_Format],[HSM]
	,[HSM_Urban],[HSM_Rural],[HSM_NTVT],[HSM_Urban_NTVT],[HSM_Rural_NTVT],[Hour], [State], [Live_Coverage]) 
	VALUES(@upload_id, @Assign_Min, @Channel, @Story, @Sub_Story, @Story_Genre_1, @Story_Genre_2, @Pgm_Name, @Pgm_Start_Time
	, @Pgm_End_Time, @Clip_Start_Time, @Clip_End_Time, @Pgm_Date, @Week, @Week_Day, @Pgm_Hour, @Geography, @Duration, @Duration_Seconds
	, @Duration2, @Personality, @Guest, @Anchor, @Reporter, @Logistics, @Telecast_Format, @Assist_Used, @Split, @Story_Format, @HSM
	, @HSM_Urban, @HSM_Rural, @HSM_NTVT, @HSM_Urban_NTVT, @HSM_Rural_NTVT, @Hour, @State, @Live_Coverage) 
	SELECT SCOPE_IDENTITY()
END




GO
/****** Object:  StoredProcedure [dbo].[proc_TelecastFormatROI]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROC [dbo].[proc_TelecastFormatROI]
AS
--SELECT Channel,Telecast_Format,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Logistics ORDER BY Channel,Telecast_Format DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6)
SELECT Channel,Telecast_Format,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #Telecast_FormatData
FROM tb_RawData_Master GROUP BY Channel,Telecast_Format 
--SELECT * FROM #Telecast_FormatData ORDER BY Channel,Telecast_Format
SELECT DISTINCT Telecast_Format  into #Telecast_FormatTable FROM #Telecast_FormatData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #Telecast_FormatData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #Telecast_FormatTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Telecast_FormatTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Telecast_FormatTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Telecast_FormatTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #Telecast_FormatTable lt JOIN 
					(SELECT Channel,Telecast_Format,Duration2,HSM_NTVT FROM #Telecast_FormatData) tbl1
					ON lt.[Telecast_Format]=tbl1.Telecast_Format and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #Telecast_FormatData WHERE Channel=@Channel GROUP BY Channel
			SET @sql = 'update #Telecast_FormatTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		print @sql
		EXEC(@sql)
		--SELECT SUM(Duration2),SUM(HSM_NTVT) FROM #Telecast_FormatData WHERE Channel=@Channel GROUP BY Channel
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels

SELECT * FROM #Telecast_FormatTable ORDER BY Telecast_Format
DROP TABLE #Telecast_FormatTable




GO
/****** Object:  StoredProcedure [dbo].[proc_Top10News]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROC [dbo].[proc_Top10News]
AS
--SELECT Channel,Story,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Story ORDER BY Channel,Story DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6),@OrderColumns NVARCHAR(1000)=''
SELECT Channel,Story,SUM(HSM_NTVT) AS HSM_NTVT,SUM(Duration2) Duration2 
into #StoryData
FROM tb_RawData_Master GROUP BY Channel,Story 
--SELECT * FROM #StoryData ORDER BY Channel,Story
SELECT DISTINCT Story  into #StoryTable FROM #StoryData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #StoryData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #StoryTable ADD ['+@Channel+'_HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #StoryTable ADD ['+@Channel+'_Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
	SET @OrderColumns=@OrderColumns+'['+@Channel+'_Duration2] +'
	--print @sql
  EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_HSM NTVT] = tbl1.HSM_NTVT, ['+@Channel+'_Duration2]=tbl1.Duration2
		            from #StoryTable lt JOIN 
					(SELECT Channel,Story,Duration2,HSM_NTVT FROM #StoryData) tbl1
					ON lt.[Story]=tbl1.Story and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_NTVT=SUM(HSM_NTVT),@SUM_Duration2=SUM(Duration2) FROM #StoryData WHERE Channel=@Channel GROUP BY Channel
		print @SUM_HSM_NTVT
		print @SUM_Duration2
		SET @sql = 'update #StoryTable set ['+@Channel+'_SOV] = ['+@Channel+'_Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'
		,['+@Channel+'_Efficiency Index (EI)] = (['+@Channel+'_HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		--print @sql
		EXEC(@sql)
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels
SET @OrderColumns=SUBSTRING(@OrderColumns,1,LEN(@OrderColumns)-1)
--print @OrderColumns
SET @sql = 'SELECT TOP 10 * FROM #StoryTable ORDER BY ('+@OrderColumns+') DESC'
print @sql
EXEC(@sql)
DROP TABLE #StoryTable





GO
/****** Object:  StoredProcedure [dbo].[proc_Uploaded_Files]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROC [dbo].[proc_Uploaded_Files]
@action NVARCHAR(50)
, @upload_id INT
, @Selected_date NVARCHAR(50)=null
, @original_file_name NVARCHAR(MAX)=''
, @system_file_name NVARCHAR(50)=''
, @upload_date DATETIME=null
, @upload_by INT=0
, @synchronize_status INT=0
, @synchronize_date DATETIME=null
, @synchronize_by INT =0

AS

IF (@action='select')
BEGIN
	SELECT upload_id,CONVERT(NVARCHAR(50),Selected_date,106) AS Selected_date, original_file_name, system_file_name
	, upload_date AS upload_date, upload_by
	, synchronize_status
	, synchronize_date, synchronize_by   FROM tb_Uploaded_Files WHERE upload_by=case when @upload_by=0 then upload_by else @upload_by end
	ORDER BY upload_date DESC 
END
ELSE
IF (@action='insert')
BEGIN
	INSERT INTO tb_Uploaded_Files(Selected_date, original_file_name, system_file_name, upload_date, upload_by, synchronize_status)
	VALUES(@Selected_date, @original_file_name, @system_file_name, GETDATE(), @upload_by, 0) 
	SELECT SCOPE_IDENTITY()
END
ELSE
IF (@action='sync-completed' AND @synchronize_status=2)
BEGIN
	--Bulk Data Insert
	INSERT INTO tb_RawData_Master([RecordID],[Assign_Min],[Channel],[Story],[Sub_Story],[Story_Genre_1],[Story_Genre_2],[Pgm_Name],[Pgm_Start_Time]
	,[Clip_Start_Time],[Clip_End_Time],[Pgm_Date],[Week],[Week_Day],[Pgm_Hour],[Geography]
	,[Duration],[Duration_Seconds],[Duration2],[Personality],[Guest],[Anchor],[Reporter],[Logistics],[Telecast_Format]
	,[Assist_Used],[Split],[Story_Format],[HSM],[HSM_Urban],[HSM_Rural],[HSM_NTVT],[HSM_Urban_NTVT],
	[HSM_Rural_NTVT],[Hour],[State],[Live_Coverage]) 
	SELECT [RecordID],CAST([Assign_Min]AS time),[Channel],[Story],[Sub_Story],[Story_Genre_1],[Story_Genre_2],[Pgm_Name],CAST([Pgm_Start_Time] AS TIME)
	,CAST([Clip_Start_Time] AS TIME),CAST([Clip_End_Time] AS TIME),CAST([Pgm_Date] AS DATE),CAST([Week] AS INT),[Week_Day],CAST([Pgm_Hour] AS INT),[Geography]
	,CAST([Duration] AS TIME),CAST([Duration_Seconds] AS INT),CAST([Duration2] AS INT),[Personality],[Guest],[Anchor],[Reporter],[Logistics],[Telecast_Format]
	,[Assist_Used],CAST([Split] AS INT),[Story_Format]	,CAST(CAST(REPLACE([HSM],',','') as xml).value('. cast as xs:decimal?','decimal(18,6)') AS DECIMAL(18,6))
	,CAST(CAST(REPLACE(REPLACE([HSM_Urban],',',''),',','') as xml).value('. cast as xs:decimal?','decimal(18,6)') AS DECIMAL(18,6))
	,CAST(CAST(REPLACE(REPLACE([HSM_Rural],',',''),',','') as xml).value('. cast as xs:decimal?','decimal(18,6)') AS DECIMAL(18,6))
	,CAST(CAST(REPLACE(REPLACE([HSM_NTVT],',',''),',','') as xml).value('. cast as xs:decimal?','decimal(18,6)') AS DECIMAL(18,6))
	,CAST(CAST(REPLACE(REPLACE([HSM_Urban_NTVT],',',''),',','') as xml).value('. cast as xs:decimal?','decimal(18,6)') AS DECIMAL(18,6))
	,CAST(CAST(REPLACE(REPLACE([HSM_Rural_NTVT],',',''),',','') as xml).value('. cast as xs:decimal?','decimal(18,6)') AS DECIMAL(18,6))
	,[Hour],[State],[Live_Coverage] FROM tb_RawData WHERE upload_id=@upload_id
	--End of Bulk Data Insert
	UPDATE tb_Uploaded_Files SET Selected_date=case when @Selected_date IS NULL then Selected_date else @Selected_date  end
	, original_file_name=case when @original_file_name='' then original_file_name else @original_file_name  end
	, system_file_name=case when @system_file_name='' then system_file_name else @system_file_name  end
	, upload_date=case when @upload_date IS NULL then upload_date else @upload_date  end
	, upload_by=case when @upload_by=0 then upload_by else @upload_by  end
	, synchronize_status=case when @synchronize_status=0 then synchronize_status else @synchronize_status  end
	, synchronize_date=case when @synchronize_date IS NULL then synchronize_date else @synchronize_date end
	, synchronize_by=case when @synchronize_by=0 then synchronize_by else @synchronize_by  end 
	WHERE upload_id=@upload_id
	
	SELECT 'Synchronization Completed Successfully For '+ original_file_name  AS Result,system_file_name FROM tb_Uploaded_Files WHERE upload_id=@upload_id	
END
ELSE
IF (@action='sync-start' AND @synchronize_status=1)
BEGIN
	DELETE tb_RawData WHERE upload_id=@upload_id
	
	UPDATE tb_Uploaded_Files SET 
	Selected_date=case when @Selected_date IS NULL then Selected_date else @Selected_date  end
	, original_file_name=case when @original_file_name='' then original_file_name else @original_file_name  end
	, system_file_name=case when @system_file_name='' then system_file_name else @system_file_name  end
	, upload_date=case when @upload_date IS NULL then upload_date else @upload_date  end
	, upload_by=case when @upload_by=0 then upload_by else @upload_by  end
	, synchronize_status=case when @synchronize_status=0 then synchronize_status else @synchronize_status  end
	, synchronize_date=case when @synchronize_date IS NULL then synchronize_date else @synchronize_date end
	, synchronize_by=case when @synchronize_by=0 then synchronize_by else @synchronize_by  end 
	WHERE upload_id=@upload_id
	
	SELECT 'Synchronization Started For '+original_file_name AS Result,system_file_name FROM tb_Uploaded_Files WHERE upload_id=@upload_id
END
ELSE
IF (@action='delete')
BEGIN
	DELETE tb_RawData WHERE upload_id=@upload_id
	
	DELETE tb_Uploaded_Files WHERE upload_id=@upload_id
	
	SELECT 'Selected File Deleted Successfully' AS Result
END



GO
/****** Object:  StoredProcedure [dbo].[proc_URSplitRoI]    Script Date: 29 May 2018 18:24:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROC [dbo].[proc_URSplitRoI]
AS
--SELECT Channel,Pgm_Name,SUM(Duration2) DurationTotal,SUM(HSM_NTVT) FROM tb_RawData_Master GROUP BY Channel,Logistics ORDER BY Channel,Pgm_Name DESC
declare @Channel nvarchar(50),@sql VARCHAR(4000),@SUM_HSM_NTVT DECIMAL(18,6),@SUM_Duration2 DECIMAL(18,6),
@SUM_HSM_Urban_NTVT DECIMAL(18,6),@SUM_HSM_Rural_NTVT DECIMAL(18,6)
SELECT Channel,Pgm_Name,SUM(HSM_NTVT) AS HSM_NTVT,SUM(HSM_Urban_NTVT) HSM_Urban_NTVT,SUM(HSM_Rural_NTVT) AS HSM_Rural_NTVT
,SUM(Duration2) Duration2 into #Pgm_NameData
FROM tb_RawData_Master GROUP BY Channel,Pgm_Name 
--SELECT * FROM #Pgm_NameData ORDER BY Channel,Pgm_Name
SELECT DISTINCT Pgm_Name  into #Pgm_NameTable FROM #Pgm_NameData

DECLARE c_channels CURSOR LOCAL FAST_FORWARD FOR
SELECT DISTINCT Channel FROM #Pgm_NameData ORDER  BY Channel
OPEN c_channels
FETCH c_channels INTO @Channel
WHILE (@@FETCH_STATUS = 0)
Begin
--print @Channel-- HSM NTVT	 Duration2
SET @Sql = 'ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Urban HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES			
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Urban Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES			
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Urban SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Urban Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Rural HSM NTVT] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES	
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Rural Duration2] DECIMAL(18,0) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Rural SOV] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES
			ALTER TABLE #Pgm_NameTable ADD ['+@Channel+'_Rural Efficiency Index (EI)] DECIMAL(18,6) DEFAULT 0.00 WITH VALUES'
			--print @sql
EXEC (@Sql)
SET @Sql = 'update lt set ['+@Channel+'_Urban HSM NTVT] = tbl1.HSM_Urban_NTVT ,['+@Channel+'_Urban Duration2] = tbl1.Duration2
						 ,['+@Channel+'_Rural HSM NTVT] = tbl1.HSM_Rural_NTVT,['+@Channel+'_Rural Duration2] = tbl1.Duration2  
						 from #Pgm_NameTable lt JOIN (SELECT Channel,Pgm_Name,Duration2,HSM_Urban_NTVT,HSM_Rural_NTVT FROM #Pgm_NameData) tbl1
						 ON lt.[Pgm_Name]=tbl1.Pgm_Name and Tbl1.channel='''+@Channel+''''
					--print @sql
		EXEC (@Sql)
		SELECT @SUM_HSM_Urban_NTVT=SUM(HSM_Urban_NTVT),@SUM_HSM_Rural_NTVT=SUM(HSM_Rural_NTVT),@SUM_Duration2=SUM(Duration2) 
		FROM #Pgm_NameData WHERE Channel=@Channel GROUP BY Channel			
		SET @sql = 'update #Pgm_NameTable set ['+@Channel+'_Urban SOV] = ['+@Channel+'_Urban Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+
		',['+@Channel+'_Rural SOV] = ['+@Channel+'_Rural Duration2]/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+
		',['+@Channel+'_Urban Efficiency Index (EI)] = (['+@Channel+'_Urban HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Urban Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_Urban_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)
		 ,['+@Channel+'_Rural Efficiency Index (EI)] = (['+@Channel+'_Rural HSM NTVT]/ISNULL(NULLIF(['+@Channel+'_Rural Duration2],0),1))/ISNULL(NULLIF(('+CAST(@SUM_HSM_Rural_NTVT AS NVARCHAR(50))+'/'+CAST(ISNULL(NULLIF(@SUM_Duration2,0),1) AS NVARCHAR(50))+'),0),1)'	
		print @sql		
		EXEC(@sql)
FETCH c_channels INTO @Channel
end
CLOSE c_channels
DEALLOCATE c_channels

SELECT * FROM #Pgm_NameTable ORDER BY Pgm_Name
DROP TABLE #Pgm_NameTable
DROP TABLE #Pgm_NameData
GO

