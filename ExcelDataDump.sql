
the purpose of this proc is to populate the data required for excel dump for an employee for a specific period of time. This 
would dynamically change the number of columns in the final result set. 


USE #####
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


    ----------------- exec [####].[usp_rpt_####_data_ExcelData_final] '54380',4090,'CABLE\vpulij001c'
CREATE proc [####].[usp_rpt_####_data_ExcelData_final] 
@Employeeid as varchar(40) 
,@Period as int
,@NTLogin Varchar(20)
as

Set NOCOUNT ON;

--Declare @Employeeid as varchar(40) = 56829 ,@Period as int = 4091 ,@NTLogin Varchar(20) = 'CABLE\vpulij001c'

Declare @RequestorEmployeeID int, @RequestorRole Varchar(20)
Declare @Startdate Date, @StopDate Date

Select @StartDate = Startdate, @StopDate = Enddate from 
coe_dw.ncoe.dimfiscalperiods where Fiscalperiodid = @Period

--Drop table ######_Temp_ExcelDataDumpTable_FallOut_Orders

------------------------------Select @Startdate, @Stopdate

IF OBJECT_ID('tempdb..#Details') IS NOT NULL		
DROP TABLE #Details
IF OBJECT_ID('tempdb..######_Temp_ExcelDataDumpTable_Final') IS NOT NULL		
Drop Table ######_Temp_ExcelDataDumpTable_Final
--IF OBJECT_ID('tempdb..#ReportDetails ') IS NOT NULL		
--DROP TABLE #ReportDetails 
IF OBJECT_ID('tempdb..##calc1') IS NOT NULL		
DROP TABLE ##calc1


select             
@RequestorEmployeeID=employeeid,
@RequestorRole =EmployeeRole
From ##Schema.###EmployeeTable
where NTLoginID=@NTLogin and @stopdate Between Startdate and isnull(StopDate,@StopDate)


;with emp as
(
Select Distinct epm.employeeserviceid, es.teamid
from ####.Table#### epm (nolock) inner join 
##Schema.###EmployeeTable es (nolock) on epm.employeeserviceid = es.employeeserviceid
where es.employeeid = @Employeeid 
and epm.fiscalperiodid = @Period
and 1 = case when @RequestorRole = 'Agent' and es.EmployeeID = @RequestorEmployeeID then 1 else case when  @RequestorRole <> 'Agent' then 1 else 0 end END 
)

--Select * from emp

,datadumpid as
(
Select Distinct de.ExcelDataDumpid , de.teamid from 
####.##ExcelTable de inner join emp on de.teamid = emp.teamid
and @period between FiscalPeriodId_start and isnull(fiscalperiodid_Stop, @Period)
)

--Select * from datadumpid 

,Details as
(
Select Distinct 
Sno = ROW_NUMBER() over (order by dd.Measure),
dd.Measure, Query, d.teamid,FieldList, CreateTable
--,concat('######_Temp_ExcelDataDumpTable_' ,dd.Measure) TableName
,TempTableName TableName
,whereclause , GroupByClause 
from datadumpid d inner join ####.##ExcelDetails dd on d.ExcelDataDumpid = dd.ExcelDataDumpid
)

Select * into #Details from details

--Select * From #Details 

Declare @Sno int ,@teamid int, @Measure varchar(200), @Query NVARCHAR(4000), @SQL NVARCHAR(4000) , 
@Startdate_Sql varchar(max), @StopDate_Sql varchar(max) , @teamid_Sql varchar(max)  , @CreateTable nvarchar(max)
,@Whereclause NVARCHAR(max) , @WhereclauseNow NVARCHAR(max),@WhereclauseNowSql NVARCHAR(max)
, @GroupByClause NVARCHAR(max) = '', @GroupByClauseNow NVARCHAR(max),@GroupByClauseNowSql NVARCHAR(max)


Declare @TableSearch int , @tablename varchar(200), @SelectQuery Nvarchar(max), @TruncateQuery Nvarchar(max), @Droptable Nvarchar(max)
    

Set @Startdate_Sql = @startdate
set @StopDate_Sql = @stopdate
--set @Teamid_SQl = @teamid


declare cur CURSOR LOCAL for
    Select Sno from #Details --where sno = 5
open cur
fetch next from cur into @Sno
while @@FETCH_STATUS = 0 
BEGIN

--Print @sno

--Set @Startdate_Sql = @startdate
--set @StopDate_Sql = @stopdate

    
    Select @Measure = Measure, @Query = Query, @Teamid_SQl = teamid , @CreateTable = CreateTable,@TableName  = tablename  
    , @whereclause = whereclause , @GroupbyClause = GroupByClause
    from #Details where Sno = @Sno

--    set @tableName = concat('######_Temp_ExcelDataDumpTable_' ,@Measure)
    Set @SelectQuery = concat('Select * from ######_Temp_ExcelDataDumpTable_',@Measure)
    Set @TruncateQuery  = concat('Truncate table ######_Temp_ExcelDataDumpTable_',@Measure)

                    -----------Select @tablename
                    ------------Select @SelectQuery 

    set @Droptable = concat('IF OBJECT_ID(','''tempdb..',@tableName,'''',') IS NOT NULL		Drop Table ',@tableName)     

                
    EXECUTE sp_ExecuteSQL @Droptable
    
EXECUTE sp_ExecuteSQL @CreateTable
    ----            end
   
    --------------------            Select @Whereclause 
    
 if @Whereclause = 1
    begin 
       Set @WhereclauseNow = CONCAT('''', @Startdate_Sql ,''' and ','''', @StopDate_Sql ,''' and es.employeeid = ', @Employeeid , ' and teamid = ', @Teamid_SQl)
    end
else if @whereclause = 2
    begin
        Set @WhereclauseNow = concat(' isnull(e.StopDate,','''',@StopDate_Sql,''') >= ','''',@Startdate_Sql,''' and StartDate <= ','''',@StopDate_Sql,''' and es.employeeid = ',@Employeeid)
    end
else if @whereclause = 3
    begin
      Set @WhereclauseNow = CONCAT('''', @Startdate_Sql ,''' and ','''', @StopDate_Sql ,''' and es.employeeid = ', @Employeeid)
    end
else if @whereclause = 4
    begin
      Set @WhereclauseNow = CONCAT('''', @Startdate_Sql ,''' and ','''', @StopDate_Sql ,''' and a.employeeid = ', @Employeeid)
    end

    Set @GroupByClauseNow  = Concat(' ',@GroupbyClause,'')
    
    Set @SQL = N''+ @Query + @WhereclauseNow+@GroupByClauseNow  

--Select  @SQL
----------------------------    Select @TruncateQuery  

EXECUTE sp_ExecuteSQL @TruncateQuery  
EXECUTE sp_ExecuteSQL @SQL

----------------------------    EXECUTE sp_ExecuteSQL @SelectQuery

fetch next from cur into @Sno
END
close cur
deallocate cur


/******************** part 2 ******************************/


----------------------------    @Sno int, 
----------------------------    ,@CreateTable nvarchar(max)
----------------------------    @TableName varchar(200) , @MeasureT varchar(100),

Declare @ColumNames nvarchar(max),      @FirstRowFieldList varchar(max),    @ColumnCount int = 0
Declare  @FiledList nvarchar(max)  , @MeasureT  varchar(2000)     
Declare @MeasureFieldList as Table (Sno int,Row_Num int, TableName varchar(200), Measure varchar(200),Field varchar(500),MeasureField varchar(500)
,SelectStatement nvarchar(4000), FirstRowList varchar(4000))
Declare @MeasureFieldonly as Table (Row_Num int, Field varchar(50))
Declare   @FieldList varchar(max) --, @SelectFieldList varchar(max)
Declare @SelectFieldList varchar(max) , @insertFieldList varchar(max) ,@CreateFieldList varchar(max)
    

declare cur CURSOR LOCAL for
    Select Sno from #Details --where sno = 5
open cur
fetch next from cur into @Sno
while @@FETCH_STATUS = 0 
BEGIN
     --set @sno = 1
    Select @TableName = tablename , @MeasureT = Measure , @FieldList = FieldList from #Details where sno = @Sno

    Set @ColumNames = concat('SELECT Distinct Column_id,Name FROM tempdb.sys.columns  WHERE object_id = OBJECT_ID(''','tempdb..',@tablename,''')',' Order by Column_id')
-----------------------Select @ColumNames 
    Insert into  @MeasureFieldonly
    EXECUTE sp_ExecuteSQL @ColumNames

        ----------------------------    SELECT * FROM @MeasureFieldonly
        ----------------------------    Select @ColumnCount 
        ----------------------------    select max(Row_Num) from  @MeasureFieldonly

    insert into @MeasureFieldList (sno, Row_num, tablename, measure, field,MeasureField )--,SelectStatement)
    Select @sno, Row_num, @TableName,@measuret,Field,concat('[',@measuret,'_',Field,']', ' varchar(500)') 
       from @MeasureFieldonly Order By Row_num

        ----------------------------    Select * from @MeasureFieldList 
        ----------------------------    Select @FirstRowFieldList=STUFF((SELECT  ','+concat('''[',@sno,'_',Field,']''') 
        
        Select @FirstRowFieldList=STUFF((SELECT  ','+concat('''',Field,'''') 
                from @MeasureFieldonly --where sno = @sno
                    group by Field ,Row_num
                    order by Row_num
            FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)') 
       ,1,1,'')
        
            ----------------------------    Select @FirstRowFieldList

        Select @ColumnCount = case when @ColumnCount > max(Row_Num) then @ColumnCount else  max(Row_Num) end from  @MeasureFieldonly

    Select @SelectFieldList =STUFF((SELECT  ','+ Field
                        from @MeasureFieldList where sno = @sno
                    group by MeasureField ,Field ,Row_num
                    order by Row_num
            FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)') 
       ,1,1,'')

            ----------------------------    Select @SelectFieldList

    Select @insertFieldList =STUFF((SELECT ','+concat('[Column','_',ROW_num,']')
                        from @MeasureFieldList where sno = @sno
                    group by Row_num
                    order by Row_num
            FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)') 
       ,1,1,'') 
    
            ----------------------------Select @insertFieldList
    Select @CreateFieldList  =STUFF((SELECT ','+concat('[Column','_',ROW_num,'] varchar(500)')
                        from @MeasureFieldList where sno = @sno
                    group by Row_num
                    order by Row_num
            FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)') 
       ,1,1,'') 
    

        Select @FiledList = case when @ColumnCount > max(Row_Num) then @FiledList else  @CreateFieldList end from  @MeasureFieldonly
   ------------------Select @ColumnCount ,@FiledList

    update @MeasureFieldList 
    set SelectStatement = concat('insert into ######_Temp_ExcelDataDumpTable_Final (sno,measure,Periodid,employeeid,',@insertFieldList,')' ,' Select ',@Sno,',''',@measuret,''',',@period,',',@employeeid,',',@SelectFieldList, ' from ',  @tablename)
    , FirstRowList= concat('insert into ######_Temp_ExcelDataDumpTable_Final  (sno,measure,Periodid,employeeid,',@insertFieldList,')' ,' Select ',@Sno,',''',@measuret,''',',@period,',',@employeeid,',', @FirstRowFieldList)
  where sno = @sno

            ---Select SelectStatement  from  @MeasureFieldList


Delete from @MeasureFieldonly

fetch next from cur into @Sno
END
close cur
deallocate cur

Set @createtable = concat('create table ######_Temp_ExcelDataDumpTable_Final (ExcelDataDumpTable_Finalid Int IDENTITY(1,1) NOT NULL,sno int,measure varchar(200),periodid int,employeeid varchar(50), ',@FiledList,')')

EXECUTE sp_ExecuteSQL @createtable 

Declare @SelectStatement_firsrow nvarchar(max), @SelectStatement_calc1 nvarchar(max), @FinalSelect_list Nvarchar(max), @FinalSelect Nvarchar(max)

declare cur CURSOR LOCAL for
    Select Distinct Sno from @MeasureFieldList --where sno = 3
open cur
fetch next from cur into @Sno
while @@FETCH_STATUS = 0 
BEGIN

    Set @SelectStatement_firsrow  = (Select Distinct  FirstRowList   from  @MeasureFieldList where sno = @Sno)
    Set @SelectStatement_calc1  = (Select Distinct  SelectStatement   from  @MeasureFieldList where sno = @Sno)

   EXECUTE sp_ExecuteSQL @SelectStatement_firsrow
   EXECUTE sp_ExecuteSQL @SelectStatement_calc1  


fetch next from cur into @Sno
END
close cur
deallocate cur


/******************* Part 3 **********************/


Select * from ######_Temp_ExcelDataDumpTable_Final
Order by ExcelDataDumpTable_Finalid
