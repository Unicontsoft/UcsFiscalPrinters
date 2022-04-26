IF OBJECT_ID('usp_sys_ServiceBrokerSetupService') IS NOT NULL DROP PROC usp_sys_ServiceBrokerSetupService
GO
-- exec usp_sys_ServiceBrokerSetupService @Mode = 'DROP_EXISTING'
-- exec usp_sys_ServiceBrokerSetupService @Mode = 'DROP_ONLY'
-- exec usp_sys_ServiceBrokerSetupService @Mode = 'DROP_ONLY', @ProcessID = 59
-- exec usp_sys_ServiceBrokerSetupService @Mode = 'DROP_ONLY', @ProcessID = 58
-- exec usp_sys_ServiceBrokerSetupService @Mode = 'DROP_ONLY', @ProcessID = 55

CREATE PROC usp_sys_ServiceBrokerSetupService (
            @QueueName      SYSNAME         = NULL OUTPUT 
            , @SvcName      SYSNAME         = NULL OUTPUT 
            , @Mode         VARCHAR(20)     = NULL
) WITH EXECUTE AS OWNER AS
/*------------------------------------------------------------------------
'
' UcsFPHub (c) 2019-2020 by Unicontsoft
'
' Unicontsoft Fiscal Printers Hub
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
*/------------------------------------------------------------------------
SET         NOCOUNT ON

DECLARE     @SQL            NVARCHAR(MAX)
            , @Handle       UNIQUEIDENTIFIER
            , @CrsClean     CURSOR
            , @Name         SYSNAME
            --- used to be params
            , @Prefix       SYSNAME
            , @Suffix       SYSNAME
            , @ProcessID    INT

SELECT      @Mode = COALESCE(@Mode, '')
            , @ProcessID = COALESCE(@ProcessID, @@SPID)
            , @Prefix = COALESCE(@Prefix, N'UcsFpInitiator')
            , @Suffix = COALESCE(@Suffix, CONVERT(NVARCHAR(50), @ProcessID))
            , @QueueName = COALESCE(@QueueName, @Prefix + N'Queue/' + @Suffix)
            , @SvcName = COALESCE(@SvcName, @Prefix + N'Service/' + @Suffix)

IF EXISTS (SELECT 0 FROM sys.databases WHERE database_id = DB_ID() AND is_broker_enabled = 0)
BEGIN
            SET         @SQL = N'BEGIN TRY ALTER DATABASE ' + QUOTENAME(DB_NAME()) + N' SET ENABLE_BROKER WITH ROLLBACK IMMEDIATE END TRY BEGIN CATCH END CATCH'
            EXEC        (@SQL)
END

IF EXISTS (SELECT 0 FROM sys.services WHERE name = @SvcName) AND @Mode IN ('DROP_SERVICE', 'DROP_EXISTING', 'DROP_ONLY')
BEGIN
            SET         @CrsClean = CURSOR FAST_FORWARD FOR 
            SELECT      conversation_handle 
            FROM        sys.conversation_endpoints
            WHERE       service_id IN (SELECT service_id FROM sys.services WHERE name = @SvcName)

            OPEN        @CrsClean

            WHILE       1=1
            BEGIN
                        FETCH NEXT  FROM @CrsClean INTO @Handle
                        IF          @@FETCH_STATUS <> 0 BREAK

                        ; END       CONVERSATION @Handle WITH CLEANUP
            END

            CLOSE       @CrsClean
            DEALLOCATE  @CrsClean

            IF  EXISTS (SELECT 0 FROM sys.services WHERE name = @SvcName)
            BEGIN
                        SET         @SQL = N'DROP SERVICE ' + QUOTENAME(@SvcName)
                        EXEC        (@SQL)
            END
END

IF EXISTS (SELECT 0 FROM sys.service_queues WHERE SCHEMA_NAME(schema_id) = N'dbo' AND name = @QueueName) AND @Mode IN ('DROP_EXISTING', 'DROP_ONLY')
BEGIN
            SET         @CrsClean = CURSOR FAST_FORWARD FOR 
            SELECT      name
            FROM        sys.services
            WHERE       service_queue_id IN (SELECT object_id FROM sys.service_queues WHERE SCHEMA_NAME(schema_id) = N'dbo' AND name = @QueueName)

            OPEN        @CrsClean

            WHILE       1=1
            BEGIN
                        FETCH NEXT  FROM @CrsClean INTO @Name
                        IF          @@FETCH_STATUS <> 0 BREAK

                        IF  EXISTS (SELECT 0 FROM sys.services WHERE name = @Name)
                        BEGIN
                                    SET         @SQL = N'DROP SERVICE ' + QUOTENAME(@Name)
                                    EXEC        (@SQL)
                        END
            END

            CLOSE       @CrsClean
            DEALLOCATE  @CrsClean

            IF EXISTS (SELECT 0 FROM sys.service_queues WHERE SCHEMA_NAME(schema_id) = N'dbo' AND name = @QueueName)
            BEGIN
                        SET         @SQL = N'DROP QUEUE dbo.' + QUOTENAME(@QueueName)
                        EXEC        (@SQL)
            END
END

IF          @Mode NOT IN ('DROP_ONLY')
BEGIN
            IF NOT EXISTS (SELECT 0 FROM sys.service_queues WHERE SCHEMA_NAME(schema_id) = N'dbo' AND name = @QueueName)
            BEGIN
                        SET         @SQL = N'CREATE QUEUE dbo.' + QUOTENAME(@QueueName)
                        EXEC        (@SQL)

                        SET         @SQL = N'GRANT RECEIVE ON dbo.' + QUOTENAME(@QueueName) + N' TO public'
                        EXEC        (@SQL)
            END

            IF NOT EXISTS (SELECT 0 FROM sys.services WHERE name = @SvcName)
            BEGIN
                        SET         @SQL = N'CREATE SERVICE ' + QUOTENAME(@SvcName) + N' ON QUEUE dbo.' + QUOTENAME(@QueueName) + N' ([DEFAULT])'
                        EXEC        (@SQL)
            END
END
GO
