IF OBJECT_ID('usp_sys_ServiceBrokerSend') IS NOT NULL DROP PROC usp_sys_ServiceBrokerSend
GO
/*
DECLARE     @Handle1    UNIQUEIDENTIFIER
            , @Result   INT
            , @Response NVARCHAR(MAX)

EXEC        dbo.usp_sys_ServiceBrokerSetupService

EXEC        @Result = dbo.usp_sys_ServiceBrokerSend 'UcsFpTargetService/DT123456', 'test_status_request', @Response = @Response OUTPUT, @Handle = @Handle1 OUTPUT
SELECT      @Result, @Response

-- close conversation
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend 'UcsFpTargetService/DT123456',  @Handle = @Handle1 OUTPUT

*/

/*
DECLARE     @Handle2    UNIQUEIDENTIFIER
            , @Result   INT
            , @Response NVARCHAR(MAX)

EXEC        dbo.usp_sys_ServiceBrokerSetupService


WHILE       1=1
BEGIN
            SET         @Handle2 = NULL
            EXEC        @Result = dbo.usp_sys_ServiceBrokerSend 'UcsFpTargetService/DT234567', 'test_status_request', @Response = @Response OUTPUT, @Handle = @Handle2 OUTPUT
            SELECT      @Result, @Response

            EXEC        @Result = dbo.usp_sys_ServiceBrokerSend 'UcsFpTargetService/DT234567',  @Handle = @Handle2 OUTPUT

            WAITFOR DELAY '0:0:1'
END
*/

CREATE PROC usp_sys_ServiceBrokerSend (
            @TargetSvc      SYSNAME
            , @Request      NVARCHAR(MAX)       = NULL
            , @Response     NVARCHAR(MAX)       = NULL OUTPUT
            , @QueueName    SYSNAME             = NULL
            , @SvcName      SYSNAME             = NULL
            , @Timeout      INT                 = NULL
            , @Handle       UNIQUEIDENTIFIER    = NULL OUTPUT
) AS
/*------------------------------------------------------------------------
'
' UcsFPHub (c) 2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Hub
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
*/------------------------------------------------------------------------
SET         NOCOUNT ON

DECLARE     @RetVal     INT
            , @SQL      NVARCHAR(MAX)
            , @MsgType  SYSNAME
            , @Lifetime INT

SELECT      @QueueName = COALESCE(@QueueName, 'UcsFpInitiator' + 'Queue/' + CONVERT(VARCHAR(50), @@SPID))
            , @SvcName = COALESCE(@SvcName, 'UcsFpInitiator' + 'Service/' + CONVERT(VARCHAR(50), @@SPID))
            , @Timeout = COALESCE(@Timeout, 30000)
            , @Lifetime = 4 * (@Timeout / 1000) -- 30 sec -> 2 min
            , @RetVal = 0

BEGIN TRY
            IF          @Handle IS NULL
            BEGIN
                        BEGIN DIALOG CONVERSATION @Handle
                        FROM        SERVICE @SvcName
                        TO          SERVICE @TargetSvc, 'CURRENT DATABASE'
                        WITH        ENCRYPTION = OFF, LIFETIME = @Lifetime
            END

            IF          @Request IS NULL
            BEGIN
                        SET         @Request = N'__FIN__'
            END

            ; SEND ON   CONVERSATION @Handle (@Request)

            SET         @SQL = N'
            WAITFOR (   RECEIVE     TOP (1) @Response = CONVERT(NVARCHAR(MAX), message_body) 
                                    , @MsgType = message_type_name
                        FROM        ' + QUOTENAME(@QueueName) + N'  ), TIMEOUT ' + CONVERT(NVARCHAR(50), @Timeout)

RepeatWait:
            SELECT      @Response = NULL, @MsgType = NULL
            EXEC        dbo.sp_executesql @SQL
                            , N'@Response NVARCHAR(MAX) OUTPUT, @MsgType SYSNAME OUTPUT'
                            , @Response OUTPUT, @MsgType OUTPUT

            IF          @MsgType IS NULL
            BEGIN
                        --PRINT { fn CURRENT_TIMESTAMP } + ': Timeout'

                        IF          @Request = '__FIN__'
                        BEGIN
                                    ; END       CONVERSATION @Handle
                        END

                        SELECT      @RetVal = 99
                        GOTO        QH
            END

            --PRINT { fn CURRENT_TIMESTAMP } + ': @MsgType=' + @MsgType

            IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/EndDialog'
            BEGIN
                        ; END       CONVERSATION @Handle

                        SELECT      @RetVal = 1
                        GOTO        QH
            END

            IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/Error'
            BEGIN
                        ; END       CONVERSATION @Handle

                        SELECT      @RetVal = 1
                        GOTO        QH
            END

            IF          @Request IS NOT NULL
            BEGIN
                        IF          @Response <> N'__ACK__'
                        BEGIN
                                    GOTO        RepeatWait
                        END

                        --- get response only when not terminating conversatiopn
                        IF          @Request <> '__FIN__'
                        BEGIN
                                    SET         @Request = NULL
                                    GOTO        RepeatWait
                        END

                        ; END       CONVERSATION @Handle
            END
END TRY
BEGIN CATCH
            --PRINT { fn CURRENT_TIMESTAMP } + ': ERROR_MESSAGE=' + ERROR_MESSAGE()

            ; END       CONVERSATION @Handle

            SELECT      @RetVal = 2
            GOTO        QH
END CATCH

QH:
RETURN      @RetVal
GO
