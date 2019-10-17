IF OBJECT_ID('usp_sys_ServiceBrokerWaitRequest') IS NOT NULL DROP PROC usp_sys_ServiceBrokerWaitRequest
GO
/*
DECLARE     @QueueName      SYSNAME
            , @Handle       UNIQUEIDENTIFIER
            , @Request      NVARCHAR(MAX)
            , @MsgType      SYSNAME
            , @SvcName      SYSNAME
            , @ErrorText    NVARCHAR(255)
            , @Result       INT

SELECT      @QueueName = 'UcsFpTargetQueue/POS2-PC'

EXEC        dbo.usp_sys_ServiceBrokerSetupService @QueueName, 'UcsFpTargetService/DT123456', 'DROP_EXISTING'
EXEC        dbo.usp_sys_ServiceBrokerSetupService @QueueName, 'UcsFpTargetService/DT518315', 'DROP_SERVICE'

WHILE       1=1
BEGIN
            EXEC        @Result = dbo.usp_sys_ServiceBrokerWaitRequest @QueueName, 5000, @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT, @ErrorText OUTPUT
            SELECT      @Result AS Result, @Handle AS Handle, @Request AS Request, @MsgType AS MsgType, @SvcName AS SvcName, @ErrorText AS ErrorText

            RAISERROR ('Result=%d', 10, 0, @Result) WITH NOWAIT
END
*/

CREATE PROC usp_sys_ServiceBrokerWaitRequest (
            @QueueName      SYSNAME             = NULL
            , @Timeout      INT                 = NULL
            , @Handle       UNIQUEIDENTIFIER    = NULL OUTPUT
            , @Request      NVARCHAR(MAX)       = NULL OUTPUT
            , @MsgType      SYSNAME             = NULL OUTPUT
            , @SvcName      SYSNAME             = NULL OUTPUT
            , @ErrorText    NVARCHAR(255)       = NULL OUTPUT
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

DECLARE     @RetVal         INT
            , @SQL          NVARCHAR(MAX)

SELECT      @RetVal = 0
            , @Timeout = COALESCE(@Timeout, 5000)

SELECT      @SQL = N'
WAITFOR (   RECEIVE     TOP (1) @Handle = conversation_handle 
                        , @Request = CONVERT(NVARCHAR(MAX), message_body)
                        , @MsgType = message_type_name
                        , @SvcName = service_name
            FROM        dbo.' + QUOTENAME(@QueueName) + N'  ), TIMEOUT ' + CONVERT(NVARCHAR(50), @Timeout)
RepeatWait:
SELECT      @Handle = NULL, @Request = NULL, @MsgType = NULL, @SvcName = NULL, @ErrorText = NULL
EXEC        dbo.sp_executesql @SQL
                , N'@Handle UNIQUEIDENTIFIER OUTPUT, @Request NVARCHAR(MAX) OUTPUT, @MsgType SYSNAME OUTPUT, @SvcName SYSNAME OUTPUT'
                , @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT

IF          @Handle IS NULL
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP() } + ': Timeout'

            SELECT      @RetVal = 99
                        , @ErrorText = N'Timeout'
            GOTO        QH
END
                        
IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/EndDialog'
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP() } + ': Closed by ' + @MsgType

            ; END       CONVERSATION @Handle

            GOTO        RepeatWait
END

IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/Error'
BEGIN
            ; END       CONVERSATION @Handle

            SELECT      @RetVal = 1
                        , @ErrorText = LEFT(CONVERT(XML, @Request).value('declare namespace ns="http://schemas.microsoft.com/SQL/ServiceBroker/Error";
                                                                            (//ns:Description)[1]', 'NVARCHAR(MAX)'), 255)
            GOTO        QH
END

IF          @Request = N'__FIN__'
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP() } + ': Conversation closed'

            ; END       CONVERSATION @Handle

            GOTO        RepeatWait
END

IF          @Request = N'__PING__'
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP() } + ': Ping reply send'

            ; SEND ON CONVERSATION @Handle (N'__PONG__')

            GOTO        RepeatWait
END

; SEND ON   CONVERSATION @Handle (N'__ACK__')

QH:
RETURN      @RetVal
GO
