IF OBJECT_ID('usp_sys_ServiceBrokerWaitRequest') IS NOT NULL DROP PROC usp_sys_ServiceBrokerWaitRequest
GO
/*
DECLARE     @QueueName      SYSNAME
            , @Handle       UNIQUEIDENTIFIER
            , @Request      NVARCHAR(MAX)
            , @MsgType      SYSNAME
            , @SvcName      SYSNAME
            , @Result       INT

SELECT      @QueueName = 'UcsFpTargetQueue/POS2-PC'

EXEC        dbo.usp_sys_ServiceBrokerSetupService @QueueName, 'UcsFpTargetService/DT123456', 'DROP_EXISTING'
EXEC        dbo.usp_sys_ServiceBrokerSetupService @QueueName, 'UcsFpTargetService/DT234567', 'DROP_SERVICE'

WHILE       1=1
BEGIN
            SELECT      @Handle = NULL, @Request = NULL, @MsgType = NULL
            EXEC        @Result = dbo.usp_sys_ServiceBrokerWaitRequest @QueueName, 5000, @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT

            IF          @Result = 0 AND @Handle IS NOT NULL
                        SELECT      @Result AS Result, @Handle AS Handle, @Request AS Request, @MsgType AS MsgType, @SvcName AS SvcName

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
            FROM        ' + QUOTENAME(@QueueName) + N'  ), TIMEOUT ' + CONVERT(NVARCHAR(50), @Timeout)
EXEC        dbo.sp_executesql @SQL
                , N'@Handle UNIQUEIDENTIFIER OUTPUT, @Request NVARCHAR(MAX) OUTPUT, @MsgType SYSNAME OUTPUT, @SvcName SYSNAME OUTPUT'
                , @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT
            
IF          @Handle IS NULL
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP() } + ': Timeout'

            SELECT      @RetVal = 99
            GOTO        QH
END
                        
IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/EndDialog'
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP() } + ': Closed by ' + @MsgType

            END         CONVERSATION @Handle --WITH CLEANUP

            SELECT      @RetVal = 1
            GOTO        QH
END

; SEND ON   CONVERSATION @Handle (N'__ACK__')

IF          @Request = N'__FIN__'
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP() } + ': Conversation closed'

            END         CONVERSATION @Handle --WITH CLEANUP

            SELECT      @RetVal = 1
            GOTO        QH
END

IF          @Request = N'__PING__'
BEGIN
            ; SEND ON CONVERSATION @Handle (N'__PONG__')

            SELECT      @RetVal = 1
            GOTO        QH
END

QH:
RETURN      @RetVal
GO
