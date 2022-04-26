IF OBJECT_ID('usp_sys_ServiceBrokerSend') IS NOT NULL DROP PROC usp_sys_ServiceBrokerSend
GO
/*
DECLARE     @Handle1    UNIQUEIDENTIFIER
            , @Result   INT
            , @Response NVARCHAR(MAX)

--- setup receiving queue for current connection once on re-connect
EXEC        dbo.usp_sys_ServiceBrokerSetupService

--- open conversation and send request
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend '{ "Url": "/printers/DT518315/status" }', @Response OUTPUT, @TargetSvc = 'UcsFpTargetService/DT518315', @Handle = @Handle1 OUTPUT
SELECT      @Result AS Result, @Response AS Response
--{  
--   "Ok":true,
--   "DeviceStatus":"",
--   "DeviceDateTime":"2019-07-23 17:17:18"
--}

--- send another request on the same conversation (result in XML)
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend '{ "Url": "/printers?format=xml" }', @Response OUTPUT, @Handle = @Handle1
SELECT      @Result AS Result, CONVERT(XML, @Response) AS Response
--<Root>
--  <Ok __json__bool="1">1</Ok>
--  <Count>2</Count>
--  <DT240349>
--    <DeviceSerialNo>DT240349</DeviceSerialNo>
--    <FiscalMemoryNo>02240349</FiscalMemoryNo>
--    <DeviceProtocol>DATECS</DeviceProtocol>
--    <DeviceModel>FP-3530?</DeviceModel>
--    <FirmwareVersion>4.10BG 10MAR08 1130</FirmwareVersion>
--    <CharsPerLine>30</CharsPerLine>
--    <TaxNo>0000000000</TaxNo>
--    <TaxCaption>БУЛСТАТ</TaxCaption>
--    <DeviceString>Protocol=DATECS;Port=COM1;Speed=9600</DeviceString>
--  </DT240349>
--  <DT518315>
--    <DeviceSerialNo>DT518315</DeviceSerialNo>
--    <FiscalMemoryNo>02518315</FiscalMemoryNo>
--    <DeviceProtocol>DATECS</DeviceProtocol>
--    <DeviceModel>DP-25</DeviceModel>
--    <FirmwareVersion>263453 08Nov18 1312</FirmwareVersion>
--    <CharsPerLine>30</CharsPerLine>
--    <TaxNo>НЕЗАДАДЕН</TaxNo>
--    <TaxCaption>ЕИК</TaxCaption>
--    <DeviceString>Protocol=DATECS;Port=COM2;Speed=115200</DeviceString>
--  </DT518315>
--  <Aliases>
--    <Count>1</Count>
--    <PrinterID1>
--      <DeviceSerialNo>DT518315</DeviceSerialNo>
--    </PrinterID1>
--  </Aliases>
--</Root>

--- send another request on the same conversation (request and result in XML)
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend '<Request><Url>/printers/DT518315/deposit?format=xml</Url><Amount>10.55</Amount></Request>', @Response OUTPUT, @Handle = @Handle1
SELECT      @Result AS Result, CONVERT(XML, @Response) AS Response
--<Root>
--  <Ok __json__bool="1">1</Ok>
--  <ReceiptNo>0000076</ReceiptNo>
--  <ReceiptDateTime>2019-07-23 16:37:43</ReceiptDateTime>
--  <TotalAvailable>475.2</TotalAvailable>
--  <TotalDeposits>488.63</TotalDeposits>
--  <TotalWithdraws>236.56</TotalWithdraws>
--</Root>

--- close conversation
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend @Handle = @Handle1
*/

CREATE PROC usp_sys_ServiceBrokerSend (
            @Request        NVARCHAR(MAX)       = NULL
            , @Response     NVARCHAR(MAX)       = NULL OUTPUT
            , @TargetSvc    SYSNAME             = NULL
            , @Handle       UNIQUEIDENTIFIER    = NULL OUTPUT
            , @QueueName    SYSNAME             = NULL
            , @SvcName      SYSNAME             = NULL
            , @Timeout      INT                 = NULL
            , @Retry        INT                 = NULL
) AS
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

DECLARE     @RetVal     INT
            , @SQL      NVARCHAR(MAX)
            , @MsgType  SYSNAME
            , @Attempt  INT
            , @IsAck    INT

SELECT      @QueueName = COALESCE(@QueueName, N'UcsFpInitiator' + N'Queue/' + CONVERT(NVARCHAR(50), @@SPID))
            , @SvcName = COALESCE(@SvcName, N'UcsFpInitiator' + N'Service/' + CONVERT(NVARCHAR(50), @@SPID))
            , @Timeout = COALESCE(@Timeout, 30000)
            , @Retry = COALESCE(@Retry, 3)
            , @Attempt = 0
            , @IsAck = 0
            , @RetVal = 0

WHILE       @Attempt < @Retry
BEGIN
            BEGIN TRY
                        SET         @Attempt = @Attempt + 1

                        IF          @TargetSvc IS NOT NULL
                        BEGIN
                                    IF          @Handle IS NOT NULL
                                    BEGIN
                                                ; END       CONVERSATION @Handle
                                    END

                                    BEGIN DIALOG CONVERSATION @Handle
                                    FROM        SERVICE @SvcName
                                    TO          SERVICE @TargetSvc, N'CURRENT DATABASE'
                                    WITH        ENCRYPTION = OFF
                        END

                        IF          @Request IS NULL
                        BEGIN
                                    IF EXISTS ( SELECT      *
                                                FROM        sys.conversation_endpoints
                                                WHERE       conversation_handle = @Handle AND state NOT IN ('CD', 'ER') )
                                    BEGIN
                                                ; SEND ON   CONVERSATION @Handle (N'__FIN__')
                                                ; END       CONVERSATION @Handle
                                    END

                                    GOTO        QH
                        END

                        ; SEND ON   CONVERSATION @Handle (N'__PING__')

                        SET         @SQL = N'
                        WAITFOR (   RECEIVE     TOP (1) @Response = CONVERT(NVARCHAR(MAX), message_body) 
                                                , @MsgType = message_type_name
                                    FROM        dbo.' + QUOTENAME(@QueueName) + N'  ), TIMEOUT 100'

                        EXEC        dbo.sp_executesql @SQL, N'@Response NVARCHAR(MAX) OUTPUT, @MsgType SYSNAME OUTPUT',
                                        @Response OUTPUT, @MsgType OUTPUT

                        IF          @Response = N'__PONG__'
                        BEGIN
                                    IF          @Request = N'__PING__'
                                    BEGIN
                                                GOTO        QH
                                    END

                                    SET         @IsAck = 0
                                    ; SEND ON   CONVERSATION @Handle (@Request)

                                    SET         @SQL = N'
                                    WAITFOR (   RECEIVE     TOP (1) @Response = CONVERT(NVARCHAR(MAX), message_body) 
                                                            , @MsgType = message_type_name
                                                FROM        dbo.' + QUOTENAME(@QueueName) + N'  ), TIMEOUT ' + CONVERT(NVARCHAR(50), @Timeout)
                        RepeatWait:
                                    SELECT      @Response = NULL, @MsgType = NULL
                                    EXEC        dbo.sp_executesql @SQL, N'@Response NVARCHAR(MAX) OUTPUT, @MsgType SYSNAME OUTPUT',
                                                    @Response OUTPUT, @MsgType OUTPUT

                                    IF          @MsgType = 'DEFAULT'
                                    BEGIN
                                                IF          @IsAck = 0 OR LEFT(@Response, 2) = N'__'
                                                BEGIN
                                                            IF          @Response = N'__ACK__'
                                                            BEGIN
                                                                        SET         @IsAck = 1
                                                            END

                                                            GOTO        RepeatWait
                                                END

                                                BREAK
                                    END
                        END
            END TRY
            BEGIN CATCH
                        --PRINT { fn CURRENT_TIMESTAMP } + ': ERROR_MESSAGE=' + ERROR_MESSAGE()

                        IF          @Handle IS NOT NULL AND ERROR_NUMBER() <> 8426 -- The conversation handle "%s" is not found.
                        BEGIN
                                    ; END       CONVERSATION @Handle
                        END

                        IF          @Attempt >= @Retry
                        BEGIN
                                    SELECT      @RetVal = 2
                                                , @Response = LEFT(ERROR_MESSAGE(), 255)
                                    GOTO        QH
                        END
            END CATCH
END

IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/EndDialog'
BEGIN
            ; END       CONVERSATION @Handle

            SELECT      @RetVal = 1
                        , @Response = N'Conversation ended'
            GOTO        QH
END

IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/Error'
BEGIN
            ; END       CONVERSATION @Handle

            SELECT      @RetVal = 1
                        , @Response = LEFT(CONVERT(XML, @Response).value('declare namespace ns="http://schemas.microsoft.com/SQL/ServiceBroker/Error";
                                                                            (//ns:Description)[1]', 'NVARCHAR(MAX)'), 255)
            GOTO        QH
END

IF          @MsgType IS NULL
BEGIN
            --PRINT { fn CURRENT_TIMESTAMP } + ': Timeout'

            SELECT      @RetVal = 99
                        , @Response = N'Timeout'
            GOTO        QH
END

QH:
RETURN      @RetVal
GO
