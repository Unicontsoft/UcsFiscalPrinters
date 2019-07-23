IF OBJECT_ID('usp_sys_ServiceBrokerSend') IS NOT NULL DROP PROC usp_sys_ServiceBrokerSend
GO
/*
DECLARE     @Handle1    UNIQUEIDENTIFIER
            , @Result   INT
            , @Response NVARCHAR(MAX)
            , @ErrorText NVARCHAR(255)

--- setup receiving queue for current connection once on re-connect
EXEC        dbo.usp_sys_ServiceBrokerSetupService

--- open conversation and send request
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend '{ "Url": "/printers/DT518315/status" }', @Response OUTPUT, @TargetSvc = 'UcsFpTargetService/DT518315', @Handle = @Handle1 OUTPUT, @ErrorText = @ErrorText OUTPUT
SELECT      @Result AS Result, @Response AS Response, @ErrorText AS ErrorText
--{  
--   "Ok":true,
--   "DeviceStatus":"",
--   "DeviceDateTime":"2019-07-23 17:17:18"
--}

--- send another request on the same conversation (result in XML)
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend '{ "Url": "/printers?format=xml" }', @Response OUTPUT, @Handle = @Handle1, @ErrorText = @ErrorText OUTPUT
SELECT      @Result AS Result, CONVERT(XML, @Response) AS Response, @ErrorText AS ErrorText
--<Root>
--  <Ok __json__bool="1">1</Ok>
--  <Count>2</Count>
--  <DT240349>
--    <DeviceSerialNo>DT240349</DeviceSerialNo>
--    <FiscalMemoryNo>02240349</FiscalMemoryNo>
--    <DeviceProtocol>DATECS FP/ECR</DeviceProtocol>
--    <DeviceModel>FP-3530?</DeviceModel>
--    <FirmwareVersion>4.10BG 10MAR08 1130</FirmwareVersion>
--    <CharsPerLine>30</CharsPerLine>
--    <TaxNo>0000000000</TaxNo>
--    <TaxCaption>БУЛСТАТ</TaxCaption>
--    <DeviceString>Protocol=DATECS FP/ECR;Port=COM1;Speed=9600</DeviceString>
--  </DT240349>
--  <DT518315>
--    <DeviceSerialNo>DT518315</DeviceSerialNo>
--    <FiscalMemoryNo>02518315</FiscalMemoryNo>
--    <DeviceProtocol>DATECS FP/ECR</DeviceProtocol>
--    <DeviceModel>DP-25</DeviceModel>
--    <FirmwareVersion>263453 08Nov18 1312</FirmwareVersion>
--    <CharsPerLine>30</CharsPerLine>
--    <TaxNo>НЕЗАДАДЕН</TaxNo>
--    <TaxCaption>ЕИК</TaxCaption>
--    <DeviceString>Protocol=DATECS FP/ECR;Port=COM2;Speed=115200</DeviceString>
--  </DT518315>
--  <Aliases>
--    <Count>1</Count>
--    <PrinterID1>
--      <DeviceSerialNo>DT518315</DeviceSerialNo>
--    </PrinterID1>
--  </Aliases>
--</Root>

--- send another request on the same conversation (request and result in XML)
EXEC        @Result = dbo.usp_sys_ServiceBrokerSend '<Request><Url>/printers/DT518315/deposit?format=xml</Url><Amount>10.55</Amount></Request>', @Response OUTPUT, @Handle = @Handle1, @ErrorText = @ErrorText OUTPUT
SELECT      @Result AS Result, CONVERT(XML, @Response) AS Response, @ErrorText AS ErrorText
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
            , @ErrorText    NVARCHAR(255)       = NULL OUTPUT
            , @QueueName    SYSNAME             = NULL
            , @SvcName      SYSNAME             = NULL
            , @Timeout      INT                 = NULL
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
            IF          @TargetSvc IS NOT NULL
            BEGIN
                        IF          @Handle IS NOT NULL
                        BEGIN
                                    ; END       CONVERSATION @Handle
                        END

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
                                    , @ErrorText = 'Timeout'
                        GOTO        QH
            END

            --PRINT { fn CURRENT_TIMESTAMP } + ': @MsgType=' + @MsgType

            IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/EndDialog'
            BEGIN
                        ; END       CONVERSATION @Handle

                        SELECT      @RetVal = 1
                                    , @ErrorText = 'Conversation ended'
                        GOTO        QH
            END

            IF          @MsgType = 'http://schemas.microsoft.com/SQL/ServiceBroker/Error'
            BEGIN
                        ; END       CONVERSATION @Handle

                        SELECT      @RetVal = 1
                                    , @ErrorText = LEFT(CONVERT(XML, @Response).value('declare namespace ns="http://schemas.microsoft.com/SQL/ServiceBroker/Error"; (ns:Error/ns:Description)[1]', 'NVARCHAR(MAX)'), 255)
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
                        , @ErrorText = LEFT(ERROR_MESSAGE(), 255)
            GOTO        QH
END CATCH

QH:
RETURN      @RetVal
GO
