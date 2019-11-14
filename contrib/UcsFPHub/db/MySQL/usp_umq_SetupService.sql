SET SQL_SAFE_UPDATES = 0;
DROP PROCEDURE IF EXISTS `usp_umq_SetupService`;
/*
CALL `usp_umq_SetupService`(NULL, NULL, NULL);
CALL `usp_umq_SetupService`(NULL, NULL, 'DROP_ONLY');
CALL `usp_umq_SetupService`('UcsFpTargetQueue/ZK123456', 'UcsFpTargetService/ZK123456', NULL);
*/
DELIMITER $$
CREATE PROCEDURE `usp_umq_SetupService` (
            IN `@queue_name`    VARCHAR(128)
            , IN `@svc_name`    VARCHAR(128)
            , IN `@mode`        VARCHAR(20)
) BEGIN
/*------------------------------------------------------------------------
'
' UcsFPHub (c) 2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Hub
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
------------------------------------------------------------------------*/
SET         `@queue_name` = COALESCE(NULLIF(`@queue_name`, ''), CONCAT('UcsFpInitiator', 'Queue', '/', CONNECTION_ID()))
            , `@svc_name` = COALESCE(NULLIF(`@svc_name`, ''), CONCAT('UcsFpInitiator', 'Service', '/', CONNECTION_ID()))
            , `@mode` = COALESCE(`@mode`, '');

IF EXISTS (SELECT 0 FROM `umq_Services` WHERE `name` = `@svc_name`) AND `@mode` IN ('DROP_SERVICE', 'DROP_EXISTING', 'DROP_ONLY') THEN
            DELETE FROM `umq_Conversations`
            WHERE       `service_id` IN (SELECT `id` FROM `umq_Services` WHERE `name` = `@svc_name`);
            
            DELETE FROM `umq_Messages`
            WHERE       `service_id` IN (SELECT `id` FROM `umq_Queues` WHERE `name` = `@svc_name`);
END IF;

IF EXISTS (SELECT 0 FROM `umq_Queues` WHERE `name` = `@queue_name`) AND `@mode` IN ('DROP_EXISTING', 'DROP_ONLY') THEN
            DELETE FROM `umq_Services`
            WHERE       `queue_id` IN (SELECT `id` FROM `umq_Queues` WHERE `name` = `@queue_name`);
            
            DELETE FROM `umq_Queues`
            WHERE       `name` = `@queue_name`;
END IF;

IF NOT EXISTS (SELECT 0 FROM `umq_Queues` WHERE `name` = `@queue_name`) AND `@mode` NOT IN ('DROP_ONLY') THEN
            INSERT INTO `umq_Queues`(`name`)
            SELECT      `@queue_name`;
END IF;

IF NOT EXISTS (SELECT 0 FROM `umq_Services` WHERE `name` = `@svc_name`) AND `@mode` NOT IN ('DROP_ONLY') THEN
            INSERT INTO `umq_Services`(`name`, `queue_id`)
            SELECT      `@svc_name`
                        , (SELECT `id` FROM `umq_Queues` WHERE `name` = `@queue_name`) AS `queue_id`;
END IF;

-- cleanup complete conversations and delivered messages after 30 minutes
DELETE FROM `umq_Messages` WHERE `status` <> 0 AND CURRENT_TIMESTAMP(3) > TIMESTAMPADD(MINUTE, 30, `created_at`);
DELETE FROM `umq_Conversations` WHERE `status` <> 0 AND CURRENT_TIMESTAMP(3) > TIMESTAMPADD(MINUTE, 30, `created_at`);

END $$
DELIMITER ;
