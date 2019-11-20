SET SQL_SAFE_UPDATES = 0;
DROP PROCEDURE IF EXISTS `usp_umq_setup_service`;
/*
CALL `usp_umq_setup_service`(NULL, NULL, NULL);
CALL `usp_umq_setup_service`(NULL, NULL, 'DROP_ONLY');
CALL `usp_umq_setup_service`('UcsFpTargetQueue/ZK123456', 'UcsFpTargetService/ZK123456', NULL);
CALL `usp_umq_setup_service`('UcsFpTargetQueue/WQW-PC/6D4F07/wqw-pc/UcsFPHub', 'UcsFpTargetService/WQW-PC/6D4F07/wqw-pc/UcsFPHub', 'DROP_ONLY');
*/
DELIMITER $$
CREATE PROCEDURE `usp_umq_setup_service` (
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
DECLARE     `@service_id` INT;
DECLARE     `@queue_id` INT;

SET         `@queue_name` = COALESCE(NULLIF(`@queue_name`, ''), CONCAT('UcsFpInitiator', 'Queue', '/', CONNECTION_ID()));
SET         `@svc_name` = COALESCE(NULLIF(`@svc_name`, ''), CONCAT('UcsFpInitiator', 'Service', '/', CONNECTION_ID()));
SET         `@mode` = COALESCE(`@mode`, '');
SET         `@service_id` = (SELECT `id` FROM `umq_services` WHERE `name` = `@svc_name`);
SET         `@queue_id` = (SELECT `id` FROM `umq_queues` WHERE `name` = `@queue_name`);

IF          `@service_id` IS NOT NULL AND `@mode` IN ('DROP_SERVICE', 'DROP_EXISTING', 'DROP_ONLY')
THEN
            DELETE FROM `umq_messages`
            WHERE       `service_id` = `@service_id`;
            
            DELETE FROM `umq_messages`
            WHERE       `conversation_id` IN (SELECT `id` FROM `umq_conversations` WHERE `service_id` = `@service_id`);
            
            DELETE FROM `umq_messages`
            WHERE       `conversation_id` IN (SELECT `id` FROM `umq_conversations` WHERE `far_service_id` = `@service_id`);
            
            DELETE FROM `umq_conversations`
            WHERE       `service_id` = `@service_id`;
            
            DELETE FROM `umq_conversations`
            WHERE       `far_service_id` = `@service_id`;
            
            DELETE FROM `umq_services`
            WHERE       `id` = `@service_id`;
END IF;

IF          `@queue_id` IS NOT NULL AND `@mode` IN ('DROP_EXISTING', 'DROP_ONLY')
THEN
            DELETE FROM `umq_messages`
            WHERE       `queue_id` = `@queue_id`;
            
            DELETE FROM `umq_messages`
            WHERE       `service_id` IN (SELECT `id` FROM `umq_services` WHERE `queue_id` = `@queue_id`);
            
            DELETE FROM `umq_messages`
            WHERE       `conversation_id` IN (SELECT `id` FROM `umq_conversations` WHERE `service_id` IN (SELECT `id` FROM `umq_services` WHERE `queue_id` = `@queue_id`));
            
            DELETE FROM `umq_messages`
            WHERE       `conversation_id` IN (SELECT `id` FROM `umq_conversations` WHERE `far_service_id` IN (SELECT `id` FROM `umq_services` WHERE `queue_id` = `@queue_id`));
            
            DELETE FROM `umq_conversations`
            WHERE       `service_id` IN (SELECT `id` FROM `umq_services` WHERE `queue_id` = `@queue_id`);
            
            DELETE FROM `umq_conversations`
            WHERE       `far_service_id` IN (SELECT `id` FROM `umq_services` WHERE `queue_id` = `@queue_id`);
            
            DELETE FROM `umq_services`
            WHERE       `queue_id` = `@queue_id`;
            
            DELETE FROM `umq_queues`
            WHERE       `id` = `@queue_id`;
END IF;

IF          NOT EXISTS (SELECT 0 FROM `umq_queues` WHERE `name` = `@queue_name`)
            AND `@mode` NOT IN ('DROP_ONLY')
THEN
            INSERT INTO `umq_queues`(`name`)
            SELECT      `@queue_name`;
END IF;

IF          NOT EXISTS (SELECT 0 FROM `umq_services` WHERE `name` = `@svc_name`)
            AND `@mode` NOT IN ('DROP_ONLY')
THEN
            INSERT INTO `umq_services`(`name`, `queue_id`)
            SELECT      `@svc_name`
                        , (SELECT `id` FROM `umq_queues` WHERE `name` = `@queue_name`) AS `queue_id`;
END IF;

-- cleanup complete conversations and delivered messages after 30 minutes
DELETE FROM `umq_messages` WHERE `status` <> 0 AND CURRENT_TIMESTAMP(3) > TIMESTAMPADD(MINUTE, 30, `created_at`);
DELETE FROM `umq_conversations` WHERE `status` <> 0 AND CURRENT_TIMESTAMP(3) > TIMESTAMPADD(MINUTE, 30, `created_at`);

END $$
DELIMITER ;
