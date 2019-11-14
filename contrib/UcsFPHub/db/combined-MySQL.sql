-- This is an amalgamation of all MySQL scripts

--- User Message Queues tables
DROP TABLE IF EXISTS `umq_Messages`, `umq_Conversations`, `umq_Services`, `umq_Queues`;

CREATE TABLE `umq_Queues` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `name`            VARCHAR(255)    NOT NULL
    , PRIMARY KEY (`id`)
    , UNIQUE KEY (`name`)
) ENGINE=InnoDB;

CREATE TABLE `umq_Services` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `name`            VARCHAR(255)    NOT NULL
    , `queue_id`        INT             NOT NULL
    , PRIMARY KEY (`id`)
    , UNIQUE KEY (`name`)
    , FOREIGN KEY (`queue_id`) REFERENCES `umq_Queues`(`id`)
) ENGINE=InnoDB;

CREATE TABLE `umq_Conversations` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `status`          INT             NOT NULL DEFAULT (0)
    , `service_id`      INT             NOT NULL 
    , `far_service_id`  INT             NOT NULL 
    , `created_at`      TIMESTAMP(3)    NOT NULL
    , PRIMARY KEY (`id`)
    , FOREIGN KEY (`service_id`) REFERENCES `umq_Services`(`id`)
    , FOREIGN KEY (`far_service_id`) REFERENCES `umq_Services`(`id`)
) ENGINE=InnoDB;

CREATE TABLE `umq_Messages` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `status`          INT             NOT NULL DEFAULT (0)
    , `conversation_id` INT             NOT NULL
    , `service_id`      INT             NOT NULL
    , `queue_id`        INT             NOT NULL
    , `message_type`    VARCHAR(255)    NOT NULL
    , `message_body`    LONGTEXT        NOT NULL
    , `created_at`      TIMESTAMP(3)    NOT NULL
    , PRIMARY KEY (`id`)
    , FOREIGN KEY (`service_id`) REFERENCES `umq_Services`(`id`)
    , FOREIGN KEY (`queue_id`) REFERENCES `umq_Queues`(`id`)
) ENGINE=InnoDB;

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

SET SQL_SAFE_UPDATES=0;
DROP PROCEDURE IF EXISTS `usp_umq_Send`;
/*
CALL `usp_umq_SetupService`(NULL, NULL, NULL);
CALL `usp_umq_Send`('test', @response, 'UcsFpTargetService/ZK123456', 3000, @handle, @result);
SELECT @result, @handle, @response;
*/
DELIMITER $$
CREATE PROCEDURE `usp_umq_Send` (
            IN `@request`       LONGTEXT
            , OUT `@response`   LONGTEXT
            , IN `@target_svc`  VARCHAR(128)
            , IN `@timeout`     INT
            , OUT `@handle`     INT
            , OUT `@retval`     INT
) PROC:BEGIN
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
DECLARE     `@queue_name`   VARCHAR(128);
DECLARE     `@svc_name`     VARCHAR(128);
DECLARE     `@service_id`   INT;
DECLARE     `@queue_id`     INT;
DECLARE     `@far_service_id` INT;
DECLARE     `@far_queue_id` INT;
DECLARE     `@conv_id`      INT;
DECLARE     `@msg_id`       INT;
DECLARE     `@msg_type`     VARCHAR(255);
DECLARE     `@msg_body`     LONGTEXT;
DECLARE     `@start_time`   TIMESTAMP(3);

SET         `@retval` = 0;
SET         `@queue_name` = COALESCE(NULLIF(`@queue_name`, ''), CONCAT('UcsFpInitiator', 'Queue', '/', CONNECTION_ID()));
SET         `@svc_name` = COALESCE(NULLIF(`@svc_name`, ''), CONCAT('UcsFpInitiator', 'Service', '/', CONNECTION_ID()));
SET         `@service_id` = (SELECT `id` FROM `umq_Services` WHERE `name` = `@svc_name`);
SET         `@queue_id` = (SELECT `queue_id` FROM `umq_Services` WHERE `name` = `@svc_name`);
SET         `@far_service_id` = (SELECT `id` FROM `umq_Services` WHERE `name` = `@target_svc`);
SET         `@far_queue_id` = (SELECT `queue_id` FROM `umq_Services` WHERE `name` = `@target_svc`);

IF          `@handle` IS NOT NULL
THEN
            SET         `@conv_id` = `@handle`;
ELSE
            -- begin conversation
            INSERT      `umq_Conversations`(`service_id`, `far_service_id`, `created_at`)
            SELECT      `@service_id`           AS `service_id`
                        , `@far_service_id`     AS `far_service_id`
                        , CURRENT_TIMESTAMP(3)  AS `created_at`;
                        
            SET         `@conv_id` = LAST_INSERT_ID();
END IF;

IF          `@request` IS NULL
THEN
            -- send on conversation '__FIN__'
            INSERT      `umq_Messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
            SELECT      `@conv_id`              AS `conversation_id`
                        , `@far_service_id`     AS `service_id`
                        , `@far_queue_id`       AS `queue_id`
                        , 'DEFAULT'             AS `message_type`
                        , '__FIN__'             AS `message_body`
                        , CURRENT_TIMESTAMP(3)  AS `created_at`;
            
            -- end conversation
            UPDATE      `umq_Conversations`
            SET         `status` = 1
            WHERE       `id` = `@conv_id`;
            
            LEAVE       PROC;
END IF;

-- send on conversation '__PING__'
INSERT      `umq_Messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
SELECT      `@conv_id`              AS `conversation_id`
            , `@far_service_id`     AS `service_id`
            , `@far_queue_id`       AS `queue_id`
            , 'DEFAULT'             AS `message_type`
            , '__PING__'            AS `message_body`
            , CURRENT_TIMESTAMP(3)  AS `created_at`;

-- waitfor receive '__PONG__'
SET         `@msg_body` = '';

WHILE       `@msg_body` <> '__PONG__'
DO
            SET         `@start_time` = CURRENT_TIMESTAMP(3);
            SET         `@msg_id` = NULL;
            SET         `@msg_type` = '';
            SET         `@msg_body` = '';

            WHILE       `@msg_id` IS NULL
                        AND CURRENT_TIMESTAMP(3) < TIMESTAMPADD(SECOND, 0.1, `@start_time`)
            DO
                        SELECT      `id`, `message_type`, `message_body` 
                        INTO        `@msg_id`, `@msg_type`, `@msg_body`
                        FROM        `umq_Messages`
                        WHERE       `queue_id` = `@queue_id`
                                    AND `status` = 0
                        ORDER BY    `id`
                        LIMIT       1 FOR UPDATE SKIP LOCKED;
                        
                        IF          `@msg_id` IS NULL
                        THEN
                                    DO SLEEP(0.001);
                        END IF;
            END WHILE;

            IF          `@msg_id` IS NULL 
            THEN
                        -- end conversation
                        UPDATE      `umq_Conversations`
                        SET         `status` = 1
                        WHERE       `id` = `@conv_id`;
                        
                        SET         `@response` = 'Timeout #1';
                        SET         `@retval` = 1;
                        LEAVE       PROC;
            END IF;

            UPDATE      `umq_Messages`
            SET         `status` = 1
            WHERE       `id` = `@msg_id`;
END WHILE;

-- send on conversation `@request`
INSERT      `umq_Messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
SELECT      `@conv_id`          AS `conversation_id`
            , `@far_service_id` AS `service_id`
            , `@far_queue_id`   AS `queue_id`
            , 'DEFAULT'         AS `message_type`
            , `@request`        AS `message_body`
            , CURRENT_TIMESTAMP(3) AS `created_at`;

-- waitfor receive `@response`
SET         `@start_time` = CURRENT_TIMESTAMP(3);
SET         `@msg_id` = NULL;
SET         `@msg_type` = '';
SET         `@msg_body` = '';

WHILE       `@msg_id` IS NULL
            AND CURRENT_TIMESTAMP(3) < TIMESTAMPADD(SECOND, `@timeout` / 1000, `@start_time`)
DO
            SELECT      `id`, `message_type`, `message_body` 
            INTO        `@msg_id`, `@msg_type`, `@msg_body`
            FROM        `umq_Messages`
            WHERE       `queue_id` = `@queue_id`
                        AND `status` = 0
            ORDER BY    `id`
            LIMIT       1 FOR UPDATE SKIP LOCKED;
            
            IF          `@msg_id` IS NULL
            THEN
                        DO SLEEP(0.001);
            END IF;
END WHILE;

IF          `@msg_id` IS NULL 
THEN
            -- end conversation
            UPDATE      `umq_Conversations`
            SET         `status` = 1
            WHERE       `id` = `@conv_id`;
            
            SET         `@response` = 'Timeout #2';
            SET         `@retval` = 1;
            LEAVE       PROC;
END IF;

UPDATE      `umq_Messages`
SET         `status` = 1
WHERE       `id` = `@msg_id`;

-- can signal transport error
IF          `@msg_type` <> 'DEFAULT'
THEN
            UPDATE      `umq_Conversations`
            SET         `status` = 1
            WHERE       `id` = `@conv_id`;
            
            SET         `@response` = `@msg_body`;
            SET         `@retval` = 1;
            LEAVE       PROC;
END IF;

SET         `@response` = `@msg_body`;
SET         `@handle` = `@conv_id`;
LEAVE       PROC;

END $$
DELIMITER ;

SET SQL_SAFE_UPDATES=0;
DROP PROCEDURE IF EXISTS `usp_umq_WaitRequest`;
/*
CALL `usp_umq_SetupService`('UcsFpTargetQueue/DEV-PC/1234', 'UcsFpTargetService/ZK123456', NULL);
CALL `usp_umq_WaitRequest`('UcsFpTargetQueue/DEV-PC/1234', 5000, @handle, @request, @svc_name, @error_text, @result);
            
INSERT      `umq_Messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
SELECT      @handle              AS `conversation_id`
            , (SELECT `id` FROM umq_Services WHERE `name` = @svc_name) AS `service_id`
            , (SELECT `queue_id` FROM umq_Services WHERE `name` = @svc_name) AS `queue_id`
            , 'DEFAULT'             AS `message_type`
            , CONCAT('Time is ', CURRENT_TIMESTAMP(3)) AS `message_body`
            , CURRENT_TIMESTAMP(3)  AS `created_at`;
*/
DELIMITER $$
CREATE PROCEDURE `usp_umq_WaitRequest` (
            IN `@queue_name`    VARCHAR(128)
            , IN `@timeout`     INT
            , OUT `@handle`     INT
            , OUT `@request`    LONGTEXT
            , OUT `@svc_name`   VARCHAR(128)
            , OUT `@error_text` VARCHAR(256)
            , OUT `@retval`     INT
) PROC:BEGIN
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
DECLARE     `@queue_id`     INT;
DECLARE     `@msg_id`       INT;
DECLARE     `@msg_type`     VARCHAR(255);
DECLARE     `@msg_body`     LONGTEXT;
DECLARE     `@conv_id`      INT;
DECLARE     `@start_time`   TIMESTAMP(3);
DECLARE     `@conv_status`  INT;
DECLARE     `@conv_service_id` INT;

SET         `@retval` = 0;
SET         `@queue_id` = (SELECT `id` FROM `umq_Queues` WHERE `name` = `@queue_name`);

IF          `@queue_id` IS NULL
THEN
            SET         `@error_text` = 'Queue not found';
            SET         `@retval` = 1;
            
            LEAVE       PROC;
END IF;

WHILE       1=1
DO
            -- waitfor receive `@response`
            SET         `@start_time` = CURRENT_TIMESTAMP(3);
            SET         `@msg_id` = NULL;
            SET         `@msg_type` = '';
            SET         `@msg_body` = '';

            WHILE       `@msg_id` IS NULL
                        AND CURRENT_TIMESTAMP(3) < TIMESTAMPADD(SECOND, `@timeout` / 1000, `@start_time`)
            DO
                        SELECT      `id`, `message_type`, `message_body`, `conversation_id`
                        INTO        `@msg_id`, `@msg_type`, `@msg_body`, `@conv_id`
                        FROM        `umq_Messages`
                        WHERE       `queue_id` = `@queue_id`
                                    AND `status` = 0
                        ORDER BY    `id`
                        LIMIT       1 FOR UPDATE SKIP LOCKED;
                        
                        IF          `@msg_id` IS NULL
                        THEN
                                    DO SLEEP(0.001);
                        END IF;
            END WHILE;
            
            IF          `@msg_id` IS NULL
            THEN
                        SET         `@error_text` = 'Timeout';
                        SET         `@retval` = 1;
                        LEAVE       PROC;            
            END IF;

            UPDATE      `umq_Messages`
            SET         `status` = 1
            WHERE       `id` = `@msg_id`;
            
            IF          `@msg_body` = '__FIN__'
            THEN
                        -- do nothing (repeat waitfor)
                        SET         `@msg_body` = NULL;
            ELSE
                        SET         `@conv_status` = NULL;
            
                        SELECT      `status`, `service_id`
                        INTO        `@conv_status`, `@conv_service_id`
                        FROM        `umq_Conversations` 
                        WHERE       `id` = `@conv_id`;
                        
                        -- SELECT `@msg_id`, `@conv_status`, `@conv_service_id`, `@msg_type`, `@msg_body`;
            
                        IF          `@conv_status` = 0
                        THEN
                                    IF          `@msg_body` = '__PING__'
                                    THEN
                                                -- send on conversation '__PONG__'
                                                INSERT      `umq_Messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
                                                SELECT      `@conv_id`              AS `conversation_id`
                                                            , `@conv_service_id`    AS `service_id`
                                                            , (SELECT `queue_id` FROM `umq_Services` WHERE `id` = `@conv_service_id`) AS `queue_id`
                                                            , 'DEFAULT'             AS `message_type`
                                                            , '__PONG__'            AS `message_body`
                                                            , CURRENT_TIMESTAMP(3)  AS `created_at`;
                                    ELSE
                                                SET         `@handle` = `@conv_id`;
                                                SET         `@svc_name` = (SELECT `name` FROM `umq_Services` WHERE `id` = `@conv_service_id`);

                                                IF          `@msg_type` <> 'DEFAULT'
                                                THEN
                                                            SET         `@error_text` = `@msg_body`;
                                                            SET         `@retval` = 1;
                                                            
                                                            LEAVE       PROC;
                                                END IF;

                                                SET         `@request` = `@msg_body`;
                                                LEAVE       PROC;
                                    END IF;
                        END IF;
            END IF;
END WHILE;
END $$
DELIMITER ;

SET SQL_SAFE_UPDATES=0;
DROP PROCEDURE IF EXISTS `usp_umq_Test`;
-- CALL `usp_umq_SetupService`('UcsFpTargetQueue/DEV-PC/1234', 'UcsFpTargetService/ZK123456', NULL);
-- CALL `usp_umq_Test`

DELIMITER $$
CREATE PROCEDURE `usp_umq_Test` ()
BEGIN
WHILE       1=1
DO
            CALL        `usp_umq_WaitRequest`('UcsFpTargetQueue/DEV-PC/1234', 5000, @handle, @request, @svc_name, @error_text, @result);
            
            IF          @result = 0
            THEN
                        INSERT      `umq_Messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
                        SELECT      @handle              AS `conversation_id`
                                    , (SELECT `id` FROM umq_Services WHERE `name` = @svc_name) AS `service_id`
                                    , (SELECT `queue_id` FROM umq_Services WHERE `name` = @svc_name) AS `queue_id`
                                    , 'DEFAULT'             AS `message_type`
                                    , CONCAT('Time is ', CURRENT_TIMESTAMP(3)) AS `message_body`
                                    , CURRENT_TIMESTAMP(3)  AS `created_at`;
            END IF;
END WHILE;
END $$
DELIMITER ;
