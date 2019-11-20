-- This is an amalgamation of all MySQL scripts

--- User Message Queues tables
DROP TABLE IF EXISTS `umq_messages`, `umq_conversations`, `umq_services`, `umq_queues`;

CREATE TABLE `umq_queues` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `name`            VARCHAR(255)    NOT NULL
    , PRIMARY KEY (`id`)
    , UNIQUE KEY (`name`)
) ENGINE=InnoDB;

CREATE TABLE `umq_services` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `name`            VARCHAR(255)    NOT NULL
    , `queue_id`        INT             NOT NULL
    , PRIMARY KEY (`id`)
    , UNIQUE KEY (`name`)
    , FOREIGN KEY (`queue_id`) REFERENCES `umq_queues`(`id`)
) ENGINE=InnoDB;

CREATE TABLE `umq_conversations` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `status`          INT             NOT NULL DEFAULT (0)
    , `service_id`      INT             NOT NULL 
    , `far_service_id`  INT             NOT NULL 
    , `created_at`      TIMESTAMP(3)    NOT NULL
    , PRIMARY KEY (`id`)
    , FOREIGN KEY (`service_id`) REFERENCES `umq_services`(`id`)
    , FOREIGN KEY (`far_service_id`) REFERENCES `umq_services`(`id`)
) ENGINE=InnoDB;

CREATE TABLE `umq_messages` (
    `id`                INT             NOT NULL AUTO_INCREMENT
    , `status`          INT             NOT NULL DEFAULT (0)
    , `conversation_id` INT             NOT NULL
    , `service_id`      INT             NOT NULL
    , `queue_id`        INT             NOT NULL
    , `message_type`    VARCHAR(255)    NOT NULL
    , `message_body`    LONGTEXT        NOT NULL
    , `created_at`      TIMESTAMP(3)    NOT NULL
    , PRIMARY KEY (`id`)
    , FOREIGN KEY (`conversation_id`) REFERENCES `umq_conversations`(`id`)
    , FOREIGN KEY (`service_id`) REFERENCES `umq_services`(`id`)
    , FOREIGN KEY (`queue_id`) REFERENCES `umq_queues`(`id`)
) ENGINE=InnoDB;


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

SET SQL_SAFE_UPDATES=0;
DROP PROCEDURE IF EXISTS `usp_umq_send`;
/*
CALL `usp_umq_setup_service`(NULL, NULL, NULL);
CALL `usp_umq_send`('{ "Url": "/printers" }', @response, 'UcsFpTargetService/ZK133759', 30000, @handle, @result);
SELECT @result, @handle, @response;
*/
DELIMITER $$
CREATE PROCEDURE `usp_umq_send` (
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
SET         `@service_id` = (SELECT `id` FROM `umq_services` WHERE `name` = `@svc_name`);
SET         `@queue_id` = (SELECT `queue_id` FROM `umq_services` WHERE `name` = `@svc_name`);
SET         `@far_service_id` = (SELECT `id` FROM `umq_services` WHERE `name` = `@target_svc`);
SET         `@far_queue_id` = (SELECT `queue_id` FROM `umq_services` WHERE `name` = `@target_svc`);

IF          `@handle` IS NOT NULL
THEN
            SET         `@conv_id` = `@handle`;
ELSE
            -- begin conversation
            INSERT      `umq_conversations`(`service_id`, `far_service_id`, `created_at`)
            SELECT      `@service_id`           AS `service_id`
                        , `@far_service_id`     AS `far_service_id`
                        , CURRENT_TIMESTAMP(3)  AS `created_at`;
                        
            SET         `@conv_id` = LAST_INSERT_ID();
END IF;

IF          `@request` IS NULL
THEN
            -- send on conversation '__FIN__'
            INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
            SELECT      `@conv_id`              AS `conversation_id`
                        , `@far_service_id`     AS `service_id`
                        , `@far_queue_id`       AS `queue_id`
                        , 'DEFAULT'             AS `message_type`
                        , '__FIN__'             AS `message_body`
                        , CURRENT_TIMESTAMP(3)  AS `created_at`;
            
            -- end conversation
            UPDATE      `umq_conversations`
            SET         `status` = 1
            WHERE       `id` = `@conv_id`;
            
            LEAVE       PROC;
END IF;

-- send on conversation '__PING__'
INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
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
                        FROM        `umq_messages`
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
                        UPDATE      `umq_conversations`
                        SET         `status` = 1
                        WHERE       `id` = `@conv_id`;
                        
                        SET         `@response` = 'Timeout #1';
                        SET         `@retval` = 1;
                        LEAVE       PROC;
            END IF;

            UPDATE      `umq_messages`
            SET         `status` = 1
            WHERE       `id` = `@msg_id`;
END WHILE;

-- send on conversation `@request`
INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
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
            FROM        `umq_messages`
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
            UPDATE      `umq_conversations`
            SET         `status` = 1
            WHERE       `id` = `@conv_id`;
            
            SET         `@response` = 'Timeout #2';
            SET         `@retval` = 1;
            LEAVE       PROC;
END IF;

UPDATE      `umq_messages`
SET         `status` = 1
WHERE       `id` = `@msg_id`;

-- can signal transport error
IF          `@msg_type` <> 'DEFAULT'
THEN
            UPDATE      `umq_conversations`
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
DROP PROCEDURE IF EXISTS `usp_umq_wait_request`;
/*
CALL `usp_umq_setup_service`('UcsFpTargetQueue/DEV-PC/1234', 'UcsFpTargetService/ZK123456', NULL);
CALL `usp_umq_wait_request`('UcsFpTargetQueue/DEV-PC/1234', 1000, @handle, @request, @svc_name, @error_text, @result);
            
INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
SELECT      @handle              AS `conversation_id`
            , (SELECT `id` FROM umq_services WHERE `name` = @svc_name) AS `service_id`
            , (SELECT `queue_id` FROM umq_services WHERE `name` = @svc_name) AS `queue_id`
            , 'DEFAULT'             AS `message_type`
            , CONCAT('Time is ', CURRENT_TIMESTAMP(3)) AS `message_body`
            , CURRENT_TIMESTAMP(3)  AS `created_at`;
*/
DELIMITER $$
CREATE PROCEDURE `usp_umq_wait_request` (
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
SET         `@queue_id` = (SELECT `id` FROM `umq_queues` WHERE `name` = `@queue_name`);

BODY:BEGIN
            IF          `@queue_id` IS NULL
            THEN
                        SET         `@error_text` = 'Queue not found';
                        SET         `@retval` = 98;
                        
                        LEAVE       BODY;
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
                                    FROM        `umq_messages`
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
                                    SET         `@retval` = 99;
                                    LEAVE       BODY;
                        END IF;

                        UPDATE      `umq_messages`
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
                                    FROM        `umq_conversations` 
                                    WHERE       `id` = `@conv_id`;
                                    
                                    -- SELECT `@msg_id`, `@conv_status`, `@conv_service_id`, `@msg_type`, `@msg_body`;
                        
                                    IF          `@conv_status` = 0
                                    THEN
                                                IF          `@msg_body` = '__PING__'
                                                THEN
                                                            -- send on conversation '__PONG__'
                                                            INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
                                                            SELECT      `@conv_id`              AS `conversation_id`
                                                                        , `@conv_service_id`    AS `service_id`
                                                                        , (SELECT `queue_id` FROM `umq_services` WHERE `id` = `@conv_service_id`) AS `queue_id`
                                                                        , 'DEFAULT'             AS `message_type`
                                                                        , '__PONG__'            AS `message_body`
                                                                        , CURRENT_TIMESTAMP(3)  AS `created_at`;
                                                ELSE
                                                            SET         `@handle` = `@conv_id`;
                                                            SET         `@svc_name` = (SELECT `name` FROM `umq_services` WHERE `id` = `@conv_service_id`);

                                                            IF          `@msg_type` <> 'DEFAULT'
                                                            THEN
                                                                        SET         `@error_text` = `@msg_body`;
                                                                        SET         `@retval` = 1;
                                                                        
                                                                        LEAVE       BODY;
                                                            END IF;

                                                            SET         `@request` = `@msg_body`;
                                                            LEAVE       BODY;
                                                END IF;
                                    END IF;
                        END IF;
            END WHILE;
END;

SELECT      `@handle` AS Handle, `@request` AS Request, `@svc_name` AS SvcName, `@error_text` AS ErrorText, `@retval` AS Result;
SET         `@msg_body` = CONCAT('@retval=', `@retval`);
SIGNAL      SQLSTATE '01000' SET MESSAGE_TEXT = `@msg_body`, MYSQL_ERRNO = 1000;

END $$
DELIMITER ;

SET SQL_SAFE_UPDATES=0;
DROP PROCEDURE IF EXISTS `usp_umq_test`;
-- CALL `usp_umq_setup_service`('UcsFpTargetQueue/DEV-PC/1234', 'UcsFpTargetService/ZK123456', NULL);
-- CALL `usp_umq_test`

DELIMITER $$
CREATE PROCEDURE `usp_umq_test` ()
BEGIN
WHILE       1=1
DO
            CALL        `usp_umq_wait_request`('UcsFpTargetQueue/DEV-PC/1234', 5000, @handle, @request, @svc_name, @error_text, @result);
            
            IF          @result = 0
            THEN
                        INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)
                        SELECT      @handle              AS `conversation_id`
                                    , (SELECT `id` FROM umq_services WHERE `name` = @svc_name) AS `service_id`
                                    , (SELECT `queue_id` FROM umq_services WHERE `name` = @svc_name) AS `queue_id`
                                    , 'DEFAULT'             AS `message_type`
                                    , CONCAT('Time is ', CURRENT_TIMESTAMP(3)) AS `message_body`
                                    , CURRENT_TIMESTAMP(3)  AS `created_at`;
            END IF;
END WHILE;
END $$
DELIMITER ;
