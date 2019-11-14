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
