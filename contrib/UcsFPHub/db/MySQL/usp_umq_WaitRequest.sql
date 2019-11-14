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
