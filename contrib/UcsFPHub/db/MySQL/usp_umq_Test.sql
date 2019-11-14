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
