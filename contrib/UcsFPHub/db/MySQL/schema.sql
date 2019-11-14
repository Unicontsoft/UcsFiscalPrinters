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
