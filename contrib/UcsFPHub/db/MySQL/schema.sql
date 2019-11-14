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

