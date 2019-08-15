-- This is an amalgamation of all MySQL scripts

--- User Message Queues tables
CREATE TABLE `umq_Queues` (
    `id`              int8            NOT NULL AUTO_INCREMENT
    , `name`          varchar(255)    NOT NULL
    , PRIMARY KEY (`id`)
) ENGINE=InnoDB;

CREATE TABLE `umq_Services` (
    `id`              int8            NOT NULL AUTO_INCREMENT
    , `name`          varchar(255)    NOT NULL
    , `queue_id`      int8            NOT NULL
    , PRIMARY KEY (`id`)
) ENGINE=InnoDB;

CREATE TABLE `umq_Messages` (
    `id`              int8            NOT NULL AUTO_INCREMENT
    , `status`        int4            NOT NULL
    , `service_id`    int8            NOT NULL
    , `queue_id`      int8            NOT NULL
    , `message_type`  varchar(255)    NOT NULL
    , `message_body`  text            NOT NULL
    , `created_at`    datetime        NOT NULL
    , PRIMARY KEY (`id`)
) ENGINE=InnoDB;
