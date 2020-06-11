import logging
import logging.config


def get_logger(logger_name):
    logger = logging.getLogger(logger_name)
    return logger


def set_log_config(config):
    log_config = config['logConfig']
    log_level = config['LOG_LEVEL']
    loggers = log_config['loggers'].keys()
    logging.config.dictConfig(log_config)
    for logger_name in loggers:
        logger = get_logger(logger_name)
        logger.setLevel(log_level)


# logging object for tagging
tagging_logger = get_logger("tagging")

# logging object for extract-info
extract_info_logger = get_logger("extractinfo")
