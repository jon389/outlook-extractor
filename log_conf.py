import logging.config, time, datetime, pathlib

proj_name = pathlib.Path(__file__).parent.name  # parent directory
logfilename = pathlib.Path(__file__).parent / 'logs' / f'{proj_name}_{datetime.datetime.now():%Y-%m-%d_%H%M%S}.log'


class UTCFormatter(logging.Formatter):
    converter = time.gmtime


class PrefixLoggerAdapter(logging.LoggerAdapter):
    """
    from log_conf import logger  # needed since logger is re-defined below
    logger.info('log message without prefix')
    logger = PrefixLoggerAdapter(logger, {'prefix': 'this is a prefix for all msgs'})
    logger.debug('log message with prefix')
    """
    def process(self, msg, kwargs):
        return '[{}] {}'.format(self.extra['prefix'], msg), kwargs


log_dict_conf = dict(
    version=1,
    disable_existing_loggers=False,
    formatters=dict(
        # https://docs.python.org/3/library/logging.html#logrecord-attributes
        verbose_utc={
            '()': UTCFormatter,
            'format': '%(asctime)s [%(levelname)-5.5s] %(filename)-12.12s:%(lineno)4d:%(funcName)-12.12s %(message)s',
            'datefmt': '%Y-%m-%d %H:%M:%S Z'
        },
        verbose_local={
            'format': '%(asctime)s [%(levelname)-5.5s] %(filename)-12.12s:%(lineno)4d:%(funcName)-12.12s %(message)s',
            'datefmt': '%Y-%m-%d %I:%M:%S%p'
        },
        simple_utc={
            '()': UTCFormatter,
            'format': '%(asctime)s[%(levelname)-5.5s] %(message)s',
            'datefmt': '%H:%M:%S'
        },
        simple_local={
            'format': '%(asctime)s[%(levelname)-5.5s] %(message)s',
            'datefmt': '%H:%M:%S'
        },
    ),
    handlers=dict(
        console={
            'class': 'logging.StreamHandler',
            'formatter': 'verbose_local',
        },
        logfile={
            'class': 'logging.handlers.RotatingFileHandler',
            'formatter': 'verbose_local',
            'filename': logfilename,
            'maxBytes': 1024*1024,
            'backupCount': 10,
        }
    ),
    loggers={
        proj_name: dict(
            handlers='console logfile'.split(),
            level='DEBUG',
        ),
    }
)

logging.config.dictConfig(log_dict_conf)
logger = logging.getLogger(proj_name)
logger.debug('logging configured')

