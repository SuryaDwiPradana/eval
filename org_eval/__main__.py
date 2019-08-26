from modules.prepare import main_prepare
# from modules.process import main_process
# from modules.report import main_report
from common import Configuration

if __name__ == '__main__':
    config = Configuration()
    # print(config.data)
    main_prepare(config)
    # main_process(config)
    # main_report(config)