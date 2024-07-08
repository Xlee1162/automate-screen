# from PopupChecker import *
from PopupChecker_2 import *
from ScreenMonitor import *

template_dir = 'SMD/Template'
input_dir = 'SMD/Input'
csv_file = 'SMD/csv/info.csv'
app_name = 'EVKey64.exe'

print('Hello anh em')
#screen_monitor = ScreenMonitor(template_dir, input_dir, interval=5, cleanup_days=7)
#screen_monitor.run()

# popup_checker = PopupChecker(app_name, template_dir, csv_file, input_dir, cleanup_days=7)
# popup_checker.run()

checker = PopupChecker(app_name=app_name, template_dir=template_dir, csv_file=csv_file, input_dir=input_dir)
checker.run()