
from const.color import color
from const.meta import meta
from core.send import SendMail

import sys 
import os

class ArgvParser:
   def __print_color_title(self, text: str):
      print(color.BOLD + color.BLUE + text + color.END)

   def __print_color_description(self, preText: str, postText: str):
      print('  ' + color.BOLD + color.DARKCYAN + '--' + preText + ':' + color.END, postText)

   # 從 excel / csv 取得並寄送 email
   def send(self, path: str):
      sender_csv = '/resource/sender_info.csv'
      mail_format = '/resource/mail_template.html'

      full_sender_csv = os.getcwd() + sender_csv
      full_mail_format = os.getcwd() + mail_format
      full_receiver_path = os.getcwd() + '/' + path

      if False == os.path.isfile(full_receiver_path):
         print('Not exist file:', full_receiver_path)
         return
      if False == os.path.isfile(full_sender_csv):
         print('Missing file:', full_sender_csv)
         return
      if False == os.path.isfile(full_mail_format):
         print('Missing file:', full_mail_format)
         return

      file_ext = path.split('.')[-1]

      if 'xlsx' == file_ext:
         SendMail.from_excel(full_sender_csv, full_mail_format, full_receiver_path)
      elif 'csv' == file_ext:
         SendMail.from_csv(full_sender_csv, full_mail_format, full_receiver_path)
      else:
         print('Just support file format .excel and .csv')
         return

   # 關於此程式的說明
   def about(self):
      print('會發送固定版型的 email 到指定的多個信箱')

   # 關於此程式的開發資訊
   def info(self):
      print('Build Date     :', meta.BUILD_DATE)
      print('Build Version  :', 'v' + meta.BUILD_VERSION)
      print('Developer Name :', meta.DEVERPER_NAME)
      print('Developer Email:', meta.DEVERPER_EMAIL)

   # 未給任何參數
   def none(self):
      self.__print_color_title('指令說明')
      self.__print_color_description(self.send.__name__ + ' [*receiver csv]', '發送 email')
      self.__print_color_description(self.about.__name__, '關於此程式的說明')
      self.__print_color_description(self.info.__name__, '關於此程式的開發資訊')

# 判斷輸入的參數指令
def __argv_is_cmd(fn_name: str) -> bool:
   if 2 <= len(sys.argv):
      return ('--' + fn_name) == sys.argv[1]
   return True

# 處理 argv 的參數
def __parse_argv():
   parser = ArgvParser()

   if 2 == len(sys.argv):
      if __argv_is_cmd(parser.about.__name__):
         return parser.about()
      elif __argv_is_cmd(parser.info.__name__):
         return parser.info()

   if 3 == len(sys.argv):
      if __argv_is_cmd(parser.send.__name__):
         return parser.send(sys.argv[2])

   return parser.none()

# 主程式進入口
if __name__ == '__main__':
   __parse_argv()




