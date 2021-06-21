
from .constant import constant

class __ConstColor(object):
   @constant
   def PURPLE() -> str:
      return '\033[95m'

   @constant
   def CYAN() -> str:
      return '\033[96m'

   @constant
   def DARKCYAN() -> str:
      return '\033[36m'

   @constant
   def BLUE() -> str:
      return '\033[94m'

   @constant
   def GREEN() -> str:
      return '\033[92m'

   @constant
   def YELLOW() -> str:
      return '\033[93m'

   @constant
   def RED() -> str:
      return '\033[91m'

   @constant
   def BOLD() -> str:
      return '\033[1m'

   @constant
   def UNDERLINE() -> str:
      return '\033[4m'

   @constant
   def END() -> str:
      return '\033[0m'

color = __ConstColor()

def output_error(text: str):
   print(color.RED + text + color.END)