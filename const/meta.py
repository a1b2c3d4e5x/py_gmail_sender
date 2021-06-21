
from .constant import constant

class __ConstMeta(object):
   @constant
   def BUILD_DATE() -> str:
      return '2021-06-20'

   @constant
   def BUILD_VERSION() -> str:
      return '0.0.1'

   @constant
   def DEVERPER_NAME() -> str:
      return 'Guo Hao Chen'

   @constant
   def DEVERPER_EMAIL() -> str:
      return 'a1b2c3d4e5x@gmail.com'

meta = __ConstMeta()
