# ERP Exchange
## convert-APway_CURRENT.py
This is the original converter, which is creating hyperlinks for IE 11 favorites based on P2way.xml. It currently requires ElementTree and pywin32.
Note: It creates .lnk files, rather than .url files. Realizing this issue made me pick up this old problem again. I will soon fix this!

Update 2018-02-09:
The latest version in master branch creates .url-links instead of .lnk-links. pywin32 is not needed anymore! Tested with IE 11 on Windows 8.1 using P2way.xml from APplus version 4.3.0.
This is an example of the result:
https://github.com/lucas3F59/erp-exchange/blob/master/APplus_IE11.png
