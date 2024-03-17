set c="C:\Program Files\Git\usr\bin\cp.exe"
set dst=C:\Users\Game\Code\xlsx_merge\xlsx_merge.dist
set libsrc=C:\Users\Game\AppData\Local\Programs\Python\Python311\Lib
set dllsrc=C:\Users\Game\AppData\Local\Programs\Python\Python311\DLLs
set srcdir=C:\Users\Game\AppData\Local\Programs\Python\Python311\Lib\site-packages


%c%  %srcdir%\openpyxl %dst% -rf 
%c%  %srcdir%\et_xmlfile %dst% -rf 
%c%  %libsrc%\xml  %dst% -rf
%c%  %dllsrc%\pyexpat.pyd %dst% -rf 