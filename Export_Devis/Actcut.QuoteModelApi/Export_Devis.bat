SET DEVIS=3
rem %date%
set folder=C:\AlmaCAM\Bin\Actcut.QuoteModelApi.exe
set file=C:\AlmaCAM\Bin\AlmaCam_Clipper\_Clipper\Devis\Trans_%DEVIS%.txt
set order=%DEVIS%
set Db=AlmaCAM_C
set Db=AlmaCAM10

    	echo %folder%
	echo %file%
pause

%folder% Action=ExportQuote Db=%Db% QuoteNumber=%DEVIS% OrderNumber=%order% ExportFile=%file% 