:: chooser.bat
:: launches a File... Open sort of file chooser and outputs choice(s) to the console
:: https://stackoverflow.com/a/15885133/1683264

@echo off
setlocal

set PATH=%PATH%;%HOMEPATH%\AppData\Local\Programs\Python\Python36-32;%PYTHONPATH%
set "open_cert_dir="(new-object -COM 'Shell.Application')^
.BrowseForFolder(0,'Select the folder with certificate files.',0,0).self.path""

set "open_recipient_list=Add-Type -AssemblyName System.Windows.Forms;$f = new-object Windows.Forms.OpenFileDialog;$f.InitialDirectory = pwd;$f.Filter = 'Text Files (*.csv)|*.csv|All Files (*.*)|*.*';$f.ShowHelp = $true;$f.Multiselect = $false;$f.Title = 'Select the file with recipient info';[void]$f.ShowDialog();if ($f.Multiselect) { $f.FileNames } else { $f.FileName }"

for /f "delims=" %%I in ('powershell -noprofile %open_cert_dir%') do (
	for /f "delims=" %%R in ('powershell -noprofile "%open_recipient_list%"') do (
		echo %%~I
				
		python scripts\prepare_attachments.py "%%~I" "%%~R"
	)
)
goto :EOF