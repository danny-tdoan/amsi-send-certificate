<# : chooser.bat
:: launches a File... Open sort of file chooser and outputs choice(s) to the console
:: https://stackoverflow.com/a/15885133/1683264

@echo off
setlocal
set PATH=%PATH%;%HOMEPATH%\AppData\Local\Programs\Python\Python36-32;%PYTHONPATH%

set "open_message_template=Add-Type -AssemblyName System.Windows.Forms;$f = new-object Windows.Forms.OpenFileDialog;$f.InitialDirectory = pwd;$f.Filter = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*';$f.ShowHelp = $true;$f.Multiselect = $false;$f.Title = 'Select main content of the email';[void]$f.ShowDialog();if ($f.Multiselect) { $f.FileNames } else { $f.FileName }"
set "open_recipients=Add-Type -AssemblyName System.Windows.Forms;$f = new-object Windows.Forms.OpenFileDialog;$f.InitialDirectory = pwd;$f.Filter = 'Text Files (*.csv)|*.csv|All Files (*.*)|*.*';$f.ShowHelp = $true;$f.Multiselect = $false;$f.Title = 'Select recipient list';[void]$f.ShowDialog();if ($f.Multiselect) { $f.FileNames } else { $f.FileName }"

for /f "delims=" %%I in ('powershell -noprofile "%open_message_template%"') do (
	for /f "delims=" %%R in ('powershell -noprofile "%open_recipients%"') do (
		echo Sending email with %%~I and %%~R
		
		python scripts\save_email_draft.py "%%~I" "%%~R"
	)
)
goto :EOF