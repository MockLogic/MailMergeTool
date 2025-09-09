@echo off
echo Starting Mail Merge Tool...
python "%~dp0MailMerge.py" %*
echo Exit code: %errorlevel%
pause