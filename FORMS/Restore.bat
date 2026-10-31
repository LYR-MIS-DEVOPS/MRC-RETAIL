@echo off
echo Restoring your BACKUP for mrcwarehouse.rar, Please wait.....
F:\Dropbox\MRCWAREHOUSE\rar e -plaborday mrcwarehouse.rar
echo Please wait...Updating server databases
F:\Dropbox\MRCWAREHOUSE\mysql -uroot -plaborday mrcwarehouse < mrcwarehouse.SQL
echo Data Successfully loaded
pause
