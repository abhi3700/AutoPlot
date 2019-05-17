@echo off
rem ___________________________________________________
rem The program installs python packages for `AutoPlot`. 
rem NOTE: First, install Anaconda with terminal activated by enabling `Not Recommended` during the installation process.
rem ___________________________________________________
rem Clean all cache, unused packages
conda clean --all --index-cache --packages --yes --tempfiles 
pause