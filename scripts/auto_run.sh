#!/bin/bash
#===================================================================================================================
# This script does 2 tasks for all logbooks:
# 1. Check if .xlsm file changed
# 2. If yes: `$ python run.py`, else do nothing.
#===================================================================================================================
# TASK 1: Navigate to the directory
cd "Z:\SECTIONS\DRY ETCH\QC Log Book\Final QC Log Book"

# -----------------------------------------------------------------------------------------------------------------------------------
# TASK 2: Run python charts
# ASFE1
cd ./ASH_09_10_LOG_BOOK/ASH09_QC_LOG_BOOK/
python run.py
cd ..
cd ..


# ASBE1
cd ./ASH_09_10_LOG_BOOK/ASH09_QC_LOG_BOOK/
python run.py
cd ..
cd ..


# REOX1A
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK/
python run.py
cd ..
cd ..
cd ..


# REOX1B
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK/
python run.py
cd ..
cd ..
cd ..


# REOX1C
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK/
python run.py
cd ..
cd ..
cd ..


# REPL1A
cd ./CNT_01_LOG_BOOK/CNT01_Ch_A_QC_LOG_BOOK/
python run.py
cd ..
cd ..


# REPL1B
cd ./CNT_01_LOG_BOOK/CNT01_Ch_B_QC_LOG_BOOK/
python run.py
cd ..
cd ..

# REML1
cd CNT_02_LOG_BOOK/CNT02_QC_LOG_BOOK/
python run.py
cd ..
cd ..

# RESP1A
cd UNT_02_LOG_BOOK/UNT02_Ch_A_QC_LOG_BOOK/
python run.py
cd ..
cd ..

# RESP1B
cd UNT_02_LOG_BOOK/UNT02_Ch_B_QC_LOG_BOOK/
python run.py
cd ..
cd ..

# -----------------------------------------------------------------------------------------------------------------------------------

# TASK 3: Check file changes (if any)
git add .
git commit -m "Auto Update QC Charts | $(date +"%D %T")"
git push origin master
