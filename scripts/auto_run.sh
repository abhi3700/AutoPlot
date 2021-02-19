#!/bin/bash
#===================================================================================================================
# This script does update charts for all logbooks:
#===================================================================================================================
# TASK 1: Navigate to the home directory of QC logbooks i.e "Final QC Log Book"
cd "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\"

# -----------------------------------------------------------------------------------------------------------------------------------
# TASK 2: Run python charts
# ASBE1
cd ./ASH_09_10_LOG_BOOK/ASH10_QC_LOG_BOOK/
python ASH10_QC_LOG_BOOK.py
tskill excel
cd ../../

# ASFE1
cd ./ASH_09_10_LOG_BOOK/ASH09_QC_LOG_BOOK/
python ASH09_QC_LOG_BOOK.py
tskill excel
cd ../../

# REML1
cd ./CNT_02_LOG_BOOK/CNT02_QC_LOG_BOOK/
python CNT02_QC_LOG_BOOK.py
tskill excel
cd ../../

# REOX1A
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK/
python CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK.py
tskill excel
cd ../../../

# REOX1B
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK/
python CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK.py
tskill excel
cd ../../../

# REOX1C
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK/
python CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK.py
tskill excel
cd ../../../

# REPL1A
cd ./CNT_01_LOG_BOOK/CNT01_Ch_A_QC_LOG_BOOK/
python CNT01_Ch_A_QC_LOG_BOOK.py
tskill excel
cd ../../


# REPL1B
cd ./CNT_01_LOG_BOOK/CNT01_Ch_B_QC_LOG_BOOK/
python CNT01_Ch_B_QC_LOG_BOOK.py
tskill excel
cd ../../


# RESP1A
cd UNT_02_LOG_BOOK/UNT02_Ch_A_QC_LOG_BOOK/
python UNT02_Ch_A_QC_LOG_BOOK.py
tskill excel
cd ../../

# RESP1B
cd UNT_02_LOG_BOOK/UNT02_Ch_B_QC_LOG_BOOK/
python UNT02_Ch_B_QC_LOG_BOOK.py
tskill excel
cd ../../

#===================================================================================================================
#init
function pause() {
	read -p "$*"
}

#....
# call it
pause 'Press [Enter] key to continue...'
# rest of the script
# ....
# -----------------------------------------------------------------------------------------------------------------------------------

# TASK 3: Check file changes (if any)
# git add .
# git commit -m "Auto Update QC Charts | $(date +"%D %T")"
# git push origin master
