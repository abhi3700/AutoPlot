#!/bin/bash
#===================================================================================================================
# This script does update charts for all logbooks:
#===================================================================================================================
# TASK 1: Navigate to the home directory of QC logbooks i.e "Final QC Log Book"
cd "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\"

# -----------------------------------------------------------------------------------------------------------------------------------
# TASK 2: Run python charts
# ASBE1
echo "ASBE1: Opening charts..."
cd ./ASH_09_10_LOG_BOOK/ASH10_QC_LOG_BOOK/
python ASH10_QC_LOG_BOOK.py
tskill excel
cd ../../
echo "ASBE1: Charts opened successfully."

echo "----------------------------------"

# ASFE1
echo "ASFE1: Opening charts..."
cd ./ASH_09_10_LOG_BOOK/ASH09_QC_LOG_BOOK/
python ASH09_QC_LOG_BOOK.py
tskill excel
echo "ASFE1: Charts opened successfully."
cd ../../

echo "----------------------------------"

# REML1
echo "REML1: Opening charts..."
cd ./CNT_02_LOG_BOOK/CNT02_QC_LOG_BOOK/
python CNT02_QC_LOG_BOOK.py
tskill excel
echo "REML1: Charts opened successfully."
cd ../../

echo "----------------------------------"

# REOX1A
echo "REOX1A: Opening charts..."
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK/
python CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK.py
tskill excel
echo "REOX1A: Charts opened successfully."
cd ../../../

echo "----------------------------------"

# REOX1B
echo "REOX1B: Opening charts..."
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK/
python CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK.py
tskill excel
echo "REOX1B: Charts opened successfully."
cd ../../../

echo "----------------------------------"

# REOX1C
echo "REOX1C: Opening charts..."
cd ./CNE_02_LOG_BOOK/NEW\ LOGBOOK\ WITH\ MORE\ DATA\ POINTS/CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK/
python CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK.py
tskill excel
echo "REOX1C: Charts opened successfully."
cd ../../../

echo "----------------------------------"

# REPL1A
echo "REPL1A: Opening charts..."
cd ./CNT_01_LOG_BOOK/CNT01_Ch_A_QC_LOG_BOOK/
python CNT01_Ch_A_QC_LOG_BOOK.py
tskill excel
echo "REPL1A: Charts opened successfully."
cd ../../

echo "----------------------------------"

# REPL1B
echo "REPL1B: Opening charts..."
cd ./CNT_01_LOG_BOOK/CNT01_Ch_B_QC_LOG_BOOK/
python CNT01_Ch_B_QC_LOG_BOOK.py
tskill excel
echo "REPL1B: Charts opened successfully."
cd ../../

echo "----------------------------------"

# RESP1A
echo "RESP1A: Opening charts..."
cd UNT_02_LOG_BOOK/UNT02_Ch_A_QC_LOG_BOOK/
python UNT02_Ch_A_QC_LOG_BOOK.py
tskill excel
echo "RESP1A: Charts opened successfully."
cd ../../

echo "----------------------------------"

# RESP1B
echo "RESP1B: Opening charts..."
cd UNT_02_LOG_BOOK/UNT02_Ch_B_QC_LOG_BOOK/
python UNT02_Ch_B_QC_LOG_BOOK.py
tskill excel
echo "RESP1B: Charts opened successfully."
cd ../../

echo "----------------------------------"
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
