#!/bin/bash
#===================================================================================================================
# 1. go to the respective directory &
# 2. cp the QC logbook
# 3. Return back to area's home directory
#===================================================================================================================
# 1. Go to home directory of "Dry Etch"
cd ../examples/dry_etch/

#===================================================================================================================
# Repeat this for all dry_etch equipments
# 2-a. Go to respective equipments folder
# 2-b. Copy the the QC logbook from intranet
# 2-c. return back to to home directory
# ------------------------------------------------------------------------------------------------------------------
# ASBE1
cd ./ASH10_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\ASH_09_10_LOG_BOOK\\ASH10_QC_LOG_BOOK\\ASH10_QC_LOG_BOOK.xlsm" .
cd ../

# ------------------------------------------------------------------------------------------------------------------
# ASFE1
cd ./ASH09_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\ASH_09_10_LOG_BOOK\\ASH09_QC_LOG_BOOK\\ASH09_QC_LOG_BOOK.xlsm" .
cd ../

# ------------------------------------------------------------------------------------------------------------------
# REML1
cd ./CNT02_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_02_LOG_BOOK\\CNT02_QC_LOG_BOOK\\CNT02_QC_LOG_BOOK.xlsm" .
cd ../

# REOX1A
cd ./CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNE_02_LOG_BOOK\\NEW LOGBOOK WITH MORE DATA POINTS\\CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK\\CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK.xlsm" .
cd ../

# REOX1B
cd ./CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNE_02_LOG_BOOK\\NEW LOGBOOK WITH MORE DATA POINTS\\CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK\\CNE02_Ch_B_CSVIA_NEW_QC_LOG_BOOK.xlsm" .
cd ../

# REOX1C
cd ./CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNE_02_LOG_BOOK\\NEW LOGBOOK WITH MORE DATA POINTS\\CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK\\CNE02_Ch_C_RAA_NEW_QC_LOG_BOOK.xlsm" .
cd ../

# REPL1A
cd ./CNT01_Ch_A_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_Ch_A_QC_LOG_BOOK\\CNT01_Ch_A_QC_LOG_BOOK.xlsm" .
cd ../

# REPL1B
cd ./CNT01_Ch_B_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_Ch_B_QC_LOG_BOOK\\CNT01_Ch_B_QC_LOG_BOOK.xlsm" .
cd ../

# RESP1A
cd ./UNT02_Ch_A_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\UNT_02_LOG_BOOK\\UNT02_Ch_A_QC_LOG_BOOK\\UNT02_Ch_A_QC_LOG_BOOK.xlsm" .
cd ../

# RESP1B
cd ./UNT02_Ch_B_QC_LOG_BOOK/
cp -r "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\UNT_02_LOG_BOOK\\UNT02_Ch_B_QC_LOG_BOOK\\UNT02_Ch_B_QC_LOG_BOOK.xlsm" .
cd ../

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