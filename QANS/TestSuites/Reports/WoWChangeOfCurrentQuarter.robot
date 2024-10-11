*** Settings ***
Resource    ../../Pages/Reports/WoWChangePage.robot

*** Test Cases ***
Verify Prev Q Ship for the OEM East table
    [Tags]  WoWChange_0001
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
    Check Data For The OEM East Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBPreQFilePath}  posOfColOnWoWChange=2  posOfColOnSGWeeklyActionDB=7   nameOfCol=Pre Q Ships

Verify Current Q Budget for the OEM East table
    [Tags]  WoWChange_0002
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table
    Check Data For The OEM East Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=3   posOfColOnSGWeeklyActionDB=3   nameOfCol=Current Q Budget

Verify LW Commit for the OEM East table
     [Tags]  WoWChange_0003
     [Documentation]     Verify the data of LW Commit column for the OEM East table
     Check The Commit Or Comment Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   posOfColOnWoWChange=4   nameOfCol=LW Commit   table=OEM East

Verify TW Commit for the OEM East table
    [Tags]  WoWChange_0004
    [Documentation]     Verify the data of TW Commit column for the OEM East table
    Check The Commit Or Comment Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   posOfColOnWoWChange=5   nameOfCol=TW Commit   table=OEM East

Verify Ships for the OEM East table
    [Tags]  WoWChange_0005
    [Documentation]     Verify the data of Ships column for the OEM East table
    Check Data For The OEM East Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=6   posOfColOnSGWeeklyActionDB=5   nameOfCol=Ships

Verify WoW of Ships for the OEM East table
    [Tags]  WoWChange_0006
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM East table
    Check The WoW Data  wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   table=OEM East   posOfColOnWoWChange=7   nameOfCol=WoW Of Ships

Verify Backlog for the OEM East table
    [Tags]  WoWChange_0007
    [Documentation]     Verify the data of Backlog column for the OEM East table
    Check Data For The OEM East Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=8   posOfColOnSGWeeklyActionDB=6   nameOfCol=Backlog

Verify LOS for the OEM East table
    [Tags]  WoWChange_0008
    [Documentation]     Verify the data of LOS column for the OEM East table
    Check Data For The OEM East Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=9   posOfColOnSGWeeklyActionDB=7   nameOfCol=LOS

Verify WoW of LOS for the OEM East table
    [Tags]  WoWChange_0009
    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM East table
    Check The WoW Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   table=OEM East   posOfColOnWoWChange=10   nameOfCol=WoW Of LOS

Verify GAP for the OEM East table
     [Tags]  WoWChange_0010
     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM East table
     Check The GAP Data   wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}     sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}   wowChangeFilePath=${wowChangeFilePath}    table=OEM East   posOfColOnWoWChange=11   nameOfCol=GAP

Veify the data of Comments column for the OEM East table
     [Tags]  WoWChange_0011
     [Documentation]     Verify the data of Comments column for the OEM East table
     Check The Commit Or Comment Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   posOfColOnWoWChange=12   nameOfCol=Comments   table=OEM East

Verify Prev Quarter Ship for the OEM West table
    [Tags]  WoWChange_0012
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
    Check Data For The OEM West Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBPreQFilePath}  posOfColOnWoWChange=2  posOfColOnSGWeeklyActionDB=7   nameOfCol=Pre Q Ships

Verify Current Quarter Budget for the OEM West table
    [Tags]  WoWChange_0013
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table
    Check Data For The OEM West Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=3   posOfColOnSGWeeklyActionDB=3   nameOfCol=Current Q Budget

Verify LW Commit for the OEM West table
    [Tags]  WoWChange_0014
    [Documentation]     Verify the data of LW Commit column for the OEM East table
    Check The Commit Or Comment Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   posOfColOnWoWChange=4   nameOfCol=LW Commit   table=OEM West

Verify TW Commit for the OEM West table
    [Tags]  WoWChange_0015
    [Documentation]     Verify the data of TW Commit column for the OEM West table
    Check The Commit Or Comment Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   posOfColOnWoWChange=5   nameOfCol=TW Commit   table=OEM West

Verify Ships for the OEM West table
    [Tags]  WoWChange_0016
    [Documentation]     Verify the data of Ships column for the OEM West table
    Check Data For The OEM West Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=6   posOfColOnSGWeeklyActionDB=5   nameOfCol=Ships

Verify WoW of Ships for the OEM West table
    [Tags]  WoWChange_0017
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM West table
    Check The WoW Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   table=OEM West   posOfColOnWoWChange=7   nameOfCol=WoW Of Ships

Verify Backlog for the OEM West table
    [Tags]  WoWChange_0018
    [Documentation]     Verify the data of Backlog column for the OEM West table
    Check Data For The OEM West Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=8   posOfColOnSGWeeklyActionDB=6   nameOfCol=Backlog

Verify LOS for the OEM West table
    [Tags]  WoWChange_0019
    [Documentation]     Verify the data of LOS column for the OEM West table
    Check Data For The OEM West Table  wowChangeFilePath=${wowChangeFilePath}      sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}  posOfColOnWoWChange=9   posOfColOnSGWeeklyActionDB=7   nameOfCol=LOS

Verify WoW of LOS for the OEM West table
    [Tags]  WoWChange_0020
    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM West table
    Check The WoW Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   table=OEM West   posOfColOnWoWChange=10   nameOfCol=WoW Of LOS

Verify GAP for the OEM West table
     [Tags]  WoWChange_0021
     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM West table
     Check The GAP Data   wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}     sgWeeklyActionDBFilePath=${sgWeeklyActionDBCurQFilePath}   wowChangeFilePath=${wowChangeFilePath}    table=OEM West   posOfColOnWoWChange=11   nameOfCol=GAP

Verify Comments for the OEM West table
    [Tags]  WoWChange_0022
    [Documentation]     Verify the data of Comments column for the OEM West table
    Check The Commit Or Comment Data   wowChangeFilePath=${wowChangeFilePath}     wowChangeOnVDCFilePath=${wowChangeOnVDCFilePath}   posOfColOnWoWChange=12   nameOfCol=Comments   table=OEM West










