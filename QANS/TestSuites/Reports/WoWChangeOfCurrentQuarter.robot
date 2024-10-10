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

#Verify the data of Backlog column for the OEM East table
#    [Tags]  WoWChangeReport_0018
#    [Documentation]     Verify the data of Backlog column for the OEM East table
#    Check Data For The OEM East Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog
#
#Verify the data of LOS column for the OEM East table
#    [Tags]  WoWChangeReport_0019
#    [Documentation]     Verify the data of LOS column for the OEM East table
#    Check Data For The OEM East Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS
#
#Verify the data of WoW(WoW of LOS column) column for the OEM East table
#    [Tags]  WoWChangeReport_0020
#    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM East table
#    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM East   10   WoW Of LOS
#
#Verify the data of GAP column for the OEM East table
#     [Tags]  WoWChangeReport_0021
#     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM East table
#     Check The GAP Data   ${wowChangeReportOnVDCFilePath}     ${sgWeeklyActionDBReportCurQFilePath}   ${wowChangeReportFilePath}    OEM East   11   GAP
#
#Veify the data of Comments column for the OEM East table
#     [Tags]  WoWChangeReport_0022
#     [Documentation]     Verify the data of Comments column for the OEM East table
#     Check The Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   OEM East
#
#Verify the data of Prev Quarter Ship column for the OEM West table
#    [Tags]  WoWChangeReport_0023
#    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
#    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}  2  7   Pre Q Ships
#
#Verify the data of Current Quarter Budget column for the OEM West table
#    [Tags]  WoWChangeReport_0024
#    [Documentation]     Verify the data of Current Q Budget column for the OEM East table
#    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  3   3   Current Q Budget
#
#Verify the data of LW Commit column for the OEM West table
#    [Tags]  WoWChangeReport_0025
#    [Documentation]     Verify the data of LW Commit column for the OEM East table
#    Check The Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   4   LW Commit   OEM West
#
#Verify the data of TW Commit column for the OEM West table
#    [Tags]  WoWChangeReport_0026
#    [Documentation]     Verify the data of TW Commit column for the OEM West table
#    Check The Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   5   TW Commit   OEM West
#
#Verify the data of Ships column for the OEM West table
#    [Tags]  WoWChangeReport_0027
#    [Documentation]     Verify the data of Ships column for the OEM West table
#    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  6   5   Ships
#
#Verify the data of WoW(WoW of Ships column) column for the OEM West table
#    [Tags]  WoWChangeReport_0028
#    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM West table
#    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM West   7   WoW Of Ships
#
#Verify the data of Backlog column for the OEM West table
#    [Tags]  WoWChangeReport_0029
#    [Documentation]     Verify the data of Backlog column for the OEM West table
#    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog
#
#Verify the data of LOS column for the OEM West table
#    [Tags]  WoWChangeReport_0030
#    [Documentation]     Verify the data of LOS column for the OEM West table
#    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS
#
#Verify the data of WoW(WoW of LOS column) column for the OEM West table
#    [Tags]  WoWChangeReport_0031
#    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM West table
#    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM West   10   WoW Of LOS
#
#Verify the data of GAP column for the OEM West table
#     [Tags]  WoWChangeReport_0032
#     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM West table
#     Check The GAP Data   ${wowChangeReportOnVDCFilePath}     ${sgWeeklyActionDBReportCurQFilePath}   ${wowChangeReportFilePath}    OEM West   11   GAP
#
#Verify the data of Comments column for the OEM West table
#    [Tags]  WoWChangeReport_0033
#    [Documentation]     Verify the data of Comments column for the OEM West table
#    Check The Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   OEM West










