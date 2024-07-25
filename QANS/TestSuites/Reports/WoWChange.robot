*** Settings ***
Resource    ../../Pages/Reports/WoWChangePage.robot

*** Variables ***
${wowChangeReportFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeReportOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${sgWeeklyActionDBReportPreQFilePath}   C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Pre Quarter.xlsx
${sgWeeklyActionDBReportCurQFilePath}   C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Current Quarter.xlsx

*** Test Cases ***
Verify The Prev Quarter Ship Data For The Strategic Table
    [Tags]  WoWChangeReport_0001
    [Documentation]     Verify the data of Prev Q Ships column for the Strategic table
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}  2   7   Pre Q Ships

Verify The Current Quarter Budget Data For The Strategic Table
    [Tags]  WoWChangeReport_0002
    [Documentation]     Verify the data of Current Q Budget column for the Strategic table
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  3   3   Current Q Budget

Verify The LW Commit Data For The Strategic Table
    [Tags]  WoWChangeReport_0003
    [Documentation]     Verify the data of LW Commit column for the Strategic table
    Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   4   LW Commit   Strategic

Verify The TW Commit Data For The Strategic Table
    [Tags]  WoWChangeReport_0004
    [Documentation]     Verify the data of TW Commit column for the Strategic table
    Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   5   TW Commit   Strategic

Verify The Ships Data For The Strategic Table
    [Tags]  WoWChangeReport_0005
    [Documentation]     Verify the data of Ships column for the Strategic table
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  6   5   Ships

Verify The WoW Data Of Ships For The Strategic Table
    [Tags]  WoWChangeReport_0006
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the Strategic table
    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   Strategic   7   WoW Of Ships

Verify The Backlog Data For The Strategic Table
    [Tags]  WoWChangeReport_0007
    [Documentation]     Verify the data of Backlog column for the Strategic table
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog

Verify The LOS Data For The Strategic Table
    [Tags]  WoWChangeReport_0008
    [Documentation]     Verify the data of LOS column for the Strategic table
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS

Verify The WoW Data Of LOS For The Strategic Table
    [Tags]  WoWChangeReport_0009
    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the Strategic table
    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   Strategic   10   WoW Of LOS

Verify The GAP Data For The Strategic Table
    [Tags]  WoWChangeReport_0010
    [Documentation]     Verify the data of GAP(LOS - Commit) column for the Strategic table
    Log To Console    To do
#    Check The GAP Data   ${wowChangeReportFilePath}     Strategic   11   GAP

Verify The Comments Data For The Strategic Table
    [Tags]  WoWChangeReport_0011
    [Documentation]     Verify the data of Comments column for the Strategic table
    Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   Strategic

Verify The Prev Quarter Ship Data For The OEM East Table
    [Tags]  WoWChangeReport_0012
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
    Check Data For The OEM East Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}  2  7   Pre Q Ships

Verify The Current Quarter Budget Data For The OEM East Table
    [Tags]  WoWChangeReport_0013
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table
    Check Data For The OEM East Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  3   3   Current Q Budget

 Verify The LW Commit Data For The OEM East Table
     [Tags]  WoWChangeReport_0014
     [Documentation]     Verify the data of LW Commit column for the OEM East table
     Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   4   LW Commit   OEM East

Verify The TW Commit Data For The OEM East Table
    [Tags]  WoWChangeReport_0015
    [Documentation]     Verify the data of TW Commit column for the OEM East table
    Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   5   TW Commit   OEM East

Verify The Ships Data For The OEM East Table
    [Tags]  WoWChangeReport_0016
    [Documentation]     Verify the data of Ships column for the OEM East table
    Check Data For The OEM East Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  6   5   Ships

Verify The WoW Data Of Ships For The OEM East Table
    [Tags]  WoWChangeReport_0017
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM East table
    Check The WoW Data  ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM East   7   WoW Of Ships

Verify The Backlog Data For The OEM East Table
    [Tags]  WoWChangeReport_0018
    [Documentation]     Verify the data of Backlog column for the OEM East table
    Check Data For The OEM East Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog

Verify The LOS Data For The OEM East Table
    [Tags]  WoWChangeReport_0019
    [Documentation]     Verify the data of LOS column for the OEM East table
    Check Data For The OEM East Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS

Verify The WoW Data Of LOS For The OEM East Table
    [Tags]  WoWChangeReport_0020
    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM East table
    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM East   10   WoW Of LOS

 Verify The GAP Data For The OEM East Table
     [Tags]  WoWChangeReport_0021
     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM East table
     Log To Console    To do
 #    Check The GAP Data   ${wowChangeReportFilePath}     OEM East   11   GAP

 Verify The Comments Data For The OEM East Table
     [Tags]  WoWChangeReport_0022
     [Documentation]     Verify the data of Comments column for the OEM East table
     Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   OEM East

Verify The Prev Quarter Ship Data For The OEM West Table
    [Tags]  WoWChangeReport_0023
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}  2  7   Pre Q Ships

Verify The Current Quarter Budget Data For The OEM West Table
    [Tags]  WoWChangeReport_0024
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table
    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  3   3   Current Q Budget

Verify The LW Commit Data For The OEM West Table
    [Tags]  WoWChangeReport_0025
    [Documentation]     Verify the data of LW Commit column for the OEM East table
    Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   4   LW Commit   OEM West

Verify The TW Commit Data For The OEM West Table
    [Tags]  WoWChangeReport_0026
    [Documentation]     Verify the data of TW Commit column for the OEM West table
    Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   5   TW Commit   OEM West

Verify The Ships Data For The OEM West Table
    [Tags]  WoWChangeReport_0027
    [Documentation]     Verify the data of Ships column for the OEM West table
    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  6   5   Ships

Verify The WoW Data Of Ships For The OEM West Table
    [Tags]  WoWChangeReport_0028
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM West table
    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM West   7   WoW Of Ships

Verify The Backlog Data For The OEM West Table
    [Tags]  WoWChangeReport_0029
    [Documentation]     Verify the data of Backlog column for the OEM West table
    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog

Verify The LOS Data For The OEM West Table
    [Tags]  WoWChangeReport_0030
    [Documentation]     Verify the data of LOS column for the OEM West table
    Check Data For The OEM West Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS

Verify The WoW Data Of LOS For The OEM West Table
    [Tags]  WoWChangeReport_0031
    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM West table
    Check The WoW Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM West   10   WoW Of LOS

Verify The GAP Data For The OEM West Table
     [Tags]  WoWChangeReport_0032
     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM West table
     Log To Console    To do
 #    Check The GAP Data   ${wowChangeReportFilePath}     OEM West   11   GAP

Verify The Comments Data For The OEM West Table
    [Tags]  WoWChangeReport_0033
    [Documentation]     Verify the data of Comments column for the OEM West table
    Check The LW Commit Or Comment Data   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   OEM West










