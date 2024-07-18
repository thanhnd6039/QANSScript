*** Settings ***
Resource    ../../Pages/Reports/WoWChangePage.robot

*** Variables ***
${wowChangeReportFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeReportOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${sgWeeklyActionDBReportPreQFilePath}   C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Pre Quarter.xlsx
${sgWeeklyActionDBReportCurQFilePath}   C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Current Quarter.xlsx

*** Test Cases ***
Verify The Prev Quarter Ship Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0001
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}  2   7   Pre Q Ships

Verify The Current Quarter Budget Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0002
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  3   3   Current Q Budget

Verify The Ships Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0003
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  6   5   Ships

Verify The Backlog Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0004
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog

Verify The LOS Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0005
    Check Data For The Strategic Table  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS

Verify The Prev Quarter Ship Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0006
    Compare Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}  2  7   Pre Q Ships

Verify The Current Quarter Budget Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0007
    Compare Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  3   3   Current Q Budget

Verify The Ships Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0008
    Compare Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  6   5   Ships

Verify The Backlog Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0009
    Compare Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog

Verify The LOS Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0010
    Compare Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS

Verify The Prev Quarter Ship Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0011
    Compare Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}  2  7   Pre Q Ships

Verify The Current Quarter Budget Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0012
    Compare Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  3   3   Current Q Budget

Verify The Ships Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0013
    Compare Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  6   5   Ships

Verify The Backlog Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0014
    Compare Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  8   6   Backlog

Verify The LOS Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0015
    Compare Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}  9   7   LOS

Verify The LW Commit Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0016
    Compare The LW Commit Or Comment Data Between WoW Change Report And WoW Change Report On VDC   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   4   LW Commit   Strategic

Verify The LW Commit Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0017
    Compare The LW Commit Or Comment Data Between WoW Change Report And WoW Change Report On VDC   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   4   LW Commit   OEM East

Verify The LW Commit Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0018
    Compare The LW Commit Or Comment Data Between WoW Change Report And WoW Change Report On VDC   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   4   LW Commit   OEM West

Verify The Comments Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0019
    Compare The LW Commit Or Comment Data Between WoW Change Report And WoW Change Report On VDC   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   Strategic

Verify The Comments Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0020
    Compare The LW Commit Or Comment Data Between WoW Change Report And WoW Change Report On VDC   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   OEM East

Verify The Comments Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0021
    Compare The LW Commit Or Comment Data Between WoW Change Report And WoW Change Report On VDC   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   12   Comments   OEM West

Verify The WoW Data Of Ships For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0022
    Verify The WoW Data On WoW Change Report   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   Strategic   7   WoW Of Ships

Verify The WoW Data Of Ships For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0023
    Verify The WoW Data On WoW Change Report   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM East   7   WoW Of Ships

Verify The WoW Data Of Ships For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0024
    Verify The WoW Data On WoW Change Report   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM West   7   WoW Of Ships

Verify The WoW Data Of LOS For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0025
    Verify The WoW Data On WoW Change Report   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   Strategic   10   WoW Of LOS

Verify The WoW Data Of LOS For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0026
    Verify The WoW Data On WoW Change Report   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM East   10   WoW Of LOS

Verify The WoW Data Of LOS For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0027
    Verify The WoW Data On WoW Change Report   ${wowChangeReportFilePath}     ${wowChangeReportOnVDCFilePath}   OEM West   10   WoW Of LOS