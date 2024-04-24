*** Settings ***
Resource    ../../Pages/Reports/WoWChangePage.robot

*** Variables ***
${wowChangeReportFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${sgWeeklyActionDBReportPreQFilePath}   C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Pre Quarter.xlsx
${sgWeeklyActionDBReportCurQFilePath}   C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Current Quarter.xlsx

*** Test Cases ***
Verify The Prev Quarter Ship Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0001
    Compare The Prev Quarter Ship Data For The Strategic Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}

Verify The Prev Quarter Ship Data For The OEM East Table On The WoW Change Report
    [Tags]  WoWChangeReport_0002
    Compare The Prev Quarter Ship Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}

Verify The Prev Quarter Ship Data For The OEM West Table On The WoW Change Report
    [Tags]  WoWChangeReport_0003
    Compare The Prev Quarter Ship Data For The OEM West Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}

Verify The Current Quarter Budget Data For The Strategic Table On The WoW Change Report
    [Tags]  WoWChangeReport_0004
    Compare The Current Quarter Budget Data For The Strategic Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportCurQFilePath}
