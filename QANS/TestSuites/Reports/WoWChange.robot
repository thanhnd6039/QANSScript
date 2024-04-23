*** Settings ***
Resource    ../../Pages/Reports/WoWChangePage.robot

*** Variables ***
${wowChangeReportFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${sgWeeklyActionDBReportPreQFilePath}   C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Pre Quarter.xlsx

*** Test Cases ***
#Verify The Prev Quarter Ship Data For The Strategic Table On The WoW Change Report
#    Compare The Prev Quarter Ship Data For The Strategic Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}

Verify The Prev Quarter Ship Data For The OEM East Table On The WoW Change Report
    Compare The Prev Quarter Ship Data For The OEM East Table Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportPreQFilePath}
