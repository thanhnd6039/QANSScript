*** Settings ***
Resource    ../../Pages/Reports/WoWChangePage.robot

*** Test Cases ***
Verify The Prev Quarter Ship Data For The WoW Change Report
    ${wowChangeReportFilePath}    Set Variable    C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
    ${sgWeeklyActionDBReportFilePath}   Set Variable    C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB Q1.xlsx
    Compare The Prev Quarter Ship Data Between WoW Change Report And SG Weekly Action DB Report  ${wowChangeReportFilePath}      ${sgWeeklyActionDBReportFilePath}


