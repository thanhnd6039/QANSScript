*** Settings ***
Resource    ../../Pages/Reports/DWValueREVByCalendarQuarterPage.robot

*** Variables ***
${DWValueRevByCalendarQuarterReportFilePath}         C:\\RobotFramework\\Downloads\\Design Win Value - Rev by Calendar Quarter.xlsx
${ssMasterOPPFilePath}                               C:\\RobotFramework\\Downloads\\SS Master OPP.xlsx


*** Test Cases ***
Verify The OPP NO Data On The DW Value REV By Calendar Quarter Report
    Compare The OPP NO Data Between The DW Value REV By Calendar Quarter Report And SS Master OPP   ${DWValueRevByCalendarQuarterReportFilePath}    ${ssMasterOPPFilePath}

