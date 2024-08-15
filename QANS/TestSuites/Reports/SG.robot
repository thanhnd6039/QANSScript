*** Settings ***
Resource    ../../Pages/Reports/SGPage.robot

*** Variables ***
${sgReportFilePath}               C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${ssRevenueCostDumpFilePath}      C:\\RobotFramework\\Downloads\\RevenueCostDump.xlsx

*** Test Cases ***
Verify REV for every quarter by OEM Group
    Check data for every quarter by OEM Group     ${sgReportFilePath}   ${ssRevenueCostDumpFilePath}     2024    1   REV

