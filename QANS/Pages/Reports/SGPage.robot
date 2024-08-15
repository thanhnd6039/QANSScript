*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
Check data for every quarter by OEM Group
    [Arguments]     ${sgReportFilePath}   ${ssRevenueCostDumpFilePath}     ${year}     ${quarter}   ${nameOfCol}
    Log To Console    test