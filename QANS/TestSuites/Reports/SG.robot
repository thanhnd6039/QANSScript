*** Settings ***
Resource    ../../Pages/Reports/SGPage.robot

*** Variables ***
${sgFilePath}                               C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx
${ssRCDFilePath}                            C:\\RobotFramework\\Downloads\\RevenueCostDump.xlsx

*** Test Cases ***
Verify REV for every quarter by OEM Group
    Check Data For Every Quarter By OEM Group     ${sgFilePath}   ${ssRCDFilePath}     2024    3    AMOUNT   REV

