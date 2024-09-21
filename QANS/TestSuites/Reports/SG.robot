*** Settings ***
Resource    ../../Pages/Reports/SGPage.robot

*** Variables ***
${sgFilePath}                               C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx
${ssRCDFilePath}                            C:\\RobotFramework\\Downloads\\RevenueCostDump.xlsx

*** Test Cases ***
Testcase1
    Convert SS RCD To Pivot And Export To Excel    ssRCDFilePath=${ssRCDFilePath}   2024    1

#Verify REV for every quarter by OEM Group
#    Check Data For Every Quarter By OEM Group     ${sgFilePath}   ${ssRCDFilePath}     2024    3    AMOUNT   REV

