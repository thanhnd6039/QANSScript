*** Settings ***
Resource    ../../Pages/Reports/SGPage.robot

*** Variables ***
${sgFilePath}                               C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx
${ssRCDFilePath}                            C:\\RobotFramework\\Downloads\\RevenueCostDump.xlsx
${ssRCDForPivotFilePath}                    C:\\RobotFramework\\Downloads\\RevenueCostDumpForPivot.xlsx


*** Test Cases ***
Verify REV QTY for every quarter by OEM Group
    ${listOfOEMGroupFromSSRCD}     Get List Of OEM Groups From SS RCD    ssRCDFilePath=${ssRCDFilePath}    year=2024    quarter=1    attribute=REVQTY
    FOR    ${oemGroupFromSSRCD}    IN    @{listOfOEMGroupFromSSRCD}

    END


#Verify REV for every quarter by OEM Group
#    Check Data For Every Quarter By OEM Group     ${sgFilePath}   ${ssRCDFilePath}     2024    3    AMOUNT   REV

