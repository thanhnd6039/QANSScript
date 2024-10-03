*** Settings ***
Resource    ../../Pages/Reports/SGPage.robot

*** Variables ***
${flatSGFilePath}                           C:\\RobotFramework\\Downloads\\Flat Sales Gap.xlsx
${ssRCDFilePath}                            C:\\RobotFramework\\Downloads\\RevenueCostDump.xlsx
${ssRCDForPivotFilePath}                    C:\\RobotFramework\\Downloads\\RevenueCostDumpForPivot.xlsx


*** Test Cases ***
Verify REV QTY for every quarter by OEM Group
    @{listOfOEMGroupGetWrongData}   Create List 
    ${listOfOEMGroupFromSSRCD}     Get List Of OEM Groups From SS RCD    ssRCDFilePath=${ssRCDFilePath}    year=2024    quarter=1    attribute=REVQTY
    FOR    ${oemGroupFromSSRCD}    IN    @{listOfOEMGroupFromSSRCD}
        ${REVQTYByOEMGroupFromSSRCD}    Get Value By OEM Group From SS RCD     ssRCDFilePath=${ssRCDFilePath}      year=2024    quarter=1    oemGroup=${oemGroupFromSSRCD}    attribute=REVQTY
        ${oemGroupFromSSRCD}    Convert To Upper Case    ${oemGroupFromSSRCD}
        ${REVQTYByOEMGroupFromFlatSG}   Get Value By OEM Group From Flat SG    flatSGFilePath=${flatSGFilePath}    year=2024    quarter=1    oemGroup=${oemGroupFromSSRCD}    attribute=REVQTY
        ${diffREVQTY}   Evaluate    abs(${REVQTYByOEMGroupFromSSRCD}-${REVQTYByOEMGroupFromFlatSG})
        Log To Console    DIFF:${diffREVQTY}
        IF    ${diffREVQTY} >= 1
             Append To List    ${listOfOEMGroupGetWrongData}    ${oemGroupFromSSRCD}
        END
    END



