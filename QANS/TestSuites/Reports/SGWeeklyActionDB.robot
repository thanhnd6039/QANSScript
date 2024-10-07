*** Settings ***
Resource    ../../Pages/Reports/SGWeeklyActionDBPage.robot

*** Variables ***
${flatSGFilePath}                           C:\\RobotFramework\\Downloads\\Flat Sales Gap.xlsx
${SGWeeklyActionDBFilePath}                 C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB.xlsx
${ssApprovedBudgetFilePath}                 C:\\RobotFramework\\Downloads\\SSApprovedBudget.xlsx

*** Test Cases ***
Verify BGT for every quarter
#    Create SS Approved Budget As Table Pivot    ssApprovedBudgetFilePath=${ssApprovedBudgetFilePath}    year=2024    quarter=4
    Create Flat SG As Table Pivot    flatSGFilePath=${flatSGFilePath}   year=2024   quarter=4   attribute=R
#    ${listOfOEMGroupFromSGWeeklyActionDB}   Get List Of OEM Group From SG Weekly Action DB    SGWeeklyActionDBFilePath=${SGWeeklyActionDBFilePath}
#    FOR    ${oemGroupFromSGWeeklyActionDB}    IN    @{listOfOEMGroupFromSGWeeklyActionDB}
#        ${bgtByOEMGroupFromSGWeeklyActionDB}  Get Value By OEM Group From SG Weekly Action DB    SGWeeklyActionDBFilePath=${SGWeeklyActionDBFilePath}    quarter=4    oemGroup=${oemGroupFromSGWeeklyActionDB}    attribute=BGT
#        ${bgtByOEMGroupFromApprovedBudget}    Get Value By OEM Group From Approved Budget    approvedBudgetFilePath=${approvedBudgetFilePath}    year=2024    quarter=4    oemGroup=${oemGroupFromSGWeeklyActionDB}
#        ${diffBGT}  Evaluate    abs(${bgtByOEMGroupFromSGWeeklyActionDB}-${bgtByOEMGroupFromApprovedBudget})
#        IF    ${diffBGT} >= 1
#             Log To Console    OEMGroup:${oemGroupFromSGWeeklyActionDB}
#        END
#
#    END



