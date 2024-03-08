*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
Compare Data Between WoW Change Report And SG Weekly Action DB Report And SG Report
    [Arguments]     ${wowChangeSourceFilePath}  ${sgWeeklyActionDBFilePath}
    Create A Source File For WoW Change Report      ${wowChangeSourceFilePath}      ${sgWeeklyActionDBFilePath}

Create A Source File For WoW Change Report
    [Arguments]     ${wowChangeSourceFilePath}    ${sgWeeklyActionDBFilePath}
    
    File Should Exist    ${sgWeeklyActionDBFilePath}
    Open Excel Document    ${sgWeeklyActionDBFilePath}    SG_Weekly_Action_DB
    ${numOfRowsOnSGWeeklyActionDBFile}  Get Number Of Rows In Excel    ${sgWeeklyActionDBFilePath}
    
    FOR    ${rowIndexOnSGWeeklyActionDBFile}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBFile}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBFile}    col_num=1
        ${mainSaleRepCol}   Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBFile}    col_num=2
    END





