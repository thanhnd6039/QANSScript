*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
Compare The Prev Quarter Ship Data Between WoW Change Report And SG Weekly Action DB Report
    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}
    
    File Should Exist    ${wowChangeReportFilePath}
    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport
    
    File Should Exist    ${sgWeeklyActionDBReportFilePath}
    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport
    Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
    ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
    
    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    6
        Switch Current Excel Document    doc_id=WoWChangeReport
        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
        ${prevQShipColOnWoWChangeReport}     Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=2
        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
            ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
            ${revColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=5

        END
        
    END







