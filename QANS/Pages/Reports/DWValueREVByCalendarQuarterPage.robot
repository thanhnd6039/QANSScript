*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
Compare The OPP NO Data Between The DW Value REV By Calendar Quarter Report And SS Master OPP
    [Arguments]     ${DWValueRevByCalendarQuarterReportFilePath}    ${ssMasterOPPFilePath}

    File Should Exist    ${DWValueRevByCalendarQuarterReportFilePath}
    Open Excel Document    ${DWValueRevByCalendarQuarterReportFilePath}    doc_id=DWValueREVByCalendarQuarterReport

    File Should Exist    ${ssMasterOPPFilePath}
    Open Excel Document    ${ssMasterOPPFilePath}    doc_id=SSMasterOPP

    Switch Current Excel Document    doc_id=SSMasterOPP
    ${numOfRowsOnSSMasterOPP}   Get Number Of Rows In Excel    ${ssMasterOPPFilePath}

    Switch Current Excel Document    doc_id=DWValueREVByCalendarQuarterReport
    ${numOfRowsOfDWValueREVByCalendarQuarterReport}     Get Number Of Rows In Excel    ${DWValueRevByCalendarQuarterReportFilePath}

    FOR    ${rowIndexOnDWValueREVByCalendarQuarterReport}    IN RANGE    3    ${numOfRowsOfDWValueREVByCalendarQuarterReport}+1
        ${oemColOnDWValueREVByCalendarQuarterReport}    Read Excel Cell    row_num=${rowIndexOnDWValueREVByCalendarQuarterReport}    col_num=2
        ${pnColOnDWValueREVByCalendarQuarterReport}     Read Excel Cell    row_num=${rowIndexOnDWValueREVByCalendarQuarterReport}    col_num=3
        ${oppNoColOnDWValueREVByCalendarQuarterReport}  Read Excel Cell    row_num=${rowIndexOnDWValueREVByCalendarQuarterReport}    col_num=4
        Switch Current Excel Document    doc_id=SSMasterOPP  
        Log To Console    OPP:${oppNoColOnDWValueREVByCalendarQuarterReport}      
        IF    '${oppNoColOnDWValueREVByCalendarQuarterReport}' == 'None'
            IF    '${oemColOnDWValueREVByCalendarQuarterReport}' == 'None'
                 BREAK
            ELSE
                FOR    ${rowIndexOnSSMasterOPP}    IN RANGE    2    ${numOfRowsOnSSMasterOPP}+1
                    ${oemColOnSSMasterOPP}  Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=6
                    ${pnColOnSSMasterOPP}   Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=12
                    ${oppColOnSSMasterOPP}  Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=2
                    IF    '${oemColOnDWValueREVByCalendarQuarterReport}' == '${oemColOnSSMasterOPP}' and '${pnColOnDWValueREVByCalendarQuarterReport}' == '${pnColOnSSMasterOPP}'
                         Log To Console    The OPP ${oppColOnSSMasterOPP} is not found on the DW Value REV by Calendar Quarter report
                         BREAK
                    END
                END
            END
        ELSE
            ${isSearchOPP}  Set Variable    ${False}
            FOR    ${rowIndexOnSSMasterOPP}    IN RANGE    2    ${numOfRowsOnSSMasterOPP}+1
                ${oemColOnSSMasterOPP}  Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=6
                ${pnColOnSSMasterOPP}   Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=12
                ${oppColOnSSMasterOPP}  Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=2
                IF    '${oppNoColOnDWValueREVByCalendarQuarterReport}' == '${oppColOnSSMasterOPP}'
                     ${isSearchOPP}  Set Variable   ${True}
                     IF    '${oemColOnDWValueREVByCalendarQuarterReport}' == '${oemColOnSSMasterOPP}' and '${pnColOnDWValueREVByCalendarQuarterReport}' == '${pnColOnSSMasterOPP}'
                          BREAK
                     ELSE
                         Log To Console    OPP:${oppNoColOnDWValueREVByCalendarQuarterReport}
                     END
                END
            END
            IF    '${isSearchOPP}' == '${False}'
                 Log To Console    The OPP ${oppNoColOnDWValueREVByCalendarQuarterReport} is not found on NS
            END
        END
        Switch Current Excel Document    doc_id=DWValueREVByCalendarQuarterReport
         
    END