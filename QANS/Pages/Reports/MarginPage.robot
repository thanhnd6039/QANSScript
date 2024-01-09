*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Keywords ***
Create Table From The SS Revenue Cost Dump For Margin Report Source
    [Arguments]     ${ssRevenueCostDumpFilePath}   ${type}    ${year}     ${quarter}
    @{table}    Create List
#    @{listOfOEMGroups}  Create List
#    @{listOfPNs}  Create List

#    ${listOfOEMGroups}      Get List Of OEM Groups From SS Revenue Cost Dump    ${ssRevenueCostDumpFilePath}
#    ${listOfPNs}    Get List Of PNS From SS Revenue Cost Dump    ${ssRevenueCostDumpFilePath}
    
    File Should Exist    ${ssRevenueCostDumpFilePath}
    Open Excel Document    ${ssRevenueCostDumpFilePath}    SSRevenueCostDump
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssRevenueCostDumpFilePath}

    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
        ${quarterColOnSS}          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
        IF    '${quarterColOnSS}' == 'Q${quarter}-${year}'
            ${parentClassColOnSS}      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
            IF    '${parentClassColOnSS}' == 'MEM' or '${parentClassColOnSS}' == 'STORAGE' or '${parentClassColOnSS}' == 'COMPONENTS' or '${parentClassColOnSS}' == 'NI'
                 ${oemGroupColOnSS}         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
                 ${pnColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
                 IF    '${type}' == 'R'
                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=30
                      ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=28
                      IF    '${oemGroupColOnSS}' == '${oemGroup}' and '${pnColOnSS}' == '${pn}'
                           ${sumQty}    Evaluate    ${sumQty}+${qtyColOnSS}
                      END
#                                  Log To Console    OEM Group: ${oemGroupColOnSS}; PN: ${pnColOnSS}; QTY: ${qtyColOnSS}; REV: ${revColOnSS}; COST: ${costColOnSS}
                 END
            END
        END
    END

#    FOR    ${oemGroup}    IN    @{listOfOEMGroups}
#        FOR    ${pn}    IN    @{listOfPNs}
#              ${sumQty}   Set Variable    0
#              FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
#                    ${quarterColOnSS}          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
#                    IF    '${quarterColOnSS}' == 'Q${quarter}-${year}'
#                        ${parentClassColOnSS}      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
#                        IF    '${parentClassColOnSS}' == 'MEM' or '${parentClassColOnSS}' == 'STORAGE' or '${parentClassColOnSS}' == 'COMPONENTS' or '${parentClassColOnSS}' == 'NI'
#                             ${oemGroupColOnSS}         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
#                             ${pnColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
#                             IF    '${type}' == 'R'
#                                  ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
#                                  ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=30
#                                  ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=28
#                                  IF    '${oemGroupColOnSS}' == '${oemGroup}' and '${pnColOnSS}' == '${pn}'
#                                       ${sumQty}    Evaluate    ${sumQty}+${qtyColOnSS}
#                                  END
##                                  Log To Console    OEM Group: ${oemGroupColOnSS}; PN: ${pnColOnSS}; QTY: ${qtyColOnSS}; REV: ${revColOnSS}; COST: ${costColOnSS}
#                             END
#                        END
#                    END
#              END
#              Log To Console    OEM Group: ${oemGroup}; PN: ${pn}; QTY: ${sumQty}
#        END
#    END

    


    
    [Return]    ${table}

Get List Of OEM Groups From SS Revenue Cost Dump
    [Arguments]     ${ssRevenueCostDumpFilePath}
    @{listOfOEMGroups}    Create List

    File Should Exist    ${ssRevenueCostDumpFilePath}
    Open Excel Document    ${ssRevenueCostDumpFilePath}    SSRevenueCostDump
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssRevenueCostDumpFilePath}

    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
        ${oemGroupColOnSS}         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
        IF    '${oemGroupColOnSS}' != '${EMPTY}'
             Append To List    ${listOfOEMGroups}    ${oemGroupColOnSS}
        END
    END
    ${listOfOEMGroups}      Remove Duplicates    ${listOfOEMGroups}
    Close All Excel Documents
    [Return]    ${listOfOEMGroups}

Get List Of PNS From SS Revenue Cost Dump
    [Arguments]     ${ssRevenueCostDumpFilePath}
    @{listOfPNs}    Create List

    File Should Exist    ${ssRevenueCostDumpFilePath}
    Open Excel Document    ${ssRevenueCostDumpFilePath}    SSRevenueCostDump
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssRevenueCostDumpFilePath}

    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
        ${pnColOnSS}         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
        IF    '${pnColOnSS}' != '${EMPTY}'
             Append To List    ${listOfPNs}    ${pnColOnSS}
        END

    END
    ${listOfPNs}      Remove Duplicates    ${listOfPNs}
    Close All Excel Documents
    [Return]    ${listOfPNs}
    
Create Table For Margin Report
    [Arguments]     ${reportFilePath}   ${type}
    @{table}    Create List
    
    File Should Exist    ${reportFilePath}
    Open Excel Document    ${reportFilePath}    MarginReport
    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}
    ${oemGroupColOnReportTemp}  Set Variable    ${EMPTY}

    FOR    ${rowIndexOnReport}    IN RANGE    7    ${numOfRowsOnReport}+1
        ${oemGroupColOnReport}      Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
        IF    '${oemGroupColOnReport}' == 'None'
             ${oemGroupColOnReport}     Set Variable    ${oemGroupColOnReportTemp}
        ELSE
             ${oemGroupColOnReportTemp}     Set Variable    ${oemGroupColOnReport}
        END
        ${pnColOnReport}             Read Excel Cell    row_num=${rowIndexOnReport}    col_num=2
        IF    '${type}' == 'R'
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=5
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=6
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=7
        END
        IF    '${type}' == 'B'
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=10
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=11
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=12
        END
        IF    '${type}' == 'CF'
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=15
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=16
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=17
        END

        ${rowOnTable}   Create List
        ...             ${oemGroupColOnReport}
        ...             ${pnColOnReport}
        ...             ${qtyColOnReport}
        ...             ${revColOnReport}
        ...             ${costColOnReport}
        Append To List    ${table}   ${rowOnTable}      
    END
    [Return]    ${table}
    
Compare Data Between Margin Report And SS On NS
    [Arguments]     ${reportFilePath}   ${ssRevenueCostDumpFilePath}
    ${result}   Set Variable    ${True}
    @{reportTable}       Create List
    @{ssTable}           Create List
    ${type}     Set Variable    R

#    ${reportTable}  Create Table For Margin Report    reportFilePath=${reportFilePath}    type=${type}
#    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
    ${ssTable}  Create Table From The SS Revenue Cost Dump For Margin Report Source    ssRevenueCostDumpFilePath=${ssRevenueCostDumpFilePath}     type=${type}     year=2024   quarter=1

    [Return]    ${result}