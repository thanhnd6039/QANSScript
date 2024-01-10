*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Keywords ***
Create Table From The SS Revenue Cost Dump For Margin Report Source
    [Arguments]     ${ssRevenueCostDumpFilePath}   ${type}    ${year}     ${quarter}
    @{table}    Create List
    File Should Exist    ${ssRevenueCostDumpFilePath}
    Open Excel Document    ${ssRevenueCostDumpFilePath}    SSRevenueCostDump
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssRevenueCostDumpFilePath}
    ${sumQty}    Set Variable    0
    ${sumRev}    Set Variable    0
    ${sumCost}   Set Variable    0

    ${sumQty}   Convert To Integer    ${sumQty}
    ${sumRev}   Convert To Integer    ${sumRev}
    ${sumCost}   Convert To Integer    ${sumCost}

    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
        ${quarterColOnSS}          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
        IF    '${quarterColOnSS}' == 'Q${quarter}-${year}'
            ${parentClassColOnSS}      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
            IF    '${parentClassColOnSS}' == 'MEM' or '${parentClassColOnSS}' == 'STORAGE' or '${parentClassColOnSS}' == 'COMPONENTS' or '${parentClassColOnSS}' == 'NI'
                 ${oemGroupColOnSS}         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
                 ${pnColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
                 ${nextRowIndexOnSS}    Evaluate    ${rowIndexOnSS}+1
                 ${nextOemGroupColOnNS}     Read Excel Cell    row_num=${nextRowIndexOnSS}    col_num=2
                 ${nextPNColOnNS}           Read Excel Cell    row_num=${nextRowIndexOnSS}    col_num=11
                 IF    '${type}' == 'R'
                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=30
                      ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=28
                 END
                 IF    '${type}' == 'B'
                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=31
                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=33
                      ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=32
                      IF    '${costColOnSS}' == '${EMPTY}'
                           ${costColOnSS}   Set Variable    0
                      END

                 END
                 IF    '${type}' == 'CF'
                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=40
                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=42
                      ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=41
                 END
                 ${sumQty}      Evaluate    ${sumQty}+${qtyColOnSS}
                 ${sumRev}      Evaluate    ${sumRev}+${revColOnSS}
                 ${sumCost}     Evaluate    ${sumCost}+${costColOnSS}
                 IF    '${oemGroupColOnSS}' == '${nextOemGroupColOnNS}' and '${pnColOnSS}' == '${nextPNColOnNS}'
                      Continue For Loop
                 END

                 ${rowOnTable}   Create List
                 ...             ${oemGroupColOnSS}
                 ...             ${pnColOnSS}
                 ...             ${sumQty}
                 ...             ${sumRev}
                 ...             ${sumCost}
                 Append To List    ${table}   ${rowOnTable}
                 ${sumQty}     Set Variable    0
                 ${sumRev}     Set Variable    0
                 ${sumCost}    Set Variable    0

            END
        END
    END

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
    [Arguments]     ${reportFilePath}   ${type}     ${year}     ${quarter}
    @{table}    Create List
    ${posOfCol}         Set Variable    0
    ${posOfQtyCol}      Set Variable    0
    ${posOfRevCol}      Set Variable    0
    ${posOfCostCol}     Set Variable    0
    ${rowIndexTemp}     Set Variable    4
    ${rowIndexTemp}     Convert To Integer    ${rowIndexTemp}
    
    File Should Exist    ${reportFilePath}
    Open Excel Document    ${reportFilePath}    MarginReport
    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}
    ${oemGroupColOnReportTemp}  Set Variable    ${EMPTY}

    IF    '${type}' == 'R'
         ${searchStr}    Set Variable    ${year} Q${quarter} Actual
    END

    IF    '${type}' == 'B'
         ${searchStr}    Set Variable    ${year} Q${quarter} Backlog
    END

    IF    '${type}' == 'CF'
         ${searchStr}    Set Variable    ${year} Q${quarter} Customer Forecast
    END
    ${posOfCol}  Get Position Of Column    ${reportFilePath}    ${rowIndexTemp}     ${searchStr}
    ${posOfQtyCol}     Set Variable    ${posOfCol}
    ${posOfRevCol}     Evaluate    ${posOfCol}+1
    ${posOfCostCol}    Evaluate    ${posOfCol}+2

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
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfQtyCol}
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfRevCol}
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfCostCol}

        END
        IF    '${type}' == 'CF'
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=15
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=16
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=17
        END
        IF    '${pnColOnReport}' == 'Total'
             Continue For Loop
        END

        IF    '${qtyColOnReport}' == 'None'
             ${qtyColOnReport}  Set Variable    0
        END
        IF    '${revColOnReport}' == 'None'
             ${revColOnReport}  Set Variable    0
        END
        IF    '${costColOnReport}' == 'None'
             ${costColOnReport}     Set Variable    0
        END
        IF    '${qtyColOnReport}' == '0' and '${revColOnReport}' == '0' and '${costColOnReport}' == '0'
              Continue For Loop
        END

        ${rowOnTable}   Create List
        ...             ${oemGroupColOnReport}
        ...             ${pnColOnReport}
        ...             ${qtyColOnReport}
        ...             ${revColOnReport}
        ...             ${costColOnReport}
        Append To List    ${table}   ${rowOnTable}      
    END
    Close All Excel Documents
    [Return]    ${table}
    
Compare Data Between Margin Report And SS On NS
    [Arguments]     ${reportFilePath}   ${ssRevenueCostDumpFilePath}

    ${result}   Set Variable    ${True}
    @{reportTable}       Create List
    @{ssTable}           Create List
    ${type}     Set Variable    B
    ${year}     Set Variable    2024
    ${quarter}  Set Variable    1

#    ${reportTable}  Create Table For Margin Report    reportFilePath=${reportFilePath}    type=${type}  year=${year}   quarter=${quarter}
#     Write The Report Table To Excel    ${reportTable}
#    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
    ${ssTable}  Create Table From The SS Revenue Cost Dump For Margin Report Source    ssRevenueCostDumpFilePath=${ssRevenueCostDumpFilePath}     type=${type}     year=${year}   quarter=${quarter}
#    Write The Report Table To Excel    ${ssTable}

    [Return]    ${result}

 Write The Report Table To Excel
    [Arguments]     ${reportTable}
    ${reportTableFilePath}     Set Variable    ${DOWNLOAD_DIR}\\Report Table.xlsx
    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
    Open Excel Document    ${reportTableFilePath}    ReportTable
    FOR    ${rowIndexOnReportTable}    IN RANGE    0    ${numOfRowsOnReportTable}
        ${oemGroupColOnReportTable}                   Set Variable        ${reportTable}[${rowIndexOnReportTable}][0]
        ${pnColOnReportTable}                         Set Variable        ${reportTable}[${rowIndexOnReportTable}][1]
        ${qtyColOnReportTable}                        Set Variable        ${reportTable}[${rowIndexOnReportTable}][2]
        ${revColOnReportTable}                        Set Variable        ${reportTable}[${rowIndexOnReportTable}][3]
        ${costColOnReportTable}                       Set Variable        ${reportTable}[${rowIndexOnReportTable}][4]
        ${rowIndexTemp}    Evaluate    ${rowIndexOnReportTable}+2
        Write Excel Cell    row_num=${rowIndexTemp}    col_num=1    value=${oemGroupColOnReportTable}
        Write Excel Cell    row_num=${rowIndexTemp}    col_num=2    value=${pnColOnReportTable}
        Write Excel Cell    row_num=${rowIndexTemp}    col_num=3    value=${qtyColOnReportTable}
        Write Excel Cell    row_num=${rowIndexTemp}    col_num=4    value=${revColOnReportTable}
        Write Excel Cell    row_num=${rowIndexTemp}    col_num=5    value=${costColOnReportTable}
        Save Excel Document    ${reportTableFilePath}
    END
    Close All Excel Documents