*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${RESULT_FILE_PATH}     ${RESULT_DIR}\\MarginReport\\MarginReportResult.xlsx

*** Keywords ***
Create Table From The SS Revenue Cost Dump For Margin Report Source
    [Arguments]     ${ssRevenueCostDumpFilePath}   ${type}    ${year}     ${quarter}
    @{table}    Create List
    @{dataTableFromSSRevenueCostDump}   Create List
    ${sumQty}    Set Variable    0
    ${sumRev}    Set Variable    0
    ${sumCost}   Set Variable    0

    ${sumQty}   Convert To Integer    ${sumQty}
    ${sumRev}   Convert To Integer    ${sumRev}
    ${sumCost}   Convert To Integer    ${sumCost}
    
    ${dataTableFromSSRevenueCostDump}   Get Data From The SS Revenue Cost Dump For Every Quarter    ssRevenueCostDumpFilePath=${ssRevenueCostDumpFilePath}    type=${type}    year=${year}    quarter=${quarter}
    ${numOfRowsOfDataTable}     Get Length    ${dataTableFromSSRevenueCostDump}

    FOR    ${rowIndexOnDataTable}    IN RANGE    0    ${numOfRowsOfDataTable}
        ${oemGroupColOnDaTable}           Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][0]
        ${pnColOnDaTable}                 Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][1]
        ${qtyColOnDaTable}                Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][2]
        ${revColOnDaTable}                Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][3]
        ${costColOnDaTable}               Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][4]

        ${lastRow}                  Evaluate    ${numOfRowsOfDataTable}-1
        IF    '${rowIndexOnDataTable}' < '${lastRow}'
             ${nextRowIndexOnDaTable}    Evaluate    ${rowIndexOnDataTable}+1
             ${nextOEMGroupColOnDaTable}         Set Variable   ${dataTableFromSSRevenueCostDump}[${nextRowIndexOnDaTable}][0]
             ${nextPNColOnDaTable}               Set Variable   ${dataTableFromSSRevenueCostDump}[${nextRowIndexOnDaTable}][1]
        END


        ${sumQty}      Evaluate    ${sumQty}+${qtyColOnDaTable}
        ${sumRev}      Evaluate    ${sumRev}+${revColOnDaTable}
        ${sumCost}     Evaluate    ${sumCost}+${costColOnDaTable}

        IF    '${rowIndexOnDataTable}' < '${lastRow}'
             IF    '${oemGroupColOnDaTable}' == '${nextOEMGroupColOnDaTable}' and '${pnColOnDaTable}' == '${nextPNColOnDaTable}'
                    Continue For Loop
             END
        END
        ${sumRev}   Evaluate  "%.2f" % ${sumRev}
        ${sumRev}   Evaluate  "%.2f" % ${sumRev}

        ${rowOnTable}   Create List
        ...             ${oemGroupColOnDaTable}
        ...             ${pnColOnDaTable}
        ...             ${sumQty}
        ...             ${sumRev}
        ...             ${sumCost}
        Append To List    ${table}   ${rowOnTable}
        ${sumQty}    Set Variable    0
        ${sumRev}    Set Variable    0
        ${sumCost}   Set Variable    0

    END

    [Return]    ${table}

Get Data From The SS Revenue Cost Dump For Every Quarter
    [Arguments]     ${ssRevenueCostDumpFilePath}    ${type}    ${year}     ${quarter}
    @{table}    Create List

    File Should Exist      ${ssRevenueCostDumpFilePath}
    Open Excel Document    ${ssRevenueCostDumpFilePath}    SSRevenueCostDump
    ${numOfRowsOnSS}       Get Number Of Rows In Excel    ${ssRevenueCostDumpFilePath}

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
                 IF    '${qtyColOnSS}' == '0' and '${revColOnSS}' == '0' and '${costColOnSS}' == '0'
                      Continue For Loop
                 END
                 ${rowOnTable}   Create List
                 ...             ${oemGroupColOnSS}
                 ...             ${pnColOnSS}
                 ...             ${qtyColOnSS}
                 ...             ${revColOnSS}
                 ...             ${costColOnSS}
                 Append To List    ${table}   ${rowOnTable}
            END
        END
    END
    Close All Excel Documents
    [Return]    ${table}

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
    @{sourceTable}       Create List
    @{listOfItemsChecked}      Create List
    ${type}     Set Variable    R
    ${year}     Set Variable    2024
    ${quarter}  Set Variable    1

    ${reportTable}  Create Table For Margin Report    reportFilePath=${reportFilePath}    type=${type}  year=${year}   quarter=${quarter}
    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
#     Write The Report Table To Excel    ${reportTable}
    ${sourceTable}  Create Table From The SS Revenue Cost Dump For Margin Report Source    ssRevenueCostDumpFilePath=${ssRevenueCostDumpFilePath}     type=${type}     year=${year}   quarter=${quarter}
    ${numOfRowsOnSourceTable}   Get Length    ${sourceTable}
    Log To Console    numOfRowsOnReportTable: ${numOfRowsOnReportTable}; numOfRowsOnSourceTable: ${numOfRowsOnSourceTable}
    Open Excel Document    ${RESULT_FILE_PATH}    MarginReportResult
    FOR    ${rowIndexOnSourceTable}    IN RANGE    0    ${numOfRowsOnSourceTable}
        ${oemGroupColOnSourceTable}    Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][0]
        ${pnColOnSourceTable}          Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][1]
        ${qtyColOnSourceTable}         Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][2]
        ${revColOnSourceTable}         Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][3]
        ${costColOnSourceTable}        Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][4]

        ${countTemp}    Set Variable    0

        FOR    ${rowIndexOnReportTable}    IN RANGE    0    ${numOfRowsOnReportTable}
            ${oemGroupColOnReportTable}    Set Variable   ${reportTable}[${rowIndexOnReportTable}][0]
            ${pnColOnReportTable}          Set Variable   ${reportTable}[${rowIndexOnReportTable}][1]
            ${qtyColOnReportTable}         Set Variable   ${reportTable}[${rowIndexOnReportTable}][2]
            ${revColOnReportTable}         Set Variable   ${reportTable}[${rowIndexOnReportTable}][3]
            ${costColOnReportTable}        Set Variable   ${reportTable}[${rowIndexOnReportTable}][4]

            IF    '${oemGroupColOnSourceTable}' == '${oemGroupColOnReportTable}' and '${pnColOnSourceTable}' == '${pnColOnReportTable}'
                 ${diffQty}    Evaluate    ${qtyColOnSourceTable}-${qtyColOnReportTable}
                 IF    '${diffQty}' >= '1'
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
                      IF    '${type}' == 'R'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
                      END
                      IF    '${type}' == 'B'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
                      END
                      IF    '${type}' == 'CF'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
                      END
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=QTY
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=6    value=${qtyColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=7    value=${qtyColOnSourceTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${diffRev}     Evaluate    ${revColOnSourceTable}-${revColOnReportTable}
                 IF    '${diffRev}' >= '1'
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
                      IF    '${type}' == 'R'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
                      END
                      IF    '${type}' == 'B'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
                      END
                      IF    '${type}' == 'CF'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
                      END
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=REVENUE
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=6    value=${revColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=7    value=${revColOnSourceTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${diffCost}    Evaluate    ${costColOnSourceTable}-${costColOnReportTable}
                 IF    '${diffCost}' >= '1'

                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}

                      IF    '${type}' == 'R'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
                      END
                      IF    '${type}' == 'B'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
                      END
                      IF    '${type}' == 'CF'
                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
                      END
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=COST
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=6    value=${costColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=7    value=${costColOnSourceTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 BREAK
            END
            ${checkedItems}     Set Variable    ${oemGroupColOnSourceTable}_${pnColOnSourceTable}
            Append To List    ${listOfItemsChecked}
            ${countTemp}    Evaluate    ${countTemp}+1
        END
        IF    '${countTemp}' == '${numOfRowsOnReportTable}'
              ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
              ${nextRow}     Evaluate    ${latestRowInResultFile}+1
              Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
              IF    '${type}' == 'R'
                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
              END
              IF    '${type}' == 'B'
                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
              END
              IF    '${type}' == 'CF'
                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
              END
              Write Excel Cell    row_num=${nextRow}    col_num=3    value=This Item is not updated in Margin Report
              Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
              Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
              Write Excel Cell    row_num=${nextRow}    col_num=6    value=${EMPTY}
              Write Excel Cell    row_num=${nextRow}    col_num=7    value=${EMPTY}
              Save Excel Document    ${RESULT_FILE_PATH}
        END
    END
    FOR    ${rowIndexOnReportTable}    IN RANGE    0    ${numOfRowsOnReportTable}
        ${oemGroupColOnReportTable}    Set Variable   ${reportTable}[${rowIndexOnReportTable}][0]
        ${pnColOnReportTable}          Set Variable   ${reportTable}[${rowIndexOnReportTable}][1]
        ${itemsInReportTable}   Set Variable    ${oemGroupColOnReportTable}_${pnColOnReportTable}
        ${countTemp}    Set Variable    0
        FOR    ${itemChecked}    IN    @{listOfItemsChecked}
            IF    '${itemsInReportTable}' == '${itemChecked}'
                 BREAK
            END
            ${countTemp}    Evaluate    ${countTemp}+1
        END
        IF    '${countTemp}' == '${numOfRowsOnReportTable}'
             ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
              ${nextRow}     Evaluate    ${latestRowInResultFile}+1
              Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
              IF    '${type}' == 'R'
                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
              END
              IF    '${type}' == 'B'
                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
              END
              IF    '${type}' == 'CF'
                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
              END
              Write Excel Cell    row_num=${nextRow}    col_num=3    value=This Item is not updated in Margin Report
              Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnReportTable}
              Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnReportTable}
              Write Excel Cell    row_num=${nextRow}    col_num=6    value=${EMPTY}
              Write Excel Cell    row_num=${nextRow}    col_num=7    value=${EMPTY}
              Save Excel Document    ${RESULT_FILE_PATH}
        END
    END
    Close All Excel Documents
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