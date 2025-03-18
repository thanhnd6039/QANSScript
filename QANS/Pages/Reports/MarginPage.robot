*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${RESULT_FILE_PATH}         ${RESULT_DIR}\\MarginReport\\MarginReportResult.xlsx
${marginFilePath}           C:\\RobotFramework\\Downloads\\Margin Reporting By OEM Part.xlsx

${startRowOnMargin}                 7
${rowIndexForSearchColOnMargin}     4
#${rowIndexForSearchColOnMargin}     3
${posOfOEMGroupColOnMargin}         1
${posOfPNColOnMargin}               2

*** Keywords ***
Create Table For Margin Report
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}
    @{table}    Create List
    ${searchStr}    Set Variable    ${EMPTY}

    IF    '${transType}' == 'REVENUE'
         ${searchStr}   Set Variable    ${year} Q${quarter} Actual
#         ${searchStr}   Set Variable    ${year} Q${quarter}
    ELSE IF     '${transType}' == 'BACKLOG'
         ${searchStr}   Set Variable    ${year} Q${quarter} Backlog
    ELSE IF     '${transType}' == 'CUSTOMER FORECAST'
         ${searchStr}   Set Variable    ${year} Q${quarter} Customer Forecast
    ELSE
         Fail    The TransType parameter ${transType} is invalid. Please contact with the Administrator for supporting
    END

    ${posOfValueCol}     Get Position Of Column    filePath=${marginFilePath}    rowIndex=${rowIndexForSearchColOnMargin}    searchStr=${searchStr}
    IF    '${attribute}' == 'QTY'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+0
    ELSE IF     '${attribute}' == 'AMOUNT'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+1
    ELSE IF     '${attribute}' == 'COST'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+2
    ELSE IF     '${attribute}' == '% MARGIN'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+3
    ELSE IF     '${attribute}' == 'AVG MM'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+4
    ELSE
        Fail    The Attribute parameter ${attribute} is invalid. Please contact with the Administrator for supporting
    END

    File Should Exist      path=${marginFilePath}
    Open Excel Document    filename=${marginFilePath}    doc_id=Margin
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${marginFilePath}
    ${oemGroup}     Set Variable    ${EMPTY}
    FOR    ${rowIndex}    IN RANGE    ${startRowOnMargin}    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnMargin}
        ${pnCol}           Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnMargin}
        ${valueCol}        Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}

        IF    '${oemGroupCol}' != 'None'
             ${oemGroup}    Set Variable    ${oemGroupCol}
        END

        IF    '${pnCol}' == 'Total'
             Continue For Loop
        END

        IF    '${valueCol}' == 'None' or '${valueCol}' == '0' or '${valueCol}' == '${EMPTY}'
             Continue For Loop
        END

        ${rowOnTable}   Create List
        ...             ${oemGroup}
        ...             ${pnCol}
        ...             ${valueCol}
        Append To List    ${table}   ${rowOnTable}
    END
    Close Current Excel Document


    [Return]    ${table}

#Create Table From The SS Revenue Cost Dump For Margin Report Source
#    [Arguments]     ${ssRevenueCostDumpFilePath}   ${type}    ${year}     ${quarter}
#    @{table}    Create List
#    @{dataTableFromSSRevenueCostDump}   Create List
#
#    ${sumQty}   Set Variable    0
#    ${sumRev}   Set Variable    0
#    ${sumCost}  Set Variable    0
#
#    ${dataTableFromSSRevenueCostDump}   Get Data From The SS Revenue Cost Dump For Every Quarter    ssRevenueCostDumpFilePath=${ssRevenueCostDumpFilePath}    type=${type}    year=${year}    quarter=${quarter}
#    ${numOfRowsOfDataTable}     Get Length    ${dataTableFromSSRevenueCostDump}
#    ${lastRow}     Evaluate    ${numOfRowsOfDataTable}-1
#
#    FOR    ${rowIndexOnDataTable}    IN RANGE    0    ${numOfRowsOfDataTable}
#        ${oemGroupColOnDaTable}           Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][0]
#        ${pnColOnDaTable}                 Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][1]
#        ${qtyColOnDaTable}                Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][2]
#        ${revColOnDaTable}                Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][3]
#        ${costColOnDaTable}               Set Variable   ${dataTableFromSSRevenueCostDump}[${rowIndexOnDataTable}][4]
#
#        IF    ${rowIndexOnDataTable} < ${lastRow}
#             ${nextRowIndexOnDaTable}    Evaluate    ${rowIndexOnDataTable}+1
#             ${nextOEMGroupColOnDaTable}         Set Variable   ${dataTableFromSSRevenueCostDump}[${nextRowIndexOnDaTable}][0]
#             ${nextPNColOnDaTable}               Set Variable   ${dataTableFromSSRevenueCostDump}[${nextRowIndexOnDaTable}][1]
#        END
#
#        ${sumQty}      Evaluate    ${sumQty}+${qtyColOnDaTable}
#        ${sumRev}      Evaluate    ${sumRev}+${revColOnDaTable}
#        ${sumCost}     Evaluate    ${sumCost}+${costColOnDaTable}
#
#        IF    ${rowIndexOnDataTable} < ${lastRow}
#             IF    '${oemGroupColOnDaTable}' == '${nextOEMGroupColOnDaTable}' and '${pnColOnDaTable}' == '${nextPNColOnDaTable}'
#                    Continue For Loop
#             END
#        END
#
#        ${sumRev}   Evaluate  "%.2f" % ${sumRev}
#        ${sumCost}   Evaluate  "%.2f" % ${sumCost}
#
#        ${rowOnTable}   Create List
#        ...             ${oemGroupColOnDaTable}
#        ...             ${pnColOnDaTable}
#        ...             ${sumQty}
#        ...             ${sumRev}
#        ...             ${sumCost}
#        Append To List    ${table}   ${rowOnTable}
#        ${sumQty}    Set Variable    0
#        ${sumRev}    Set Variable    0
#        ${sumCost}   Set Variable    0
#
#    END
#
#    [Return]    ${table}
#
#Get Data From The SS Revenue Cost Dump For Every Quarter
#    [Arguments]     ${ssRevenueCostDumpFilePath}    ${type}    ${year}     ${quarter}
#    @{table}    Create List
#
#    File Should Exist      ${ssRevenueCostDumpFilePath}
#    Open Excel Document    ${ssRevenueCostDumpFilePath}    SSRevenueCostDump
#    ${numOfRowsOnSS}       Get Number Of Rows In Excel    ${ssRevenueCostDumpFilePath}
#
#    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
#        ${quarterColOnSS}          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
#        IF    '${quarterColOnSS}' == 'Q${quarter}-${year}'
#            ${parentClassColOnSS}      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
#            IF    '${parentClassColOnSS}' == 'MEM' or '${parentClassColOnSS}' == 'STORAGE' or '${parentClassColOnSS}' == 'COMPONENTS' or '${parentClassColOnSS}' == 'NI'
#                 ${oemGroupColOnSS}         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
#                 ${pnColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
#                 IF    '${type}' == 'R'
#                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
#                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=30
#                      ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=28
#                 END
#                 IF    '${type}' == 'B'
#                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=31
#                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=33
#                      ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=32
#                      IF    '${costColOnSS}' == '${EMPTY}'
#                           ${costColOnSS}   Set Variable    0
#                      END
#                 END
#                 IF    '${type}' == 'CF'
#                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=40
#                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=42
#                      ${costColOnSS}              Read Excel Cell    row_num=${rowIndexOnSS}    col_num=41
#                 END
#                 IF    ${qtyColOnSS} == 0 and ${revColOnSS} == 0 and ${costColOnSS} == 0
#                      Continue For Loop
#                 END
#                 ${rowOnTable}   Create List
#                 ...             ${oemGroupColOnSS}
#                 ...             ${pnColOnSS}
#                 ...             ${qtyColOnSS}
#                 ...             ${revColOnSS}
#                 ...             ${costColOnSS}
#                 Append To List    ${table}   ${rowOnTable}
#            END
#        END
#    END
#    Close All Excel Documents
#    ${table}    Sort Table By Column    ${table}    1
#    [Return]    ${table}
#
#Create Table For Margin Report
#    [Arguments]     ${reportFilePath}   ${type}     ${year}     ${quarter}
#    @{table}    Create List
#    ${posOfCol}         Set Variable    0
#    ${posOfQtyCol}      Set Variable    0
#    ${posOfRevCol}      Set Variable    0
#    ${posOfCostCol}     Set Variable    0
#    ${rowIndexTemp}     Set Variable    4
#    ${rowIndexTemp}     Convert To Integer    ${rowIndexTemp}
#
#    File Should Exist    ${reportFilePath}
#    Open Excel Document    ${reportFilePath}    MarginReport
#    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}
#    ${oemGroupColOnReportTemp}  Set Variable    ${EMPTY}
#
#    IF    '${type}' == 'R'
#         ${searchStr}    Set Variable    ${year} Q${quarter} Actual
#    END
#
#    IF    '${type}' == 'B'
#         ${searchStr}    Set Variable    ${year} Q${quarter} Backlog
#    END
#
#    IF    '${type}' == 'CF'
#         ${searchStr}    Set Variable    ${year} Q${quarter} Customer Forecast
#    END
#    ${posOfCol}  Get Position Of Column    ${reportFilePath}    ${rowIndexTemp}     ${searchStr}
#    ${posOfQtyCol}     Set Variable    ${posOfCol}
#    ${posOfRevCol}     Evaluate    ${posOfCol}+1
#    ${posOfCostCol}    Evaluate    ${posOfCol}+2
#
#    FOR    ${rowIndexOnReport}    IN RANGE    7    ${numOfRowsOnReport}+1
#        ${oemGroupColOnReport}      Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
#        IF    '${oemGroupColOnReport}' == 'None'
#             ${oemGroupColOnReport}     Set Variable    ${oemGroupColOnReportTemp}
#        ELSE
#             ${oemGroupColOnReportTemp}     Set Variable    ${oemGroupColOnReport}
#        END
#        ${pnColOnReport}             Read Excel Cell    row_num=${rowIndexOnReport}    col_num=2
#
#        IF    '${type}' == 'R'
#             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfQtyCol}
#             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfRevCol}
#             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfCostCol}
#        END
#        IF    '${type}' == 'B'
#             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfQtyCol}
#             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfRevCol}
#             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfCostCol}
#
#        END
#        IF    '${type}' == 'CF'
#             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfQtyCol}
#             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfRevCol}
#             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=${posOfCostCol}
#        END
#        IF    '${pnColOnReport}' == 'Total'
#             Continue For Loop
#        END
#
#        IF    '${qtyColOnReport}' == 'None'
#             ${qtyColOnReport}  Set Variable    0
#        END
#        IF    '${revColOnReport}' == 'None'
#             ${revColOnReport}  Set Variable    0
#        END
#        IF    '${costColOnReport}' == 'None'
#             ${costColOnReport}     Set Variable    0
#        END
#        IF    ${qtyColOnReport} == 0 and ${revColOnReport} == 0 and ${costColOnReport} == 0
#              Continue For Loop
#        END
#
#        ${revColOnReport}   Evaluate  "%.2f" % ${revColOnReport}
#        ${costColOnReport}   Evaluate  "%.2f" % ${costColOnReport}
#        ${rowOnTable}   Create List
#        ...             ${oemGroupColOnReport}
#        ...             ${pnColOnReport}
#        ...             ${qtyColOnReport}
#        ...             ${revColOnReport}
#        ...             ${costColOnReport}
#        Append To List    ${table}   ${rowOnTable}
#    END
#    Close All Excel Documents
#    [Return]    ${table}
#
#Compare Data Between Margin Report And SS On NS
#    [Arguments]     ${reportFilePath}   ${ssRevenueCostDumpFilePath}
#
#    ${result}   Set Variable    ${True}
#    @{reportTable}       Create List
#    @{sourceTable}       Create List
#    @{listOfOEMGRoupAndPNChecked}      Create List
#
#    ${type}     Set Variable    B
#    ${year}     Set Variable    2024
#    ${quarter}  Set Variable    1
#
#    ${reportTable}  Create Table For Margin Report    reportFilePath=${reportFilePath}    type=${type}  year=${year}   quarter=${quarter}
#    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
##    ${reportTableFilePath}     Set Variable    ${DOWNLOAD_DIR}\\Report Table.xlsx
##    Write The Report Table To Excel    ${reportTable}    ${reportTableFilePath}
#
#    ${sourceTable}  Create Table From The SS Revenue Cost Dump For Margin Report Source    ssRevenueCostDumpFilePath=${ssRevenueCostDumpFilePath}     type=${type}     year=${year}   quarter=${quarter}
#    ${numOfRowsOnSourceTable}   Get Length    ${sourceTable}
##    ${sourceTableFilePath}     Set Variable    ${DOWNLOAD_DIR}\\Source Table.xlsx
##    Write The Report Table To Excel    ${sourceTable}    ${sourceTableFilePath}
#
#    Open Excel Document    ${RESULT_FILE_PATH}    MarginReportResult
#    FOR    ${rowIndexOnSourceTable}    IN RANGE    0    ${numOfRowsOnSourceTable}
#        ${oemGroupColOnSourceTable}    Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][0]
#        ${pnColOnSourceTable}          Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][1]
#        ${qtyColOnSourceTable}         Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][2]
#        ${revColOnSourceTable}         Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][3]
#        ${costColOnSourceTable}        Set Variable   ${sourceTable}[${rowIndexOnSourceTable}][4]
#
#        ${countTemp}    Set Variable    0
#
#        FOR    ${rowIndexOnReportTable}    IN RANGE    0    ${numOfRowsOnReportTable}
#            ${oemGroupColOnReportTable}    Set Variable   ${reportTable}[${rowIndexOnReportTable}][0]
#            ${pnColOnReportTable}          Set Variable   ${reportTable}[${rowIndexOnReportTable}][1]
#            ${qtyColOnReportTable}         Set Variable   ${reportTable}[${rowIndexOnReportTable}][2]
#            ${revColOnReportTable}         Set Variable   ${reportTable}[${rowIndexOnReportTable}][3]
#            ${costColOnReportTable}        Set Variable   ${reportTable}[${rowIndexOnReportTable}][4]
#
#            IF    '${oemGroupColOnSourceTable}' == '${oemGroupColOnReportTable}' and '${pnColOnSourceTable}' == '${pnColOnReportTable}'
#                 ${diffQty}    Evaluate    ${qtyColOnSourceTable}-${qtyColOnReportTable}
#                 ${diffQty}    Evaluate     abs(${diffQty})
#                 IF    ${diffQty} >= 1
#                      ${result}   Set Variable    ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
#                      IF    '${type}' == 'R'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
#                      END
#                      IF    '${type}' == 'B'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
#                      END
#                      IF    '${type}' == 'CF'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
#                      END
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=QTY
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=6    value=${qtyColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=7    value=${qtyColOnSourceTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 ${diffRev}     Evaluate    ${revColOnSourceTable}-${revColOnReportTable}
#                 ${diffRev}     Evaluate    abs(${diffRev})
#                 IF    ${diffRev} >= 1
#                      ${result}   Set Variable    ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
#                      IF    '${type}' == 'R'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
#                      END
#                      IF    '${type}' == 'B'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
#                      END
#                      IF    '${type}' == 'CF'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
#                      END
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=REVENUE
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=6    value=${revColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=7    value=${revColOnSourceTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 ${diffCost}    Evaluate    ${costColOnSourceTable}-${costColOnReportTable}
#                 ${diffCost}    Evaluate    abs(${diffCost})
#                 IF    ${diffCost} >= 1
#                      ${result}   Set Variable    ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
#
#                      IF    '${type}' == 'R'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
#                      END
#                      IF    '${type}' == 'B'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
#                      END
#                      IF    '${type}' == 'CF'
#                           Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
#                      END
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=COST
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=6    value=${costColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=7    value=${costColOnSourceTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#                 BREAK
#            END
#            ${countTemp}    Evaluate    ${countTemp}+1
#        END
#        #Check whether the Items is in Report or not
#        IF    ${countTemp} == ${numOfRowsOnReportTable}
#              ${result}   Set Variable    ${False}
#              ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#              ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#              Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
#              IF    '${type}' == 'R'
#                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
#              END
#              IF    '${type}' == 'B'
#                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
#              END
#              IF    '${type}' == 'CF'
#                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
#              END
#              Write Excel Cell    row_num=${nextRow}    col_num=3    value=Margin Report is missing this item
#              Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSourceTable}
#              Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnSourceTable}
#              Write Excel Cell    row_num=${nextRow}    col_num=6    value=${EMPTY}
#              Write Excel Cell    row_num=${nextRow}    col_num=7    value=${EMPTY}
#              Save Excel Document    ${RESULT_FILE_PATH}
#        END
#        ${oemGroupAndPNChecked}     Set Variable    ${oemGroupColOnSourceTable}_${pnColOnSourceTable}
#        Append To List    ${listOfOEMGRoupAndPNChecked}     ${oemGroupAndPNChecked}
#    END
#    FOR    ${rowIndexOnReportTable}    IN RANGE    0    ${numOfRowsOnReportTable}
#        ${oemGroupColOnReportTable}      Set Variable    ${reportTable}[${rowIndexOnReportTable}][0]
#        ${pnColOnReportTable}            Set Variable    ${reportTable}[${rowIndexOnReportTable}][1]
#        ${qtyColOnReportTable}           Set Variable    ${reportTable}[${rowIndexOnReportTable}][2]
#        ${revColOnReportTable}           Set Variable    ${reportTable}[${rowIndexOnReportTable}][3]
#        ${costColOnReportTable}          Set Variable    ${reportTable}[${rowIndexOnReportTable}][4]
#
#        ${oemGroupAndPNInReportTable}    Set Variable    ${oemGroupColOnReportTable}_${pnColOnReportTable}
#        ${isOEMGroupAndPNChecked}        Set Variable    ${False}
#
#        FOR    ${oemGroupAndPNChecked}    IN    @{listOfOEMGRoupAndPNChecked}
#            IF    '${oemGroupAndPNInReportTable}' == '${oemGroupAndPNChecked}'
#                 ${isOEMGroupAndPNChecked}      Set Variable    ${True}
#                 BREAK
#            END
#        END
#        IF    '${isOEMGroupAndPNChecked}' == '${False}'
#             ${result}   Set Variable    ${False}
#             ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#              ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#              Write Excel Cell    row_num=${nextRow}    col_num=1    value=Q${quarter}-${year}
#              IF    '${type}' == 'R'
#                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=ACTUAL
#              END
#              IF    '${type}' == 'B'
#                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=BACKLOG
#              END
#              IF    '${type}' == 'CF'
#                   Write Excel Cell    row_num=${nextRow}    col_num=2    value=CUSTOMER FORECAST
#              END
#              Write Excel Cell    row_num=${nextRow}    col_num=3    value=The SS Revenue Cost Dump is missing this item
#              Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnReportTable}
#              Write Excel Cell    row_num=${nextRow}    col_num=5    value=${pnColOnReportTable}
#              Write Excel Cell    row_num=${nextRow}    col_num=6    value=${EMPTY}
#              Write Excel Cell    row_num=${nextRow}    col_num=7    value=${EMPTY}
#              Save Excel Document    ${RESULT_FILE_PATH}
#        END
#    END
#    Close All Excel Documents
#    IF    '${result}' == '${False}'
#         Fail   The data between the Margin Report and SS Revenue Cost Dump is different
#    END
#    [Return]    ${result}
#
# Write The Report Table To Excel
#    [Arguments]     ${reportTable}      ${outputFilePath}
#
#    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
#    Open Excel Document    ${outputFilePath}    ReportTable
#    FOR    ${rowIndexOnReportTable}    IN RANGE    0    ${numOfRowsOnReportTable}
#        ${oemGroupColOnReportTable}                   Set Variable        ${reportTable}[${rowIndexOnReportTable}][0]
#        ${pnColOnReportTable}                         Set Variable        ${reportTable}[${rowIndexOnReportTable}][1]
#        ${qtyColOnReportTable}                        Set Variable        ${reportTable}[${rowIndexOnReportTable}][2]
#        ${revColOnReportTable}                        Set Variable        ${reportTable}[${rowIndexOnReportTable}][3]
#        ${costColOnReportTable}                       Set Variable        ${reportTable}[${rowIndexOnReportTable}][4]
#        ${rowIndexTemp}    Evaluate    ${rowIndexOnReportTable}+2
#        Write Excel Cell    row_num=${rowIndexTemp}    col_num=1    value=${oemGroupColOnReportTable}
#        Write Excel Cell    row_num=${rowIndexTemp}    col_num=2    value=${pnColOnReportTable}
#        Write Excel Cell    row_num=${rowIndexTemp}    col_num=3    value=${qtyColOnReportTable}
#        Write Excel Cell    row_num=${rowIndexTemp}    col_num=4    value=${revColOnReportTable}
#        Write Excel Cell    row_num=${rowIndexTemp}    col_num=5    value=${costColOnReportTable}
#        Save Excel Document    ${outputFilePath}
#    END
#    Close All Excel Documents