*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/SaveSearchPage.robot
Resource    ../../Pages/NS/LoginPage.robot

*** Variables ***
${MARGIN_FILE_PATH}                      ${OUTPUT_DIR}\\Margin Reporting By OEM Part.xlsx
${MARGIN_RESULT_FILE_PATH}               ${OUTPUT_DIR}\\MarginResult.xlsx
${TEST_DATA_FOR_MARGIN_FILE_PATH}        ${TEST_DATA_DIR}\\TestDataForMargin.xlsx

${ROW_INDEX_FOR_SEARCH_COL_ON_MARGIN}           3
${ROW_INDEX_FOR_SEARCH_TOTAL_VALUE_ON_MARGIN}   5
${START_ROW_ON_MARGIN}                          6
${POS_OEM_GROUP_COL_ON_MARGIN}                  1
${POS_PN_COL_ON_MARGIN}                         3

*** Keywords ***
Setup Test Environment For Margin Report
    [Arguments]     ${browser}  
    Remove All Files In Specified Directory    dirPath=${OUTPUT_DIR}
    Create Excel File     filePath=${MARGIN_RESULT_FILE_PATH}
    Wait Until Created    path=${MARGIN_RESULT_FILE_PATH}
    @{emptyTable}   Create List
    @{listNameOfColsForHeader}   Create List
    Append To List    ${listNameOfColsForHeader}  QUARTER
    Append To List    ${listNameOfColsForHeader}  TRANS TYPE
    Append To List    ${listNameOfColsForHeader}  OEM GROUP
    Append To List    ${listNameOfColsForHeader}  PN
    Append To List    ${listNameOfColsForHeader}  ON Margin
    Append To List    ${listNameOfColsForHeader}  ON NS
    Write Table To Excel    filePath=${MARGIN_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=@{emptyTable}  hasHeader=${True}
    Setup    browser=${browser} 
    Navigate To Report    reportLink=/NetSuite+Reports/Finance/Margin+Reporting+By+OEM+Part&rs:Command=Render
    Export Report To      option=Excel
    Wait Until Created    path=${MARGIN_FILE_PATH}    timeout=${TIMEOUT}   
    Login To NS With Account    account=PRODUCTION
    Navigate To SS Revenue Cost Dump
    Export SS To CSV
    Sleep    180s
    ${fullyFileNameOfSSRCD}     Get Fully File Name From Given Name    givenName=RevenueCostDump    dirPath=${OUTPUT_DIR}
    Convert Csv To Xlsx    csvFilePath=${OUTPUT_DIR}\\${fullyFileNameOfSSRCD}    xlsxFilePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    Sleep    5s
    Close Browser

Comparing Data For Every PN Between Margin And SS RCD
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}   ${nameOfColOnSSRCD}
    @{tableError}   Create List

    ${tableMargin}      Create Table For Margin Report           transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}   
    ${tableSSRCD}       Create Table For SS Revenue Cost Dump    nameOfCol=${nameOfColOnSSRCD}    year=${year}    quarter=${quarter}

    ${totalValueOnMargin}       Get Total Value On Margin Report           table=${tableMargin}   
    ${totalValueOnSSRCD}        Get Total Value On SS Revenue Cost Dump    table=${tableSSRCD}  

    IF    '${attribute}' == 'AMOUNT' or '${attribute}' == 'COST'
         ${totalValueOnMargin}         Evaluate  "%.2f" % ${totalValueOnMargin}
         ${totalValueOnSSRCD}          Evaluate  "%.2f" % ${totalValueOnSSRCD}
    END

    ${diff}     Evaluate    abs(${totalValueOnMargin}-${totalValueOnSSRCD})

     IF    ${diff} > 1
          FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
               ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
               ${oemGroupColOnSSRCD}       Convert To Upper Case    ${oemGroupColOnSSRCD}
               ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
               ${valueColOnSSRCD}           Set Variable    ${rowOnSSRCD[2]}              
               ${isFoundOEMGroupAndPN}     Set Variable    ${False}
               FOR    ${rowOnMargin}    IN    @{tableMargin}
                    ${oemGroupColOnMargin}      Set Variable    ${rowOnMargin[0]}
                    ${oemGroupColOnMargin}      Convert To Upper Case    ${oemGroupColOnMargin}
                    ${pnColOnMargin}            Set Variable    ${rowOnMargin[1]}
                    ${valueColOnMargin}         Set Variable    ${rowOnMargin[2]}                   
                    IF    '${oemGroupColOnSSRCD}' == '${oemGroupColOnMargin}' and '${pnColOnSSRCD}' == '${pnColOnMargin}'
                         ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                         IF    '${attribute}' == 'AMOUNT' or '${attribute}' == 'COST'
                              ${valueColOnSSRCD}      Evaluate  "%.2f" % ${valueColOnSSRCD}
                              ${valueColOnMargin}     Evaluate  "%.2f" % ${valueColOnMargin}
                         END
                         IF    ${valueColOnSSRCD} != ${valueColOnMargin}
                              @{rowOnTableError}   Create List
                              Append To List    ${rowOnTableError}    Q${quarter}-${year}
                              Append To List    ${rowOnTableError}    ${transType}-${attribute}
                              Append To List    ${rowOnTableError}    ${oemGroupColOnSSRCD}
                              Append To List    ${rowOnTableError}    ${pnColOnSSRCD}
                              Append To List    ${rowOnTableError}    ${valueColOnMargin}
                              Append To List    ${rowOnTableError}    ${valueColOnSSRCD}
                              Append To List    ${tableError}     ${rowOnTableError}
                         END
                         BREAK
                    END
               END
               IF    '${isFoundOEMGroupAndPN}' == '${False}'
                      @{rowOnTableError}   Create List
                      Append To List    ${rowOnTableError}    Q${quarter}-${year}
                      Append To List    ${rowOnTableError}    ${transType}-${attribute}
                      Append To List    ${rowOnTableError}    ${oemGroupColOnSSRCD}
                      Append To List    ${rowOnTableError}    ${pnColOnSSRCD}
                      Append To List    ${rowOnTableError}    ${EMPTY}
                      Append To List    ${rowOnTableError}    ${valueColOnSSRCD}
                      Append To List    ${tableError}     ${rowOnTableError}
               END
          END

          FOR    ${rowOnMargin}    IN    @{tableMargin}
               ${oemGroupColOnMargin}      Set Variable    ${rowOnMargin[0]}
               ${oemGroupColOnMargin}      Convert To Upper Case    ${oemGroupColOnMargin}
               ${pnColOnMargin}            Set Variable    ${rowOnMargin[1]}
               ${valueColOnMargin}         Set Variable    ${rowOnMargin[2]}              
               ${isFoundOEMGroupAndPN}     Set Variable    ${False}
               FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
                    ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
                    ${oemGroupColOnSSRCD}     Convert To Upper Case    ${oemGroupColOnSSRCD}
                    ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
                    ${valueColOnSSRCD}        Set Variable    ${rowOnSSRCD[2]}                   
                    IF    '${oemGroupColOnMargin}' == '${oemGroupColOnSSRCD}' and '${pnColOnMargin}' == '${pnColOnSSRCD}'
                         ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                         BREAK
                    END
               END
               IF    '${isFoundOEMGroupAndPN}' == '${False}'
                    @{rowOnTableError}   Create List
                    Append To List    ${rowOnTableError}    Q${quarter}-${year}
                    Append To List    ${rowOnTableError}    ${transType}-${attribute}
                    Append To List    ${rowOnTableError}    ${oemGroupColOnMargin}
                    Append To List    ${rowOnTableError}    ${pnColOnMargin}
                    Append To List    ${rowOnTableError}    ${valueColOnMargin}
                    Append To List    ${rowOnTableError}    ${EMPTY}
                    Append To List    ${tableError}     ${rowOnTableError}
               END
          END

          ${lengthTableError}  Get Length    ${tableError}
          IF    ${lengthTableError} > 0
               @{listNameOfColsForHeader}   Create List
               Append To List    ${listNameOfColsForHeader}  QUARTER
               Append To List    ${listNameOfColsForHeader}  TRANS TYPE
               Append To List    ${listNameOfColsForHeader}  OEM GROUP
               Append To List    ${listNameOfColsForHeader}  PN
               Append To List    ${listNameOfColsForHeader}  ON MARGIN
               Append To List    ${listNameOfColsForHeader}  ON NS
               Write Table To Excel    filePath=${MARGIN_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}
               Fail   The data is different between Margin report and SS Revenue Cost Dump
          END
     END
    
Get Total Value On Margin Report
    [Arguments]     ${table}
    ${totalValue}   Set Variable    0

    FOR    ${rowOnTable}    IN    @{table}
        ${valueCol}     Set Variable    ${rowOnTable[2]}
        ${totalValue}   Evaluate    ${totalValue}+${valueCol} 
    END

    [Return]    ${totalValue}

Create Table For Margin Report
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}
    @{table}    Create List
    ${searchStr}    Set Variable    ${EMPTY}

    IF    '${transType}' == 'REVENUE'
         ${searchStr}   Set Variable    ${year}.Q${quarter} R
    ELSE IF     '${transType}' == 'BACKLOG'
         ${searchStr}   Set Variable    ${year}.Q${quarter} B
    ELSE IF     '${transType}' == 'BACKLOG FORECAST'
         ${searchStr}   Set Variable    ${year}.Q${quarter} BF
    ELSE IF     '${transType}' == 'CUSTOMER FORECAST'
         ${searchStr}   Set Variable    ${year}.Q${quarter} CF
    ELSE
         Fail    The TransType parameter ${transType} is invalid.
    END
    ${posOfValueCol}     Get Position Of Column    filePath=${marginFilePath}    rowIndex=${rowIndexForSearchColOnMargin}    searchStr=${searchStr}

    IF    '${attribute}' == 'QTY'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+0
    ELSE IF     '${attribute}' == 'AMOUNT'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+2
    ELSE IF     '${attribute}' == 'COST'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+3
    ELSE IF     '${attribute}' == '% MARGIN'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+4
    ELSE IF     '${attribute}' == 'AVG MM'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+5
    ELSE
        Fail    The Attribute parameter ${attribute} is invalid.
    END

    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of ${searchStr} column
    END
    File Should Exist      path=${MARGIN_FILE_PATH}
    Open Excel Document    filename=${MARGIN_FILE_PATH}    doc_id=Margin
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${MARGIN_FILE_PATH}
    ${oemGroup}     Set Variable    ${EMPTY}
    FOR    ${rowIndex}    IN RANGE    ${START_ROW_ON_MARGIN}    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_MARGIN}
        ${pnCol}           Read Excel Cell    row_num=${rowIndex}    col_num=${POS_PN_COL_ON_MARGIN}
        ${valueCol}        Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        IF    '${oemGroupCol}' != 'None'
             ${oemGroup}    Set Variable    ${oemGroupCol}
        END
        IF    '${pnCol}' == 'None' or '${pnCol}' == '${EMPTY}'
            Continue For Loop
        END
        IF    '${valueCol}' == 'None' or '${valueCol}' == '${EMPTY}'
             Continue For Loop
        END
        ${tempValue}    Set Variable    ${valueCol}
        ${tempValue}     Convert To Integer    ${tempValue}
        IF    '${tempValue}' == '0'
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