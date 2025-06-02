*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${MARGIN_FILE_PATH}                 ${OUTPUT_DIR}\\Margin Reporting By OEM Part.xlsx
${MARGIN_RESULT_FILE_PATH}          ${OUTPUT_DIR}\\MarginResult.xlsx
${TEST_DATA_FOR_MARGIN_FILE}        ${OUTPUT_DIR}\\TestDataForMarginReport.xlsx

${ROW_INDEX_FOR_SEARCH_COL_ON_MARGIN}           3
${ROW_INDEX_FOR_SEARCH_TOTAL_VALUE_ON_MARGIN}   5
${START_ROW_ON_MARGIN}                          6
${POS_OEM_GROUP_COL_ON_MARGIN}                  1
${POS_PN_COL_ON_MARGIN}                         3

*** Keywords ***
Comparing Data For Every PN Between Margin And SS RCD
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}   ${nameOfColOnSSRCD}
    @{tableError}   Create List

    ${tableMargin}      Create Table For Margin Report           transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    ${tableSSRCD}       Create Table For SS Revenue Cost Dump    nameOfCol=${nameOfColOnSSRCD}    year=${year}    quarter=${quarter}
    
    FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
        ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
        ${oemGroupColOnSSRCD}       Convert To Upper Case    ${oemGroupColOnSSRCD}
        ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
        ${valueOnSSRCD}           Set Variable    ${rowOnSSRCD[2]}
        ${valueOnSSRCD}    Convert To Integer    ${valueOnSSRCD}
        ${isFoundOEMGroupAndPN}     Set Variable    ${False}
        FOR    ${rowOnMargin}    IN    @{tableMargin}
            ${oemGroupColOnMargin}      Set Variable    ${rowOnMargin[0]}
            ${oemGroupColOnMargin}      Convert To Upper Case    ${oemGroupColOnMargin}
            ${pnColOnMargin}            Set Variable    ${rowOnMargin[1]}
            ${valueOnMargin}            Set Variable    ${rowOnMargin[2]}
            ${valueOnMargin}   Convert To Integer    ${valueOnMargin}
            IF    '${oemGroupColOnSSRCD}' == '${oemGroupColOnMargin}' and '${pnColOnSSRCD}' == '${pnColOnMargin}'
                 ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                 IF    ${valueOnSSRCD} != ${valueOnMargin}
                      @{rowOnTableError}   Create List
                      Append To List    ${rowOnTableError}    Q${quarter}-${year}
                      Append To List    ${rowOnTableError}    ${oemGroupColOnSSRCD}
                      Append To List    ${rowOnTableError}    ${pnColOnSSRCD}
                      Append To List    ${rowOnTableError}    ${valueOnMargin}
                      Append To List    ${rowOnTableError}    ${valueOnSSRCD}
                      Append To List    ${tableError}     ${rowOnTableError}
                 END
                 BREAK
            END
        END
        IF    '${isFoundOEMGroupAndPN}' == '${False}'
            @{rowOnTableError}   Create List
            Append To List    ${rowOnTableError}    Q${quarter}-${year}
            Append To List    ${rowOnTableError}    ${oemGroupColOnSSRCD}
            Append To List    ${rowOnTableError}    ${pnColOnSSRCD}
            Append To List    ${rowOnTableError}    ${EMPTY}
            Append To List    ${rowOnTableError}    ${valueOnSSRCD}
            Append To List    ${tableError}     ${rowOnTableError}
        END
    END

    FOR    ${rowOnMargin}    IN    @{tableMargin}
        ${oemGroupColOnMargin}      Set Variable    ${rowOnMargin[0]}
        ${oemGroupColOnMargin}      Convert To Upper Case    ${oemGroupColOnMargin}
        ${pnColOnMargin}            Set Variable    ${rowOnMargin[1]}
        ${valueOnMargin}            Set Variable    ${rowOnMargin[2]}
        ${valueOnMargin}            Convert To Integer    ${valueOnMargin}
        ${isFoundOEMGroupAndPN}     Set Variable    ${False}
        FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
            ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
            ${oemGroupColOnSSRCD}     Convert To Upper Case    ${oemGroupColOnSSRCD}
            ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
            ${valueOnSSRCD}           Set Variable    ${rowOnSSRCD[2]}
            ${valueOnSSRCD}           Convert To Integer    ${valueOnSSRCD}
            IF    '${oemGroupColOnMargin}' == '${oemGroupColOnSSRCD}' and '${pnColOnMargin}' == '${pnColOnSSRCD}'
                 ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                 BREAK
            END
        END
        IF    '${isFoundOEMGroupAndPN}' == '${False}'
            @{rowOnTableError}   Create List
            Append To List    ${rowOnTableError}    Q${quarter}-${year}
            Append To List    ${rowOnTableError}    ${oemGroupColOnMargin}
            Append To List    ${rowOnTableError}    ${pnColOnMargin}
            Append To List    ${rowOnTableError}    ${valueOnMargin}
            Append To List    ${rowOnTableError}    ${EMPTY}
            Append To List    ${tableError}     ${rowOnTableError}
        END
    END

    ${lengthTableError}  Get Length    ${tableError}
    IF    ${lengthTableError} > 0
         @{listNameOfColsForHeader}   Create List
         Append To List    ${listNameOfColsForHeader}  QUARTER
         Append To List    ${listNameOfColsForHeader}  OEM GROUP
         Append To List    ${listNameOfColsForHeader}  PN
         Append To List    ${listNameOfColsForHeader}  ON MARGIN
         Append To List    ${listNameOfColsForHeader}  ON NS
         Write Table To Excel    filePath=${marginResultFilePath}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}
         Fail   The data is different between Margin report and SS Revenue Cost Dump
    END

Get Total Value On Margin Report
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}
    ${totalValue}   Set Variable    0
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
    ${posOfValueCol}     Get Position Of Column    filePath=${MARGIN_FILE_PATH}    rowIndex=${ROW_INDEX_FOR_SEARCH_COL_ON_MARGIN}    searchStr=${searchStr}

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
    ${totalValue}   Read Excel Cell    row_num=${ROW_INDEX_FOR_SEARCH_TOTAL_VALUE_ON_MARGIN}    col_num=${posOfValueCol}
    Close Current Excel Document
    Log To Console    totalValue:${totalValue}

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
    ELSE IF     '${attribute}' == 'REV'
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