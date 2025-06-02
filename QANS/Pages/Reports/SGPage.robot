*** Settings ***
Resource    ../CommonPage.robot
Resource    ../../Pages/NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${POS_OEM_GROUP_COL_ON_SG_TABLE}         0
${POS_MAIN_SALES_REP_COL_ON_SG_TABLE}    1
${POS_PN_COL_ON_SG_TABLE}                2
${POS_VALUE_COL_ON_SG_TABLE}             3

${START_ROW_ON_SG}      6
${ROW_INDEX_FOR_SEARCH_POS_COL_ON_SG}    3
${POS_OEM_GROUP_COL_ON_SG}               2
${POS_MAIN_SALES_REP_COL_ON_SG}          3
${POS_PN_COL_ON_SG}                      4

${SG_RESULT_FILE_PATH}              ${OUTPUT_DIR}\\SGResult.xlsx
${SG_FILE_PATH}                     ${OUTPUT_DIR}\\Sales Gap Report NS With SO Forecast.xlsx
${TEST_DATA_FOR_SG_FILE_PATH}       ${TEST_DATA_DIR}\\TestDataForSG.xlsx

${lstParentClass}   //button[@id='ReportViewerControl_ctl04_ctl07_ctl01']


*** Keywords ***
Setup Test Environment For SG Report
    [Arguments]     ${browser}
    Remove All Files In Specified Directory    dirPath=${OUTPUT_DIR}
    Create Excel File     filePath=${SG_RESULT_FILE_PATH}
    Wait Until Created    path=${SG_RESULT_FILE_PATH}
    @{emptyTable}   Create List
    @{listNameOfColsForHeader}   Create List
    Append To List    ${listNameOfColsForHeader}  QUARTER
    Append To List    ${listNameOfColsForHeader}  TRANS TYPE
    Append To List    ${listNameOfColsForHeader}  OEM GROUP
    Append To List    ${listNameOfColsForHeader}  PN
    Append To List    ${listNameOfColsForHeader}  ON SG
    Append To List    ${listNameOfColsForHeader}  ON NS
    Write Table To Excel    filePath=${SG_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=@{emptyTable}  hasHeader=${True}
    Setup    browser=${browser}
    Navigate To Report    configFileName=SGConfig.json
    Export Report To      option=Excel
    Wait Until Created    path=${SG_FILE_PATH}    timeout=${TIMEOUT}
    Login To NS With Account    account=PRODUCTION
    Navigate To SS Revenue Cost Dump
    Export SS To CSV
    Sleep    180s
    ${fullyFileNameOfSSRCD}     Get Fully File Name From Given Name    givenName=RevenueCostDump    dirPath=${OUTPUT_DIR}
    Convert Csv To Xlsx    csvFilePath=${OUTPUT_DIR}\\${fullyFileNameOfSSRCD}    xlsxFilePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
     Navigate To SS Approved Sales Forecast
     Expand Filters On SS
     ${currentYear}     Get Current Year
     Set Year On SS Approved Sales Forecast    year=${currentYear}
     Export SS To CSV
     Sleep    10s
     ${fullyFileNameOfSSApprovedSalesFC}     Get Fully File Name From Given Name    givenName=ApprovedSalesForecast    dirPath=${OUTPUT_DIR}
     Convert Csv To Xlsx    csvFilePath=${OUTPUT_DIR}\\${fullyFileNameOfSSApprovedSalesFC}    xlsxFilePath=${OUTPUT_DIR}\\SS Approved Sales Forecast.xlsx
     Sleep    5s
     Close Browser


Select Parent Class On SG Report
    [Arguments]     ${options}

    Wait Until Element Is Visible    locator=${lstParentClass}  timeout=${TIMEOUT}
    Click Element    locator=${lstParentClass}
    ${chkSelectAllXpath}     Set Variable   //label[normalize-space()='(Select All)']
    Wait Until Element Is Visible    locator=${chkSelectAllXpath}   timeout=${TIMEOUT}
    Wait Until Element Is Enabled    locator=${chkSelectAllXpath}   timeout=${TIMEOUT}
    Click Element    locator=${chkSelectAllXpath}
    Click Element    locator=${chkSelectAllXpath}
    FOR    ${option}    IN    @{options}
        ${chkOptionXpath}  Set Variable    //label[normalize-space()='${option}']
        Wait Until Element Is Visible    locator=${chkOptionXpath}  timeout=${TIMEOUT}
        Wait Until Element Is Enabled    locator=${chkOptionXpath}  timeout=${TIMEOUT}
        Click Element    locator=${chkOptionXpath}
    END
    Click Element    locator=${lstParentClass}


Comparing Data For Every PN Between SG And SS Approved SF
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}   ${nameOfColOnSSApprovedSF}
    @{tableError}   Create List
    
    ${tableSG}              Create Table For SG Report    transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    ${tableSSApprovedSF}    Create Table For SS Approved Sales Forecast    nameOfCol=${nameOfColOnSSApprovedSF}    year=${year}    quarter=${quarter}
    
    ${totalValueOnSG}              Get Total Value On SG Report    table=${tableSG}
    ${totalValueOnSSApprovedSF}    Get Total Value On SS Approved Sales Forecast    table=${tableSSApprovedSF}
    IF    '${attribute}' == 'AMOUNT'
         ${totalValueOnSG}                Evaluate  "%.2f" % ${totalValueOnSG}
         ${totalValueOnSSApprovedSF}      Evaluate  "%.2f" % ${totalValueOnSSApprovedSF}
    END
    ${diff}     Evaluate    abs(${totalValueOnSG}-${totalValueOnSSApprovedSF})
    IF    ${diff} > 1
         FOR    ${rowOnSSApprovedSF}    IN    @{tableSSApprovedSF}
            ${oemGroupColOnSSApprovedSF}       Set Variable    ${rowOnSSApprovedSF[0]}
            ${oemGroupColOnSSApprovedSF}       Convert To Upper Case    ${oemGroupColOnSSApprovedSF}
            ${pnColOnSSApprovedSF}             Set Variable    ${rowOnSSApprovedSF[1]}
            ${valueColOnSSApprovedSF}          Set Variable    ${rowOnSSApprovedSF[2]}
            ${isFoundOEMGroupAndPN}     Set Variable    ${False}
            FOR    ${rowOnSG}    IN    @{tableSG}
                ${oemGroupColOnSG}      Set Variable    ${rowOnSG[${POS_OEM_GROUP_COL_ON_SG_TABLE}]}
                ${oemGroupColOnSG}      Convert To Upper Case    ${oemGroupColOnSG}
                ${pnColOnSG}            Set Variable    ${rowOnSG[${POS_PN_COL_ON_SG_TABLE}]}
                ${valueColOnSG}         Set Variable    ${rowOnSG[${POS_VALUE_COL_ON_SG_TABLE}]}
                IF    '${oemGroupColOnSSApprovedSF}' == '${oemGroupColOnSG}' and '${pnColOnSSApprovedSF}' == '${pnColOnSG}'
                    ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                    IF    '${attribute}' == 'AMOUNT'
                         ${valueColOnSSApprovedSF}      Evaluate  "%.2f" % ${valueColOnSSApprovedSF}
                         ${valueColOnSG}                Evaluate  "%.2f" % ${valueColOnSG}
                    END
                    IF    ${valueColOnSSApprovedSF} != ${valueColOnSG}
                        @{rowOnTableError}   Create List
                        Append To List    ${rowOnTableError}    Q${quarter}-${year}
                        Append To List    ${rowOnTableError}    ${transType}-${attribute}
                        Append To List    ${rowOnTableError}    ${oemGroupColOnSSApprovedSF}
                        Append To List    ${rowOnTableError}    ${pnColOnSSApprovedSF}
                        Append To List    ${rowOnTableError}    ${valueColOnSG}
                        Append To List    ${rowOnTableError}    ${valueColOnSSApprovedSF}
                        Append To List    ${tableError}     ${rowOnTableError}
                    END
                    BREAK
                END
            END
            IF    '${isFoundOEMGroupAndPN}' == '${False}'
                @{rowOnTableError}   Create List
                Append To List    ${rowOnTableError}    Q${quarter}-${year}
                Append To List    ${rowOnTableError}    ${transType}-${attribute}
                Append To List    ${rowOnTableError}    ${oemGroupColOnSSApprovedSF}
                Append To List    ${rowOnTableError}    ${pnColOnSSApprovedSF}
                Append To List    ${rowOnTableError}    ${EMPTY}
                Append To List    ${rowOnTableError}    ${valueColOnSSApprovedSF}
                Append To List    ${tableError}     ${rowOnTableError}
            END
         END
         FOR    ${rowOnSG}    IN    @{tableSG}
            ${oemGroupColOnSG}      Set Variable    ${rowOnSG[${POS_OEM_GROUP_COL_ON_SG_TABLE}]}
            ${oemGroupColOnSG}      Convert To Upper Case    ${oemGroupColOnSG}
            ${pnColOnSG}            Set Variable    ${rowOnSG[${POS_PN_COL_ON_SG_TABLE}]}
            ${valueColOnSG}         Set Variable    ${rowOnSG[${POS_VALUE_COL_ON_SG_TABLE}]}
            ${isFoundOEMGroupAndPN}     Set Variable    ${False}
            FOR    ${rowOnSSApprovedSF}    IN    @{tableSSApprovedSF}
                ${oemGroupColOnSSApprovedSF}     Set Variable    ${rowOnSSApprovedSF[0]}
                ${oemGroupColOnSSApprovedSF}     Convert To Upper Case    ${oemGroupColOnSSApprovedSF}
                ${pnColOnSSApprovedSF}           Set Variable    ${rowOnSSApprovedSF[1]}
                IF    '${oemGroupColOnSG}' == '${oemGroupColOnSSApprovedSF}' and '${pnColOnSG}' == '${pnColOnSSApprovedSF}'
                    ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                    BREAK
                END
            END
            IF    '${isFoundOEMGroupAndPN}' == '${False}'
                @{rowOnTableError}   Create List
                Append To List    ${rowOnTableError}    Q${quarter}-${year}
                Append To List    ${rowOnTableError}    ${transType}-${attribute}
                Append To List    ${rowOnTableError}    ${oemGroupColOnSG}
                Append To List    ${rowOnTableError}    ${pnColOnSG}
                Append To List    ${rowOnTableError}    ${valueColOnSG}
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
             Append To List    ${listNameOfColsForHeader}  ON SG
             Append To List    ${listNameOfColsForHeader}  ON NS
             Write Table To Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}  hasHeader=${False}
         END
    END

Comparing Data For Every PN Between SG And SS RCD
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}   ${nameOfColOnSSRCD}
    @{tableError}   Create List

    ${tableSG}              Create Table For SG Report    transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    ${tableSSRCD}           Create Table For SS Revenue Cost Dump    nameOfCol=${nameOfColOnSSRCD}    year=${year}    quarter=${quarter}
    ${totalValueOnSG}       Get Total Value On SG Report    table=${tableSG}
    ${totalValueOnSSRCD}    Get Total Value On SS Revenue Cost Dump    table=${tableSSRCD}
    IF    '${attribute}' == 'AMOUNT'
         ${totalValueOnSG}         Evaluate  "%.2f" % ${totalValueOnSG}
         ${totalValueOnSSRCD}      Evaluate  "%.2f" % ${totalValueOnSSRCD}
    END
    
    ${diff}     Evaluate    abs(${totalValueOnSG}-${totalValueOnSSRCD})
    IF    ${diff} > 1
         FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
            ${oemGroupColOnSSRCD}       Set Variable    ${rowOnSSRCD[0]}
            ${oemGroupColOnSSRCD}       Convert To Upper Case    ${oemGroupColOnSSRCD}
            ${pnColOnSSRCD}             Set Variable    ${rowOnSSRCD[1]}
            ${valueColOnSSRCD}          Set Variable    ${rowOnSSRCD[2]}
            ${isFoundOEMGroupAndPN}     Set Variable    ${False}
            FOR    ${rowOnSG}    IN    @{tableSG}
                ${oemGroupColOnSG}      Set Variable    ${rowOnSG[${POS_OEM_GROUP_COL_ON_SG_TABLE}]}
                ${oemGroupColOnSG}      Convert To Upper Case    ${oemGroupColOnSG}
                ${pnColOnSG}            Set Variable    ${rowOnSG[${POS_PN_COL_ON_SG_TABLE}]}
                ${valueColOnSG}         Set Variable    ${rowOnSG[${POS_VALUE_COL_ON_SG_TABLE}]}
                IF    '${oemGroupColOnSSRCD}' == '${oemGroupColOnSG}' and '${pnColOnSSRCD}' == '${pnColOnSG}'
                    ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                    IF    '${attribute}' == 'AMOUNT'
                         ${valueColOnSSRCD}      Evaluate  "%.2f" % ${valueColOnSSRCD}
                         ${valueColOnSG}         Evaluate  "%.2f" % ${valueColOnSG}
                    END
                    IF    ${valueColOnSSRCD} != ${valueColOnSG}
                        @{rowOnTableError}   Create List
                        Append To List    ${rowOnTableError}    Q${quarter}-${year}
                        Append To List    ${rowOnTableError}    ${transType}-${attribute}
                        Append To List    ${rowOnTableError}    ${oemGroupColOnSSRCD}
                        Append To List    ${rowOnTableError}    ${pnColOnSSRCD}
                        Append To List    ${rowOnTableError}    ${valueColOnSG}
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
         FOR    ${rowOnSG}    IN    @{tableSG}
            ${oemGroupColOnSG}      Set Variable    ${rowOnSG[${POS_OEM_GROUP_COL_ON_SG_TABLE}]}
            ${oemGroupColOnSG}      Convert To Upper Case    ${oemGroupColOnSG}
            ${pnColOnSG}            Set Variable    ${rowOnSG[${POS_PN_COL_ON_SG_TABLE}]}
            ${valueColOnSG}         Set Variable    ${rowOnSG[${POS_VALUE_COL_ON_SG_TABLE}]}
            ${isFoundOEMGroupAndPN}     Set Variable    ${False}
            FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
                ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
                ${oemGroupColOnSSRCD}     Convert To Upper Case    ${oemGroupColOnSSRCD}
                ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
                IF    '${oemGroupColOnSG}' == '${oemGroupColOnSSRCD}' and '${pnColOnSG}' == '${pnColOnSSRCD}'
                    ${isFoundOEMGroupAndPN}     Set Variable    ${True}
                    BREAK
                END
            END
            IF    '${isFoundOEMGroupAndPN}' == '${False}'
                @{rowOnTableError}   Create List
                Append To List    ${rowOnTableError}    Q${quarter}-${year}
                Append To List    ${rowOnTableError}    ${transType}-${attribute}
                Append To List    ${rowOnTableError}    ${oemGroupColOnSG}
                Append To List    ${rowOnTableError}    ${pnColOnSG}
                Append To List    ${rowOnTableError}    ${valueColOnSG}
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
             Append To List    ${listNameOfColsForHeader}  ON SG
             Append To List    ${listNameOfColsForHeader}  ON NS
             Write Table To Excel    filePath=${SG_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}  hasHeader=${False}
         END
    END

Get Total Value On SG Report
    [Arguments]     ${table}
    ${totalValue}   Set Variable    0

    FOR    ${rowOnTable}    IN    @{table}
        ${valueCol}     Set Variable    ${rowOnTable[3]}
        ${totalValue}   Evaluate    ${totalValue}+${valueCol} 
    END

    [Return]    ${totalValue}
    
Create Table For SG Report
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}
    @{table}        Create List
    ${searchStr}    Set Variable    ${EMPTY}
    IF    '${transType}' == 'REVENUE'
         ${searchStr}   Set Variable    ${year}.Q${quarter} R
    ELSE IF     '${transType}' == 'BACKLOG'
         ${searchStr}   Set Variable    ${year}.Q${quarter} B
    ELSE IF     '${transType}' == 'BACKLOG FORECAST'
         ${searchStr}   Set Variable    ${year}.Q${quarter} BF
    ELSE IF     '${transType}' == 'CUSTOMER FORECAST'
         ${searchStr}   Set Variable    ${year}.Q${quarter} CF
    ELSE IF     '${transType}' == 'BUDGET'
         ${searchStr}   Set Variable    ${year}.Q${quarter} BGT
    ELSE
         Fail    The TransType parameter ${transType} is invalid.
    END

    ${posOfValueCol}    Get Position Of Column    filePath=${SG_FILE_PATH}    rowIndex=${ROW_INDEX_FOR_SEARCH_POS_COL_ON_SG}    searchStr=${searchStr}
    IF    ${posOfValueCol} == 0
         Fail   Not found the position of ${searchStr} column
    END
    IF    '${attribute}' == 'QTY'
        ${posOfValueCol}     Evaluate    ${posOfValueCol}+0
    ELSE IF   '${attribute}' == 'AMOUNT'
        ${posOfValueCol}     Evaluate    ${posOfValueCol}+2
    ELSE
        Fail    The Attribute parameter ${attribute} in invalid
    END

    File Should Exist    path=${SG_FILE_PATH}
    Open Excel Document    filename=${SG_FILE_PATH}    doc_id=SG
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${SG_FILE_PATH}
    ${oemGroupTemp}         Set Variable    ${EMPTY}
    ${mainSalesRepTemp}     Set Variable    ${EMPTY}
    FOR    ${rowIndex}    IN RANGE    ${START_ROW_ON_SG}    ${numOfRows}+1
        ${oemGroupCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_SG}
        ${mainSalesRepCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${POS_MAIN_SALES_REP_COL_ON_SG}
        ${pnCol}                Read Excel Cell    row_num=${rowIndex}    col_num=${POS_PN_COL_ON_SG}
        IF    '${oemGroupCol}' != 'None'
             ${oemGroupTemp}        Set Variable    ${oemGroupCol}
             ${mainSalesRepTemp}    Set Variable    ${mainSalesRepCol}
        END
        IF    '${pnCol}' != '${EMPTY}'
            ${valueCol}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
            IF    '${valueCol}' == 'None' or '${valueCol}' == '${EMPTY}'
                Continue For Loop
            END
            ${tempValue}     Set Variable    ${valueCol}
            ${tempValue}     Convert To Integer    ${tempValue}
            IF    ${tempValue} == 0
                 Continue For Loop
            END
            IF    '${attribute}' == 'AMOUNT'
                 ${valueCol}      Evaluate  "%.2f" % ${valueCol}
            END
            ${rowOnTable}   Create List
            ...             ${oemGroupTemp}
            ...             ${mainSalesRepTemp}
            ...             ${pnCol}
            ...             ${valueCol}
            Append To List    ${table}   ${rowOnTable}
        END
    END

    Close Current Excel Document
    [Return]    ${table}

Get Value By OEM Group On SG Report
    [Arguments]     ${tableOnSG}    ${oemGroup}
    ${value}    Set Variable    0

    FOR    ${rowOnSG}    IN    @{tableOnSG}
        ${oemGroupCol}  Set Variable    ${rowOnSG[${POS_OEM_GROUP_COL_ON_SG_TABLE}]}
        ${valueCol}     Set Variable    ${rowOnSG[${POS_VALUE_COL_ON_SG_TABLE}]}
        IF    '${oemGroupCol}' == '${oemGroup}'
             ${value}   Evaluate    ${value}+${valueCol}
        END
    END

    [Return]    ${value}

    









    
