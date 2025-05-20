*** Settings ***
Resource    ../CommonPage.robot
Resource    ../../Pages/NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Keywords ***
Setup Test Environment For SG Report
    [Arguments]     ${browser}
    Remove All Files In Specified Directory    dirPath=${OUTPUT_DIR}
    Create Excel File     filePath=${OUTPUT_DIR}\\SGResult.xlsx
    Wait Until Created    path=${OUTPUT_DIR}\\SGResult.xlsx
    Setup    browser=${browser}
    Navigate To Report    configFileName=SGConfig.json
    Export Report To      option=Excel
    ${SGFilePath}   Set Variable    ${OUTPUT_DIR}\\Sales Gap Report NS With SO Forecast.xlsx
    Wait Until Created    path=${SGFilePath}    timeout=${TIMEOUT}
    Login To NS With Account    account=PRODUCTION
    Navigate To SS Revenue Cost Dump
    Export SS To CSV
    Sleep    120s
    ${fullyFileNameOfSSRCD}     Get Fully File Name From Given Name    givenName=RevenueCostDump    dirPath=${OUTPUT_DIR}
    Convert Csv To Xlsx    csvFilePath=${OUTPUT_DIR}\\${fullyFileNameOfSSRCD}    xlsxFilePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    @{emptyTable}   Create List
    @{listNameOfColsForHeader}   Create List
     Append To List    ${listNameOfColsForHeader}  QUARTER
     Append To List    ${listNameOfColsForHeader}  TRANS TYPE
     Append To List    ${listNameOfColsForHeader}  OEM GROUP
     Append To List    ${listNameOfColsForHeader}  PN
     Append To List    ${listNameOfColsForHeader}  ON SG
     Append To List    ${listNameOfColsForHeader}  ON SS RCD
     Write Table To Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx    listNameOfCols=${listNameOfColsForHeader}    table=@{emptyTable}  hasHeader=${True}
     Close Browser

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
    Log To Console    totalValueOnSG:${totalValueOnSG};totalValueOnSSRCD:${totalValueOnSSRCD}
    ${diff}     Evaluate    ${totalValueOnSG}-${totalValueOnSSRCD}
    IF    ${diff} > 1
         FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
            ${oemGroupColOnSSRCD}       Set Variable    ${rowOnSSRCD[0]}
            ${oemGroupColOnSSRCD}       Convert To Upper Case    ${oemGroupColOnSSRCD}
            ${pnColOnSSRCD}             Set Variable    ${rowOnSSRCD[1]}
            ${valueColOnSSRCD}          Set Variable    ${rowOnSSRCD[2]}
            ${isFoundOEMGroupAndPN}     Set Variable    ${False}
            FOR    ${rowOnSG}    IN    @{tableSG}
                ${oemGroupColOnSG}      Set Variable    ${rowOnSG[0]}
                ${oemGroupColOnSG}      Convert To Upper Case    ${oemGroupColOnSG}
                ${pnColOnSG}            Set Variable    ${rowOnSG[2]}
                ${valueColOnSG}         Set Variable    ${rowOnSG[3]}
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
            ${oemGroupColOnSG}      Set Variable    ${rowOnSG[0]}
            ${oemGroupColOnSG}      Convert To Upper Case    ${oemGroupColOnSG}
            ${pnColOnSG}            Set Variable    ${rowOnSG[2]}
            ${valueColOnSG}         Set Variable    ${rowOnSG[3]}
            ${isFoundOEMGroupAndPN}     Set Variable    ${False}
            FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
                ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
                ${oemGroupColOnSSRCD}     Convert To Upper Case    ${oemGroupColOnSSRCD}
                ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
                IF    '${oemGroupColOnsg}' == '${oemGroupColOnSSRCD}' and '${pnColOnsg}' == '${pnColOnSSRCD}'
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
             Append To List    ${listNameOfColsForHeader}  ON SS RCD
             Write Table To Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}  hasHeader=${False}
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
    ${SGFilePath}   Set Variable    ${OUTPUT_DIR}\\Sales Gap Report NS With SO Forecast.xlsx

    ${configFileObject}           Load Json From File    file_name=${CONFIG_DIR}\\SGConfig.json
    ${rowIndexForSearchPosOfCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.rowIndexForSearchPosOfCol
    ${rowIndexForSearchPosOfCol}  Set Variable    ${rowIndexForSearchPosOfCol[0]}

    ${posOfValueCol}    Get Position Of Column    filePath=${SGFilePath}    rowIndex=${rowIndexForSearchPosOfCol}    searchStr=${searchStr}
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

    ${startRow}  Get Value From Json    json_object=${configFileObject}    json_path=$.startRow
    ${startRow}  Set Variable    ${startRow[0]}
    ${posOfOEMGroupCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfOEMGroupCol
    ${posOfOEMGroupCol}  Set Variable    ${posOfOEMGroupCol[0]}
    ${posOfMainSalesRepCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfMainSalesRepCol
    ${posOfMainSalesRepCol}  Set Variable    ${posOfMainSalesRepCol[0]}
    ${posOfPNCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfPNCol
    ${posOfPNCol}  Set Variable    ${posOfPNCol[0]}
    File Should Exist    path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${SGFilePath}
    ${oemGroupTemp}         Set Variable    ${EMPTY}
    ${mainSalesRepTemp}     Set Variable    ${EMPTY}
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${numOfRows}+1
        ${oemGroupCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupCol}
        ${mainSalesRepCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfMainSalesRepCol}
        ${pnCol}                Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNCol}
        IF    '${oemGroupCol}' != 'None'
             ${oemGroupTemp}    Set Variable    ${oemGroupCol}
             ${mainSalesRepTemp}    Set Variable    ${mainSalesRepCol}
        END
        IF    '${pnCol}' != '${EMPTY}'
            ${valueCol}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
            IF    '${valueCol}' == 'None' or '${valueCol}' == '${EMPTY}'
                Continue For Loop
            END
            ${tempValue}    Set Variable    ${valueCol}
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


    









    
