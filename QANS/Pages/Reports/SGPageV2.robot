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
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${numOfRows}+1
        ${oemGroupCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupCol}
        IF    '${oemGroupCol}' != 'None'
            ${mainSalesRepCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfMainSalesRepCol}
            ${pnCol}                Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNCol}
            ${valueCol}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
            IF    '${valueCol}' == 'None' or '${valueCol}' == '${EMPTY}'
                 Continue For Loop
            END
            ${valueCol}     Convert To Integer    ${valueCol}
            IF    ${valueCol} == 0
                 Continue For Loop
            END

            ${rowOnTable}   Create List
            ...             ${oemGroupCol}
            ...             ${mainSalesRepCol}
            ...             ${pnCol}
            ...             ${valueCol}
            Append To List    ${table}   ${rowOnTable}
        END
    END

    Close Current Excel Document
    [Return]    ${table}


    









    
