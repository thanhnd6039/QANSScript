*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***   
${iconExportSSToCSV}           //div[@title='Export - CSV']
${iconFilters}                 //*[@aria-label='Expand/Collapse filters']
${txtYearFilterOnSSApprovedSalesForecast}   //input[@id='CUSTRECORD_APP_SF_YEAR']

${ROW_INDEX_FOR_SEARCH_POS_COL_ON_SS_APPROVED_SF}   1
${START_ROW_ON_SS_APPROVED_SF}                      2
${POS_OEM_GROUP_COL_ON_SS_APPROVED_SF}              2
${POS_PN_COL_ON_SS_APPROVED_SF}                     3
${POS_YEAR_COL_ON_SS_APPROVED_SF}                   4
${POS_QUARTER_COL_ON_SS_APPROVED_SF}                5
${POS_OEM_GROUP_COL_ON_SS_APPROVED_SF_TABLE}        0
${POS_PN_COL_ON_SS_APPROVED_SF_TABLE}               1
${POS_VALUE_COL_ON_SS_APPROVED_SF_TABLE}            2

${ROW_INDEX_FOR_SEARCH_POS_COL_ON_SS_RCD}           1
${START_ROW_ON_SS_RCD}                              2
${POS_OEM_GROUP_COL_ON_SS_RCD}                      2
${POS_PARENT_CLASS_COL_ON_SS_RCD}                   9
${POS_PN_COL_ON_SS_RCD}                             11
${POS_QUARTER_COL_ON_SS_RCD}                        19
${POS_OEM_GROUP_COL_ON_SS_RCD_TABLE}                0
${POS_PN_COL_ON_SS_RCD_TABLE}                       1
${POS_VALUE_COL_ON_SS_RCD_TABLE}                    2

${URL_SS_APPROVED_SF}       https://4499123.app.netsuite.com/app/common/custom/custrecordentrylist.nl?rectype=548
${URL_SS_RCD}               https://4499123.app.netsuite.com/app/common/search/searchredirect.nl?id=4436

*** Keywords ***
Navigate To SS Approved Sales Forecast
    Go To    url=${URL_SS_APPROVED_SF}
    SS Should Contain Title    title=Approved Sales Forecast

Navigate To SS Revenue Cost Dump
    Run Keyword And Ignore Error    Go To    url=${URL_SS_RCD}
    SS Should Contain Title    title=Revenue Cost Dump

Get Total Value On SS Revenue Cost Dump
    [Arguments]     ${table}
    ${totalValue}   Set Variable    0

    FOR    ${rowOnTable}    IN    @{table}
        ${valueCol}     Set Variable    ${rowOnTable[2]}
        ${totalValue}   Evaluate    ${totalValue}+${valueCol}
    END

    [Return]    ${totalValue}

Get Total Value On SS Approved Sales Forecast
    [Arguments]     ${table}
    ${totalValue}   Set Variable    0

    FOR    ${rowOnTable}    IN    @{table}
        ${valueCol}     Set Variable    ${rowOnTable[2]}
        ${totalValue}   Evaluate    ${totalValue}+${valueCol}
    END

    [Return]    ${totalValue}

Create Table For SS Approved Sales Forecast
    [Arguments]     ${nameOfCol}    ${year}     ${quarter}
    @{table}    Create List
    ${approvedSFFilePath}   Set Variable    ${OUTPUT_DIR}\\SS Approved Sales Forecast.xlsx
    ${SSApprovedSFConfigFilePath}   Set Variable    ${CONFIG_DIR}\\SSApprovedSFConfig.json

    
    File Should Exist    path=${approvedSFFilePath}
    Open Excel Document    filename=${approvedSFFilePath}    doc_id=ApprovedSF
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${approvedSFFilePath}

    ${configFileObject}     Load Json From File    file_name=${SSApprovedSFConfigFilePath}
    ${startRow}  Get Value From Json    json_object=${configFileObject}    json_path=$.startRow
    ${startRow}  Set Variable    ${startRow[0]}
    ${posOfOEMGroupCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfOEMGroupCol
    ${posOfOEMGroupCol}  Set Variable    ${posOfOEMGroupCol[0]}
    ${posOfPNCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfPNCol
    ${posOfPNCol}  Set Variable    ${posOfPNCol[0]}
    ${posOfYearCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfYearCol
    ${posOfYearCol}  Set Variable    ${posOfYearCol[0]}
    ${posOfQuarterCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfQuarterCol
    ${posOfQuarterCol}  Set Variable    ${posOfQuarterCol[0]}
    ${rowIndexForSearchPosOfCol}  Get Value From Json    json_object=${configFileObject}    json_path=$.rowIndexForSearchPosOfCol
    ${rowIndexForSearchPosOfCol}  Set Variable    ${rowIndexForSearchPosOfCol[0]}

    ${posOfValueCol}     Get Position Of Column    filePath=${approvedSFFilePath}   rowIndex=${rowIndexForSearchPosOfCol}    searchStr=${nameOfCol}
    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of ${nameOfCol} column
    END

    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${numOfRows}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupCol}
        ${pnCol}            Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNCol}
        ${yearCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfYearCol}
        ${quarterCol}       Read Excel Cell    row_num=${rowIndex}    col_num=${posOfQuarterCol}
        ${valueCol}         Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        IF    '${yearCol}' == '${year}' and '${quarterCol}' == '${quarter}'
             ${tempValue}   Set Variable    ${valueCol}
             ${tempValue}   Convert To Integer    ${tempValue}
             IF    ${tempValue} == 0
                  Continue For Loop
             END
             ${rowOnTable}   Create List
             ...             ${oemGroupCol}
             ...             ${pnCol}
             ...             ${valueCol}
             Append To List    ${table}   ${rowOnTable}
        END
         
    END
    Close All Excel Documents
    [Return]    ${table}
    
Create Table For SS Revenue Cost Dump
    [Arguments]     ${nameOfCol}    ${year}     ${quarter}
    @{table}    Create List

    ${listOEMGroupAndPN}    Get List OEM GROUP And PN For Every Quarter     year=${year}    quarter=${quarter}
    ${allTransactions}      Get All Transactions On SS RCD For Every Quarter    nameOfCol=${nameOfCol}    year=${year}    quarter=${quarter}
    FOR    ${oemGroupAndPN}    IN    @{listOEMGroupAndPN}
        ${oemGroup}     Set Variable    ${oemGroupAndPN[0]}
        ${pn}           Set Variable    ${oemGroupAndPN[1]}
        ${value}        Set Variable    0
        FOR    ${transaction}    IN    @{allTransactions}
            ${oemGroupOnTransaction}      Set Variable    ${transaction[0]}
            ${pnOnTransaction}            Set Variable    ${transaction[1]}
            ${valueOnTransaction}         Set Variable    ${transaction[2]}
            IF    '${oemGroup}' == '${oemGroupOnTransaction}' and '${pn}' == '${pnOnTransaction}'
                 ${value}    Evaluate    ${value}+${valueOnTransaction}
            END        
        END
        ${tempValue}    Set Variable    ${value}
        ${tempValue}    Convert To Integer    ${tempValue}
        IF    '${tempValue}' == '0'
             Continue For Loop
        END

        ${rowOnTable}   Create List
        ...             ${oemGroup}
        ...             ${pn}
        ...             ${value}
        Append To List    ${table}   ${rowOnTable}
    END
    [Return]    ${table}
    
Get All Transactions On SS RCD For Every Quarter
    [Arguments]     ${nameOfCol}    ${year}     ${quarter}
    @{table}       Create List
    ${quarterStr}  Set Variable    Q${quarter}-${year}

    ${configFileObject}     Load Json From File    file_name=${CONFIG_DIR}\\SSRevenueCostDumpConfig.json
    ${rowIndexForSearchColOnSSRCD}  Get Value From Json    json_object=${configFileObject}    json_path=$.rowIndexForSearchPosOfCol
    ${rowIndexForSearchColOnSSRCD}  Set Variable    ${rowIndexForSearchColOnSSRCD[0]}
    ${startRowOnSSRCD}  Get Value From Json    json_object=${configFileObject}    json_path=$.startRow
    ${startRowOnSSRCD}  Set Variable    ${startRowOnSSRCD[0]}
    ${posOfParentClassColOnSSRCD}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfParentClassCol
    ${posOfParentClassColOnSSRCD}  Set Variable    ${posOfParentClassColOnSSRCD[0]}
    ${posOfQuarterColOnSSRCD}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfQuarterCol
    ${posOfQuarterColOnSSRCD}  Set Variable    ${posOfQuarterColOnSSRCD[0]}
    ${posOfOEMGroupColOnSSRCD}   Get Value From Json    json_object=${configFileObject}    json_path=$.posOfOEMGroupCol
    ${posOfOEMGroupColOnSSRCD}  Set Variable    ${posOfOEMGroupColOnSSRCD[0]}
    ${posOfPNColOnSSRCD}   Get Value From Json    json_object=${configFileObject}    json_path=$.posOfPNCol
    ${posOfPNColOnSSRCD}  Set Variable    ${posOfPNColOnSSRCD[0]}
    ${posOfValueCol}     Get Position Of Column    filePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx    rowIndex=${rowIndexForSearchColOnSSRCD}    searchStr=${nameOfCol}
    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of ${nameOfCol} column
    END

    File Should Exist      path=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    Open Excel Document    filename=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx    doc_id=SSRCD
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    ${listParentClass}  Get List Parent Class
    FOR    ${rowIndex}    IN RANGE    ${startRowOnSSRCD}    ${numOfRows}+1
        ${parentClassCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfParentClassColOnSSRCD}
        IF    '${parentClassCol}' in ${listParentClass}
            ${quarterCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfQuarterColOnSSRCD}
            IF    '${quarterCol}' == '${quarterStr}'
                ${oemGroupCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnSSRCD}
                ${pnCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnSSRCD}
                ${valueCol}       Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
                ${tempValue}    Set Variable    ${valueCol}
                ${tempValue}    Convert To Integer    ${tempValue}
                IF    '${tempValue}' == '0'
                     Continue For Loop
                END
                ${rowOnTable}   Create List
                ...             ${oemGroupCol}
                ...             ${pnCol}
                ...             ${valueCol}
                Append To List    ${table}   ${rowOnTable}
            END
        END
    END
    Close Current Excel Document

    [Return]    ${table}

Get List OEM GROUP And PN For Every Quarter
    [Arguments]     ${year}     ${quarter}
    @{listOEMGroupAndPN}    Create List
    ${quarterStr}  Set Variable    Q${quarter}-${year}

    File Should Exist      path=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    Open Excel Document    filename=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx    doc_id=SSRCD
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    ${listParentClass}  Get List Parent Class

    ${configFileObject}     Load Json From File    file_name=${CONFIG_DIR}\\SSRevenueCostDumpConfig.json
    ${startRowOnSSRCD}  Get Value From Json    json_object=${configFileObject}    json_path=$.startRow
    ${startRowOnSSRCD}  Set Variable    ${startRowOnSSRCD[0]}
    ${posOfParentClassColOnSSRCD}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfParentClassCol
    ${posOfParentClassColOnSSRCD}  Set Variable    ${posOfParentClassColOnSSRCD[0]}
    ${posOfQuarterColOnSSRCD}  Get Value From Json    json_object=${configFileObject}    json_path=$.posOfQuarterCol
    ${posOfQuarterColOnSSRCD}  Set Variable    ${posOfQuarterColOnSSRCD[0]}
    ${posOfOEMGroupColOnSSRCD}   Get Value From Json    json_object=${configFileObject}    json_path=$.posOfOEMGroupCol
    ${posOfOEMGroupColOnSSRCD}  Set Variable    ${posOfOEMGroupColOnSSRCD[0]}
    ${posOfPNColOnSSRCD}   Get Value From Json    json_object=${configFileObject}    json_path=$.posOfPNCol
    ${posOfPNColOnSSRCD}  Set Variable    ${posOfPNColOnSSRCD[0]}

    FOR    ${rowIndex}    IN RANGE    ${startRowOnSSRCD}    ${numOfRows}+1
        ${parentClassCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfParentClassColOnSSRCD}
        IF    '${parentClassCol}' in ${listParentClass}
             ${quarterCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfQuarterColOnSSRCD}
             IF    '${quarterCol}' == '${quarterStr}'
                  ${oemGroupCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnSSRCD}
                  ${pnCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnSSRCD}
                  ${rowOnTable}   Create List
                  ...             ${oemGroupCol}
                  ...             ${pnCol}
                  Append To List    ${listOEMGroupAndPN}   ${rowOnTable}
             END
        END
    END   
    ${listOEMGroupAndPN}    Remove Duplicates    ${listOEMGroupAndPN}
    Close Current Excel Document
    [Return]    ${listOEMGroupAndPN}

Get List Parent Class
    @{listParentClass}   Create List
    Append To List    ${listParentClass}    COMPONENTS
    Append To List    ${listParentClass}    MEM
    Append To List    ${listParentClass}    NI ITEMS
    Append To List    ${listParentClass}    SERVICE
    Append To List    ${listParentClass}    STORAGE
    Append To List    ${listParentClass}    DOC

    [Return]    ${listParentClass}

#Get List Of OPP JOIN ID On SS Master OPP
#    @{listOfOPPJoinID}  Create List
#
#    File Should Exist    path=${SSMasterOPPFilePath}
#    Open Excel Document    filename=${SSMasterOPPFilePath}    doc_id=SSMasterOPP
#    ${numOfRowsOnSSMasterOPP}    Get Number Of Rows In Excel    ${SSMasterOPPFilePath}
#
#    FOR    ${rowIndex}    IN RANGE    ${startRowOnSSMasterOPP}    ${numOfRowsOnSSMasterOPP}+1
#        ${oppJoinIDCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOPPJoinIDColOnSSMasterOPP}
#        Append To List    ${listOfOPPJoinID}    ${oppJoinIDCol}
#    END
#
#    [Return]    ${listOfOPPJoinID}

#Check The OPP Join ID Data Is Exist On SS Master OPP By OEM Group And PN
#    [Arguments]     ${oemGroup}     ${pn}   ${oppJoinID}
#
#    ${result}   Set Variable    ${False}
#
#    File Should Exist    path=${SSMasterOPPFilePath}
#    Open Excel Document    filename=${SSMasterOPPFilePath}    doc_id=SSMasterOPP
#    ${numOfRowsOnSSMasterOPP}    Get Number Of Rows In Excel    ${SSMasterOPPFilePath}
#
#    FOR    ${rowIndex}    IN RANGE    ${startRowOnSSMasterOPP}    ${numOfRowsOnSSMasterOPP}+1
##        ${oppJoinIDCol}     Set Variable    None
#        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnSSMasterOPP}
#        ${pnCol}            Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnSSMasterOPP}
#        ${oppJoinIDCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOPPJoinIDColOnSSMasterOPP}
##        IF    '${oemGroupCol}' == '${oemGroup}' and '${pnCol}' == '${pn}'
##             ${oppJoinIDCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOPPJoinIDColOnSSMasterOPP}
##        END
#        IF    '${oppJoinIDCol}' == '${oppJoinID}'
#             ${result}  Set Variable    ${True}
#             BREAK
#        END
#    END
#    Close Current Excel Document
#    [Return]    ${result}

SS Should Contain Title
    [Arguments]     ${title}
    ${titleXpath}   Set Variable     //h1[contains(text(),'${title}')]
    Wait Until Element Is Visible    ${titleXpath}      ${TIMEOUT}

Export SS To CSV
    Wait Until Element Is Visible    ${iconExportSSToCSV}   ${TIMEOUT}
    Click Element    ${iconExportSSToCSV}

Expand Filters On SS
    Wait Until Element Is Visible    ${iconFilters}     ${TIMEOUT}
    Click Element    ${iconFilters}

Set Year On SS Approved Sales Forecast
    [Arguments]     ${year}
    Wait Until Element Is Visible    locator=${txtYearFilterOnSSApprovedSalesForecast}      timeout=${TIMEOUT}
    Wait Until Element Is Enabled    locator=${txtYearFilterOnSSApprovedSalesForecast}      timeout=${TIMEOUT}
    Input Text    locator=${txtYearFilterOnSSApprovedSalesForecast}    text=${year}
    Press Keys     None    TAB






