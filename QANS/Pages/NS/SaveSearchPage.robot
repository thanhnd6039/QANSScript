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
    # ${SSApprovedSFConfigFilePath}   Set Variable    ${CONFIG_DIR}\\SSApprovedSFConfig.json

    
    File Should Exist    path=${approvedSFFilePath}
    Open Excel Document    filename=${approvedSFFilePath}    doc_id=ApprovedSF
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${approvedSFFilePath}
    ${posOfValueCol}     Get Position Of Column    filePath=${approvedSFFilePath}   rowIndex=${ROW_INDEX_FOR_SEARCH_POS_COL_ON_SS_APPROVED_SF}    searchStr=${nameOfCol}
    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of ${nameOfCol} column
    END

    FOR    ${rowIndex}    IN RANGE    ${START_ROW_ON_SS_APPROVED_SF}    ${numOfRows}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_SS_APPROVED_SF}
        ${pnCol}            Read Excel Cell    row_num=${rowIndex}    col_num=${POS_PN_COL_ON_SS_APPROVED_SF}
        ${yearCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_YEAR_COL_ON_SS_APPROVED_SF}
        ${quarterCol}       Read Excel Cell    row_num=${rowIndex}    col_num=${POS_QUARTER_COL_ON_SS_APPROVED_SF}
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
        IF    '${tempValue}' == '0'
             Continue For Loop
        END
        
        IF    '${oemGroup}' == 'Lockheed Martin Corporation : Lockheed Martin Missiles and Fire Control'
            ${oemGroup}    Set Variable    Lockheed Martin Missiles and Fire Control
        ELSE IF    '${oemGroup}' == 'AMETEK Programmable Power : VTI Instruments'
            ${oemGroup}    Set Variable    VTI Instruments
        ELSE IF    '${oemGroup}' == 'BAE SYSTEMS APPLIED INTELLIGENCE : BAE systems'
            ${oemGroup}    Set Variable    BAE SYSTEMS APPLIED INTELLIGENCE
        ELSE IF    '${oemGroup}' == 'Boeing : Argon ST, Inc.'
            ${oemGroup}    Set Variable    Argon ST, Inc.
        ELSE IF    '${oemGroup}' == 'INFINERA : FABRINET CO., LTD.'
            ${oemGroup}    Set Variable    INFINERA
        ELSE IF    '${oemGroup}' == 'SANMINA-SCI : SCI TECHNOLOGY, INC.'
            ${oemGroup}    Set Variable    SCI TECHNOLOGY, INC.
        ELSE IF    '${oemGroup}' == 'SM ELECTRONIC TECHNOLOGIES PVT. LTD. : GS TECHNOLOGY PTE LTD'
            ${oemGroup}    Set Variable    GS TECHNOLOGY PTE LTD
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

    ${posOfValueCol}     Get Position Of Column    filePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx    rowIndex=${ROW_INDEX_FOR_SEARCH_POS_COL_ON_SS_RCD}    searchStr=${nameOfCol}
    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of ${nameOfCol} column
    END

    File Should Exist      path=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    Open Excel Document    filename=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx    doc_id=SSRCD
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SS Revenue Cost Dump.xlsx
    ${listParentClass}  Get List Parent Class
    FOR    ${rowIndex}    IN RANGE    ${startRowOnSSRCD}    ${numOfRows}+1
        ${parentClassCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${POS_PARENT_CLASS_COL_ON_SS_RCD}
        IF    '${parentClassCol}' in ${listParentClass}
            ${quarterCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${POS_QUARTER_COL_ON_SS_RCD}
            IF    '${quarterCol}' == '${quarterStr}'
                ${oemGroupCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_SS_RCD}
                ${pnCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_PN_COL_ON_SS_RCD}
                ${valueCol}       Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
                ${tempValue}    Set Variable    ${valueCol}
                               
                IF    '${tempValue}' == '0' or '${tempValue}' == '${EMPTY}' or '${tempValue}' == 'None'
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

    FOR    ${rowIndex}    IN RANGE    ${startRowOnSSRCD}    ${numOfRows}+1
        ${parentClassCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${POS_PARENT_CLASS_COL_ON_SS_RCD}
        IF    '${parentClassCol}' in ${listParentClass}
             ${quarterCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${POS_QUARTER_COL_ON_SS_RCD}
             IF    '${quarterCol}' == '${quarterStr}'
                  ${oemGroupCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_SS_RCD}
                  ${pnCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_PN_COL_ON_SS_RCD}
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






