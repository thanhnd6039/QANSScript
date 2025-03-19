*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${iconExportDataSSToCSV}       //*[@title='Export - CSV']
${iconFilters}                 //*[@aria-label='Expand/Collapse filters']
${txtDateCreatedFrom}          //input[@id='BaseTran_DATECREATEDfrom']
${txtDateCreateTo}             //input[@id='BaseTran_DATECREATEDto']

${SSMasterOPPFilePath}         C:\\RobotFramework\\Downloads\\SS Master OPP.xlsx
${SSRCDFilePath}               C:\\RobotFramework\\Downloads\\SS Revenue Cost Dump.xlsx
${startRowOnSSRCD}                             2
${rowIndexForSearchColOnSSRCD}                 1
${posOfOEMGroupColOnSSRCD}                     2
${posOfParentClassColOnSSRCD}                  9
${posOfPNColOnSSRCD}                           11
${posOfQuarterColOnSSRCD}                      18
${startRowOnSSMasterOPP}                       2
${posOfOPPJoinIDColOnSSMasterOPP}              3
${posOfOEMGroupColOnSSMasterOPP}               6
${posOfPNColOnSSMasterOPP}                     7

*** Keywords ***
Create Table For SS Revenue Cost Dump
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}
    @{table}    Create List
    ${quarterStr}  Set Variable    Q${quarter}-${year}
    IF    '${transType}' == 'REVENUE'
         ${searchStr}   Set Variable    ${year} Q${quarter} Actual
    ELSE IF     '${transType}' == 'BACKLOG'
         ${searchStr}   Set Variable    ${year} Q${quarter} Backlog
    ELSE IF     '${transType}' == 'TOTAL BACKLOG'
         ${searchStr}   Set Variable    ${year} Q${quarter} Backlog
    ELSE IF     '${transType}' == 'CUSTOMER FORECAST'
         ${searchStr}   Set Variable    ${year} Q${quarter} Customer Forecast
    ELSE
         Fail    The TransType parameter ${transType} is invalid. Please contact with the Administrator for supporting
    END
    ${posOfValueCol}     Get Position Of Column    filePath=${SSRCDFilePath}    rowIndex=${rowIndexForSearchColOnSSRCD}    searchStr=${searchStr}
    ${listOEMGroupAndPN}    Get List OEM GROUP And PN For Every Quarter    year=${year}    quarter=${quarter}
    ${listParentClass}  Get List Parent Class
    File Should Exist      path=${SSRCDFilePath}
    Open Excel Document    filename=${SSRCDFilePath}    doc_id=SSRCD
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${SSRCDFilePath}
    
    FOR    ${oemGroupAndPN}    IN    @{listOEMGroupAndPN}
        ${value}    Set Variable    0
#        FOR    ${rowIndex}    IN RANGE    ${startRowOnSSRCD}    ${numOfRows}+1
#            ${parentClassCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfParentClassColOnSSRCD}
#            IF    '${parentClassCol}' in ${listParentClass}
#                ${quarterCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfQuarterColOnSSRCD}
#                IF    '${quarterCol}' == '${quarterStr}'
#                    ${oemGroupCol}    Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnSSRCD}
#                    ${pnCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnSSRCD}
#                    IF    '${oemGroupCol}' == '${oemGroupAndPN[0]}' and '${pnCol}' == '${oemGroupAndPN[1]}'
#                        ${valueCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
#                        ${value}   Evaluate    ${value}+${valueCol}
#                    END
#                END
#            END
#        END
        Log To Console    OEM GRoup: ${oemGroupAndPN[0]}; PN: ${oemGroupAndPN[1]}; Value: ${value}
         
    END

    Close Current Excel Document
    [Return]    ${table}

Get List OEM GROUP And PN For Every Quarter
    [Arguments]     ${year}     ${quarter}
    @{listOEMGroupAndPN}    Create List
    ${quarterStr}  Set Variable    Q${quarter}-${year}
    File Should Exist      path=${SSRCDFilePath}
    Open Excel Document    filename=${SSRCDFilePath}    doc_id=SSRCD
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${SSRCDFilePath}
    ${listParentClass}  Get List Parent Class
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

The Title Of Save Search Should Contain
    [Arguments]     ${title}
    ${titleXpath}   Set Variable     //h1[contains(text(),'${title}')]
    Wait Until Element Is Visible    ${titleXpath}      ${TIMEOUT}

Export SS Data To CSV
    Wait Until Element Is Visible    ${iconExportDataSSToCSV}   ${TIMEOUT}
    Click Element    ${iconExportDataSSToCSV}

Expand Filters On SS
    Wait Until Element Is Visible    ${iconFilters}     ${TIMEOUT}
    Click Element    ${iconFilters}

Set Date Create On SS
    [Arguments]     ${from}     ${to}
    IF    '${from}' != '${EMPTY}'
         Wait Until Element Is Visible    ${txtDateCreatedFrom}  ${TIMEOUT}
         Input Text    ${txtDateCreatedFrom}    ${from}
         Press Keys     None    TAB
    END
    IF    '${to}' != '${EMPTY}'
         Wait Until Element Is Visible    ${txtDateCreateTo}    ${TIMEOUT}
         Input Text    ${txtDateCreateTo}    ${to}
         Press Keys     None    TAB
    END





