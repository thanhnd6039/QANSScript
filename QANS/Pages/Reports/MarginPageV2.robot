*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${marginFilePath}           C:\\RobotFramework\\Downloads\\Margin Reporting By OEM Part_Rebuild.xlsx
${rowIndexForSearchColOnMargin}     3
${startRowOnMargin}                 6
${posOfOEMGroupColOnMargin}         2
${posOfPNColOnMargin}               4

*** Keywords ***
Comparing Data Between Margin And SS RCD For Every PN

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
        ${totalValue}   Evaluate    ${totalValue}+${valueCol}
    END
    ${totalValue}     Convert To Integer    ${totalValue}

    Close Current Excel Document
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