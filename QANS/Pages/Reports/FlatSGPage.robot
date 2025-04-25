*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${flatSGFilePath}           C:\\RobotFramework\\Downloads\\Flat Sales Gap.xlsx

${startRowOnFlatSG}                 5
${rowIndexForSearchColOnFlatSG}     4
${posOfOEMGroupColOnFlatSG}         1
${posOfPNColOnFlatSG}               2

*** Keywords ***
Comparing Data For Every PN Between Flat SG and SS RCD
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}  ${nameOfColOnSSRCD}
    ${result}   Set Variable    ${True}

    ${tableFlatSG}   Create Table For Flat SG    transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    ${tableSSRCD}    Create Table For SS Revenue Cost Dump    nameOfCol=${nameOfColOnSSRCD}    year=${year}    quarter=${quarter}

    FOR    ${rowOnFlatSG}    IN    @{tableFlatSG}
        ${oemGroupColOnFlatSG}      Set Variable    ${rowOnFlatSG[0]}
        ${pnColOnFlatSG}            Set Variable    ${rowOnFlatSG[1]}
        ${valueOnFlatSG}            Set Variable    ${rowOnFlatSG[2]}
        FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
            ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
            ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
            ${valueOnSSRCD}           Set Variable    ${rowOnSSRCD[2]}
            IF    '${oemGroupColOnFlatSG}' == '${oemGroupColOnSSRCD}' and '${pnColOnFlatSG}' == '${pnColOnSSRCD}'
                 IF    '${valueOnFlatSG}' != '${valueOnSSRCD}'

                 END
            END
        END
    END

Create Table For Flat SG
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
    ELSE IF     '${transType}' == 'SALE FORECAST'
         ${searchStr}   Set Variable    ${year}.Q${quarter} SF
    ELSE
         Fail    The TransType parameter ${transType} is invalid.
    END

    ${posOfValueCol}       Get Position Of Column    filePath=${flatSGFilePath}    rowIndex=${rowIndexForSearchColOnFlatSG}    searchStr=${searchStr}

    IF    '${attribute}' == 'REV'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}+0
    ELSE IF   '${attribute}' == 'QTY'
         ${posOfValueCol}   Evaluate    ${posOfValueCol}-2
    ELSE
        Fail    The Attribute parameter ${attribute} is invalid.
    END

    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of ${searchStr} column
    END

    File Should Exist      path=${flatSGFilePath}
    Open Excel Document    filename=${flatSGFilePath}    doc_id=FlatSG
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${flatSGFilePath}

    FOR    ${rowIndex}    IN RANGE    ${startRowOnFlatSG}    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnFlatSG}
        ${pnCol}           Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnFlatSG}
        ${valueCol}        Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        IF    '${valueCol}' == 'None'
             Continue For Loop
        END
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
    Close Current Excel Document
    [Return]    ${table}




