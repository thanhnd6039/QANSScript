*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${flatSGFilePath}           C:\\RobotFramework\\Downloads\\Flat Sales Gap.xlsx
${flatSGResultFilePath}     C:\\RobotFramework\\Results\\FlatSGResult.xlsx

${startRowOnFlatSG}                 5
${rowIndexForSearchColOnFlatSG}     4
${posOfOEMGroupColOnFlatSG}         1
${posOfPNColOnFlatSG}               2

*** Keywords ***
Comparing Data For Every PN Between Flat SG and SS RCD
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}  ${nameOfColOnSSRCD}
    @{tableError}   Create List

    ${tableFlatSG}   Create Table For Flat SG    transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    ${tableSSRCD}    Create Table For SS Revenue Cost Dump    nameOfCol=${nameOfColOnSSRCD}    year=${year}    quarter=${quarter}

    FOR    ${rowOnFlatSG}    IN    @{tableFlatSG}
        ${oemGroupColOnFlatSG}      Set Variable    ${rowOnFlatSG[0]}
        ${pnColOnFlatSG}            Set Variable    ${rowOnFlatSG[1]}
        ${valueOnFlatSG}            Set Variable    ${rowOnFlatSG[2]}
        ${isFoundOEMGroupAndPN}     Set Variable    ${False}
        FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
            ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
            ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
            ${valueOnSSRCD}           Set Variable    ${rowOnSSRCD[2]}
            ${oemGroupColOnFlatSG}      Convert To Upper Case    ${oemGroupColOnFlatSG}
            ${oemGroupColOnSSRCD}       Convert To Upper Case    ${oemGroupColOnSSRCD}
            IF    '${oemGroupColOnFlatSG}' == '${oemGroupColOnSSRCD}' and '${pnColOnFlatSG}' == '${pnColOnSSRCD}'
                 ${valueOnFlatSG}   Convert To Integer    ${valueOnFlatSG}
                 ${valueOnSSRCD}    Convert To Integer    ${valueOnSSRCD}
                 IF    '${valueOnFlatSG}' != '${valueOnSSRCD}'
                    @{rowOnTableError}   Create List
                    Append To List    ${rowOnTableError}    Q${quarter}-${year}
                    Append To List    ${rowOnTableError}    ${oemGroupColOnFlatSG}
                    Append To List    ${rowOnTableError}    ${pnColOnFlatSG}
                    Append To List    ${rowOnTableError}    ${valueOnFlatSG}
                    Append To List    ${rowOnTableError}    ${valueOnSSRCD}
                    Append To List    ${tableError}     ${rowOnTableError}
                 END
                 ${isFoundOEMGroupAndPN}    Set Variable    ${True}
                 BREAK
            END
        END
        IF    '${isFoundOEMGroupAndPN}' == '${False}'
            @{rowOnTableError}   Create List
            Append To List    ${rowOnTableError}    Q${quarter}-${year}
            Append To List    ${rowOnTableError}    ${oemGroupColOnFlatSG}
            Append To List    ${rowOnTableError}    ${pnColOnFlatSG}
            Append To List    ${rowOnTableError}    ${valueOnFlatSG}
            Append To List    ${rowOnTableError}    ${EMPTY}
            Append To List    ${tableError}     ${rowOnTableError}
        END
    END

    FOR    ${rowOnSSRCD}    IN    @{tableSSRCD}
        ${oemGroupColOnSSRCD}     Set Variable    ${rowOnSSRCD[0]}
        ${pnColOnSSRCD}           Set Variable    ${rowOnSSRCD[1]}
        ${valueOnSSRCD}           Set Variable    ${rowOnSSRCD[2]}
        ${isFoundOEMGroupAndPN}     Set Variable    ${False}
        FOR    ${rowOnFlatSG}    IN    @{tableFlatSG}
            ${oemGroupColOnFlatSG}      Set Variable    ${rowOnFlatSG[0]}
            ${pnColOnFlatSG}            Set Variable    ${rowOnFlatSG[1]}
            ${valueOnFlatSG}            Set Variable    ${rowOnFlatSG[2]}
            ${oemGroupColOnFlatSG}      Convert To Upper Case    ${oemGroupColOnFlatSG}
            ${oemGroupColOnSSRCD}       Convert To Upper Case    ${oemGroupColOnSSRCD}
            IF    '${oemGroupColOnSSRCD}' == '${oemGroupColOnFlatSG}' and '${pnColOnSSRCD}' == '${pnColOnFlatSG}'
                ${isFoundOEMGroupAndPN}     Set Variable    ${True}
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

    ${lengthTableError}  Get Length    ${tableError}
    IF    ${lengthTableError} > 0
         @{listNameOfColsForHeader}   Create List
         Append To List    ${listNameOfColsForHeader}  QUARTER
         Append To List    ${listNameOfColsForHeader}  OEM GROUP
         Append To List    ${listNameOfColsForHeader}  PN
         Append To List    ${listNameOfColsForHeader}  ON FLAT SG
         Append To List    ${listNameOfColsForHeader}  ON NS
         Write Table To Excel    filePath=${flatSGResultFilePath}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}
         Fail   The data is different between SG report and SS Revenue Cost Dump
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




