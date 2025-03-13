*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${flatSGFilePath}           C:\\RobotFramework\\Downloads\\Margin Reporting By OEM Part.xlsx

${startRowOnFlatSG}                 5
${rowIndexForSearchColOnFlatSG}     4
${posOfOEMGroupColOnFlatSG}         1
${posOfPNColOnFlatSG}               2

*** Keywords ***
Create Table For Flat SG Report
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}
    @{table}    Create List
    ${searchStr}    Set Variable    ${EMPTY}

    IF    '${transType}' == 'Revenue'
         ${searchStr}   Set Variable    ${year}.Q${quarter} R
    ELSE IF     '${transType}' == 'Backlog'
         ${searchStr}   Set Variable    ${year}.Q${quarter} B
    ELSE IF     '${transType}' == 'Customer Forecast'
         ${searchStr}   Set Variable    ${year}.Q${quarter} CF
    ELSE
         Fail    The TransType parameter ${transType} is invalid. Please contact with the Administrator for supporting
    END

    ${posOfValueCol}     Get Position Of Column    filePath=${flatSGFilePath}    rowIndex=${rowIndexForSearchColOnFlatSG}    searchStr=${searchStr}
    File Should Exist      path=${flatSGFilePath}
    Open Excel Document    filename=${flatSGFilePath}    doc_id=FlatSG
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${flatSGFilePath}

    FOR    ${rowIndex}    IN RANGE    ${startRowOnFlatSG}    ${numOfRows}+1
        ${oemGroupCol}                Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnFlatSG}
        ${pnCol}                      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnFlatSG}
        IF    '${transType}' == 'Backlog'
             ${backlogCol}                 Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnFlatSG}
             ${backlogForecastCol}         Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnFlatSG}
             IF    '${backlogCol}' == 'None'
                  ${backlogCol}     Set Variable    0
             END
             IF    '${backlogForecastCol}' == 'None'
                  ${backlogForecastCol}     Set Variable    0
             END
             ${valueCol}           Evaluate    ${backlogCol}+${backlogForecastCol}
        ELSE
            ${valueCol}           Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        END

    END
    Close Current Excel Document
    [Return]    ${table}




