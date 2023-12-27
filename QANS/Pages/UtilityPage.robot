*** Settings ***
Library    Collections
Library    OperatingSystem

*** Keywords ***
Get Fully File Name From Given Name
    [Arguments]     ${givenName}    ${dirPath}
    ${fullyFileName}    Set Variable    ${EMPTY}
    @{files}    List Files In Directory    ${dirPath}
    FOR    ${file}    IN    @{files}
        ${contains}     Evaluate    "${givenName}" in """${file}"""
        IF    '${contains}' == '${True}'
             ${fullyFileName}   Set Variable    ${file}
             Exit For Loop
        END
    END
    [Return]    ${fullyFileName}

Sort Table By Column
    [Arguments]     ${table}    ${colIndex}
    @{sortedRows}   Evaluate    sorted(${table}, key=lambda x: x[${colIndex}])
    ${sortedTable}  Create List
    FOR    ${row}    IN    @{sortedRows}
        Append To List    ${sortedTable}    ${row}
    END

    [Return]    ${sortedTable}


