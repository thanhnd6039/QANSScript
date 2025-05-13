*** Settings ***
Library    Collections
Library    OperatingSystem

*** Keywords ***


Sort Table By Column
    [Arguments]     ${table}    ${colIndex}
    @{sortedRows}   Evaluate    sorted(${table}, key=lambda x: x[${colIndex}])
    ${sortedTable}  Create List
    FOR    ${row}    IN    @{sortedRows}
        Append To List    ${sortedTable}    ${row}
    END

    [Return]    ${sortedTable}
    



