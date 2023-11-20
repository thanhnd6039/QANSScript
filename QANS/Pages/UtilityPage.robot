*** Settings ***
Resource    CommonPage.robot
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


