*** Settings ***
Test Template   Validate Login
Library     DataDriver  file=..\\Resources\\TestData.xlsx

*** Test Cases ***
Login Test  ${username}  ${password}



*** Keywords ***
Validate Login
    [Arguments]     ${username}     ${password}
    Log To Console    USER:${username}; PASS:${password}
