*** Settings ***
Resource    CommonPage.robot
Library    OperatingSystem

*** Keywords ***
Get File Path From Given Name
    [Arguments]     ${givenName}
    @{files}    List Files In Directory    ${DOWNLOAD_DIR}
    FOR    ${file}    IN    @{files}
        Log To Console    File Path: ${file}

         
    END


