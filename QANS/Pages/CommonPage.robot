*** Settings ***
Library     SeleniumLibrary
Library     JSONLibrary

*** Variables ***
${CONFIG_FILE}      C:\\RobotFramework\\Config\\Config.json
${TIMEOUT}          60s

*** Keywords ***
Setup
    [Arguments]     ${browser}
    Open Browser    browser=${browser}
    Maximize Browser Window

TearDown
    Close Browser

Wait Until Page Load Completed
    FOR    ${count}    IN RANGE    1    61
        ${stage}    Execute Javascript      return document.readyState
        Exit For Loop If    '${stage}' == 'complete'
        Sleep    1s
        IF    ${count} == 60
             Fail   Page is hang or crashed
        END
    END

