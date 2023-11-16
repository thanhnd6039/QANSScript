*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
The Title Of Save Search Should Contain
    [Arguments]     ${title}
    ${titleXpath}   Set Variable     //h1[contains(text(),'${title}')]
    Wait Until Element Is Visible    ${titleXpath}      ${TIMEOUT}

#Export Data To XLS
