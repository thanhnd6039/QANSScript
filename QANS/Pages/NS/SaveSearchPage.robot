*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${iconExportDataSSToCSV}      //*[@title='Export - CSV']

*** Keywords ***
The Title Of Save Search Should Contain
    [Arguments]     ${title}
    ${titleXpath}   Set Variable     //h1[contains(text(),'${title}')]
    Wait Until Element Is Visible    ${titleXpath}      ${TIMEOUT}

Export SS Data To CSV
    Wait Until Element Is Visible    ${iconExportDataSSToCSV}   ${TIMEOUT}
    Click Element    ${iconExportDataSSToCSV}

