*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${iconExportDataSSToCSV}       //*[@title='Export - CSV']
${iconFilters}                 //*[@aria-label='Expand/Collapse filters']
${txtDateCreatedFrom}          //input[@id='BaseTran_DATECREATEDfrom']
${txtDateCreateTo}             //input[@id='BaseTran_DATECREATEDto']

*** Keywords ***
The Title Of Save Search Should Contain
    [Arguments]     ${title}
    ${titleXpath}   Set Variable     //h1[contains(text(),'${title}')]
    Wait Until Element Is Visible    ${titleXpath}      ${TIMEOUT}

Export SS Data To CSV
    Wait Until Element Is Visible    ${iconExportDataSSToCSV}   ${TIMEOUT}
    Click Element    ${iconExportDataSSToCSV}

Expand Filters On SS
    Wait Until Element Is Visible    ${iconFilters}     ${TIMEOUT}
    Click Element    ${iconFilters}

Set Date Create On SS
    [Arguments]     ${from}     ${to}
    IF    '${from}' != '${EMPTY}'
         Wait Until Element Is Visible    ${txtDateCreatedFrom}  ${TIMEOUT}
         Input Text    ${txtDateCreatedFrom}    ${from}
         Press Keys     None    TAB
    END
    IF    '${to}' != '${EMPTY}'
         Wait Until Element Is Visible    ${txtDateCreateTo}    ${TIMEOUT}
         Input Text    ${txtDateCreateTo}    ${to}
         Press Keys     None    TAB
    END



