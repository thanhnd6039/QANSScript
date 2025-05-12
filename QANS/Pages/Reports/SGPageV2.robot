*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${txtTitle}     //*[contains(text(),'Sales Gap Report')]
${iconExport}   //*[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink']

*** Keywords ***
Setup Test Environment For SG Report
    [Arguments]     ${browser}
    ${configFileObject}     Load Json From File    file_name=${CONFIG_DIR}\\SGConfig.json
    ${urlSG}  Get Value From Json    json_object=${configFileObject}    json_path=$.url
    ${urlSG}  Set Variable    ${urlSG[0]}
    Open Browser    url=${urlSG}    browser=${browser}
    Maximize Browser Window
    Wait Until Element Is Visible    locator=${txtTitle}    timeout=${TIMEOUT}
    Click Element    locator=${iconExport}
    Log To Console    Finishhhhhhhhhhhhhhhh







    
