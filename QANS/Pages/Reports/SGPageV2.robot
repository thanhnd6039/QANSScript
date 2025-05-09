*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${txtTitle}     //*[contains(text(),'Sales Gap Report')]

*** Keywords ***
Navigate To SG Report
    [Arguments]     ${configFileName}
    ${configFileObject}     Load Json From File    file_name=${CONFIG_DIR}\\${configFileName}
    ${url}  Get Value From Json    json_object=${configFileObject}    json_path=$.url
    ${url}  Set Variable    ${url[0]}
    Go To    ${url}
    Wait Until Element Is Visible    locator=${txtTitle}    timeout=${TIMEOUT}




    
