*** Settings ***
Resource    ../CommonPage.robot
Resource    ../../Pages/NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Keywords ***
Setup Test Environment For SG Report
    [Arguments]     ${browser}
#    Setup    browser=${browser}
#    Navigate To Report    configFileName=SGConfig.json
#    Export Report To      option=Excel
#    ${SGFilePath}   Set Variable    ${OUTPUT_DIR}\\Sales Gap Report NS With SO Forecast.xlsx
#    Wait Until Created    path=${SGFilePath}    timeout=${TIMEOUT}
#    Login To NS With Account    account=PRODUCTION
#    Navigate To SS Revenue Cost Dump
    ${name}     Get Fully File Name From Given Name    givenName=MasterOpportunity    dirPath=${OUTPUT_DIR}
    Log To Console    NAME:${name}

    









    
