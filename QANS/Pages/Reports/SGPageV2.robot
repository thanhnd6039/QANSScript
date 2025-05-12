*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***

*** Keywords ***
Setup Test Environment For SG Report
    [Arguments]     ${browser}
    Navigate To Report    browser=${browser}    configFileName=SGConfig.json
    Export Report To      option=Excel









    
