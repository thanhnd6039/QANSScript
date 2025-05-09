*** Settings ***
Resource    ../../Pages/Reports/SGPageV2.robot

*** Test Cases ***
Testcase1
    [Setup]     Setup    chrome
    Navigate To SG Report    configFileName=SGConfig.json

