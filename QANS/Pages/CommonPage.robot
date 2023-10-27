*** Settings ***
Library     SeleniumLibrary
Library     JSONLibrary

*** Variables ***
${CONFIG_FILE}      C:\\RobotFramework\\Config\\Config.json

*** Keywords ***
Setup
    [Arguments]     ${browser}
    Open Browser    browser=${browser}
    Maximize Browser Window

TearDown
    Close Browser
