*** Settings ***
Library     SeleniumLibrary

*** Keywords ***
Setup
    [Arguments]     ${browser}
    Open Browser    browser=${browser}
    Maximize Browser Window

TearDown
    Close Browser
