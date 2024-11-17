*** Settings ***
Resource    ../../Pages/Reports/WoWChangePageV2.robot

*** Test Cases ***
Verify Prev Q Ship for the OEM East table
    [Tags]  WoWChange_0001
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
    ${searchStr}    Set Variable    2024.Q3 R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfColOnSG}  Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfColOnSG}     Evaluate    ${posOfColOnSG}+2
    Check Data For The OEM East Table    wowChangeFilePath=${wowChangeFilePath}    SGFilePath=${SGFilePath}    posOfColOnWoWChange=2       posOfColOnSG=${posOfColOnSG}    nameOfCol=Pre Q Ships
