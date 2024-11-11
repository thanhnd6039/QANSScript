*** Settings ***
Resource    ../../Pages/Reports/WoWChangePageV2.robot

*** Test Cases ***
Verify Prev Q Ship for the OEM East table
    [Tags]  WoWChange_0001
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
#    Open Excel Document    filename=${SGFilePath}    doc_id=SG
#    ${numOfRowsOnSG}    Get Number Of Rows In Excel    ${SGFilePath}
#    Log To Console    numOfRowsOnSG:${numOfRowsOnSG}
#    ${data}     Read Excel Cell    row_num=3    col_num=7
#    Log To Console    data:${data}
#    ${currentQuarter}   Get Current Quarter
#    ${currentYear}      Get Current Year
#    ${preQuarter}   Evaluate    ${currentQuarter}-1
#    ${searchStr}    Set Variable    ${currentYear}.Q${preQuarter} R
    ${searchStr}    Set Variable    2024.Q1 R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfColOnSG}  Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    Log To Console    posOfColOnSG:${posOfColOnSG}
#    Check Data For The OEM East Table    wowChangeFilePath=${wowChangeFilePath}    SGFilePath=${SGFilePath}    posOfColOnWoWChange=2       posOfColOnSG=${posOfColOnSG}    nameOfCol=Pre Q Ships
