*** Settings ***
Resource    ../../Pages/Reports/WoWChangePageV2.robot

*** Test Cases ***
Verify Prev Q Ship for the OEM East table
    [Tags]  WoWChange_0001
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table

    ${posOfColOnWoWChange}  Set Variable    2
    ${currentYear}  Get Current Year
    ${currentQuarter}   Get Current Quarter
    ${preQuarter}   Evaluate    ${currentQuarter}-1
    ${searchStr}    Set Variable    ${currentYear}.Q${preQuarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfColOnSG}  Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfColOnSG}     Evaluate    ${posOfColOnSG}+2
    Check Data For The OEM East Table    posOfColOnWoWChange=${posOfColOnWoWChange}       posOfColOnSG=${posOfColOnSG}    nameOfCol=Pre Q Ships

Verify Current Q Budget for the OEM East table
    [Tags]  WoWChange_0002
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table

    ${posOfColOnWoWChange}          Set Variable    3
    ${posOfColOnSGWeeklyActionDB}   Set Variable    3
    Check The Budget Data    posOfColOnWoWChange=${posOfColOnWoWChange}       posOfColOnSGWeeklyActionDB=${posOfColOnSGWeeklyActionDB}    nameOfCol=Current Q Budget

Verify LW Commit for the OEM East table
     [Tags]  WoWChange_0003
     [Documentation]     Verify the data of LW Commit column for the OEM East table

     Check The Commit Or Comment Data   posOfColOnWoWChange=4   nameOfCol=LW Commit   table=OEM East

Verify TW Commit for the OEM East table
    [Tags]  WoWChange_0004
    [Documentation]     Verify the data of TW Commit column for the OEM East table

    Check The Commit Or Comment Data   posOfColOnWoWChange=5   nameOfCol=TW Commit   table=OEM East

Verify Ships for the OEM East table
    [Tags]  WoWChange_0005
    [Documentation]     Verify the data of Ships column for the OEM East table

    ${posOfColOnWoWChange}  Set Variable    6
    ${currentYear}  Get Current Year
    ${currentQuarter}   Get Current Quarter
    ${searchStr}    Set Variable    ${currentYear}.Q${currentQuarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfColOnSG}  Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfColOnSG}     Evaluate    ${posOfColOnSG}+2
    Check Data For The OEM East Table    posOfColOnWoWChange=${posOfColOnWoWChange}       posOfColOnSG=${posOfColOnSG}    nameOfCol=Ships

Verify WoW of Ships for the OEM East table
    [Tags]  WoWChange_0006
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM East table

    Check The WoW Data  table=OEM East   posOfColOnWoWChange=7   nameOfCol=WoW Of Ships

Verify Backlog for the OEM East table
    [Tags]  WoWChange_0007
    [Documentation]     Verify the data of Backlog column for the OEM East table

    ${posOfColOnWoWChange}  Set Variable    8
    ${currentYear}  Get Current Year
    ${currentQuarter}   Get Current Quarter
    ${searchStr}    Set Variable    ${currentYear}.Q${currentQuarter} B
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfColOnSG}  Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfColOnSG}     Evaluate    ${posOfColOnSG}+2
    Check Data For The OEM East Table    posOfColOnWoWChange=${posOfColOnWoWChange}       posOfColOnSG=${posOfColOnSG}    nameOfCol=Backlog

Verify LOS for the OEM East table
    [Tags]  WoWChange_0008
    [Documentation]     Verify the data of LOS column for the OEM East table

    ${posOfRColOnSG}    Set Variable    0
    ${posOfBColOnSG}    Set Variable    0
    Check The LOS Data  table=OEM East  posOfColOnWoWChange=9   posOfRColOnSG=${posOfRColOnSG}   posOfBColOnSG=${posOfBColOnSG}   nameOfCol=LOS



