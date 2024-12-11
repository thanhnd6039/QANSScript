*** Settings ***
Resource    ../../Pages/Reports/WoWChangePageV2.robot

*** Test Cases ***
Verify Prev Q Ship for the OEM East table
    [Tags]  WoWChange_0001
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table

    ${posOfColOnWoWChange}      Set Variable    2
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    ${preQuarter}               Evaluate        ${currentQuarter}-1
    ${searchStr}                Set Variable    ${currentYear}.Q${preQuarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfRColOnSG}            Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfRColOnSG}            Evaluate    ${posOfRColOnSG}+2
    Check The Ship, Backlog, LOS Data    table=OEM East     nameOfCol=Pre Q Ships   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfRColOnSG=${posOfRColOnSG}

Verify Current Q Budget for the OEM East table
    [Tags]  WoWChange_0002
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table

    ${posOfColOnWoWChange}          Set Variable    3
    ${posOfColOnSGWeeklyActionDB}   Set Variable    3
    Check The Budget Data    table=OEM East     nameOfCol=Current Q Budget  posOfColOnWoWChange=${posOfColOnWoWChange}       posOfColOnSGWeeklyActionDB=${posOfColOnSGWeeklyActionDB}

Verify LW Commit for the OEM East table
     [Tags]  WoWChange_0003
     [Documentation]     Verify the data of LW Commit column for the OEM East table

     ${posOfColOnWoWChange}          Set Variable    4
     Check The Commit Or Comment Data   table=OEM East   nameOfCol=LW Commit    posOfColOnWoWChange=${posOfColOnWoWChange}

Verify TW Commit for the OEM East table
    [Tags]  WoWChange_0004
    [Documentation]     Verify the data of TW Commit column for the OEM East table

    ${posOfColOnWoWChange}          Set Variable    5
    Check The Commit Or Comment Data   table=OEM East   nameOfCol=TW Commit    posOfColOnWoWChange=${posOfColOnWoWChange}

Verify Ships for the OEM East table
    [Tags]  WoWChange_0005
    [Documentation]     Verify the data of Ships column for the OEM East table

    ${posOfColOnWoWChange}      Set Variable    6
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfRColOnSG}            Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfRColOnSG}            Evaluate    ${posOfRColOnSG}+2
    Check The Ship, Backlog, LOS Data    table=OEM East     nameOfCol=Ships   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfRColOnSG=${posOfRColOnSG}

Verify WoW of Ships for the OEM East table
    [Tags]  WoWChange_0006
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM East table

    ${posOfColOnWoWChange}      Set Variable    7
    Check The WoW Data  table=OEM East   nameOfCol=WoW Of Ships     posOfColOnWoWChange=${posOfColOnWoWChange}

Verify Backlog for the OEM East table
    [Tags]  WoWChange_0007
    [Documentation]     Verify the data of Backlog column for the OEM East table

    ${posOfColOnWoWChange}      Set Variable    8
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} B
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfBColOnSG}            Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfBColOnSG}            Evaluate    ${posOfBColOnSG}+2
    Check The Ship, Backlog, LOS Data    table=OEM East     nameOfCol=Backlog   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfBColOnSG=${posOfBColOnSG}

Verify LOS for the OEM East table
    [Tags]  WoWChange_0008
    [Documentation]     Verify the data of LOS column for the OEM East table

    ${posOfColOnWoWChange}      Set Variable    9
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfRColOnSG}    Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfRColOnSG}    Evaluate    ${posOfRColOnSG}+2
    ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} B
    ${posOfBColOnSG}    Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfBColOnSG}    Evaluate    ${posOfBColOnSG}+2
    Check The Ship, Backlog, LOS Data    table=OEM East     nameOfCol=LOS   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfRColOnSG=${posOfRColOnSG}    posOfBColOnSG=${posOfBColOnSG}

Verify WoW of LOS for the OEM East table
    [Tags]  WoWChange_0009
    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM East table

    ${posOfColOnWoWChange}      Set Variable    10
    Check The WoW Data  table=OEM East   nameOfCol=WoW Of LOS     posOfColOnWoWChange=${posOfColOnWoWChange}

Verify GAP for the OEM East table
     [Tags]  WoWChange_0010
     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM East table

     ${posOfColOnWoWChange}      Set Variable    11
     ${currentYear}              Get Current Year
     ${currentQuarter}           Get Current Quarter
     ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} R
     ${rowIndexForSearchStr}     Convert To Number    3
     ${posOfRColOnSG}    Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
     ${posOfRColOnSG}    Evaluate    ${posOfRColOnSG}+2
     ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} B
     ${posOfBColOnSG}    Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
     ${posOfBColOnSG}    Evaluate    ${posOfBColOnSG}+2
     Check The GAP Data  table=OEM East   nameOfCol=GAP     posOfColOnWoWChange=${posOfColOnWoWChange}      posOfRColOnSG=${posOfRColOnSG}    posOfBColOnSG=${posOfBColOnSG}

Veify Comments for the OEM East table
     [Tags]  WoWChange_0011
     [Documentation]     Verify the data of Comments column for the OEM East table

     ${posOfColOnWoWChange}          Set Variable    12
     Check The Commit Or Comment Data   table=OEM East   nameOfCol=Comments    posOfColOnWoWChange=${posOfColOnWoWChange}

Verify Prev Quarter Ship for the OEM West table
    [Tags]  WoWChange_0012
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table

    ${posOfColOnWoWChange}      Set Variable    2
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    ${preQuarter}               Evaluate        ${currentQuarter}-1
    ${searchStr}                Set Variable    ${currentYear}.Q${preQuarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfRColOnSG}            Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfRColOnSG}            Evaluate    ${posOfRColOnSG}+2
    Check The Ship, Backlog, LOS Data    table=OEM West     nameOfCol=Pre Q Ships   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfRColOnSG=${posOfRColOnSG}

Verify Current Quarter Budget for the OEM West table
    [Tags]  WoWChange_0013
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table

    ${posOfColOnWoWChange}          Set Variable    3
    ${posOfColOnSGWeeklyActionDB}   Set Variable    3
    Check The Budget Data    table=OEM West     nameOfCol=Current Q Budget  posOfColOnWoWChange=${posOfColOnWoWChange}       posOfColOnSGWeeklyActionDB=${posOfColOnSGWeeklyActionDB}

Verify LW Commit for the OEM West table
    [Tags]  WoWChange_0014
    [Documentation]     Verify the data of LW Commit column for the OEM East table

    ${posOfColOnWoWChange}          Set Variable    4
    Check The Commit Or Comment Data   table=OEM West   nameOfCol=LW Commit    posOfColOnWoWChange=${posOfColOnWoWChange}

Verify TW Commit for the OEM West table
    [Tags]  WoWChange_0015
    [Documentation]     Verify the data of TW Commit column for the OEM West table

    ${posOfColOnWoWChange}          Set Variable    5
    Check The Commit Or Comment Data   table=OEM West   nameOfCol=TW Commit    posOfColOnWoWChange=${posOfColOnWoWChange}

Verify Ships for the OEM West table
    [Tags]  WoWChange_0016
    [Documentation]     Verify the data of Ships column for the OEM West table

    ${posOfColOnWoWChange}      Set Variable    6
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfRColOnSG}            Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfRColOnSG}            Evaluate    ${posOfRColOnSG}+2
    Check The Ship, Backlog, LOS Data    table=OEM West     nameOfCol=Ships   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfRColOnSG=${posOfRColOnSG}

Verify WoW of Ships for the OEM West table
    [Tags]  WoWChange_0017
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM West table

    ${posOfColOnWoWChange}      Set Variable    7
    Check The WoW Data  table=OEM West   nameOfCol=WoW Of Ships     posOfColOnWoWChange=${posOfColOnWoWChange}

Verify Backlog for the OEM West table
    [Tags]  WoWChange_0018
    [Documentation]     Verify the data of Backlog column for the OEM West table

    ${posOfColOnWoWChange}      Set Variable    8
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    ${searchStr}                Set Variable    ${currentYear}.Q${currentQuarter} B
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfBColOnSG}            Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfBColOnSG}            Evaluate    ${posOfBColOnSG}+2
    Check The Ship, Backlog, LOS Data    table=OEM West     nameOfCol=Backlog   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfBColOnSG=${posOfBColOnSG}