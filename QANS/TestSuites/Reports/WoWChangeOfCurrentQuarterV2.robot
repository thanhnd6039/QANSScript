*** Settings ***
Resource    ../../Pages/Reports/WoWChangePageV2.robot
Library    DependencyLibrary

*** Test Cases ***
Verify Prev Q Ship for the OEM East table
    [Tags]  WoWChange_0001
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    IF    '${currentQuarter}' == '1'
         ${preQuarter}      Set Variable    4
         ${currentYear}     Evaluate    ${currentYear}-1
    ELSE
         ${preQuarter}               Evaluate        ${currentQuarter}-1
    END
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM East     nameOfCol=Pre Q Ships  transType=REVENUE   attribute=REV     year=${currentYear}     quarter=${preQuarter}

Verify Current Q Budget for the OEM East table
    [Tags]  WoWChange_0002
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM East     nameOfCol=Current Q Budget  transType=BUDGET   attribute=REV     year=${currentYear}     quarter=${currentQuarter}

Verify LW Commit for the OEM East table
     [Tags]  WoWChange_0003
     [Documentation]     Verify the data of LW Commit column for the OEM East table

     Check LW Commit, Comment On WoW Change   table=OEM East  nameOfCol=LW Commit

Verify TW Commit for the OEM East table
    [Tags]  WoWChange_0004
    [Documentation]     Verify the data of TW Commit column for the OEM East table

    Check TW Commit On WoW Change  table=OEM East  nameOfCol=TW Commit

Verify Ships for the OEM East table
    [Tags]  WoWChange_0005
    [Documentation]     Verify the data of Ships column for the OEM East table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM East     nameOfCol=Ships  transType=REVENUE   attribute=REV     year=${currentYear}     quarter=${currentQuarter}

Verify WoW of Ships for the OEM East table
    [Tags]  WoWChange_0006
    [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM East table

    Depends On Test    name=Verify Ships for the OEM East table
    Check WoW On WoW Change  table=OEM East     nameOfCol=WoW Of Ships

Verify Backlog for the OEM East table
    [Tags]  WoWChange_0007
    [Documentation]     Verify the data of Backlog column for the OEM East table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM East     nameOfCol=Backlog  transType=BACKLOG   attribute=REV     year=${currentYear}     quarter=${currentQuarter}

Verify LOS for the OEM East table
    [Tags]  WoWChange_0008
    [Documentation]     Verify the data of LOS column for the OEM East table

    Depends On Test    name=Verify Ships for the OEM East table
    Depends On Test    name=Verify Backlog for the OEM East table
    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM East     nameOfCol=LOS  transType=LOS   attribute=REV     year=${currentYear}     quarter=${currentQuarter}

Verify WoW of LOS for the OEM East table
    [Tags]  WoWChange_0009
    [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM East table

    Depends On Test    name=Verify LOS for the OEM East table
    Check WoW On WoW Change  table=OEM East     nameOfCol=WoW Of LOS

Verify GAP for the OEM East table
     [Tags]  WoWChange_0010
     [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM East table
     
     Depends On Test    name=Verify LOS for the OEM East table
     Depends On Test    name=Verify LW Commit for the OEM East table
     Check GAP On WoW Change    table=OEM East  nameOfCol=GAP

Veify Comments for the OEM East table
     [Tags]  WoWChange_0011
     [Documentation]     Verify the data of Comments column for the OEM East table

     Check LW Commit, Comment On WoW Change  table=OEM East     nameOfCol=Comments

Verify Prev Quarter Ship for the OEM West table
    [Tags]  WoWChange_0012
    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    IF    '${currentQuarter}' == '1'
         ${preQuarter}      Set Variable    4
         ${currentYear}     Evaluate    ${currentYear}-1
    ELSE
         ${preQuarter}               Evaluate        ${currentQuarter}-1
    END
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM West + Channel     nameOfCol=Pre Q Ships  transType=REVENUE   attribute=REV     year=${currentYear}     quarter=${preQuarter}

Verify Current Quarter Budget for the OEM West table
    [Tags]  WoWChange_0013
    [Documentation]     Verify the data of Current Q Budget column for the OEM East table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM West + Channel     nameOfCol=Current Q Budget  transType=BUDGET   attribute=REV     year=${currentYear}     quarter=${currentQuarter}
