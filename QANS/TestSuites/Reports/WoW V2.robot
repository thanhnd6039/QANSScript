*** Settings ***
Resource       ../../Pages/Reports/WoWPageV2.robot
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

    Check BGT, Ship, Backlog On WoW Change    table=OEM East     nameOfCol=Pre Q Ships  transType=REVENUE   attribute=REV     year=${currentYear}     quarter=${preQuarter}