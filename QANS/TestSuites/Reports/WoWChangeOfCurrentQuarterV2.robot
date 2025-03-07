*** Settings ***
Resource    ../../Pages/Reports/WoWChangePageV2.robot

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

#    ${searchStr}                Set Variable    ${currentYear}.Q${preQuarter} R
#    ${posOfRColOnSG}            Get Position Of Column    ${SGFilePath}    3    ${searchStr}
#    ${posOfRColOnSG}            Evaluate    ${posOfRColOnSG}+2
#    ${posOfColOnWoWChange}      Set Variable    2
    Check BGT, Ship, Backlog, LOS On WoW Change    table=OEM East     nameOfCol=Pre Q Ships     year=${currentYear}     quarter=${preQuarter}
#    Check The Ship, Backlog, LOS Data    table=OEM East     nameOfCol=Pre Q Ships   posOfColOnWoWChange=${posOfColOnWoWChange}    posOfRColOnSG=${posOfRColOnSG}