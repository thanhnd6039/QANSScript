*** Settings ***
#Suite Setup     Setup Test Environment For WoW Change Report    browser=firefox
Resource    ../../Pages/Reports/WoWChangePage.robot
Library    DependencyLibrary

*** Test Cases ***
#Verify Prev Q Ship for the OEM East table
#    [Tags]  WoWChange_0001
#    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
#
#    ${currentYear}              Get Current Year
#    ${currentQuarter}           Get Current Quarter
#    ${preQuarter}   Set Variable    ${currentQuarter}
#    Check LOS On WoW Change    nameOftable=OEM East     nameOfCol=LOS

# Verify Current Q Budget for the OEM East table
#     [Tags]  WoWChange_0002
#     [Documentation]     Verify the data of Current Q Budget column for the OEM East table

#     ${currentYear}              Get Current Year
#     ${currentQuarter}           Get Current Quarter
#     IF    '${currentQuarter}' == '4'
#          ${nextQuarter}     Set Variable    1
#          ${currentYear}     Evaluate    ${currentYear}+1
#     ELSE
#          ${nextQuarter}      Evaluate    ${currentQuarter}+1
#     END

#     Check BGT, Ship, Backlog On WoW Change    nameOftable=OEM East     nameOfCol=Current Q Budget  transType=BUDGET   attribute=AMOUNT     year=${currentYear}     quarter=${nextQuarter}

Verify LW Commit for the OEM East table
    [Tags]  WoWChange_0003
    [Documentation]     Verify the data of LW Commit column for the OEM East table

    Check LW Commit, Comment On WoW Change   nameOftable=OEM East  nameOfCol=LW Commit

Verify TW Commit for the OEM East table
   [Tags]  WoWChange_0004
   [Documentation]     Verify the data of TW Commit column for the OEM East table

   Check TW Commit On WoW Change  nameOftable=OEM East  nameOfCol=TW Commit

Verify Ships for the OEM East table
   [Tags]  WoWChange_0005
   [Documentation]     Verify the data of Ships column for the OEM East table
   @{tableError}   Create List
   ${result}   Set Variable    ${True}

   ${startRow}     Get Start Row On WoW Change    nameOftable=OEM East
   ${endRow}       Get End Row On WoW Change      nameOftable=OEM East
   File Should Exist    path=${WOW_CHANGE_FILE_PATH}
   Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange

   FOR    ${rowIndex}    IN RANGE    ${startRow}    ${endRow}+3
       ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
       ${valueColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=6
       IF    '${valueColOnWoWChange}' != '${EMPTY}'
            ${result}     Set Variable    ${False}
            @{rowOnTableError}   Create List
            Append To List    ${rowOnTableError}   OEM East
            Append To List    ${rowOnTableError}   Ships
            Append To List    ${rowOnTableError}   ${oemGroupColOnWoWChange}
            Append To List    ${rowOnTableError}   ${valueColOnWoWChange}
            Append To List    ${rowOnTableError}   ${EMPTY}
            Append To List    ${tableError}    ${rowOnTableError}
       END
   END
   IF    '${result}' == '${False}'
        @{listNameOfColsForHeader}   Create List
        Append To List    ${listNameOfColsForHeader}   TABLE
        Append To List    ${listNameOfColsForHeader}   CHECK POINT
        Append To List    ${listNameOfColsForHeader}   OEM GROUP
        Append To List    ${listNameOfColsForHeader}   VALUE ON WOW CHANGE
        Append To List    ${listNameOfColsForHeader}   VALUE ON SG
        Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}    hasHeader=${False}
        Fail   The Ships data for the OEM East table is wrong
   END
   Close Current Excel Document

Verify WoW of Ships for the OEM East table
   [Tags]  WoWChange_0006
   [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM East table

   Depends On Test    name=Verify Ships for the OEM East table
   @{tableError}   Create List
   ${result}   Set Variable    ${True}

   ${startRow}     Get Start Row On WoW Change    nameOftable=OEM East
   ${endRow}       Get End Row On WoW Change      nameOftable=OEM East
   File Should Exist    path=${WOW_CHANGE_FILE_PATH}
   Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange

   FOR    ${rowIndex}    IN RANGE    ${startRow}    ${endRow}+3
       ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
       ${valueColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=7
       IF    '${valueColOnWoWChange}' != '${EMPTY}'
            ${result}     Set Variable    ${False}
            @{rowOnTableError}   Create List
            Append To List    ${rowOnTableError}   OEM East
            Append To List    ${rowOnTableError}   WoW Of Ships
            Append To List    ${rowOnTableError}   ${oemGroupColOnWoWChange}
            Append To List    ${rowOnTableError}   ${valueColOnWoWChange}
            Append To List    ${rowOnTableError}   ${EMPTY}
            Append To List    ${tableError}    ${rowOnTableError}
       END
   END
   IF    '${result}' == '${False}'
        @{listNameOfColsForHeader}   Create List
        Append To List    ${listNameOfColsForHeader}   TABLE
        Append To List    ${listNameOfColsForHeader}   CHECK POINT
        Append To List    ${listNameOfColsForHeader}   OEM GROUP
        Append To List    ${listNameOfColsForHeader}   VALUE ON WOW CHANGE
        Append To List    ${listNameOfColsForHeader}   VALUE ON SG
        Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}    hasHeader=${False}
        Fail   The WoW Of Ships data for the OEM East table is wrong
   END
   Close Current Excel Document

Verify Backlog for the OEM East table
    [Tags]  WoWChange_0007
    [Documentation]     Verify the data of Backlog column for the OEM East table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    IF    '${currentQuarter}' == '4'
         ${nextQuarter}     Set Variable    1
         ${currentYear}     Evaluate    ${currentYear}+1
    ELSE
         ${nextQuarter}      Evaluate    ${currentQuarter}+1
    END
    Check BGT, Ship, Backlog On WoW Change    nameOftable=OEM East     nameOfCol=Backlog  transType=BACKLOG   attribute=AMOUNT     year=${currentYear}     quarter=${nextQuarter}

Verify LOS for the OEM East table
    [Tags]  WoWChange_0008
    [Documentation]     Verify the data of LOS column for the OEM East table

#     Depends On Test    name=Verify Ships for the OEM East table
    Depends On Test    name=Verify Backlog for the OEM East table

    Check LOS On WoW Change    nameOftable=OEM East     nameOfCol=LOS

Verify WoW of LOS for the OEM East table
   [Tags]  WoWChange_0009
   [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM East table

   Depends On Test    name=Verify LOS for the OEM East table
   Check WoW On WoW Change  nameOftable=OEM East     nameOfCol=WoW Of LOS

Verify GAP for the OEM East table
    [Tags]  WoWChange_0010
    [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM East table

    Depends On Test    name=Verify LOS for the OEM East table
    Depends On Test    name=Verify LW Commit for the OEM East table
    Check GAP On WoW Change    nameOftable=OEM East  nameOfCol=GAP

Veify Comments for the OEM East table
    [Tags]  WoWChange_0011
    [Documentation]     Verify the data of Comments column for the OEM East table

    Check LW Commit, Comment On WoW Change  nameOftable=OEM East     nameOfCol=Comments

#Verify Prev Quarter Ship for the OEM West table
#    [Tags]  WoWChange_0012
#    [Documentation]     Verify the data of Pre Q Ships column for the OEM East table
#
#    ${currentYear}              Get Current Year
#    ${currentQuarter}           Get Current Quarter
#    ${preQuarter}   Set Variable    ${currentQuarter}
#
#    Check BGT, Ship, Backlog On WoW Change    nameOftable=OEM West + Channel     nameOfCol=Pre Q Ships  transType=LOS   attribute=AMOUNT     year=${currentYear}     quarter=${preQuarter}

# Verify Current Quarter Budget for the OEM West table
#     [Tags]  WoWChange_0013
#     [Documentation]     Verify the data of Current Q Budget column for the OEM East table

#     ${currentYear}              Get Current Year
#     ${currentQuarter}           Get Current Quarter
#     IF    '${currentQuarter}' == '4'
#          ${nextQuarter}     Set Variable    1
#          ${currentYear}     Evaluate    ${currentYear}+1
#     ELSE
#          ${nextQuarter}      Evaluate    ${currentQuarter}+1
#     END

#     Check BGT, Ship, Backlog On WoW Change    nameOftable=OEM West + Channel     nameOfCol=Current Q Budget  transType=BUDGET   attribute=AMOUNT     year=${currentYear}     quarter=${nextQuarter}

Verify LW Commit for the OEM West table
   [Tags]  WoWChange_0014
   [Documentation]     Verify the data of LW Commit column for the OEM East table

   Check LW Commit, Comment On WoW Change   nameOftable=OEM West + Channel  nameOfCol=LW Commit

Verify TW Commit for the OEM West table
   [Tags]  WoWChange_0015
   [Documentation]     Verify the data of TW Commit column for the OEM West table

   Check TW Commit On WoW Change  nameOftable=OEM West + Channel  nameOfCol=TW Commit

Verify Ships for the OEM West table
   [Tags]  WoWChange_0016
   [Documentation]     Verify the data of Ships column for the OEM West table

   @{tableError}   Create List
   ${result}   Set Variable    ${True}

   ${startRow}     Get Start Row On WoW Change    nameOftable=OEM West + Channel
   ${endRow}       Get End Row On WoW Change      nameOftable=OEM West + Channel
   File Should Exist      path=${WOW_CHANGE_FILE_PATH}
   Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange

   FOR    ${rowIndex}    IN RANGE    ${startRow}    ${endRow}+3
       ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
       ${valueColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=6
       IF    '${valueColOnWoWChange}' != '${EMPTY}'
            ${result}     Set Variable    ${False}
            @{rowOnTableError}   Create List
            Append To List    ${rowOnTableError}   OEM West + Channel
            Append To List    ${rowOnTableError}   Ships
            Append To List    ${rowOnTableError}   ${oemGroupColOnWoWChange}
            Append To List    ${rowOnTableError}   ${valueColOnWoWChange}
            Append To List    ${rowOnTableError}   ${EMPTY}
            Append To List    ${tableError}    ${rowOnTableError}
       END
   END
   IF    '${result}' == '${False}'
        @{listNameOfColsForHeader}   Create List
        Append To List    ${listNameOfColsForHeader}   TABLE
        Append To List    ${listNameOfColsForHeader}   CHECK POINT
        Append To List    ${listNameOfColsForHeader}   OEM GROUP
        Append To List    ${listNameOfColsForHeader}   VALUE ON WOW CHANGE
        Append To List    ${listNameOfColsForHeader}   VALUE ON SG
        Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}    hasHeader=${False}
        Fail   The Ships data for the OEM West + Channel table is wrong
   END
   Close Current Excel Document

Verify WoW of Ships for the OEM West table
   [Tags]  WoWChange_0017
   [Documentation]     Verify the data of WoW(WoW of Ships column) column for the OEM West table

   Depends On Test    name=Verify Ships for the OEM West table
   @{tableError}   Create List
   ${result}   Set Variable    ${True}

   ${startRow}     Get Start Row On WoW Change    nameOftable=OEM West + Channel
   ${endRow}       Get End Row On WoW Change      nameOftable=OEM West + Channel
   File Should Exist    path=${WOW_CHANGE_FILE_PATH}
   Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange

   FOR    ${rowIndex}    IN RANGE    ${startRow}    ${endRow}+3
       ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
       ${valueColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=7
       IF    '${valueColOnWoWChange}' != '${EMPTY}'
            ${result}     Set Variable    ${False}
            @{rowOnTableError}   Create List
            Append To List    ${rowOnTableError}   OEM West + Channel
            Append To List    ${rowOnTableError}   WoW Of Ships
            Append To List    ${rowOnTableError}   ${oemGroupColOnWoWChange}
            Append To List    ${rowOnTableError}   ${valueColOnWoWChange}
            Append To List    ${rowOnTableError}   ${EMPTY}
            Append To List    ${tableError}    ${rowOnTableError}
       END
   END
   IF    '${result}' == '${False}'
        @{listNameOfColsForHeader}   Create List
        Append To List    ${listNameOfColsForHeader}   TABLE
        Append To List    ${listNameOfColsForHeader}   CHECK POINT
        Append To List    ${listNameOfColsForHeader}   OEM GROUP
        Append To List    ${listNameOfColsForHeader}   VALUE ON WOW CHANGE
        Append To List    ${listNameOfColsForHeader}   VALUE ON SG
        Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}    hasHeader=${False}
        Fail   The WoW Of Ships data for the OEM West + Channel table is wrong
   END
   Close Current Excel Document

Verify Backlog for the OEM West table
    [Tags]  WoWChange_0018
    [Documentation]     Verify the data of Backlog column for the OEM West table

    ${currentYear}              Get Current Year
    ${currentQuarter}           Get Current Quarter
    IF    '${currentQuarter}' == '4'
         ${nextQuarter}     Set Variable    1
         ${currentYear}     Evaluate    ${currentYear}+1
    ELSE
         ${nextQuarter}      Evaluate    ${currentQuarter}+1
    END

    Check BGT, Ship, Backlog On WoW Change    nameOftable=OEM West + Channel     nameOfCol=Backlog  transType=BACKLOG   attribute=AMOUNT     year=${currentYear}     quarter=${nextQuarter}

Verify LOS for the OEM West table
    [Tags]  WoWChange_0019
    [Documentation]     Verify the data of LOS column for the OEM West table

    Depends On Test    name=Verify Ships for the OEM West table
    Depends On Test    name=Verify Backlog for the OEM West table
    Check LOS On WoW Change    nameOftable=OEM West + Channel     nameOfCol=LOS

Verify WoW of LOS for the OEM West table
   [Tags]  WoWChange_0020
   [Documentation]     Verify the data of WoW(WoW of LOS column) column for the OEM West table

   Depends On Test    name=Verify LOS for the OEM West table
   Check WoW On WoW Change  nameOftable=OEM West + Channel     nameOfCol=WoW Of LOS

Verify GAP for the OEM West table
    [Tags]  WoWChange_0021
    [Documentation]     Verify the data of GAP(LOS - Commit) column for the OEM West table

    Depends On Test    name=Verify LOS for the OEM West table
    Depends On Test    name=Verify LW Commit for the OEM West table
    Check GAP On WoW Change    nameOftable=OEM West + Channel  nameOfCol=GAP

Verify Comments for the OEM West table
   [Tags]  WoWChange_0022
   [Documentation]     Verify the data of Comments column for the OEM West table

   Check LW Commit, Comment On WoW Change  nameOftable=OEM West + Channel     nameOfCol=Comments

