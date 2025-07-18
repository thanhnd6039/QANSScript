*** Settings ***
Resource    ../CommonPage.robot
Resource    SGPage.robot

*** Variables ***
${WOW_CHANGE_FILE_PATH}                   ${OUTPUT_DIR}\\Wow Change [Current Week].xlsx
${WOW_CHANGE_ON_VDC_FILE_PATH}            ${OUTPUT_DIR}\\Wow Change [Current Week] On VDC.xlsx
${WOW_CHANGE_RESULT_FILE_PATH}            ${OUTPUT_DIR}\\WoWChangeResult.xlsx

${POS_OEM_GROUP_COL_ON_WOW_CHANGE}         1
${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}   0
${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}       1

*** Keywords ***
Setup Test Environment For WoW Change Report
    [Arguments]     ${browser}
    Create Excel File    filePath=${WOW_CHANGE_RESULT_FILE_PATH}
    Wait Until Created    path=${WOW_CHANGE_RESULT_FILE_PATH}   timeout=${TIMEOUT}
    @{emptyTable}   Create List
    @{listNameOfColsForHeader}   Create List
    Append To List    ${listNameOfColsForHeader}  TABLE
    Append To List    ${listNameOfColsForHeader}  CHECK POINT
    Append To List    ${listNameOfColsForHeader}  OEM GROUP
    Append To List    ${listNameOfColsForHeader}  ON WOW CHANGE
    Append To List    ${listNameOfColsForHeader}  ON SG
    Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=@{emptyTable}  hasHeader=${True}
    Setup    browser=${browser}
    Navigate To Report    reportLink=/NetSuite%20Reports/Sales/SalesManagement/Wow%20Change%20%5BCurrent%20Week%5D
    Export Report To      option=Excel
    Wait Until Created    path=${WOW_CHANGE_FILE_PATH}  timeout=${TIMEOUT}
    Navigate To Report    reportLink=/NetSuite+Reports/Sales/Sales+Gap+Report+NS+With+SO+Forecast&rs:Command=Render
    @{parentClassOptionsOnSG}   Create List
    Append To List    ${parentClassOptionsOnSG}     MEM
    Append To List    ${parentClassOptionsOnSG}     STORAGE
    Select Parent Class On SG Report    options=${parentClassOptionsOnSG}
    Sleep    10s  
    Click On Button View Report
    Wait Until Element Is Enabled    locator=${btnViewReport}   timeout=${TIMEOUT}
    Export Report To      option=Excel
    Wait Until Created    path=${SG_FILE_PATH}
    Sleep    20s
    Close Browser

Get WoW By Formular By OEM GRoup
    [Arguments]     ${nameOftable}    ${nameOfCol}    ${oemGroup}
    ${wowByFormula}           Set Variable    0
    ${valueOnWoWChange}       Set Variable    0
    ${valueOnWoWChangeOnVDC}  Set Variable    0


    ${tableOnWoWChange}         Create Table On WoW Change           nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    ${tableOnWoWChangeOnVDC}    Create Table On WoW Change On VDC    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}

    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
        IF    '${oemGroupOnWoWChange}' == '${oemGroup}'
            ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
            BREAK
        END
    END

    FOR    ${rowOnWoWChangeOnVDC}    IN    @{tableOnWoWChangeOnVDC}
        ${oemGroupOnWoWChangeOnVDC}  Set Variable    ${rowOnWoWChangeOnVDC[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
        IF    '${oemGroupOnWoWChangeOnVDC}' == '${oemGroup}'
            ${valueOnWoWChangeOnVDC}     Set Variable    ${rowOnWoWChangeOnVDC[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
            BREAK
        END
    END
    ${wowByFormula}     Evaluate    ${valueOnWoWChange}-${valueOnWoWChangeOnVDC}
    [Return]    ${wowByFormula}

Check GAP On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}
    @{tableError}   Create List
    ${result}       Set Variable    ${True}
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0

    IF    '${nameOftable}' != 'OEM East' and '${nameOftable}' != 'OEM West + Channel'
         Fail    The table parameter ${nameOftable} is invalid.
    END

    ${startRow}     Get Start Row On WoW Change    nameOftable=${nameOftable}
    ${endRow}       Get End Row On WoW Change      nameOftable=${nameOftable}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${totalRow}+1
        ${oemGroupColOnWoWChange}       Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
        ${lwCommitColOnWoWChange}       Read Excel Cell    row_num=${rowIndex}    col_num=4
        ${losColOnWoWChange}            Read Excel Cell    row_num=${rowIndex}    col_num=9
        ${gapByFormula}                 Evaluate    ${losColOnWoWChange}-${lwCommitColOnWoWChange}
        ${gapColOnWoWChange}            Read Excel Cell    row_num=${rowIndex}    col_num=11
        ${gapByFormula}                 Evaluate  "%.2f" % ${gapByFormula}
        ${gapColOnWoWChange}            Evaluate  "%.2f" % ${gapColOnWoWChange}
        IF    ${gapColOnWoWChange} != ${gapByFormula}
             ${result}     Set Variable    ${False}
             @{rowOnTableError}   Create List
             Append To List    ${rowOnTableError}   ${nameOftable}
             Append To List    ${rowOnTableError}   ${nameOfCol}
             Append To List    ${rowOnTableError}   ${oemGroupColOnWoWChange}
             Append To List    ${rowOnTableError}   ${gapColOnWoWChange}
             Append To List    ${rowOnTableError}   ${gapByFormula}
             Append To List    ${tableError}    ${rowOnTableError}
        END

    END
    Close Current Excel Document
    IF    '${result}' == '${False}'
         @{listNameOfColsForHeader}   Create List
         Append To List    ${listNameOfColsForHeader}   TABLE
         Append To List    ${listNameOfColsForHeader}   CHECK POINT
         Append To List    ${listNameOfColsForHeader}   OEM GROUP
         Append To List    ${listNameOfColsForHeader}   VALUE ON WOW CHANGE
         Append To List    ${listNameOfColsForHeader}   VALUE ON SG
         Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}    hasHeader=${False}
         Fail   The ${nameOfCol} data for the ${nameOftable} table is wrong
    END

Check WoW On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}   ${isNextQuarter}=${False}
    ${result}   Set Variable    ${True}
    @{tableError}   Create List
    ${wowByFormula}     Set Variable    0

    ${tableOnWoWChange}      Create Table On WoW Change    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    IF    '${isNextQuarter}' == '${True}'
         FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
            ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
            ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
            IF    '${valueOnWoWChange}' != '${EMPTY}'
                 ${result}     Set Variable    ${False}
                 @{rowOnTableError}   Create List
                 Append To List    ${rowOnTableError}   ${nameOftable}
                 Append To List    ${rowOnTableError}   ${nameOfCol}
                 Append To List    ${rowOnTableError}   ${oemGroupOnWoWChange}
                 Append To List    ${rowOnTableError}   ${valueOnWoWChange}
                 Append To List    ${rowOnTableError}   ${EMPTY}
                 Append To List    ${tableError}    ${rowOnTableError}
            END
         END
    ELSE
        FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
            ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
            ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}

            IF    '${nameOfCol}' == 'WoW Of Ships'
                 ${wowByFormula}         Get WoW By Formular By OEM GRoup    nameOftable=${nameOftable}    nameOfCol=Ships    oemGroup=${oemGroupOnWoWChange}
            ELSE
                 ${wowByFormula}         Get WoW By Formular By OEM GRoup    nameOftable=${nameOftable}    nameOfCol=LOS    oemGroup=${oemGroupOnWoWChange}
            END

            ${valueOnWoWChange}      Evaluate  "%.2f" % ${valueOnWoWChange}
            ${wowByFormula}          Evaluate  "%.2f" % ${wowByFormula}

            IF    ${valueOnWoWChange} != ${wowByFormula}
                  ${result}     Set Variable    ${False}
                  @{rowOnTableError}   Create List
                  Append To List    ${rowOnTableError}   ${nameOftable}
                  Append To List    ${rowOnTableError}   ${nameOfCol}
                  Append To List    ${rowOnTableError}   ${oemGroupOnWoWChange}
                  Append To List    ${rowOnTableError}   ${valueOnWoWChange}
                  Append To List    ${rowOnTableError}   ${wowByFormula}
                  Append To List    ${tableError}    ${rowOnTableError}
            END
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
         Fail   The ${nameOfCol} data for the ${nameOftable} table is wrong
    END

Check TW Commit On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}
    @{tableError}   Create List
    ${result}   Set Variable    ${True}

    ${tableOnWoWChange}     Create Table On WoW Change    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}

    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
        ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
        IF    '${valueOnWoWChange}' == 'None'            
            ${valueOnWoWChange}    Set Variable    ${EMPTY}
        END      
        IF    '${valueOnWoWChange}' != '${EMPTY}'
             ${result}     Set Variable    ${False}
             @{rowOnTableError}   Create List
             Append To List    ${rowOnTableError}   ${nameOftable}
             Append To List    ${rowOnTableError}   ${nameOfCol}
             Append To List    ${rowOnTableError}   ${oemGroupOnWoWChange}
             Append To List    ${rowOnTableError}   ${valueOnWoWChange}
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
         Fail   The ${nameOfCol} data for the ${nameOftable} table is wrong
    END
Create Table On WoW Change On VDC
    [Arguments]     ${nameOftable}    ${nameOfCol}

    @{table}     Create List
    ${result}       Set Variable    ${True}
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0

    IF    '${nameOftable}' != 'OEM East' and '${nameOftable}' != 'OEM West + Channel'
         Fail    The table parameter ${nameOftable} is invalid
    END

    ${startRow}     Get Start Row On WoW Change    nameOftable=${nameOftable}
    ${endRow}       Get End Row On WoW Change      nameOftable=${nameOftable}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    IF    '${nameOfCol}' == 'LW Commit'
         ${posOfValueCol}    Set Variable    5
    ELSE IF   '${nameOfCol}' == 'WoW Of Ships'
         ${posOfValueCol}    Set Variable    7
    ELSE IF   '${nameOfCol}' == 'WoW Of LOS'
         ${posOfValueCol}    Set Variable    10
    ELSE
        ${posOfValueCol}     Get Position Of Column On WoW Change    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    END

    File Should Exist      path=${WOW_CHANGE_ON_VDC_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_ON_VDC_FILE_PATH}           doc_id=WoWChangeOnVDC
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${WOW_CHANGE_ON_VDC_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${totalRow}+1
        ${oemGroupColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
        ${valueColOnWoWChangeOnVDC}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        ${rowOnTable}   Create List
        ...             ${oemGroupColOnWoWChangeOnVDC}
        ...             ${valueColOnWoWChangeOnVDC}
        Append To List    ${table}   ${rowOnTable}
    END

    Close Current Excel Document
    [Return]    ${table}

Check LW Commit, Comment On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}
    ${result}       Set Variable    ${True}
    @{tableError}   Create List

    ${tableOnWoWChange}         Create Table On WoW Change           nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    ${tableOnWoWChangeOnVDC}    Create Table On WoW Change On VDC    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
        ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
        IF    '${valueOnWoWChange}' == 'None'
             ${valueOnWoWChange}    Set Variable    ${EMPTY}
        END
        IF    '${nameOfCol}' == 'LW Commit'
             IF    '${valueOnWoWChange}' == 'None'
                ${valueOnWoWChange}    Set Variable    0
             END
             ${valueOnWoWChange}      Evaluate  "%.2f" % ${valueOnWoWChange}
        END
        FOR    ${rowOnWoWChangeOnVDC}    IN    @{tableOnWoWChangeOnVDC}
            ${oemGroupOnWoWChangeOnVDC}  Set Variable    ${rowOnWoWChangeOnVDC[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
            ${valueOnWoWChangeOnVDC}     Set Variable    ${rowOnWoWChangeOnVDC[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
            IF    '${valueOnWoWChangeOnVDC}' == 'None'
                 ${valueOnWoWChangeOnVDC}   Set Variable    ${EMPTY}
            END
            IF    '${nameOfCol}' == 'LW Commit'
                ${valueOnWoWChangeOnVDC}      Evaluate  "%.2f" % ${valueOnWoWChangeOnVDC}
            END
            IF    '${oemGroupOnWoWChange}' == '${oemGroupOnWoWChangeOnVDC}'
                 IF    '${valueOnWoWChange}' != '${valueOnWoWChangeOnVDC}'
                      ${result}     Set Variable    ${False}
                      @{rowOnTableError}   Create List
                      Append To List    ${rowOnTableError}   ${nameOftable}
                      Append To List    ${rowOnTableError}   ${nameOfCol}
                      Append To List    ${rowOnTableError}   ${oemGroupOnWoWChange}
                      Append To List    ${rowOnTableError}   ${valueOnWoWChange}
                      Append To List    ${rowOnTableError}   ${valueOnWoWChangeOnVDC}
                      Append To List    ${tableError}    ${rowOnTableError}
                 END
                 BREAK
            END
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
         Fail   The ${nameOfCol} data for the ${nameOftable} table is different between the WoW Change Report and WoW Change on VDC
    END


Check BGT, Ship, Backlog On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}     ${transType}   ${attribute}   ${year}     ${quarter}
    ${result}       Set Variable    ${True}
    @{tableError}   Create List
    ${sumOfValueOfOEMGroup}     Set Variable    0


    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    ${tableOnWoWChange}     Create Table On WoW Change    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    ${tableOnSG}            Create Table For SG Report    transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    #   Verify the data for each OEM Group
    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupCol}          Set Variable    ${rowOnWoWChange[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
        ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
        IF    '${oemGroupCol}' == 'OTHERS' or '${oemGroupCol}' == 'Total'
             Continue For Loop
        END
        ${valueOnSG}    Get Value By OEM Group On SG Report     tableOnSG=${tableOnSG}    oemGroup=${oemGroupCol}

        ${sumOfValueOfOEMGroup}     Evaluate    ${sumOfValueOfOEMGroup}+${valueOnSG}
        ${valueOnWoWChange}      Evaluate  "%.2f" % ${valueOnWoWChange}
        ${valueOnSG}             Evaluate  "%.2f" % ${valueOnSG}

        IF    ${valueOnWoWChange} != ${valueOnSG}
             ${result}     Set Variable    ${False}
             @{rowOnTableError}   Create List
             Append To List    ${rowOnTableError}   ${nameOftable}
             Append To List    ${rowOnTableError}   ${nameOfCol}
             Append To List    ${rowOnTableError}   ${oemGroupCol}
             Append To List    ${rowOnTableError}   ${valueOnWoWChange}
             Append To List    ${rowOnTableError}   ${valueOnSG}
             Append To List    ${tableError}    ${rowOnTableError}
        END
    END
    #   Verify the Total data
    ${totalOnSG}    Set Variable    0
    ${valueOnSG}    Set Variable    0
    FOR    ${rawData}    IN    @{tableOnSG}
        ${mainSalesRepColOnSG}  Set Variable    ${rawData[${POS_MAIN_SALES_REP_COL_ON_SG_TABLE}]}
        IF    '${nameOftable}' == 'OEM East'
             IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
                ${valueOnSG}    Set Variable    ${rawData[${POS_VALUE_COL_ON_SG_TABLE}]}
                ${totalOnSG}    Evaluate    ${totalOnSG}+${valueOnSG}
             END
        ELSE
            IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
                ${valueOnSG}    Set Variable    ${rawData[${POS_VALUE_COL_ON_SG_TABLE}]}
                ${totalOnSG}    Evaluate    ${totalOnSG}+${valueOnSG}
            END
        END
    END
    ${totalOnWoWchange}     Set Variable    0
    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupCol}          Set Variable    ${rowOnWoWChange[${POS_OEM_GROUP_COL_ON_WOW_CHANGE_TABLE}]}
        IF    '${oemGroupCol}' == 'Total'
             ${totalOnWoWchange}    Set Variable    ${rowOnWoWChange[${POS_VALUE_COL_ON_WOW_CHANGE_TABLE}]}
             BREAK
        END
    END
    ${totalOnSG}          Evaluate  "%.2f" % ${totalOnSG}
    ${totalOnWoWchange}   Evaluate  "%.2f" % ${totalOnWoWchange}

    IF    ${totalOnWoWchange} != ${totalOnSG}
         ${result}     Set Variable    ${False}
         @{rowOnTableError}   Create List
         Append To List    ${rowOnTableError}   ${nameOftable}
         Append To List    ${rowOnTableError}   ${nameOfCol}
         Append To List    ${rowOnTableError}   ${oemGroupCol}
         Append To List    ${rowOnTableError}   ${totalOnWoWchange}
         Append To List    ${rowOnTableError}   ${totalOnSG}
         Append To List    ${tableError}    ${rowOnTableError}
    END
    #  Verify the OTHERS data
    ${othersOnSG}   Evaluate    ${totalOnSG}-${sumOfValueOfOEMGroup}
    ${othersOnWoWChange}     Set Variable    0
    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupCol}          Set Variable    ${rowOnWoWChange[0]}
        IF    '${oemGroupCol}' == 'OTHERS'
             ${othersOnWoWChange}    Set Variable    ${rowOnWoWChange[1]}
             BREAK
        END
    END
    ${othersOnSG}          Evaluate  "%.2f" % ${othersOnSG}
    ${othersOnWoWChange}   Evaluate  "%.2f" % ${othersOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSG}
         ${result}     Set Variable    ${False}
         @{rowOnTableError}   Create List
         Append To List    ${rowOnTableError}   ${nameOftable}
         Append To List    ${rowOnTableError}   ${nameOfCol}
         Append To List    ${rowOnTableError}   ${oemGroupCol}
         Append To List    ${rowOnTableError}   ${othersOnWoWChange}
         Append To List    ${rowOnTableError}   ${othersOnSG}
         Append To List    ${tableError}    ${rowOnTableError}
    END

    IF    '${result}' == '${False}'
         @{listNameOfColsForHeader}   Create List
         Append To List    ${listNameOfColsForHeader}   TABLE
         Append To List    ${listNameOfColsForHeader}   CHECK POINT
         Append To List    ${listNameOfColsForHeader}   OEM GROUP
         Append To List    ${listNameOfColsForHeader}   VALUE ON WOW CHANGE
         Append To List    ${listNameOfColsForHeader}   VALUE ON SG
         Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}    hasHeader=${False}
         Fail   The ${nameOfCol} data for the ${nameOftable} table is different between the WoW Change Report and SG Report
    END

Get Row Index For Search Col
    [Arguments]     ${nameOftable}
    ${rowIndexForSearchCol}     Set Variable    0

    File Should Exist    path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${WOW_CHANGE_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=1
        IF    '${oemGroupCol}' == '${nameOftable}'
             ${rowIndexForSearchCol}    Set Variable    ${rowIndex}
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${rowIndexForSearchCol}

Get Position Of Column On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}
    ${pos}  Set Variable    0
    ${rowIndexForSearchCol}     Set Variable    0

    ${rowIndexForSearchCol}     Get Row Index For Search Col    nameOftable=${nameOftable}
    ${pos}      Get Position Of Column    filePath=${WOW_CHANGE_FILE_PATH}    rowIndex=${rowIndexForSearchCol}    searchStr=${nameOfCol}

    [Return]    ${pos}

Get Start Row On WoW Change
    [Arguments]     ${nameOftable}

    ${startRow}     Set Variable    0
    ${posOfCol}     Get Position Of Column On WoW Change    nameOftable=${nameOftable}   nameOfCol=${nameOftable}

    File Should Exist    path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${WOW_CHANGE_FILE_PATH}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == '${nameOftable}'
             ${startRow}    Evaluate    ${rowIndex}+1
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${startRow}

Get End Row On WoW Change
    [Arguments]     ${nameOftable}

    ${endRow}   Set Variable    0
    ${count}    Set Variable    0

    ${posOfCol}     Get Position Of Column On WoW Change    nameOftable=${nameOftable}   nameOfCol=${nameOftable}

    File Should Exist      path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${WOW_CHANGE_FILE_PATH}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == 'OTHERS'
             ${count}   Evaluate    ${count}+1
             IF    '${nameOftable}' == 'OEM East'
                  IF    '${count}' == '1'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             ELSE IF     '${nameOftable}' == 'OEM West + Channel'
                  IF    '${count}' == '2'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             END
        END
    END

    Close Current Excel Document
    [Return]    ${endRow}


Check LOS On WoW Change
    [Arguments]     ${nameOftable}   ${nameOfCol}
    ${result}       Set Variable    ${True}
    @{tableError}   Create List

    ${startRow}     Get Start Row On WoW Change    nameOftable=${nameOftable}
    ${endRow}       Get End Row On WoW Change      nameOftable=${nameOftable}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    File Should Exist      path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}           doc_id=WoWChange
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${totalRow}+1
        ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
        ${shipsColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=6
        ${backlogColOnWoWChange}           Read Excel Cell    row_num=${rowIndex}    col_num=8
        ${losColOnWoWChange}               Read Excel Cell    row_num=${rowIndex}    col_num=9
        IF    '${shipsColOnWoWChange}' == '${EMPTY}' or '${shipsColOnWoWChange}' == 'None'
             ${shipsColOnWoWChange}     Set Variable    0
        END
        IF    '${backlogColOnWoWChange}' == '${EMPTY}' or '${backlogColOnWoWChange}' == 'None'
             ${backlogColOnWoWChange}     Set Variable    0
        END
        IF    '${losColOnWoWChange}' == '${EMPTY}' or '${losColOnWoWChange}' == 'None'
             ${losColOnWoWChange}     Set Variable    0
        END
        ${losValueByFormular}   Evaluate    ${shipsColOnWoWChange}+${backlogColOnWoWChange}
        ${losColOnWoWChange}          Evaluate  "%.2f" % ${losColOnWoWChange}
        ${losValueByFormular}          Evaluate  "%.2f" % ${losValueByFormular}
        IF    ${losColOnWoWChange} != ${losValueByFormular}
             ${result}       Set Variable    ${False}
             @{rowOnTableError}   Create List
             Append To List    ${rowOnTableError}   ${nameOftable}
             Append To List    ${rowOnTableError}   ${nameOfCol}
             Append To List    ${rowOnTableError}   ${oemGroupColOnWoWChange}
             Append To List    ${rowOnTableError}   ${losColOnWoWChange}
             Append To List    ${rowOnTableError}   ${losValueByFormular}
             Append To List    ${tableError}    ${rowOnTableError}
        END
    END
    Close Current Excel Document
    IF    '${result}' == '${False}'
        @{listNameOfColsForHeader}   Create List
         Append To List    ${listNameOfColsForHeader}   TABLE
         Append To List    ${listNameOfColsForHeader}   CHECK POINT
         Append To List    ${listNameOfColsForHeader}   OEM GROUP
         Append To List    ${listNameOfColsForHeader}   VALUE ON WOW CHANGE
         Append To List    ${listNameOfColsForHeader}   VALUE ON SG
         Write Table To Excel    filePath=${WOW_CHANGE_RESULT_FILE_PATH}    listNameOfCols=${listNameOfColsForHeader}    table=${tableError}    hasHeader=${False}
         Fail   The ${nameOfCol} data for the ${nameOftable} table is different between the WoW Change Report and SG Report
    END

Create Table On WoW Change
    [Arguments]     ${nameOftable}    ${nameOfCol}

    @{table}     Create List
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0

    IF    '${nameOftable}' != 'OEM East' and '${nameOftable}' != 'OEM West + Channel'
         Fail    The nameOftable parameter ${nameOftable} is invalid
    END

    ${startRow}     Get Start Row On WoW Change    nameOftable=${nameOftable}
    ${endRow}       Get End Row On WoW Change      nameOftable=${nameOftable}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    IF    '${nameOfCol}' == 'Pre Q Ships'
         ${posOfValueCol}    Set Variable    2
    ELSE IF    '${nameOfCol}' == 'Current Q Budget'
         ${posOfValueCol}    Set Variable    3
    ELSE IF    '${nameOfCol}' == 'LW Commit'
         ${posOfValueCol}    Set Variable    4
    ELSE IF  '${nameOfCol}' == 'TW Commit'
         ${posOfValueCol}    Set Variable    5
    ELSE IF  '${nameOfCol}' == 'WoW Of Ships'
         ${posOfValueCol}    Set Variable    7
    ELSE IF  '${nameOfCol}' == 'WoW Of LOS'
         ${posOfValueCol}    Set Variable    10
    ELSE
        ${posOfValueCol}     Get Position Of Column On WoW Change    nameOftable=${nameOftable}    nameOfCol=${nameOfCol}
    END

    File Should Exist      path=${WOW_CHANGE_FILE_PATH}
    Open Excel Document    filename=${WOW_CHANGE_FILE_PATH}           doc_id=WoWChange
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${totalRow}+1
        ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${POS_OEM_GROUP_COL_ON_WOW_CHANGE}
        ${valueColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        ${rowOnTable}   Create List
        ...             ${oemGroupColOnWoWChange}
        ...             ${valueColOnWoWChange}
        Append To List    ${table}   ${rowOnTable}
    END

    Close Current Excel Document
    [Return]    ${table}

Get List Of Sales Member In OEM East Table
    @{listOfSalesMember}    Create List
    Append To List    ${listOfSalesMember}      Chris Seitz
    Append To List    ${listOfSalesMember}      Daniel Schmidt
    Append To List    ${listOfSalesMember}      Eli Tiomkin
    Append To List    ${listOfSalesMember}      Michael Pauser

    [Return]    ${listOfSalesMember}

Get List Of OEM Group Shown In OEM East Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      NVIDIA/MELLANOX
    Append To List    ${listOfOEMGroup}      NOKIA/ALCATEL LUCENT WORLDWIDE
    Append To List    ${listOfOEMGroup}      CURTISS WRIGHT GROUP
    Append To List    ${listOfOEMGroup}      JUNIPER NETWORKS
    Append To List    ${listOfOEMGroup}      ERICSSON WORLDWIDE

    [Return]    ${listOfOEMGroup}

Get List Of Sales Member In OEM West Table
    @{listOfSalesMember}    Create List
    Append To List    ${listOfSalesMember}      Amy Duong
    Append To List    ${listOfSalesMember}      Caden Douglas
    Append To List    ${listOfSalesMember}      Michael Nilsson
    Append To List    ${listOfSalesMember}      Tiger Wang
    Append To List    ${listOfSalesMember}      Yoda Yasunobu
    Append To List    ${listOfSalesMember}      Nicole Lau
    Append To List    ${listOfSalesMember}      Huan Tran
    Append To List    ${listOfSalesMember}      Dave Beasley
    Append To List    ${listOfSalesMember}      Cameron Sinclair
    Append To List    ${listOfSalesMember}      Clint Stalker

    [Return]    ${listOfSalesMember}

Get List Of OEM Group Shown In OEM West Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      PALO ALTO NETWORKS
    Append To List    ${listOfOEMGroup}      ARISTA
    Append To List    ${listOfOEMGroup}      SCHWEITZER ENGINEERING LABORATORIES (SEL)
    Append To List    ${listOfOEMGroup}      PANASONIC AVIONICS
    Append To List    ${listOfOEMGroup}      ZTE KANGXUN TELECOM CO. LTD.
    Append To List    ${listOfOEMGroup}      NATIONAL INSTRUMENTS

    [Return]    ${listOfOEMGroup}