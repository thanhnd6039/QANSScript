*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${testResultOfSGReportByOEMGroupFilePath}   C:\\RobotFramework\\Results\\SGReport\\SGReportResultByOEMGroup.xlsx
${testResultOfSGReportByPNFilePath}         C:\\RobotFramework\\Results\\SGReport\\SGReportResultByPN.xlsx
${SGFilePath}                               C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx

${startRowOnSG}                      6
${rowIndexForSearchPosOfCol}         3
${posOfOEMGroupColOnSG}              2
${posOfMainSalesRepColOnSG}          3

*** Keywords ***
Create Table For SG Report
    [Arguments]     ${transType}    ${attribute}    ${year}     ${quarter}
    @{table}        Create List
    ${searchStr}    Set Variable    ${EMPTY}

    IF    '${transType}' == 'REVENUE'
         ${searchStr}   Set Variable    ${year}.Q${quarter} R
    ELSE IF     '${transType}' == 'BACKLOG'
         ${searchStr}   Set Variable    ${year}.Q${quarter} B
    ELSE IF     '${transType}' == 'BACKLOG FORECAST'
         ${searchStr}   Set Variable    ${year}.Q${quarter} BF
    ELSE IF     '${transType}' == 'CUSTOMER FORECAST'
         ${searchStr}   Set Variable    ${year}.Q${quarter} CF
    ELSE IF     '${transType}' == 'BUDGET'
         ${searchStr}   Set Variable    ${year}.Q${quarter} BGT
    ELSE IF     '${transType}' == 'LOS'
        ${searchStr}   Set Variable    ${year}.Q${quarter} R
        ${posOfRCol}    Get Position Of Column    filePath=${SGFilePath}    rowIndex=${rowIndexForSearchPosOfCol}    searchStr=${searchStr}
        IF   '${posOfRCol}' == '0'
            Fail   Not found the position of column
        END
        IF    '${attribute}' == 'QTY'
             ${posOfRCol}   Evaluate    ${posOfRCol}+0
        ELSE IF     '${attribute}' == 'AMOUNT'
             ${posOfRCol}   Evaluate    ${posOfRCol}+2
        ELSE
            Fail    The Attribute parameter ${attribute} is invalid. Please contact with the Administrator for supporting
        END
        ${searchStr}   Set Variable    ${year}.Q${quarter} B
        ${posOfBCol}    Get Position Of Column    filePath=${SGFilePath}    rowIndex=${rowIndexForSearchPosOfCol}    searchStr=${searchStr}
        IF   '${posOfBCol}' == '0'
            Fail   Not found the position of column
        END
        IF    '${attribute}' == 'QTY'
             ${posOfBCol}   Evaluate    ${posOfBCol}+0
        ELSE IF     '${attribute}' == 'AMOUNT'
             ${posOfBCol}   Evaluate    ${posOfBCol}+2
        ELSE
            Fail    The Attribute parameter ${attribute} is invalid. Please contact with the Administrator for supporting
        END
    ELSE
         Fail    The TransType parameter ${transType} is invalid. Please contact with the Administrator for supporting
    END

    IF    '${transType}' != 'LOS'
        ${posOfValueCol}    Get Position Of Column    filePath=${SGFilePath}    rowIndex=${rowIndexForSearchPosOfCol}    searchStr=${searchStr}
        IF    '${posOfValueCol}' == '0'
            Fail   Not found the position of column
        END
        IF    '${attribute}' == 'QTY'
             ${posOfValueCol}   Evaluate    ${posOfValueCol}+0
        ELSE IF     '${attribute}' == 'AMOUNT'
             ${posOfValueCol}   Evaluate    ${posOfValueCol}+2
        ELSE
            Fail    The Attribute parameter ${attribute} is invalid. Please contact with the Administrator for supporting
        END
    END

    File Should Exist    path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${SGFilePath}
    FOR    ${rowIndex}    IN RANGE    ${startRowOnSG}    ${numOfRows}+1
        ${oemGroupCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnSG}
        IF    '${oemGroupCol}' != 'None'
            ${mainSalesRepCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfMainSalesRepColOnSG}
            IF    '${transType}' != 'LOS'
                 ${valueCol}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
            ELSE
                ${valueOfRCol}       Read Excel Cell    row_num=${rowIndex}    col_num=${posOfRCol}
                IF    '${valueOfRCol}' == 'None'
                     ${valueOfRCol}     Set Variable    0
                END
                ${valueOfBCol}       Read Excel Cell    row_num=${rowIndex}    col_num=${posOfBCol}
                IF    '${valueOfBCol}' == 'None'
                     ${valueOfBCol}     Set Variable    0
                END
                ${valueCol}     Evaluate    ${valueOfRCol}+${valueOfBCol}
            END
            IF    '${valueCol}' == 'None'
                 ${valueCol}    Set Variable    0
            END
            ${valueCol}      Evaluate  "%.2f" % ${valueCol}
            IF    '${valueCol}' == '0.00'
                 Continue For Loop
            END
            ${rowOnTable}   Create List
            ...             ${oemGroupCol}
            ...             ${mainSalesRepCol}
            ...             ${valueCol}
            Append To List    ${table}   ${rowOnTable}
        END

    END

    Close Current Excel Document
    [Return]    ${table}

Get Value By OEM Group On SG Report
    [Arguments]     ${oemGroup}     ${transType}    ${year}     ${quarter}  ${attribute}
    ${value}   Set Variable    0

    ${searchStr}        Set Variable    ${year}.Q${quarter} ${transType}
    ${posOfCol}     Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchPosOfCol}    ${searchStr}
    IF    '${attribute}' == 'AMOUNT'
         ${posOfCol}    Evaluate    ${posOfCol}+2
    END

    File Should Exist    path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${SGFilePath}

    FOR    ${rowIndex}    IN RANGE    ${startRowOnSG}    ${numOfRows}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnSG}
        IF    '${oemGroupCol}' == '${oemGroup}'
             ${value}  Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
             IF    '${value}' == 'None'
                  ${value}     Set Variable    0
             END
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${value}

Get Value By Main Sales Rep On SG Report
    [Arguments]     ${mainSalesRep}     ${transType}    ${year}     ${quarter}  ${attribute}
    ${value}   Set Variable    0

    ${searchStr}        Set Variable    ${year}.Q${quarter} ${transType}
    ${posOfCol}     Get Position Of Column    ${SGFilePath}    ${rowIndexForSearchPosOfCol}    ${searchStr}
    IF    '${attribute}' == 'AMOUNT'
         ${posOfCol}    Evaluate    ${posOfCol}+2
    END

    File Should Exist      path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${SGFilePath}

    FOR    ${rowIndex}    IN RANGE    ${startRowOnSG}    ${numOfRows}+1
        ${mainSalesRepCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfMainSalesRepColOnSG}
        ${valueCol}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${valueCol}' == 'None'
             ${valueCol}    Set Variable    0
        END
        IF    '${mainSalesRepCol}' in ${mainSalesRep}
             ${value}   Evaluate    ${value}+${valueCol}
        END
    END

    Close Current Excel Document
    [Return]    ${value}
#Write Data To SS RCD For Pivot
#    [Arguments]     ${ssRCDForPivotFilePath}    ${quarter}  ${oemGroup}  ${pn}  ${tranID}   ${revQty}
#
#    File Should Exist    ${ssRCDForPivotFilePath}
#    Open Excel Document    ${ssRCDForPivotFilePath}    doc_id=SSRCDForPivot
#    Switch Current Excel Document    doc_id=SSRCDForPivot
#    ${latestRowInSSRCDForPivotFile}   Get Number Of Rows In Excel    ${ssRCDForPivotFilePath}
#    ${nextRow}    Evaluate    ${latestRowInSSRCDForPivotFile}+1
#    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${quarter}
#    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oemGroup}
#    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${pn}
#    Write Excel Cell    row_num=${nextRow}    col_num=4    value=${tranID}
#    Write Excel Cell    row_num=${nextRow}    col_num=5    value=${revQty}
#    Save Excel Document    ${ssRCDForPivotFilePath}
#    Close Current Excel Document
#
#Convert SS RCD To Pivot And Export To Excel
#    [Arguments]     ${ssRCDFilePath}    ${ssRCDForPivotFilePath}    ${year}     ${quarter}
#
#    @{table}    Create List
#    @{listParentClass}  Create List     COMPONENTS      MEM     STORAGE     NI ITEMS
#    ${startRow}     Set Variable    2
#
#    ${quarter}  Set Variable    Q${quarter}-${year}
#    File Should Exist    ${ssRCDFilePath}
#    Open Excel Document    ${ssRCDFilePath}    doc_id=SSRCD
#    ${numOfRowsOnSSRCD}  Get Number Of Rows In Excel    ${ssRCDFilePath}
#
#    FOR    ${rowIndexOnSSRCD}    IN RANGE    ${startRow}    ${numOfRowsOnSSRCD}+1
#        ${oemGroupCol}            Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
#        ${parentClassCol}         Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=9
#        ${pnCol}                  Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=11
#        ${quarterCol}             Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
#        ${tranIdCol}              Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=24
#        ${revQtyCol}              Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=29
#
#
#        ${sumREVQty}    Set Variable    ${revQtyCol}
#
#        IF    '${parentClassCol}' in ${listParentClass} and '${quarterCol}' == '${quarter}'
#            ${isDataInTable}    Set Variable    ${False}
#            FOR    ${rowOnTable}    IN    @{table}
#                IF    '${rowOnTable}[0]' == '${oemGroupCol}' and '${rowOnTable}[1]' == '${pnCol}'
#                     ${isDataInTable}   Set Variable    ${True}
#                     BREAK
#                END
#            END
#            IF    '${isDataInTable}' == '${True}'
#                 Continue For Loop
#            END
#            FOR    ${rowIndexTemp}    IN RANGE    ${startRow}+1    ${numOfRowsOnSSRCD}+1
#                      ${oemGroupColTemp}            Read Excel Cell    row_num=${rowIndexTemp}       col_num=2
#                      ${pnColTemp}                  Read Excel Cell    row_num=${rowIndexTemp}       col_num=11
#                      ${quarterColTemp}             Read Excel Cell    row_num=${rowIndexTemp}       col_num=18
#                      ${revQtyColTemp}              Read Excel Cell    row_num=${rowIndexTemp}       col_num=29
#
#                      IF    '${oemGroupColTemp}' == '${oemGroupCol}' and '${pnColTemp}' == '${pnCol}' and '${quarterColTemp}' == '${quarter}'
#                           ${sumREVQty}     Evaluate    ${sumREVQty}+${revQtyColTemp}
#                      END
#            END
#        ELSE
#           Continue For Loop
#        END
#        Log To Console    Row:${rowIndexOnSSRCD}
#        ${rowOnTable}   Create List
#        ...             ${oemGroupCol}
#        ...             ${pnCol}
#        ...             ${tranIdCol}
#        ...             ${sumREVQty}
#        Append To List    ${table}  ${rowOnTable}
#    END
#    Close All Excel Documents
#
#    [Return]  ${table}
#
#
#Write The Test Result Of SG Report By OEM Group To Excel
#    [Arguments]     ${oemGroup}     ${valueOnSGReport}   ${valueOnNS}
#    File Should Exist    ${testResultOfSGReportByOEMGroupFilePath}
#    Open Excel Document    ${testResultOfSGReportByOEMGroupFilePath}    doc_id=SGReportResult
#    ${latestRowInSGReportResultFile}   Get Number Of Rows In Excel    ${testResultOfSGReportByOEMGroupFilePath}
#    ${nextRow}    Evaluate    ${latestRowInSGReportResultFile}+1
#    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${oemGroup}
#    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${valueOnSGReport}
#    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${valueOnNS}
#    Save Excel Document    ${testResultOfSGReportByOEMGroupFilePath}
#    Close Current Excel Document
#
#Check Data For Every Quarter By OEM Group
#    [Arguments]     ${sgFilePath}   ${ssRCDFilePath}     ${year}     ${quarter}     ${attribute}    ${valueType}
#    @{listOfOEMGroupOnSSRCD}    Create List
#    @{listOfOEMGroupOnSG}       Create List
##    ${listOfOEMGroupOnSSRCD}    Get List Of OEM Groups From SS RCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarter}    attribute=${attribute}
#    ${listOfOEMGroupOnSG}       Get List Of OEM Groups From SG        sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}   attribute=${attribute}   valueType=${valueType}
#
##    ${quarterOnSSRCD}  Set Variable    Q${quarter}-${year}
##    FOR    ${oemGroupOnSSRCD}    IN    @{listOfOEMGroupOnSSRCD}
##        ${hasOEMGroupOnSG}  Set Variable    ${False}
##        ${valueOnSSRCD}    Get Value By OEM Group From SSRCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarterOnSSRCD}    oemGroup=${oemGroupOnSSRCD}    valueType=REV
##        ${oemGroupOnSSRCD}  Convert To Upper Case    ${oemGroupOnSSRCD}
##        FOR    ${oemGroupOnSG}    IN    @{listOfOEMGroupOnSG}
##            IF    '${oemGroupOnSSRCD}' == '${oemGroupOnSG}'
##                 ${valueOnSG}   Get Value By OEM Group From SG    sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}    oemGroup=${oemGroupOnSG}    valueType=REV
##                 ${hasOEMGroupOnSG}     Set Variable    ${True}
##                 ${valueOnSG}   Remove String    ${valueOnSG}   $   ,
##                 ${valueOnSSRCD}    Convert To Integer    ${valueOnSSRCD}
##                 ${valueOnSG}       Convert To Integer    ${valueOnSG}
##                 ${diff}    Evaluate    abs(${valueOnSSRCD}-${valueOnSG})
##                 IF    ${diff} > 3
##                      Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSSRCD}    valueOnSGReport=${valueOnSG}    valueOnNS=${valueOnSSRCD}
##                 END
##                 BREAK
##            END
##        END
##        IF    '${hasOEMGroupOnSG}' == '${False}'
##             Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSSRCD}    valueOnSGReport=${EMPTY}    valueOnNS=${valueOnSSRCD}
##        END
##    END
##
##    FOR    ${oemGroupOnSG}    IN    @{listOfOEMGroupOnSG}
##        ${hasOEMGroupOnSSRCD}   Set Variable    ${False}
##        FOR    ${oemGroupOnSSRCD}    IN    @{listOfOEMGroupOnSSRCD}
##            ${oemGroupOnSSRCD}  Convert To Upper Case    ${oemGroupOnSSRCD}
##            IF    '${oemGroupOnSG}' == '${oemGroupOnSSRCD}'
##                 ${hasOEMGroupOnSSRCD}  Set Variable    ${True}
##                 BREAK
##            END
##        END
##        IF    '${hasOEMGroupOnSSRCD}' == '${False}'
##             ${valueOnSG}   Get Value By OEM Group From SG    sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}    oemGroup=${oemGroupOnSG}    valueType=REV
##             Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSG}    valueOnSGReport=${valueOnSG}    valueOnNS=${EMPTY}
##        END
##    END
#
#Get Value By OEM Group From Flat SG
#    [Arguments]     ${flatSGFilePath}    ${year}     ${quarter}   ${oemGroup}    ${attribute}
#    ${value}    Set Variable    0
#
#    File Should Exist    ${flatSGFilePath}
#    Open Excel Document    filename=${flatSGFilePath}    doc_id=FlatSG
#    ${numOfRowsOnFlatSG}     Get Number Of Rows In Excel    ${flatSGFilePath}
#
#    IF    '${attribute}' == 'REVQTY'
#         ${searchStr}    Set Variable    ${year}.Q${quarter} QTY
#    END
#
#    ${rowIndexForSearchStr}     Convert To Number    4
#    ${posOfCol}  Get Position Of Column    ${flatSGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
#
#    FOR    ${rowIndexOnFlatSG}    IN RANGE    5    ${numOfRowsOnFlatSG}+1
#        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnFlatSG}    col_num=1
#        ${valCol}           Read Excel Cell    row_num=${rowIndexOnFlatSG}    col_num=${posOfCol}
##        Log To Console    OEM Group: ${oemGroupCol}; Value:${valCol}
#
#        IF    '${oemGroupCol}' == '${oemGroup}'
#             ${value}   Evaluate    ${value}+${valCol}
#        END
#    END
#
#    Close All Excel Documents
#
#    [Return]    ${value}
#
#Get Value By OEM Group From SS RCD
#    [Arguments]     ${ssRCDFilePath}    ${year}     ${quarter}   ${oemGroup}    ${attribute}
#    ${value}    Set Variable    0
#    @{listParentClass}  Create List     COMPONENTS      MEM     STORAGE     NI ITEMS
#    ${quarter}  Set Variable    Q${quarter}-${year}
#
#    File Should Exist    ${ssRCDFilePath}
#    Open Excel Document    filename=${ssRCDFilePath}    doc_id=SSRCD
#    ${numOfRowsOnSSRCD}     Get Number Of Rows In Excel    ${ssRCDFilePath}
#
#    FOR    ${rowIndexOnSSRCD}    IN RANGE    2    ${numOfRowsOnSSRCD}+1
#        ${parentClassCol}      Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=9
#        ${quarterCol}          Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
#        IF    '${parentClassCol}' in ${listParentClass} and '${quarterCol}' == '${quarter}'
#             ${oemGroupCol}         Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
#             IF    '${attribute}' == 'REVQTY'
#                  ${valCol}           Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=29
#             ELSE
#                  Fail    The value of attribute parameter ${attribute} is invalid. Please contact with Admin!
#             END
#             IF    '${oemGroupCol}' == '${oemGroup}'
#                  ${value}   Evaluate    ${value}+${valCol}
#             END
#        END
#    END
#    Close All Excel Documents
#
#    [Return]    ${value}
#
#Check Parameter Year
#    [Arguments]     ${year}
#    ${result}   Set Variable    ${True}
#
#    [Return]    ${result}
#
#Get List Of OEM Groups From SG
#    [Arguments]     ${sgFilePath}    ${year}     ${quarter}     ${attribute}    ${valueType}
#    @{listOfOEMGroup}   Create List
#
#    ${year}     Convert To Number    ${year}
#    ${year}     Convert To Integer    ${year}
#    ${currentYear}  Get Current Year
#    IF    ${year} < 0
#         Fail   The parameter year ${year} is invalid. It must be Integer number. Please contact with Admin!
#    END
#
#    ${minYear}  Set Variable    2018
#    ${maxYear}  Evaluate    ${currentYear}+1
#    IF    ${year} < ${minYear} or ${year} > ${maxYear}
#         Fail   The parameter year ${year} is invalid. The range of parameter year is between ${minYear} and ${maxYear}. Please contact with Admin!
#    END
#
#    ${quarter}  Convert To Number    ${quarter}
#    ${quarter}  Convert To Integer    ${quarter}
#    IF    ${quarter} < 0
#         Fail   The parameter quarter ${quarter} is invalid. It must be Interger number. Please contact with Admin!
#    END
#    IF    ${quarter} < 1 or ${quarter} > 4
#         Fail   The parameter quarter ${quarter} is invalid. The range of quarter is between 1 and 4. Please contact with Admin!
#    END
#
##    File Should Exist    ${sgFilePath}
##    Open Excel Document    filename=${sgFilePath}    doc_id=SG
##    ${numOfRowsOnSG}     Get Number Of Rows In Excel    ${sgFilePath}
##    ${currentYear}  Get Current Year
#
##    ${searchStr}    Set Variable    ${year}.Q${quarter} R
##    ${rowIndexForSearchStr}     Convert To Number    3
##    ${posOfREVCol}  Get Position Of Column    ${sgFilePath}    ${rowIndexForSearchStr}    ${searchStr}
##    ${posOfREVCol}  Evaluate    ${posOfREVCol}+2
##
##    FOR    ${rownIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
##        ${oemGRoupCol}  Read Excel Cell    row_num=${rownIndexOnSG}    col_num=2
##        ${revCol}       Read Excel Cell    row_num=${rownIndexOnSG}    col_num=${posOfREVCol}
##        IF    '${oemGRoupCol}' != 'None' and '${revCol}' != '${EMPTY}'
##             Append To List    ${listOfOEMGroup}    ${oemGRoupCol}
##        END
##    END
##    Close All Excel Documents
##    [Return]    ${listOfOEMGroup}
#
#Get List Of OEM Groups From SS RCD
#    [Arguments]     ${ssRCDFilePath}    ${year}     ${quarter}   ${attribute}
#    @{listOfOEMGroup}   Create List
#    @{listParentClass}  Create List     COMPONENTS      MEM     STORAGE     NI ITEMS
#
#    ${quarter}  Set Variable    Q${quarter}-${year}
#    File Should Exist    ${ssRCDFilePath}
#    Open Excel Document    filename=${ssRCDFilePath}    doc_id=SSRCD
#    ${numOfRowsOnSSRCD}     Get Number Of Rows In Excel    ${ssRCDFilePath}
#
#    FOR    ${rowIndexOnSSRCD}    IN RANGE    2    ${numOfRowsOnSSRCD}+1
#        ${parrentClassCol}  Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=9
#        ${quarterCol}       Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
#        IF    '${parrentClassCol}' in ${listParentClass} and '${quarterCol}' == '${quarter}'
#             IF    '${attribute}' == 'REVQTY'
#                  ${attributeCol}     Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=29
#             ELSE IF     '${attribute}' == 'REVAMOUNT'
#                  ${attributeCol}     Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=30
#             ELSE
#                Fail    The value of attribute parameter ${attribute} is invalid. Please contact with Admin!
#             END
#             IF    '${attributeCol}' != '0'
#                  ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
#                  Append To List    ${listOfOEMGroup}   ${oemGroupCol}
#             END
#        END
#
#    END
#    ${listOfOEMGroup}     Remove Duplicates    ${listOfOEMGroup}
#    Close All Excel Documents
#
#    [Return]    ${listOfOEMGroup}

