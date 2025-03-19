*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${dwDBFilePath}              C:\\RobotFramework\\Downloads\\Design Win Database.xlsx
${rowIndexForSearchColOnDWDB}    1
${startRowOnDWDB}                2
${posOfOEMGroupColOnDWDB}        1
${posOfPNColOnDWDB}              2
${posOfDWNoColOnDWDB}            4


*** Keywords ***
Check Total Revenue On DWDB Report
    [Arguments]     ${transType}    ${attribute}    ${year}
    ${result}   Set Variable    ${True}

    [Return]    ${result}

Get Total Revenue On DWDB Report
    ${totalRevenue}     Set Variable    0
    File Should Exist    path=${dwDBFilePath}
    Open Excel Document    filename=${dwDBFilePath}    doc_id=DWDB
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${dwDBFilePath}
    FOR    ${rowIndex}    IN RANGE    ${startRowOnDWDB}    ${numOfRows}+1
        ${valueCol}     Read Excel Cell    row_num=${rowIndex}    col_num=
    END
    [Return]    ${totalRevenue}

Get Position Of Revenue Column On DWDB Report
    ${pos}  Set Variable    0

    [Return]    ${pos}

Create Table For DWDB Report With Revenue
    [Arguments]     ${transType}    ${attribute}    ${year}
    @{table}    Create List
    ${searchStr}    Set Variable    ${EMPTY}
    
    IF    '${transType}' == 'REVENUE'
         ${searchStr}   Set Variable    ${year} REVENUE
    ELSE IF  '${transType}' == 'BACKLOG'
        ${searchStr}   Set Variable    ${year} BACKLOG
    ELSE IF  '${transType}' == 'BACKLOG FORECAST'
        ${searchStr}   Set Variable    ${year} BACKLOG FORECAST
    ELSE IF  '${transType}' == 'CUSTOMER FORECAST'
        ${searchStr}   Set Variable    ${year} CUSTOMER FORECAST
    ELSE IF  '${transType}' == 'SALES FORECAST'
        IF    '${attribute}' == 'QTY'
             ${searchStr}   Set Variable    ${year} SALES FORECAST QUANTITY
        ELSE IF  '${attribute}' == 'AMOUNT'
             ${searchStr}   Set Variable    ${year} SALES FORECAST REVENUE
        ELSE
            Fail    The attribute parameter ${attribute} is invalid
        END
    ELSE IF  '${transType}' == 'CUSTOMER FORECAST'
        ${searchStr}   Set Variable    ${year} CUSTOMER FORECAST
    ELSE IF  '${transType}' == 'BUGET'
        ${searchStr}   Set Variable    ${year} BUDGET
    ELSE
        Fail    The transtype parameter ${transType} is invalid
    END
    ${posOfValueCol}     Get Position Of Column    filePath=${dwDBFilePath}    rowIndex=${rowIndexForSearchColOnDWDB}    searchStr=${searchStr}
    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of ${searchStr} column
    END
    File Should Exist    path=${dwDBFilePath}
    Open Excel Document    filename=${dwDBFilePath}    doc_id=DWDB
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${dwDBFilePath}
    FOR    ${rowIndex}    IN RANGE    ${startRowOnDWDB}    ${numOfRows}+1
        ${oemGroupCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnDWDB}
        ${pnCol}         Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnDWDB}
        ${valueCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        IF    '${valueCol}' == 'None' or '${valueCol}' == '0.0'
             Continue For Loop
        END
        ${rowOnTable}   Create List
        ...             ${oemGroupCol}
        ...             ${pnCol}
        ...             ${valueCol}
        Append To List    ${table}   ${rowOnTable}
    END
    
    [Return]    ${table}
    
#Check The Data Of Design Win No Column On DWDB Report
#    File Should Exist    path=${dwDBFilePath}
#    Open Excel Document    filename=${dwDBFilePath}    doc_id=DWDB
#    ${numOfRowsOnDWDB}    Get Number Of Rows In Excel    ${dwDBFilePath}
#
#    @{list}     Create List
#    ${list}     Get List Of OPP JOIN ID On SS Master OPP
#
#    FOR    ${rowIndex}    IN RANGE    ${startRowOnDWDB}    ${numOfRowsOnDWDB}+1
#        Log To Console    rowIndex:${rowIndex}
#        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnDWDB}
#        ${pnCol}            Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnDWDB}
#        ${dwNoCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfDWNoColOnDWDB}
#
#        IF    '${dwNoCol}' == 'None'
#             Continue For Loop
#        END
#        FOR    ${oppJoinID}    IN    @{list}
#            IF    '${dwNoCol}' == '${oppJoinID}'
#                 BREAK
#            END
#
#        END
##        ${dwNoIsExistOnNS}   Check The OPP Join ID Data Is Exist On SS Master OPP By OEM Group And PN    oemGroup=${oemGroupCol}     pn=${pnCol}     oppJoinID=${dwNoCol}
##        IF    '${dwNoIsExistOnNS}' == '${False}'
##             Log To Console    OEM Group:${oemGroupCol}; PN:${pnCol}
##        END
#    END
#    Close Current Excel Document

Check The Duplication Of Design Win No Column
    @{listOfDWsNo}    Create List

    ${listOfDWsNo}              Get List Of DWs No
    ${uniqueListOfDWsNo}        Remove Duplicates    ${listOfDWsNo}
    ${listOfDWsNoLength}        Get Length    ${listOfDWsNo}
    ${uniqueListOfDWsNoLength}  Get Length    ${uniqueListOfDWsNo}
    Should Be Equal    ${listOfDWsNoLength}    ${uniqueListOfDWsNoLength}   The Design Win No column is duplicated

Get List Of DWs No
    @{listOfDWsNo}    Create List

    File Should Exist    path=${dwDBFilePath}
    Open Excel Document    filename=${dwDBFilePath}    doc_id=DWDB
    ${numOfRowsOnDWDB}    Get Number Of Rows In Excel    ${dwDBFilePath}
    
    FOR    ${rowIndex}    IN RANGE    ${startRowOnDWDB}    ${numOfRowsOnDWDB}+1
        ${dwNoCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfDWNoColOnDWDB}
        IF    '${dwNoCol}' == 'None'
             Continue For Loop
        END
        Append To List    ${listOfDWsNo}    ${dwNoCol}       
    END

    Close Current Excel Document
    [Return]    ${listOfDWsNo}
