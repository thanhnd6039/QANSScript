*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Variables ***
${dwDBFilePath}              C:\\RobotFramework\\Downloads\\Design Win Database.xlsx
${startRowOnDWDB}                2
${posOfOEMGroupColOnDWDB}        1
${posOfPNColOnDWDB}              2
${posOfDWNoColOnDWDB}            4


*** Keywords ***
Check The Data Of Design Win No Column On DWDB Report
    File Should Exist    path=${dwDBFilePath}
    Open Excel Document    filename=${dwDBFilePath}    doc_id=DWDB
    ${numOfRowsOnDWDB}    Get Number Of Rows In Excel    ${dwDBFilePath}
    
    @{list}     Create List
    ${list}     Get List Of OPP JOIN ID On SS Master OPP

    FOR    ${rowIndex}    IN RANGE    ${startRowOnDWDB}    ${numOfRowsOnDWDB}+1
        Log To Console    rowIndex:${rowIndex}
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnDWDB}
        ${pnCol}            Read Excel Cell    row_num=${rowIndex}    col_num=${posOfPNColOnDWDB}
        ${dwNoCol}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfDWNoColOnDWDB}

        IF    '${dwNoCol}' == 'None'
             Continue For Loop
        END
        FOR    ${oppJoinID}    IN    @{list}
            IF    '${dwNoCol}' == '${oppJoinID}'
                 BREAK
            END
             
        END
#        ${dwNoIsExistOnNS}   Check The OPP Join ID Data Is Exist On SS Master OPP By OEM Group And PN    oemGroup=${oemGroupCol}     pn=${pnCol}     oppJoinID=${dwNoCol}
#        IF    '${dwNoIsExistOnNS}' == '${False}'
#             Log To Console    OEM Group:${oemGroupCol}; PN:${pnCol}
#        END
    END
    Close Current Excel Document



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
