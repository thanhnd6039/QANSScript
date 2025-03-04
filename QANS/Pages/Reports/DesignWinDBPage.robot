*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${dwDBFilePath}              C:\\RobotFramework\\Downloads\\Design Win Database.xlsx
${startRowOnDWDB}            2
${posOfDWNoCol}              4


*** Keywords ***
Check The Duplication Of Design Win No Column
    @{listOfDWsNo}    Create List

    ${listOfDWsNo}          Get List Of DWs No
    ${uniqueListOfDWsNo}    Remove Duplicates    ${listOfDWsNo}
    ${listOfDWsNoLength}    Get Length    ${listOfDWsNo}
    ${uniqueListOfDWsNoLength}  Get Length    ${uniqueListOfDWsNo}
    Log To Console    listOfDWsNoLength:${listOfDWsNoLength}
    Log To Console    uniqueListOfDWsNoLength: ${uniqueListOfDWsNoLength}



Get List Of DWs No
    @{listOfDWsNo}    Create List

    File Should Exist    path=${dwDBFilePath}
    Open Excel Document    filename=${dwDBFilePath}    doc_id=DWDB
    ${numOfRowsOnDWDB}    Get Number Of Rows In Excel    ${dwDBFilePath}
    
    FOR    ${rowIndex}    IN RANGE    ${startRowOnDWDB}    ${numOfRowsOnDWDB}+1
        ${dwNoCol}      Read Excel Cell    row_num=${rowIndex}    col_num=${posOfDWNoCol}
        IF    '${dwNoCol}' == 'None'
             Continue For Loop
        END
        Append To List    ${listOfDWsNo}    ${dwNoCol}       
    END

    [Return]    ${listOfDWsNo}
