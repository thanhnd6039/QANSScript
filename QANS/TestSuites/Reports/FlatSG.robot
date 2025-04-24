*** Settings ***
Resource    ../../Pages/Reports/FlatSGPage.robot
Resource    ../../Pages/NS/SaveSearchPage.robot

*** Test Cases ***
Verify REV QTY for every quarter by OEM Group
    Comparing Data For Every PN Between Flat SG and SS RCD    transType=BACKLOG    attribute=QTY    year=2025    quarter=2    nameOfColOnSSRCD=BL QTY
#    ${table}    Create Table For SS Revenue Cost Dump    nameOfCol=BL QTY    year=2025    quarter=2
#    ${filePath}     Set Variable    C:\\RobotFramework\\Downloads\\test.xlsx
#    @{listNameOfCols}   Create List
#    Append To List    ${listNameOfCols}  OEM GROUP
#    Append To List    ${listNameOfCols}  PN
#    Append To List    ${listNameOfCols}  VALUE
#    Write Table To Excel    filePath=${filePath}    listNameOfCols=${listNameOfCols}    table=${table}



