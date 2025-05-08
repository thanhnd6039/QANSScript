*** Settings ***
Resource    ../../Pages/Reports/FlatSGPage.robot
Resource    ../../Pages/NS/SaveSearchPage.robot

*** Test Cases ***
Verify Backlog QTY on Flat SG report
    ${path}     Set Variable    C:\\RobotFramework\\Downloads\\test.xlsx
    @{listOfColForHeader}   Create List
    Append To List    ${listOfColForHeader}     OEM GROUP
    Append To List    ${listOfColForHeader}     PN
    Append To List    ${listOfColForHeader}     VALUE
    ${table}    Create Table For SS Revenue Cost Dump    nameOfCol=BL FC QTY    year=2025    quarter=2
    Write Table To Excel    filePath=${path}    listNameOfCols=${listOfColForHeader}    table=${table}

#    Comparing Data For Every PN Between Flat SG and SS RCD    transType=BACKLOG    attribute=QTY    year=2025    quarter=2    nameOfColOnSSRCD=BL QTY




