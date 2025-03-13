*** Settings ***
Resource    ../../Pages/Reports/MarginPage.robot

*** Test Cases ***
Testcase1
#    @{table}    Create List
#    ${table}    Create Table For Margin Report  transType=CUSTOMER FORECAST   attribute=QTY    year=2025   quarter=1
#
#    ${filePath}     Set Variable    C:\\RobotFramework\\Downloads\\Margin.xlsx
#    @{listNameOfCols}   Create List
#    Append To List    ${listNameOfCols}     OEM GROUP
#    Append To List    ${listNameOfCols}     PN
#    Append To List    ${listNameOfCols}     VALUE
#    Write Table To Excel    filePath=${filePath}    listNameOfCols=${listNameOfCols}    table=${table}



