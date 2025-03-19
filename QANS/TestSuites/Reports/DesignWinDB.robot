*** Settings ***
Resource    ../../Pages/Reports/DesignWinDBPage.robot


*** Test Cases ***
Verify Sales Forecast on Design Win DB report
    Get List OEM GROUP And PN For Every Quarter    year=2025    quarter=1
#    ${filePath}     Set Variable    C:\\RobotFramework\\Downloads\\test.xlsx
#    @{listNameOfCols}   Create List
#    Append To List    ${listNameOfCols}  OEM GROUP
#    Append To List    ${listNameOfCols}  PN
#    Append To List    ${listNameOfCols}  VALUE
#    ${table}    Create Table For DWDB Report With Revenue    transType=REVENUE    attribute=AMOUNT    year=2018
#    Write Table To Excel    filePath=${filePath}    listNameOfCols=${listNameOfCols}    table=${table}


#Verify the data of Design Win No column
#    Check The Duplication Of Design Win No Column
#    Check The Data Of Design Win No Column On DWDB Report
    
#Verify the data of Stage column
#    Log To Console    Verify the data of Design Win No column