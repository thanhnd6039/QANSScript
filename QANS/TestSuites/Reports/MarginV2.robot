*** Settings ***
Resource    ../../Pages/Reports/MarginPageV2.robot

*** Test Cases ***
Verify QTY on Margin report for every quarter
#    ${totalQTYOnMargin}   Get Total Value On Margin Report    transType=REVENUE    attribute=REV     year=2025    quarter=1
#    ${table}    Get All Transactions On SS RCD For Every Quarter    nameOfCol=REVQTY    year=2025    quarter=1
    ${table}    Create Table For SS Revenue Cost Dump   nameOfCol=REVQTY    year=2025    quarter=1
#    ${table}    Create Table For Margin Report    transType=REVENUE    attribute=COST    year=2024    quarter=4
    ${filePath}     Set Variable    C:\\RobotFramework\\Downloads\\test.xlsx
    @{listNameOfCols}   Create List
    Append To List    ${listNameOfCols}     OEM GROUP
    Append To List    ${listNameOfCols}     PN
    Append To List    ${listNameOfCols}     VALUE
    Write Table To Excel    filePath=${filePath}    listNameOfCols=${listNameOfCols}    table=${table}


    
#Verify REV on Margin report for every quarter
#Verify COST on Margin report for every quarter
#Verify % margin on Margin report for every quarter
#Verify AVG MM on Margin report for every quarter
