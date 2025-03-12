*** Settings ***
Resource    ../../Pages/Reports/MarginPage.robot

*** Test Cases ***
Testcase1
    Create Table For Margin Report  transType=Revenue   attribute=QTY    year=2025   quarter=1



