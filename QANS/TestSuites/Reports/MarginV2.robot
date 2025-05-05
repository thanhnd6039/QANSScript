*** Settings ***
Resource    ../../Pages/Reports/MarginPageV2.robot
Library   DataDriver   ../../Resources/TestDataForMarginReport.xlsx sheet_name=Sheet1

*** Test Cases ***
Verify Revenue QTY on Margin report
    Comparing Data For Every PN Between Margin And SS RCD    transType=REVENUE    attribute=QTY    year=2025    quarter=2    nameOfColOnSSRCD=REV QTY

