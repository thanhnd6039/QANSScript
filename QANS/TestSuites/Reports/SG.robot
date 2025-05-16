*** Settings ***
#Suite Setup     Setup Test Environment For SG Report    browser=firefox
Resource    ../../Pages/Reports/SGPageV2.robot

*** Test Cases ***
Verify QTY Revenue on SG Report
    Comparing Data For Every PN Between SG And SS RCD    transType=REVENUE    attribute=QTY    year=2025    quarter=2    nameOfColOnSSRCD=REV QTY



    




