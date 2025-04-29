*** Settings ***
Resource    ../../Pages/Reports/FlatSGPage.robot
Resource    ../../Pages/NS/SaveSearchPage.robot

*** Test Cases ***
Verify Backlog QTY on Flat SG report
    Comparing Data For Every PN Between Flat SG and SS RCD    transType=BACKLOG    attribute=QTY    year=2025    quarter=2    nameOfColOnSSRCD=BL QTY




