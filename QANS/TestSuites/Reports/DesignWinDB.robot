*** Settings ***
Resource    ../../Pages/Reports/DesignWinDBPage.robot


*** Test Cases ***
Verify the data of Design Win No column
#    Check The Duplication Of Design Win No Column
    Check The Data Of Design Win No Column On DWDB Report
    
#Verify the data of Stage column
#    Log To Console    Verify the data of Design Win No column