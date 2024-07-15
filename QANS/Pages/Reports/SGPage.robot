*** Settings ***

*** Keywords ***
Compare REV Data On SG Report Between Old Server And New Server
    [Arguments]     ${SGNewServerFilePath}   ${SGOldServerFilePath}     ${year}     ${quarter}
    Log To Console    test