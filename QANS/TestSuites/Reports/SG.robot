*** Settings ***
Resource    ../../Pages/Reports/SGPage.robot

*** Variables ***
${SGNewServerFilePath}      C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${SGOldServerFilePath}      C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx

*** Test Cases ***
Validating data
    Compare REV Data On SG Report Between Old Server And New Server     ${SGNewServerFilePath}   ${SGOldServerFilePath}     2024    1

