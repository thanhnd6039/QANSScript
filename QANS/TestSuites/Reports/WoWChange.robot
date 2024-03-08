*** Settings ***
Resource    ../../Pages/Reports/WoWChangePage.robot

*** Test Cases ***
Validating Data For WoW Chang Report
    ${wowChangeSourceFilePath}    Set Variable    C:\\RobotFramework\\Downloads\\Wow Change [Current Week] Source.xlsx
    ${sgWeeklyActionDBFilePath}   Set Variable    C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB.xlsx
    Compare Data Between WoW Change Report And SG Weekly Action DB Report And SG Report  ${wowChangeSourceFilePath}      ${sgWeeklyActionDBFilePath}


