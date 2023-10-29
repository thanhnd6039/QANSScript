*** Settings ***
Resource    ../../Pages/Reports/MasterOppPage.robot

*** Test Cases ***
Testcase1
    Setup    Firefox
    Navigate To Master Opp Report
    Select Opp Stage On Master Opp Report    0.Identified

    