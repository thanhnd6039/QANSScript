*** Settings ***
Resource    ../../Pages/Reports/MasterOppPage.robot
Resource    ../../Pages/NS/LoginPage.robot

*** Test Cases ***
Testcase1
    Setup    Firefox
#    Navigate To Master Opp Report
#    Wait Until Page Load Completed
#    Should See The Title Of Master Opp Report    Master Opportunity Report
#    Filter Created Date On Master Opp Report    NULL    NULL
#    Sleep    3s
#    ${multiOppStageOptions}     Create List
#    Append To List        ${multiOppStageOptions}   0.Identified
#    Append To List        ${multiOppStageOptions}   1.Opp Approved
#    Append To List        ${multiOppStageOptions}   2.Eval Submitted/Qual in Progress
#    Append To List        ${multiOppStageOptions}   3.Qual Issues
#    Append To List        ${multiOppStageOptions}   4.Qual Approved
#    Append To List        ${multiOppStageOptions}   5.First - Production PO
#    Append To List        ${multiOppStageOptions}   6.Production
#    Append To List        ${multiOppStageOptions}   7.Hold
#    Append To List        ${multiOppStageOptions}   8.Lost
#    Append To List        ${multiOppStageOptions}   9.Cancelled
#    Append To List        ${multiOppStageOptions}   9.Closed
#    Append To List        ${multiOppStageOptions}   9.Opp Disapproved
#    Select Opp Stage On Master Opp Report    ${multiOppStageOptions}
#    Click On Button View Report
#    Should See The Title Of Master Opp Report    Master Opportunity Report
#    Export Report To    Excel

    Login To NS With Account    PRODUCTION






    