*** Settings ***
Resource    ../../Pages/Reports/MasterOppPage.robot
Resource    ../../Pages/NS/LoginPage.robot
Suite Setup     Initialize Suite

*** Test Cases ***
Verify the number of OPPs on Master OPP Report
    [Tags]  MasterOPP_0001
    Check The Number Of OPPs On Master OPP Report
Verify the Line ID data of OPP on Master OPP Report
    [Tags]  MasterOPP_0002
    Check The Line ID Data On Master OPP Report
Verify the LOS data of OPP on Master OPP Report
    [Tags]  MasterOPP_0003

#Verify the SF data of OPP on Master OPP Report

#Validating The Detailed Data Of Master Opp Report
#    Remove All Files in Specified Directory    ${DOWNLOAD_DIR}\\
#    Setup    Chrome
#    Navigate To Master Opp Report
#    Wait Until Page Load Completed
#    Should See The Title Of Master Opp Report    Master Opportunity Report
#    Filter Created Date On Master Opp Report    NULL    NULL
#    Sleep    10s
#    Select All Opp Stages On Master Opp Report
#    ${multiOppStageOptions}     Create List
#    Append To List        ${multiOppStageOptions}   0.Identified
#    Append To List        ${multiOppStageOptions}   1.Opp Approved
#    Append To List        ${multiOppStageOptions}   2.Eval Submitted/Qual in Progress
#    Append To List        ${multiOppStageOptions}   3.Qual Issues
#    Append To List        ${multiOppStageOptions}   4.Qual Approved
#    Append To List        ${multiOppStageOptions}   5.First - Production PO
#    Append To List        ${multiOppStageOptions}   6.Production
#    Select Opp Stage On Master Opp Report    ${multiOppStageOptions}
#    Click On Button View Report
#    Sleep    10s
#    Should See The Title Of Master Opp Report    Master Opportunity Report
#    Export Report Data To    Excel
#    Sleep    30s
#    File Should Exist    ${DOWNLOAD_DIR}\\Opportunity Report.xlsx
#    Navigate To The Save Search Of Master Opp Report On NS
#    The Title Of Save Search Should Contain    Master Opps
#    Sleep    5s
#    Export Excel Data From The Save Search Of Master Opp Report On NS
#    Sleep    30s
#    File Should Exist    ${DOWNLOAD_DIR}\\MasterOppSource.xlsx
#    TearDown
#    Compare Data Between Master Opp Report And SS On NS     ${DOWNLOAD_DIR}\\Opportunity Report.xlsx      ${DOWNLOAD_DIR}\\MasterOppSource.xlsx


    




























    