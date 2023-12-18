*** Settings ***
Resource    ../../Pages/Reports/MasterOppPage.robot
Resource    ../../Pages/NS/LoginPage.robot





*** Test Cases ***
#Validating Detailed Data For Master Opp Report
#    Remove All Files in Specified Directory    ${DOWNLOAD_DIR}\\
#    Setup    Chrome
#    Navigate To Master Opp Report
#    Wait Until Page Load Completed
#    Should See The Title Of Master Opp Report    Master Opportunity Report
##    Filter Created Date On Master Opp Report    NULL    NULL
#    Sleep    10s
#    Select All Opp Stages On Master Opp Report
##    ${multiOppStageOptions}     Create List
##    Append To List        ${multiOppStageOptions}   0.Identified
##    Append To List        ${multiOppStageOptions}   1.Opp Approved
##    Append To List        ${multiOppStageOptions}   2.Eval Submitted/Qual in Progress
##    Append To List        ${multiOppStageOptions}   3.Qual Issues
##    Append To List        ${multiOppStageOptions}   4.Qual Approved
##    Append To List        ${multiOppStageOptions}   5.First - Production PO
##    Append To List        ${multiOppStageOptions}   6.Production
##    Select Opp Stage On Master Opp Report    ${multiOppStageOptions}
#    Click On Button View Report
#    Sleep    10s
#    Should See The Title Of Master Opp Report    Master Opportunity Report
#    Export Report Data To    Excel
#    Sleep    10s
#    File Should Exist    ${DOWNLOAD_DIR}\\Opportunity Report.xlsx
#    Navigate To The Save Search Of Master Opp Report On NS
#    The Title Of Save Search Should Contain    Master Opps
#    Export Excel Data From The Save Search Of Master Opp Report On NS
#    Sleep    5s
#    File Should Exist    ${DOWNLOAD_DIR}\\MasterOppSource.xlsx
#    Compare Data Between Master Opp Report And SS On NS     ${DOWNLOAD_DIR}\\Opportunity Report.xlsx      ${DOWNLOAD_DIR}\\MasterOppSource.xlsx
#    TearDown
#    @{listOfItems}      Create List
#    Append To List    ${listOfItems}    ID   Name
#    Append To List    ${listOfItems}    1    Thanh
#    Log To Console    ID: ${listOfItems[0][1]}
#    ${array} =    Create 2D Array    3    4
##    ${value} =    Get Value From 2D Array    ${array}    1    2
##    Insert Data    ${array}    1    2    NewValue
#    Log To Console    Data: ${array}
Create Two-Dimensional Array
    @{two_dimensional_array}    Create List
    ${row_1}    Create List    1    2    3
    ${row_2}    Create List    4    5    6
    ${row_3}    Create List    7    8    9
    Append To List    ${two_dimensional_array}    ${row_1}
    Append To List    ${two_dimensional_array}    ${row_2}
    Append To List    ${two_dimensional_array}    ${row_3}

    Log To Console    Two-Dimensional Array: ${two_dimensional_array}

    ${new_row}    Create List    10    11    12
    Append To List    ${two_dimensional_array}    ${new_row}
    Log To Console    Modified Array: ${two_dimensional_array}
    ${value}    Set Variable    ${two_dimensional_array}[3][0]
    Log To Console    Value at index (1, 1): ${value}


























    