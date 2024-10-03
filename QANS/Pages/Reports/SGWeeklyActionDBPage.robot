*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
Get Value By OEM Group From Approved Budget
    [Arguments]     ${approvedBudgetFilePath}     ${year}   ${quarter}   ${oemGroup}
    ${value}    Set Variable    0
    
    File Should Exist    ${approvedBudgetFilePath}
    Open Excel Document    filename=${approvedBudgetFilePath}    doc_id=ApprovedBudget
    ${numOfRowsOnApprovedBudget}    Get Number Of Rows In Excel    ${approvedBudgetFilePath}
    
    FOR    ${rowIndexOnApprovedBudget}    IN RANGE    2    ${numOfRowsOnApprovedBudget}+1
        ${oemGroupCol}   Read Excel Cell    row_num=${rowIndexOnApprovedBudget}    col_num=2
        ${oemGroupCol}  Convert To Upper Case    ${oemGroupCol}
        ${yearCol}       Read Excel Cell    row_num=${rowIndexOnApprovedBudget}    col_num=3
        ${quarterCol}    Read Excel Cell    row_num=${rowIndexOnApprovedBudget}    col_num=4
        ${budgetCol}     Read Excel Cell    row_num=${rowIndexOnApprovedBudget}    col_num=5
        IF    '${oemGroupCol}' == '${oemGroup}' and '${yearCol}' == '${year}' and '${quarterCol}' == '${quarter}'
             ${value}   Evaluate    ${value}+${budgetCol}
        END
         
    END
    [Return]    ${value}
    
Get Value By OEM Group From SG Weekly Action DB
    [Arguments]     ${SGWeeklyActionDBFilePath}     ${quarter}   ${oemGroup}    ${attribute}

    ${value}    Set Variable    0
    File Should Exist    ${SGWeeklyActionDBFilePath}
    Open Excel Document    filename=${SGWeeklyActionDBFilePath}    doc_id=SGWeeklyActionDB
    ${numOfRowsOnSGWeeklyActionDB}     Get Number Of Rows In Excel    ${SGWeeklyActionDBFilePath}

    IF    '${attribute}' == 'BGT'
         ${searchStr}    Set Variable    Q${quarter} BGT

    END
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfCol}  Get Position Of Column    ${SGWeeklyActionDBFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
        ${oemGroupCol}  Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
        ${oemGroupCol}  Convert To Upper Case    ${oemGroupCol}
        ${bgtCol}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == '${oemGroup}'
             ${value}   Set Variable    ${bgtCol}
             IF    '${value}' == 'None'
                  ${value}   Set Variable   0
             END
             BREAK
        END
    END

    Close All Excel Documents
    [Return]    ${value}

Get List Of OEM Group From SG Weekly Action DB
    [Arguments]     ${SGWeeklyActionDBFilePath}
    @{listOfOEMGroup}   Create List

    File Should Exist    ${SGWeeklyActionDBFilePath}
    Open Excel Document    filename=${SGWeeklyActionDBFilePath}    doc_id=SGWeeklyActionDB
    ${numOfRowsOnSGWeeklyActionDB}     Get Number Of Rows In Excel    ${SGWeeklyActionDBFilePath}

    FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
        ${oemGroupCol}  Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1  
        IF    '${oemGroupCol}' == 'Total'
             BREAK
        END
        ${oemGroupCol}  Convert To Upper Case    ${oemGroupCol}
        Append To List    ${listOfOEMGroup}     ${oemGroupCol}
    END
    Close All Excel Documents
    [Return]    ${listOfOEMGroup}
