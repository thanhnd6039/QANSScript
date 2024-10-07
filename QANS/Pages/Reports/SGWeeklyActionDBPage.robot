*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
Create Flat SG As Table Pivot
    [Arguments]     ${flatSGFilePath}   ${year}     ${quarter}   ${attribute}
    @{table}    Create List
    ${listOfOEMGroupFromFlatSG}     Get List Of OEM Group From Flat SG    flatSGFilePath=${flatSGFilePath}
    File Should Exist    path=${flatSGFilePath}
    Open Excel Document    filename=${flatSGFilePath}    doc_id=FlatSG
    ${numOfRowsOnFlatSG}   Get Number Of Rows In Excel    filePath=${flatSGFilePath}

    IF    '${attribute}' == 'R'
         ${searchStr}    Set Variable    ${year}.Q${quarter} R
    ELSE IF  '${attribute}' == 'B'
         ${searchStr}    Set Variable    ${year}.Q${quarter} B
    ELSE IF  '${attribute}' == 'BF'
         ${searchStr}    Set Variable    ${year}.Q${quarter} BF
    ELSE IF  '${attribute}' == 'CF'
         ${searchStr}    Set Variable    ${year}.Q${quarter} CF
    ELSE
        Fail    The value of attribute parameter ${attribute} is invalid. Please contact with Admin!
    END
    ${rowIndexForSearchStr}     Convert To Number    4
    ${posOfCol}  Get Position Of Column    ${flatSGFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    
    FOR    ${oemGroup}    IN    @{listOfOEMGroupFromFlatSG}
        ${sumValueOfAttribute}  Set Variable    0
        FOR    ${rowIndexOnFlatSG}    IN RANGE    5    ${numOfRowsOnFlatSG}+1
            ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnFlatSG}    col_num=1
            ${valueOfAttribute}     Read Excel Cell    row_num=${rowIndexOnFlatSG}    col_num=${posOfCol}
            IF    '${valueOfAttribute}' == '${EMPTY}'
                 Continue For Loop
            END
            ${valueOfAttribute}     Remove String   ${valueOfAttribute}     $   ,

            IF    '${oemGroupCol}' == '${oemGroup}'
                 ${sumValueOfAttribute}     Evaluate    ${sumValueOfAttribute}+${valueOfAttribute}
            END
        END
        IF    '${sumValueOfAttribute}' == '0'
             Continue For Loop
        END
        ${rowOnTable}   Create List
        ...             ${oemGroup}
        ...             ${sumValueOfAttribute}
        Append To List    ${table}   ${rowOnTable}

    END
    
    [Return]    ${table}
    
Create SS Approved Budget As Table Pivot
    [Arguments]     ${ssApprovedBudgetFilePath}     ${year}     ${quarter}
    @{table}    Create List
    ${listOfOEMGroupFromSSApprovedBudget}   Get List Of OEM Group From SS Approved Budget    ssApprovedBudgetFilePath=${ssApprovedBudgetFilePath}   year=${year}    quarter=${quarter}
   
    File Should Exist    path=${ssApprovedBudgetFilePath}
    Open Excel Document    filename=${ssApprovedBudgetFilePath}    doc_id=SSApprovedBudget
    ${numOfRowsOnSSApprovedBudget}    Get Number Of Rows In Excel    ${ssApprovedBudgetFilePath}

    FOR    ${oemGroup}    IN    @{listOfOEMGroupFromSSApprovedBudget}
        ${sumBudget}    Set Variable    0
        FOR    ${rowIndexOnSSApprovedBudget}    IN RANGE    2    ${numOfRowsOnSSApprovedBudget}+1
            ${oemGroupCol}         Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=2
            ${yearCol}             Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=3
            ${quarterCol}          Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=4
            ${budgetCol}           Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=5
            IF    '${oemGroupCol}' == '${oemGroup}' and '${yearCol}' == '${year}' and '${quarterCol}' == '${quarter}'
                 ${sumBudget}   Evaluate    ${sumBudget}+${budgetCol}
            END
        END
        IF    '${sumBudget}' == '0'
             Continue For Loop
        END
        ${rowOnTable}   Create List
        ...             ${oemGroup}        
        ...             ${sumBudget}
        Append To List    ${table}   ${rowOnTable}

    END

    [Return]    ${table}

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

Get List Of OEM Group From SS Approved Budget
    [Arguments]     ${ssApprovedBudgetFilePath}     ${year}     ${quarter}
    @{listOfOEMGroup}   Create List

    File Should Exist    path=${ssApprovedBudgetFilePath}
    Open Excel Document    filename=${ssApprovedBudgetFilePath}    doc_id=SSApprovedBudget
    ${numOfRowsOnSSApprovedBudget}     Get Number Of Rows In Excel    ${ssApprovedBudgetFilePath}

    FOR    ${rowIndexOnSSApprovedBudget}    IN RANGE    2    ${numOfRowsOnSSApprovedBudget}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=2
        ${yearCol}         Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=3
        ${quarterCol}      Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=4
        ${budgetCol}       Read Excel Cell    row_num=${rowIndexOnSSApprovedBudget}    col_num=5
        IF    '${oemGroupCol}' == 'Overall Total'
             BREAK
        END
        IF    '${yearCol}' == '${year}' and '${quarterCol}' == '${quarter}' and '${budgetCol}' != '0'
             Append To List    ${listOfOEMGroup}     ${oemGroupCol}
        END

    END
    Close All Excel Documents
    
    ${listOfOEMGroup}   Remove Duplicates    ${listOfOEMGroup}
       
    [Return]    ${listOfOEMGroup}

Get List Of OEM Group From Flat SG
    [Arguments]     ${flatSGFilePath}
    @{listOfOEMGroup}   Create List

    File Should Exist    path=${flatSGFilePath}
    Open Excel Document    filename=${flatSGFilePath}    doc_id=FlatSG
    ${numOfRowsOnFlatSG}     Get Number Of Rows In Excel    ${flatSGFilePath}

    FOR    ${rowIndexOnFlatSG}    IN RANGE    5    ${numOfRowsOnFlatSG}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndexOnFlatSG}    col_num=1
        Append To List    ${listOfOEMGroup}     ${oemGroupCol}
    END
    Close All Excel Documents

    ${listOfOEMGroup}   Remove Duplicates    ${listOfOEMGroup}

    [Return]    ${listOfOEMGroup}
