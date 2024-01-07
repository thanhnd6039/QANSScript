*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

*** Keywords ***
Create Table From The SS Of Margin Report On NS
    [Arguments]     ${ssFilePath}    ${type}    ${year}     ${quarter}
    @{table}    Create List

    File Should Exist    ${ssFilePath}
    Open Excel Document    ${ssFilePath}    MarginReportSource
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssFilePath}
    
    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
        ${quarterColOnSS}          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
        IF    '${quarterColOnSS}' == 'Q${quarter}-${year}'
            ${parentClassColOnSS}      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
            IF    '${parentClassColOnSS}' == 'MEM' or '${parentClassColOnSS}' == 'STORAGE' or '${parentClassColOnSS}' == 'COMPONENTS' or '${parentClassColOnSS}' == 'NI'
                 ${oemGroupColOnSS}         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
                 ${pnColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
                 IF    '${type}' == 'R'
                      ${qtyColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
                      ${revColOnSS}               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=30
                 END
            END


        END
#        ${parentClassColOnSS}      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
        

    END
    [Return]    ${table}
    
Create Table For Margin Report
    [Arguments]     ${reportFilePath}   ${type}
    @{table}    Create List
    
    File Should Exist    ${reportFilePath}
    Open Excel Document    ${reportFilePath}    MarginReport
    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}
    ${oemGroupColOnReportTemp}  Set Variable    ${EMPTY}

    FOR    ${rowIndexOnReport}    IN RANGE    7    ${numOfRowsOnReport}+1
        ${oemGroupColOnReport}      Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
        IF    '${oemGroupColOnReport}' == 'None'
             ${oemGroupColOnReport}     Set Variable    ${oemGroupColOnReportTemp}
        ELSE
             ${oemGroupColOnReportTemp}     Set Variable    ${oemGroupColOnReport}
        END
        ${pnColOnReport}             Read Excel Cell    row_num=${rowIndexOnReport}    col_num=2
        IF    '${type}' == 'R'
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=5
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=6
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=7
        END
        IF    '${type}' == 'B'
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=10
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=11
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=12
        END
        IF    '${type}' == 'CF'
             ${qtyColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=15
             ${revColOnReport}            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=16
             ${costColOnReport}           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=17
        END

        ${rowOnTable}   Create List
        ...             ${oemGroupColOnReport}
        ...             ${pnColOnReport}
        ...             ${qtyColOnReport}
        ...             ${revColOnReport}
        ...             ${costColOnReport}
        Append To List    ${table}   ${rowOnTable}      
    END
    [Return]    ${table}
    
Compare Data Between Margin Report And SS On NS
    [Arguments]     ${reportFilePath}   ${ssFilePath}
    ${result}   Set Variable    ${True}
    @{reportTable}       Create List
    @{ssTable}           Create List
    ${type}     Set Variable    R

#    ${reportTable}  Create Table For Margin Report    reportFilePath=${reportFilePath}    type=${type}
#    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
    ${ssTable}  Create Table From The SS Of Margin Report On NS    ssFilePath=${ssFilePath}    type=${type}     year=2024   quarter=1

    

    
    [Return]    ${result}