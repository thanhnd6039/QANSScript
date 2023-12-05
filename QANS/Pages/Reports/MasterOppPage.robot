*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

Library    XML

*** Variables ***
${txtTitleOfMasterOpp}                  //*[contains(text(),'Master')]
${lstOppStageFilter}                    //*[@id='ReportViewerControl_ctl04_ctl29_txtValue']
${chkSelectAllOfOppStageOption}         //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl00']
${chk0_IdentifyOfOppStageOption}        //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl02']
${chk1_OppApprovedOppStageOption}       //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl03']
${chk2_EvalSubmittedOppStageOption}     //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl04']
${chk3_QualIssuesOppStageOption}        //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl05']
${chk4_QualApprovedOppStageOption}        //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl06']
${chk5_FirstProductionPOOppStageOption}        //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl07']
${chk6_ProductionOppStageOption}               //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl08']
${chk7_HoldOppStageOption}                     //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl09']
${chk8_LostOppStageOption}                     //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl10']
${chk9_CancelledOppStageOption}                //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl11']
${chk9_ClosedOppStageOption}                   //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl12']
${chk9_OppDisapprovedOppStageOption}           //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl13']
${chkNullOfCreatedFromFilter}                        //*[@id='ReportViewerControl_ctl04_ctl05_cbNull']
${chkNullOfCreatedToFilter}                          //*[@id='ReportViewerControl_ctl04_ctl07_cbNull']

*** Keywords ***
Navigate To Master Opp Report
    ${configFileObject}     Load Json From File    ${CONFIG_FILE}
    ${username}             Get Value From Json    ${configFileObject}    $.accounts[0].username
    ${username}             Set Variable           ${username}[0]
    ${pass}                 Get Value From Json    ${configFileObject}    $.accounts[0].password
    ${pass}                 Set Variable           ${pass}[0]
    ${url}                  Set Variable           http://${username}:${pass}@report/ReportServer/Pages/ReportViewer.aspx?/NetSuite+Reports/Sales/Opportunity+Report&rs:Command=Render
    Go To    ${url}

Should See The Title Of Master Opp Report
    [Arguments]     ${title}
    Wait Until Element Is Visible    ${txtTitleOfMasterOpp}     ${TIMEOUT}
    Element Text Should Be    ${txtTitleOfMasterOpp}    ${title}
    
Select All Opp Stages On Master Opp Report
    Wait Until Element Is Visible    ${lstOppStageFilter}   ${TIMEOUT}
    Click Element    ${lstOppStageFilter}
    Wait Until Element Is Visible    ${chkSelectAllOfOppStageOption}    ${TIMEOUT}
    ${isCheckSelectAll}     Run Keyword And Return Status    Checkbox Should Be Selected    ${chkSelectAllOfOppStageOption}
    IF    '${isCheckSelectAll}' == '${False}'
         Click Element    ${chkSelectAllOfOppStageOption}
    END

Select Opp Stage On Master Opp Report
    [Arguments]     ${multiOppStageOptions}
    Wait Until Element Is Visible    ${lstOppStageFilter}       ${TIMEOUT}
    Click Element    ${lstOppStageFilter}
    Wait Until Element Is Visible    ${chkSelectAllOfOppStageOption}    ${TIMEOUT}
    ${isCheckSelectAll}     Run Keyword And Return Status    Checkbox Should Be Selected    ${chkSelectAllOfOppStageOption}
    IF    '${isCheckSelectAll}' == '${False}'
        Click Element    ${chkSelectAllOfOppStageOption}
        Click Element    ${chkSelectAllOfOppStageOption}
    ELSE
        Click Element    ${chkSelectAllOfOppStageOption}
    END
    
    FOR    ${oppStage}    IN    @{multiOppStageOptions}
        IF    '${oppStage}' == '0.Identified'
             Wait Until Element Is Visible    ${chk0_IdentifyOfOppStageOption}      ${TIMEOUT}
             Click Element    ${chk0_IdentifyOfOppStageOption}
        END
        IF    '${oppStage}' == '1.Opp Approved'
             Wait Until Element Is Visible    ${chk1_OppApprovedOppStageOption}      ${TIMEOUT}
             Click Element    ${chk1_OppApprovedOppStageOption}
        END
        IF    '${oppStage}' == '2.Eval Submitted/Qual in Progress'
             Wait Until Element Is Visible    ${chk2_EvalSubmittedOppStageOption}      ${TIMEOUT}
             Click Element    ${chk2_EvalSubmittedOppStageOption}
        END
        IF    '${oppStage}' == '3.Qual Issues'
             Wait Until Element Is Visible    ${chk3_QualIssuesOppStageOption}      ${TIMEOUT}
             Click Element    ${chk3_QualIssuesOppStageOption}
        END
        IF    '${oppStage}' == '4.Qual Approved'
             Wait Until Element Is Visible    ${chk4_QualApprovedOppStageOption}      ${TIMEOUT}
             Click Element    ${chk4_QualApprovedOppStageOption}
        END
        IF    '${oppStage}' == '5.First - Production PO'
             Wait Until Element Is Visible    ${chk5_FirstProductionPOOppStageOption}      ${TIMEOUT}
             Click Element    ${chk5_FirstProductionPOOppStageOption}
        END
        IF    '${oppStage}' == '6.Production'
             Wait Until Element Is Visible    ${chk6_ProductionOppStageOption}      ${TIMEOUT}
             Click Element    ${chk6_ProductionOppStageOption}
        END
        IF    '${oppStage}' == '7.Hold'
             Wait Until Element Is Visible    ${chk7_HoldOppStageOption}      ${TIMEOUT}
             Click Element    ${chk7_HoldOppStageOption}
        END
        IF    '${oppStage}' == '8.Lost'
             Wait Until Element Is Visible    ${chk8_LostOppStageOption}      ${TIMEOUT}
             Click Element    ${chk8_LostOppStageOption}
        END
        IF    '${oppStage}' == '9.Cancelled'
             Wait Until Element Is Visible    ${chk9_CancelledOppStageOption}      ${TIMEOUT}
             Click Element    ${chk9_CancelledOppStageOption}
        END
        IF    '${oppStage}' == '9.Closed'
             Wait Until Element Is Visible    ${chk9_ClosedOppStageOption}      ${TIMEOUT}
             Click Element    ${chk9_ClosedOppStageOption}
        END
        IF    '${oppStage}' == '9.Opp Disapproved'
             Wait Until Element Is Visible    ${chk9_OppDisapprovedOppStageOption}      ${TIMEOUT}
             Click Element    ${chk9_OppDisapprovedOppStageOption}
        END
    END

Filter Created Date On Master Opp Report
    [Arguments]     ${createdFrom}      ${createdTo}
    IF    '${createdFrom}' == 'NULL'
         Wait Until Element Is Visible    ${chkNullOfCreatedFromFilter}     ${TIMEOUT}
         ${isCheckCheckboxNullOfCreatedFromFilter}  Run Keyword And Return Status    Checkbox Should Be Selected    ${chkNullOfCreatedFromFilter}
         IF    '${isCheckCheckboxNullOfCreatedFromFilter}' == '${False}'
              Click Element    ${chkNullOfCreatedFromFilter}
         END
    END
    
    IF    '${createdTo}' == 'NULL'
         Wait Until Element Is Visible    ${chkNullOfCreatedToFilter}   ${TIMEOUT}
         ${isCheckCheckboxNullOfCreatedToFilter}  Run Keyword And Return Status    Checkbox Should Be Selected    ${chkNullOfCreatedToFilter}
         IF    '${isCheckCheckboxNullOfCreatedToFilter}' == '${False}'
              Click Element    ${chkNullOfCreatedToFilter}
         END
    END

Navigate To The Save Search Of Master Opp Report On NS
    ${url}      Set Variable    https://4499123.app.netsuite.com/app/common/search/searchresults.nl?searchid=4002&whence=
    Login To NS With Account    PRODUCTION
    Go To    ${url}

Export Excel Data From The Save Search Of Master Opp Report On NS
    Export SS Data To CSV
    Sleep    5s
    ${fullyFileName}    Get Fully File Name From Given Name    MasterOpps    ${DOWNLOAD_DIR}
    ${csvFilePath}      Set Variable    ${DOWNLOAD_DIR}${fullyFileName}
    ${xlsxFilePath}     Set Variable    ${DOWNLOAD_DIR}MasterOppSource.xlsx
    Convert Csv To Xlsx    ${csvFilePath}    ${xlsxFilePath}

Compare Data Between Master Opp Report And SS On NS
    [Arguments]     ${reportFilePath}   ${ssFilePath}

    ${resultFilePath}       Set Variable    ${OUTPUT_DIR}\\Results\\MasterOpp\\MasterOppResult.xlsx
    @{listOfOppsOnReport}   Create List
    @{listOfOppsOnNS}       Create List

    ${listOfOppsOnNS}       Get List Of Opps From The SS Of Master Opp Report On NS    ${ssFilePath}
    ${listOfOppsOnReport}   Get List Of Opps From The Master Opp Report    ${reportFilePath}

    ${numOfOppsOnNS}    Get Length    ${listOfOppsOnNS}
    ${numOfOppsOnReport}    Get Length    ${listOfOppsOnReport}
    Log To Console    Number of Opps on Report: ${numOfOppsOnReport}
    Log To Console    Number of Opps On NS: ${numOfOppsOnNS}
    File Should Exist    ${resultFilePath}
    Open Excel Document    ${resultFilePath}    MasterOppResult
    IF    '${numOfOppsOnReport}' != '${numOfOppsOnNS}'
        ${latestRowInResultFile}   Get Number Of Rows In Excel    ${resultFilePath}
        ${nextRow}     Evaluate    ${latestRowInResultFile}+1
        Write Excel Cell    row_num=${nextRow}    col_num=2    value=Number of Opps
        Write Excel Cell    row_num=${nextRow}    col_num=3    value=${numOfOppsOnReport}
        Write Excel Cell    row_num=${nextRow}    col_num=4    value=${numOfOppsOnNS}
        Save Excel Document    ${resultFilePath}
    END

    Open Excel Document    ${ssFilePath}    MasterOppSource
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssFilePath}
    Open Excel Document    ${reportFilePath}    MasterOppReport
    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}
    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}
        Switch Current Excel Document    doc_id=MasterOppSource
        ${oppColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
        ${trackedOppColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=3
        ${oppLinkColOnSS}                       Read Excel Cell    row_num=${rowIndexOnSS}    col_num=4
        ${oemGroupColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=6
        ${samColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=7
        ${saleRepColOnSS}                       Read Excel Cell    row_num=${rowIndexOnSS}    col_num=8
        ${tmColOnSS}                            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
        ${oppDiscoveryPersonColOnSS}            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=10
        ${bizDevSupportColOnSS}                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
        ${pnColOnSS}                            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=12
        ${qtyColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=13
        ${projectTotalColOnSS}                  Read Excel Cell    row_num=${rowIndexOnSS}    col_num=14
        ${probabilityColOnSS}                   Read Excel Cell    row_num=${rowIndexOnSS}    col_num=15
        ${oppStageColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=16
        ${oppCategoryColOnSS}                   Read Excel Cell    row_num=${rowIndexOnSS}    col_num=17
        ${expSampleShipDateColOnSS}             Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
        ${expQualApprovedDateColOnSS}           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=19
        ${expDWDateColOnSS}                     Read Excel Cell    row_num=${rowIndexOnSS}    col_num=20
        ${1PPODateColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=21
        ${DWDateColOnSS}                        Read Excel Cell    row_num=${rowIndexOnSS}    col_num=22
        ${DWStatusColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=23
        ${customerPNColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=24
        ${subSegmentColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=25
        ${programNameColOnSS}                   Read Excel Cell    row_num=${rowIndexOnSS}    col_num=26
        ${applicationColOnSS}                   Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
        ${functionColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=28
        ${countRowOnReport}     Set Variable    4
        Log To Console    OPP: ${oppColOnSS}
        FOR    ${rowIndexOnReport}    IN RANGE    5    ${numOfRowsOnReport}+1
            Switch Current Excel Document    doc_id=MasterOppReport            
            ${oppColOnReport}               Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
            IF    '${oppColOnReport}' == '${oppColOnSS}'
                 BREAK
            END
            ${countRowOnReport}     Evaluate    ${countRowOnReport}+1
        END
        Log To Console    countRow: ${countRowOnReport}
        Log To Console    Number of Rows On Report: ${numOfRowsOnReport}
        IF    '${countRowOnReport}' == '${numOfRowsOnReport}'
             Switch Current Excel Document    doc_id=MasterOppResult
             ${latestRowInResultFile}   Get Number Of Rows In Excel    ${resultFilePath}
             ${nextRow}     Evaluate    ${latestRowInResultFile}+1
             Write Excel Cell    row_num=${nextRow}    col_num=1    value=${oppColOnSS}
             Write Excel Cell    row_num=${nextRow}    col_num=2    value=OPP
             Write Excel Cell    row_num=${nextRow}    col_num=3    value=${EMPTY}
             Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppColOnSS}
             Save Excel Document    ${resultFilePath}
        END
    END
    

Get List Of Opps From The SS Of Master Opp Report On NS
    [Arguments]     ${ssFilePath}
    @{listOfOpps}   Create List

    Open Excel Document    ${ssFilePath}    MasterOppSource
    ${numOfRows}    Get Number Of Rows In Excel    ${ssFilePath}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRows}
        ${opp}  Read Excel Cell    ${rowIndex}    2
        Append To List    ${listOfOpps}     ${opp}
    END
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Close All Excel Documents
    [Return]    ${listOfOpps}

Get List Of Opps From The Master Opp Report
    [Arguments]     ${reportFilePath}

    @{listOfOpps}   Create List

    Open Excel Document    ${reportFilePath}    MasterOppReport
    ${numOfRows}    Get Number Of Rows In Excel    ${reportFilePath}
    FOR    ${rowIndex}    IN RANGE    5    ${numOfRows}+1
        ${opp}  Read Excel Cell    ${rowIndex}    1
        Append To List    ${listOfOpps}     ${opp}
    END
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Close All Excel Documents
    [Return]    ${listOfOpps}

Get List Of Opps Have Multi Items From The Master Opp Report
    [Arguments]     ${reportFilePath}
    @{listOfOpps}   Create List
    
    Open Excel Document    ${reportFilePath}    MasterOppReport
    ${numOfRows}    Get Number Of Rows In Excel    ${reportFilePath}
    FOR    ${rowIndex}    IN RANGE    5    ${numOfRows}+1
        ${currentOpp}   Read Excel Cell    ${rowIndex}    1
        ${nextRow}  Evaluate    ${rowIndex}+1
        ${nextOpp}      Read Excel Cell    ${nextRow}    1
        IF    '${nextOpp}' == '${currentOpp}'
             Append To List    ${listOfOpps}    ${currentOpp}
        END
    END
    Close All Excel Documents
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}

    [Return]    ${listOfOpps}

Get List Of Opps Have Multi Items From The SS Of Master Opp Report On NS
    [Arguments]     ${ssFilePath}
    @{listOfOpps}   Create List

    Open Excel Document    ${ssFilePath}    MasterOppSource
    ${numOfRows}    Get Number Of Rows In Excel    ${ssFilePath}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRows}
        ${currentOpp}   Read Excel Cell    ${rowIndex}    2
        ${nextRow}  Evaluate    ${rowIndex}+1
        ${nextOpp}      Read Excel Cell    ${nextRow}    2
        IF    '${nextOpp}' == '${currentOpp}'
             Append To List    ${listOfOpps}    ${currentOpp}
        END
    END
    Close All Excel Documents
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}

    [Return]    ${listOfOpps}




