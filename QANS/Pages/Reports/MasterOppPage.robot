*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

Library    XML
Library    DateTime

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
${chk9_ClosedOppStageOption}                         //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl12']
${chk9_OppDisapprovedOppStageOption}                 //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl13']
${chkNullOfCreatedFromFilter}                        //*[@id='ReportViewerControl_ctl04_ctl05_cbNull']
${chkNullOfCreatedToFilter}                          //*[@id='ReportViewerControl_ctl04_ctl07_cbNull']
${RESULT_FILE_PATH}                       ${OUTPUT_DIR}\\Results\\MasterOpp\\MasterOppResult.xlsx

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
    ${csvFilePath}      Set Variable    ${DOWNLOAD_DIR}\\${fullyFileName}
    ${xlsxFilePath}     Set Variable    ${DOWNLOAD_DIR}\\MasterOppSource.xlsx
    Convert Csv To Xlsx    ${csvFilePath}    ${xlsxFilePath}

Compare Data Between Master Opp Report And SS On NS
    [Arguments]     ${reportFilePath}   ${ssFilePath}

    ${result}   Set Variable    ${True}
#    Verify The Number Of Opps On Master Opp Report                          reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
#    Verify The Document Number Of Opp On Master Opp Report                  reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
    Verify The Data Of Opp With Only One Item On Master Opp Report          reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}

    [Return]    ${result}

Verify The Data Of Opp With Only One Item On Master Opp Report
    [Arguments]     ${reportFilePath}   ${ssFilePath}
    ${result}   Set Variable    ${True}
    @{listOfOppsCheckedOnReport}    Create List
    @{listOfOppsHaveMultiItemsOnNS}     Create List
    ${listOfOppsHaveMultiItemsOnNS}     Get List Of Opps Have Multi Items From The SS Of Master Opp Report On NS    ssFilePath=${ssFilePath}

    File Should Exist    ${ssFilePath}
    File Should Exist    ${reportFilePath}
    Open Excel Document    ${ssFilePath}    MasterOppSource
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssFilePath}
    Open Excel Document    ${reportFilePath}    MasterOppReport
    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}
    Open Excel Document    ${RESULT_FILE_PATH}    doc_id=MasterOppResult

    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
        Switch Current Excel Document    doc_id=MasterOppSource
        ${oppColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
        ${trackedOppColOnSS}                         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=3
        ${oppLinkToColOnSS}                          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=4
        ${oemGroupColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=6
        ${samColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=7
        ${saleRepColOnSS}                            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=8
        ${tmColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
        ${oppDiscoveryPersonColOnSS}                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=10
        ${bizDevSupportColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
        ${pnColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=12
        ${qtyColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=13
        ${projectTotalColOnSS}                       Read Excel Cell    row_num=${rowIndexOnSS}    col_num=14
        ${probColOnSS}                               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=15
        ${currentOppStageColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=16
        ${oppCategoryColOnSS}                        Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
        ${expSampleShipColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=19
        Log To Console    Opp: ${oppColOnSS}
        FOR    ${rowIndexOnReport}    IN RANGE    5    ${numOfRowsOnReport}+1
            Switch Current Excel Document    doc_id=MasterOppReport
            ${oppColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
            ${isOppCheckedOnReport}     Set Variable    ${False}
            FOR    ${oppCheckedOnReport}    IN    @{listOfOppsCheckedOnReport}
                IF    '${oppColOnReport}' == '${oppCheckedOnReport}'
                    ${isOppCheckedOnReport}     Set Variable    ${True}
                    BREAK
                END
            END
            IF    '${isOppCheckedOnReport}' == '${True}'
                 Continue For Loop
            END
            IF    '${oppColOnSS}' != '${oppColOnReport}'
                Continue For Loop
            END
            ${trackedOppColOnReport}                         Read Excel Cell    row_num=${rowIndexOnReport}    col_num=2
            ${oppLinkToColOnReport}                          Read Excel Cell    row_num=${rowIndexOnReport}    col_num=3
            ${oemGroupColOnReport}                           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=5
            ${samColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=7
            ${saleRepColOnReport}                            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=8
            ${tmColOnReport}                                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=9
            ${oppDiscoveryPersonColOnReport}                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=10
            ${bizDevSupportColOnReport}                      Read Excel Cell    row_num=${rowIndexOnReport}    col_num=11
            ${pnColOnReport}                                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=12
            ${qtyColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=13
            ${projectTotalColOnReport}                       Read Excel Cell    row_num=${rowIndexOnReport}    col_num=14
            ${probColOnReport}                               Read Excel Cell    row_num=${rowIndexOnReport}    col_num=15
            ${currentOppStageColOnReport}                    Read Excel Cell    row_num=${rowIndexOnReport}    col_num=16
            ${oppCategoryColOnReport}                        Read Excel Cell    row_num=${rowIndexOnReport}    col_num=18
            ${expSampleShipColOnReport}                      Read Excel Cell    row_num=${rowIndexOnReport}    col_num=19

            IF    '${oppColOnSS}' == '${oppColOnReport}'
                 IF    '${trackedOppColOnSS}' == '${EMPTY}'
                      ${trackedOppColOnSS}      Set Variable    No
                 END
                 IF    '${trackedOppColOnReport}' != '${trackedOppColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=TRACKED OPP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${trackedOppColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${trackedOppColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${oppLinkToColOnReport}' == 'None'
                      ${oppLinkToColOnReport}   Set Variable    ${EMPTY}
                 END
                 IF    '${oppLinkToColOnSS}' == 'None'
                      ${oppLinkToColOnSS}   Set Variable    ${EMPTY}
                 END
                 IF    '${oppLinkToColOnReport}' != '${oppLinkToColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP LINK TO
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppLinkToColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppLinkToColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${oemGroupColOnSS}' == 'PALO ALTO NETWORKS'
                      ${oemGroupColOnSS}    Set Variable    PALOALTO NETWORKS
                 END
                 IF    '${oemGroupColOnReport}' != '${oemGroupColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OEM GROUP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oemGroupColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${samColOnReport}' != '${samColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SAM
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${samColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${samColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${saleRepColOnSS}' == 'Sinclair, Cameron R'
                      ${saleRepColOnSS}     Set Variable    Cameron Sinclair
                 END
                 IF    '${saleRepColOnSS}' == 'Tran, Huan'
                      ${saleRepColOnSS}     Set Variable    Huan Tran
                 END
                 IF    '${saleRepColOnSS}' == 'Nilsson, Michael J'
                      ${saleRepColOnSS}     Set Variable    Michael Nilsson
                 END
                 IF    '${saleRepColOnReport}' != '${saleRepColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SALES REP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${saleRepColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${saleRepColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${tmColOnSS}' == 'Ting, Darren'
                      ${tmColOnSS}   Set Variable     Darren Ting
                 END
                 IF    '${tmColOnSS}' == 'Cook, Christopher'
                      ${tmColOnSS}   Set Variable     Christopher Cook
                 END
                 IF    '${tmColOnSS}' == 'Lawrence, Scott'
                      ${tmColOnSS}   Set Variable     Scott Lawrence
                 END
                 IF    '${tmColOnReport}' != '${tmColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=TECHNICAL MARKETING
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${tmColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${tmColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${oppDiscoveryPersonColOnReport}' == 'None'
                     ${oppDiscoveryPersonColOnReport}       Set Variable    ${EMPTY}
                 END
                 IF    '${oppDiscoveryPersonColOnSS}' == 'None'
                      ${oppDiscoveryPersonColOnSS}  Set Variable    ${EMPTY}
                 END
                 IF    '${oppDiscoveryPersonColOnSS}' == 'Sinclair, Cameron R'
                      ${oppDiscoveryPersonColOnSS}   Set Variable   Cameron Sinclair
                 END
                 IF    '${oppDiscoveryPersonColOnReport}' != '${oppDiscoveryPersonColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP DISCOVERY PERSON
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppDiscoveryPersonColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppDiscoveryPersonColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${bizDevSupportColOnReport}' == 'None'
                      ${bizDevSupportColOnReport}   Set Variable    ${EMPTY}
                 END
                 IF    '${bizDevSupportColOnSS}' == 'None'
                      ${bizDevSupportColOnSS}   Set Variable    ${EMPTY}
                 END
                 IF    '${bizDevSupportColOnReport}' != '${bizDevSupportColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=BIZ DEV SUPPORT
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${bizDevSupportColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${bizDevSupportColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${isOppInListOfOppsHaveMultiItemsOnNS}    Set Variable    ${False}
                 FOR    ${opp}    IN    @{listOfOppsHaveMultiItemsOnNS}
                     IF    '${oppColOnSS}' == '${opp}'
                          ${isOppInListOfOppsHaveMultiItemsOnNS}    Set Variable    ${True}
                          BREAK
                     END

                 END
                 IF    '${isOppInListOfOppsHaveMultiItemsOnNS}' == '${False}'
                      IF    '${pnColOnReport}' != '${pnColOnSS}'
                          Switch Current Excel Document    doc_id=MasterOppResult
                          ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                          ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                          Write Excel Cell    row_num=${nextRow}    col_num=1    value=PART NUMBER
                          Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                          Write Excel Cell    row_num=${nextRow}    col_num=3    value=${pnColOnReport}
                          Write Excel Cell    row_num=${nextRow}    col_num=4    value=${pnColOnSS}
                          Save Excel Document    ${RESULT_FILE_PATH}
                      END
                      IF    '${qtyColOnReport}' != '${qtyColOnSS}'
                          Switch Current Excel Document    doc_id=MasterOppResult
                          ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                          ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                          Write Excel Cell    row_num=${nextRow}    col_num=1    value=QTY PER YR
                          Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                          Write Excel Cell    row_num=${nextRow}    col_num=3    value=${qtyColOnReport}
                          Write Excel Cell    row_num=${nextRow}    col_num=4    value=${qtyColOnSS}
                          Save Excel Document    ${RESULT_FILE_PATH}
                      END
                 END
                 ${diffProjectTotal}    Evaluate    abs(${projectTotalColOnReport}-${projectTotalColOnSS})

                 IF    '${diffProjectTotal}' > '1'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=PROJECT TOTAL
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${projectTotalColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${projectTotalColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${probColOnReport}     Evaluate    ${probColOnReport}*100
                 ${probColOnSS}     Remove String    ${probColOnSS}  %
                 IF    '${probColOnReport}' != '${probColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=PROB
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${probColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${probColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${currentOppStageColOnReport}' != '${currentOppStageColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=CURRENT OPP STAGE
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${currentOppStageColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${currentOppStageColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${oppCategoryColOnReport}' != '${oppCategoryColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=CURRENT OPP STAGE
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${currentOppStageColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${currentOppStageColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${expSampleShipColOnSS}' == 'None'
                      ${expSampleShipColOnSS}   Set Variable    ${EMPTY}
                 END
                 IF    '${expSampleShipColOnReport}' == 'None'
                      ${expSampleShipColOnReport}   Set Variable    ${EMPTY}
                 END
                 IF    '${expSampleShipColOnSS}' != '${EMPTY}'
                      ${expSampleShipColOnSS}        Convert Date    ${expSampleShipColOnSS}         date_format=%m/%d/%Y
                      ${expSampleShipColOnSS}        Convert Date    ${expSampleShipColOnSS}         result_format=%m/%d/%Y
                 END
                 IF    '${expSampleShipColOnReport}' != '${EMPTY}'
                      ${expSampleShipColOnReport}    Convert Date    ${expSampleShipColOnReport}     result_format=%m/%d/%Y
                 END

                 IF    '${expSampleShipColOnReport}' != '${expSampleShipColOnSS}'
                      Switch Current Excel Document    doc_id=MasterOppResult
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=EXP SAMPLE SHIP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${expSampleShipColOnReport}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${expSampleShipColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 Append To List    ${listOfOppsCheckedOnReport}     ${oppColOnSS}
                 BREAK
            END
        END
    END

    [Return]    ${result}

Verify The Document Number Of Opp On Master Opp Report
    [Arguments]     ${reportFilePath}   ${ssFilePath}
    ${result}   Set Variable    ${True}
    @{listOfOppsOnReport}   Create List
    @{listOfOppsOnNS}       Create List

    ${listOfOppsOnReport}   Get List Of Opps From The Master Opp Report    ${reportFilePath}
    ${listOfOppsOnNS}       Get List Of Opps From The SS Of Master Opp Report On NS    ${ssFilePath}

    File Should Exist    ${RESULT_FILE_PATH}
    Open Excel Document    ${RESULT_FILE_PATH}    MasterOppResult

    FOR    ${oppOnNS}    IN    @{listOfOppsOnNS}
        ${posOfOppInReport}     Set Variable    0
        ${numOfOppsOnReport}    Get Length    ${listOfOppsOnReport}

        FOR    ${oppOnReport}    IN    @{listOfOppsOnReport}
            IF    '${oppOnReport}' == '${oppOnNS}'
                 Remove From List    list_=${listOfOppsOnReport}    index=${posOfOppInReport}
                 BREAK
            END
            ${posOfOppInReport}     Evaluate    ${posOfOppInReport}+1
        END
        IF    '${posOfOppInReport}' == '${numOfOppsOnReport}'
             ${result}   Set Variable    ${False}
             ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
             ${nextRow}     Evaluate    ${latestRowInResultFile}+1
             Write Excel Cell    row_num=${nextRow}    col_num=1    value=Document Number
             Write Excel Cell    row_num=${nextRow}    col_num=2    value=${EMPTY}
             Write Excel Cell    row_num=${nextRow}    col_num=3    value=${EMPTY}
             Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppOnNS}
             Save Excel Document    ${RESULT_FILE_PATH}
        END
    END
    Close All Excel Documents

    [Return]    ${result}

Verify The Number Of Opps On Master Opp Report
    [Arguments]     ${reportFilePath}   ${ssFilePath}
    ${result}   Set Variable    ${True}
    @{listOfOppsOnReport}   Create List
    @{listOfOppsOnNS}       Create List

    ${listOfOppsOnReport}   Get List Of Opps From The Master Opp Report    ${reportFilePath}
    ${listOfOppsOnNS}       Get List Of Opps From The SS Of Master Opp Report On NS    ${ssFilePath}
    ${numOfOppsOnReport}    Get Length    ${listOfOppsOnReport}
    ${numOfOppsOnNS}    Get Length    ${listOfOppsOnNS}

    IF    '${numOfOppsOnReport}' != '${numOfOppsOnNS}'
         ${result}      Set Variable    ${False}
         File Should Exist    ${RESULT_FILE_PATH}
         Open Excel Document    ${RESULT_FILE_PATH}    MasterOppResult
         ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
         ${nextRow}     Evaluate    ${latestRowInResultFile}+1
         Write Excel Cell    row_num=${nextRow}    col_num=1    value=Number of Opps
         Write Excel Cell    row_num=${nextRow}    col_num=2    value=${EMPTY}
         Write Excel Cell    row_num=${nextRow}    col_num=3    value=${numOfOppsOnReport}
         Write Excel Cell    row_num=${nextRow}    col_num=4    value=${numOfOppsOnNS}
         Save Excel Document    ${RESULT_FILE_PATH}
    END
    Close All Excel Documents
    [Return]    ${result}

Get List Of Opps From The SS Of Master Opp Report On NS
    [Arguments]     ${ssFilePath}
    @{listOfOpps}   Create List

    File Should Exist    ${ssFilePath}
    Open Excel Document    ${ssFilePath}    MasterOppSource
    ${numOfRows}    Get Number Of Rows In Excel    ${ssFilePath}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRows}+1
        ${opp}  Read Excel Cell    ${rowIndex}    2
        IF    '${opp}' != '${EMPTY}'
             Append To List    ${listOfOpps}     ${opp}
        END
    END
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Close All Excel Documents

    [Return]    ${listOfOpps}

Get List Of Opps From The Master Opp Report
    [Arguments]     ${reportFilePath}
    @{listOfOpps}   Create List

    File Should Exist    ${reportFilePath}
    Open Excel Document    ${reportFilePath}    MasterOppReport
    ${numOfRows}    Get Number Of Rows In Excel    ${reportFilePath}
    FOR    ${rowIndex}    IN RANGE    5    ${numOfRows}+1
        ${opp}  Read Excel Cell    ${rowIndex}    1
        IF    '${opp}' != '${EMPTY}'
             Append To List    ${listOfOpps}     ${opp}
        END
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




