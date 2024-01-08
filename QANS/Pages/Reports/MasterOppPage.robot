*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot


*** Variables ***
${txtTitleOfMasterOpp}                               //*[contains(text(),'Master')]
${lstOppStageFilter}                                 //*[@id='ReportViewerControl_ctl04_ctl29_txtValue']
${chkSelectAllOfOppStageOption}                      //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl00']
${chk0_IdentifyOfOppStageOption}                     //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl02']
${chk1_OppApprovedOppStageOption}                    //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl03']
${chk2_EvalSubmittedOppStageOption}                  //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl04']
${chk3_QualIssuesOppStageOption}                     //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl05']
${chk4_QualApprovedOppStageOption}                   //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl06']
${chk5_FirstProductionPOOppStageOption}              //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl07']
${chk6_ProductionOppStageOption}                     //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl08']
${chk7_HoldOppStageOption}                           //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl09']
${chk8_LostOppStageOption}                           //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl10']
${chk9_CancelledOppStageOption}                      //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl11']
${chk9_ClosedOppStageOption}                         //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl12']
${chk9_OppDisapprovedOppStageOption}                 //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl13']
${chkNullOfCreatedFromFilter}                        //*[@id='ReportViewerControl_ctl04_ctl05_cbNull']
${chkNullOfCreatedToFilter}                          //*[@id='ReportViewerControl_ctl04_ctl07_cbNull']
${RESULT_FILE_PATH}                                  ${RESULT_DIR}\\MasterOpp\\MasterOppResult.xlsx

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

    ${verifyNumOfOPPs}                          Verify The Number Of Opps On Master Opp Report                          reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
    ${verifyDocumentNumberOfOPP}                Verify The Document Number Of Opp On Master Opp Report                  reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
    ${verifyDetailedDataOfOPPWithOnlyOneItem}   Verify The Data Of Opp With Only One Item On Master Opp Report          reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
#    ${verifyOPPsHaveMultiItems}                 Verify The OPPs Have Multi Items On Master Opp Report    reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}

#    IF    '${verifyNumOfOPPs}' == '${False}' or '${verifyDocumentNumberOfOPP}' == '${False}' or '${verifyDetailedDataOfOPPWithOnlyOneItem}' == '${False}' or '${verifyOPPsHaveMultiItems}' == '${False}'
#         ${result}  Set Variable    ${False}
#         Fail   The data betwwen Master Opp Report and NS is difference
#    END
    IF    '${verifyNumOfOPPs}' == '${False}' or '${verifyDocumentNumberOfOPP}' == '${False}' or '${verifyDetailedDataOfOPPWithOnlyOneItem}' == '${False}'
         ${result}  Set Variable    ${False}
         Fail   The data betwwen Master Opp Report and NS is difference
    END

    [Return]    ${result}

Verify The OPPs Have Multi Items On Master Opp Report
    [Arguments]     ${reportFilePath}   ${ssFilePath}
    @{listOfOppsHaveMultiItemsOnNS}     Create List
    ${result}   Set Variable    ${True}
    
    ${listOfOppsHaveMultiItemsOnNS}     Get List Of Opps Have Multi Items From The SS Of Master Opp Report On NS    ssFilePath=${ssFilePath}

    File Should Exist    ${ssFilePath}
    Open Excel Document    ${ssFilePath}    doc_id=MasterOppSource
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssFilePath}

    File Should Exist    ${reportFilePath}
    Open Excel Document    ${reportFilePath}    doc_id=MasterOppReport
    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}

    File Should Exist    ${RESULT_FILE_PATH}
    Open Excel Document    ${RESULT_FILE_PATH}    doc_id=MasterOppResult


    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
            ${isFound}      Set Variable    ${False}
            Switch Current Excel Document    MasterOppSource
            ${oppColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=1
            ${pnColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
            ${qtyColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=12
            ${isOppInListOfOppsHaveMultiItemsOnNS}    Set Variable    ${False}
            
             FOR    ${opp}    IN    @{listOfOppsHaveMultiItemsOnNS}
                  IF    '${oppColOnSS}' == '${opp}'
                       ${isOppInListOfOppsHaveMultiItemsOnNS}    Set Variable    ${True}
                       BREAK
                  END
             END
             IF    '${isOppInListOfOppsHaveMultiItemsOnNS}' == '${True}'
                  Switch Current Excel Document    MasterOppReport
                  FOR    ${rowIndexOnReport}    IN RANGE    5    ${numOfRowsOnReport}+1
                       ${oppColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
                       ${pnColOnReport}                                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=12
                       ${qtyColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=13
                       IF    '${oppColOnSS}' == '${oppColOnReport}' and '${pnColOnSS}' == '${pnColOnReport}' and '${qtyColOnSS}' == '${qtyColOnReport}'
                            ${isFound}  Set Variable    ${True}
                            BREAK
                       END
                  END
                  IF    '${isFound}' == '${False}'
                      Switch Current Excel Document    MasterOppResult
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=MULTI ITEMS
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${EMPTY}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${pnColOnSS}
                      Save Excel Document    ${RESULT_FILE_PATH}
                  END
             END

    END

    [Return]    ${result}

Verify The Data Of Opp With Only One Item On Master Opp Report
    [Arguments]     ${reportFilePath}   ${ssFilePath}
    ${result}                           Set Variable    ${True}
    @{reportTable}                      Create List
    @{ssTable}                          Create List
    @{listOfOppsHaveMultiItemsOnNS}     Create List

    ${listOfOppsHaveMultiItemsOnNS}     Get List Of Opps Have Multi Items From The SS Of Master Opp Report On NS    ssFilePath=${ssFilePath}
    ${reportTable}                      Create Table For Master Opp Report    ${reportFilePath}
    ${ssTable}                          Create Table From The SS Of Master Opp Report On NS    ${ssFilePath}

    ${reportTable}  Sort Table By Column    ${reportTable}    0
    ${ssTable}      Sort Table By Column    ${ssTable}        0
    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
    ${numOfRowsOnSSTable}       Get Length    ${ssTable}

    Open Excel Document    ${RESULT_FILE_PATH}    doc_id=MasterOppResult
    ${previousOpp}  Set Variable    ${EMPTY}
    ${rowIndexOnReportTableTemp}     Set Variable    0
    FOR    ${rowIndexOnSSTable}    IN RANGE    0    ${numOfRowsOnSSTable}
        ${oppColOnSSTable}          Set Variable        ${ssTable}[${rowIndexOnSSTable}][0]
        IF    '${oppColOnSSTable}' == '${previousOpp}'
             Continue For Loop
        END
        FOR    ${rowIndexOnReportTable}    IN RANGE    ${rowIndexOnReportTableTemp}    ${numOfRowsOnReportTable}
            ${oppColOnReportTable}  Set Variable   ${reportTable}[${rowIndexOnReportTable}][0]
            IF    '${oppColOnReportTable}' == '${oppColOnSSTable}'
                 ${trackedOppColOnSSTable}                   Set Variable        ${ssTable}[${rowIndexOnSSTable}][1]
                 ${oppLinkToColOnSSTable}                    Set Variable        ${ssTable}[${rowIndexOnSSTable}][2]
                 ${oemGroupColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][3]
                 ${samColOnSSTable}                          Set Variable        ${ssTable}[${rowIndexOnSSTable}][4]
                 ${saleRepColOnSSTable}                      Set Variable        ${ssTable}[${rowIndexOnSSTable}][5]
                 ${tmColOnSSTable}                           Set Variable        ${ssTable}[${rowIndexOnSSTable}][6]
                 ${oppDiscoveryPersonColOnSSTable}           Set Variable        ${ssTable}[${rowIndexOnSSTable}][7]
                 ${bizDevSupportColOnSSTable}                Set Variable        ${ssTable}[${rowIndexOnSSTable}][8]
                 ${pnColOnSSTable}                           Set Variable        ${ssTable}[${rowIndexOnSSTable}][9]
                 ${qtyColOnSSTable}                          Set Variable        ${ssTable}[${rowIndexOnSSTable}][10]
                 ${projectTotalColOnSSTable}                 Set Variable        ${ssTable}[${rowIndexOnSSTable}][11]
                 ${probColOnSSTable}                         Set Variable        ${ssTable}[${rowIndexOnSSTable}][12]
                 ${oppStageColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][13]
                 ${oppCategoryColOnSSTable}                  Set Variable        ${ssTable}[${rowIndexOnSSTable}][14]
                 ${expSampleShipColOnSSTable}                Set Variable        ${ssTable}[${rowIndexOnSSTable}][15]
                 ${expQualApprovedColOnSSTable}              Set Variable        ${ssTable}[${rowIndexOnSSTable}][16]
                 ${expDWDateColOnSSTable}                    Set Variable        ${ssTable}[${rowIndexOnSSTable}][17]
                 ${1PPODateColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][18]
                 ${DWDateColOnSSTable}                       Set Variable        ${ssTable}[${rowIndexOnSSTable}][19]
                 ${DWColOnSSTable}                           Set Variable        ${ssTable}[${rowIndexOnSSTable}][20]
                 ${customerPNColOnSSTable}                   Set Variable        ${ssTable}[${rowIndexOnSSTable}][21]
                 ${subSegmentColOnSSTable}                   Set Variable        ${ssTable}[${rowIndexOnSSTable}][22]
                 ${programColOnSSTable}                      Set Variable        ${ssTable}[${rowIndexOnSSTable}][23]
                 ${applicationColOnSSTable}                  Set Variable        ${ssTable}[${rowIndexOnSSTable}][24]
                 ${functionColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][25]

                 ${trackedOppColOnReportTable}                   Set Variable    ${reportTable}[${rowIndexOnReportTable}][1]
                 ${oppLinkToColOnReportTable}                    Set Variable    ${reportTable}[${rowIndexOnReportTable}][2]
                 ${oemGroupColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][3]
                 ${samColOnReportTable}                          Set Variable    ${reportTable}[${rowIndexOnReportTable}][4]
                 ${saleRepColOnReportTable}                      Set Variable    ${reportTable}[${rowIndexOnReportTable}][5]
                 ${tmColOnReportTable}                           Set Variable    ${reportTable}[${rowIndexOnReportTable}][6]
                 ${oppDiscoveryPersonColOnReportTable}           Set Variable    ${reportTable}[${rowIndexOnReportTable}][7]
                 ${bizDevSupportColOnReportTable}                Set Variable    ${reportTable}[${rowIndexOnReportTable}][8]
                 ${pnColOnReportTable}                           Set Variable    ${reportTable}[${rowIndexOnReportTable}][9]
                 ${qtyColOnReportTable}                          Set Variable    ${reportTable}[${rowIndexOnReportTable}][10]
                 ${projectTotalColOnReportTable}                 Set Variable    ${reportTable}[${rowIndexOnReportTable}][11]
                 ${probColOnReportTable}                         Set Variable    ${reportTable}[${rowIndexOnReportTable}][12]
                 ${oppStageColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][13]
                 ${oppCategoryColOnReportTable}                  Set Variable    ${reportTable}[${rowIndexOnReportTable}][14]
                 ${expSampleShipColOnReportTable}                Set Variable    ${reportTable}[${rowIndexOnReportTable}][15]
                 ${expQualApprovedColOnReportTable}              Set Variable    ${reportTable}[${rowIndexOnReportTable}][16]
                 ${expDWDateColOnReportTable}                    Set Variable    ${reportTable}[${rowIndexOnReportTable}][17]
                 ${1PPODateColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][18]
                 ${DWDateColOnReportTable}                       Set Variable    ${reportTable}[${rowIndexOnReportTable}][19]
                 ${DWColOnReportTable}                           Set Variable    ${reportTable}[${rowIndexOnReportTable}][20]
                 ${customerPNColOnReportTable}                   Set Variable    ${reportTable}[${rowIndexOnReportTable}][21]
                 ${subSegmentColOnReportTable}                   Set Variable    ${reportTable}[${rowIndexOnReportTable}][22]
                 ${programColOnReportTable}                      Set Variable    ${reportTable}[${rowIndexOnReportTable}][23]
                 ${applicationColOnReportTable}                  Set Variable    ${reportTable}[${rowIndexOnReportTable}][24]
                 ${functionColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][25]

                 IF    '${trackedOppColOnReportTable}' != '${trackedOppColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=TRACKED OPP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${trackedOppColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${trackedOppColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 
                 IF    '${oppLinkToColOnReportTable}' == 'None'
                      ${oppLinkToColOnReportTable}     Set Variable     ${EMPTY}
                 END
                 IF    '${oppLinkToColOnSSTable}' == '- None -'
                      ${oppLinkToColOnSSTable}      Set Variable    ${EMPTY}
                 END
                 IF    '${oppLinkToColOnReportTable}' != '${oppLinkToColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP LINK TO
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppLinkToColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppLinkToColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${isOEMGroupOnSSTableContainsColon}  Set Variable    ${False}
                 ${isOEMGroupOnSSTableContainsColon}  Evaluate   ":" in """${oemGroupColOnSSTable}"""
                 IF    '${isOEMGroupOnSSTableContainsColon}' == '${True}'
                        ${strArrTemp}   Split String    ${oemGroupColOnSSTable}  :
                        ${oemGroupColOnSSTable}    Set Variable    ${strArrTemp}[1]
                        ${oemGroupColOnSSTable}    Set Variable    ${oemGroupColOnSSTable.strip()}
                 END
                 IF    '${oemGroupColOnReportTable}' != '${oemGroupColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OEM GROUP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oemGroupColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${samColOnReportTable}' != '${samColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SAM
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${samColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${samColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${isSaleRepOnSSTableContainsComma}  Set Variable    ${False}
                 ${isSaleRepOnSSTableContainsComma}  Evaluate   "," in """${saleRepColOnSSTable}"""
                 IF    '${isSaleRepOnSSTableContainsComma}' == '${True}'
                        ${strArrTemp}   Split String    ${saleRepColOnSSTable}
                        ${saleRepColOnSSTable}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
                        ${saleRepColOnSSTable}  Remove String    ${saleRepColOnSSTable}     ,
                        ${saleRepColOnSSTable}    Set Variable    ${saleRepColOnSSTable.strip()}
                 END
                 IF    '${saleRepColOnReportTable}' != '${saleRepColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SALES REP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${saleRepColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${saleRepColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${tmColOnSSTable}' == '- None -'
                      ${tmColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 ${isTMOnSSTableContainsComma}  Set Variable    ${False}
                 ${isTMOnSSTableContainsComma}  Evaluate   "," in """${tmColOnSSTable}"""
                 IF    '${isTMOnSSTableContainsComma}' == '${True}'
                        ${strArrTemp}   Split String    ${tmColOnSSTable}
                        ${tmColOnSSTable}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
                        ${tmColOnSSTable}  Remove String    ${tmColOnSSTable}     ,
                        ${tmColOnSSTable}    Set Variable    ${tmColOnSSTable.strip()}
                 END
                 ${tmColOnReportTable}    Set Variable    ${tmColOnReportTable.strip()}
                 IF    '${tmColOnReportTable}' != '${tmColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=TECHNICAL MARKETING
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${tmColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${tmColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END

                 IF    '${oppDiscoveryPersonColOnSSTable}' == '- None -'
                      ${oppDiscoveryPersonColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 ${isOPPDiscoveryPersonOnSSTableContainsComma}  Set Variable    ${False}
                 ${isOPPDiscoveryPersonOnSSTableContainsComma}  Evaluate   "," in """${oppDiscoveryPersonColOnSSTable}"""
                 IF    '${isOPPDiscoveryPersonOnSSTableContainsComma}' == '${True}'
                        ${strArrTemp}   Split String    ${oppDiscoveryPersonColOnSSTable}
                        ${oppDiscoveryPersonColOnSSTable}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
                        ${oppDiscoveryPersonColOnSSTable}  Remove String    ${oppDiscoveryPersonColOnSSTable}     ,
                        ${oppDiscoveryPersonColOnSSTable}    Set Variable    ${oppDiscoveryPersonColOnSSTable.strip()}
                 END
                 ${oppDiscoveryPersonColOnReportTable}   Set Variable   ${oppDiscoveryPersonColOnReportTable.strip()}
                 IF    '${oppDiscoveryPersonColOnReportTable}' != '${oppDiscoveryPersonColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP DISCOVERY PERSON
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppDiscoveryPersonColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppDiscoveryPersonColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${isOppInListOfOppsHaveMultiItemsOnNS}    Set Variable    ${False}
                 FOR    ${opp}    IN    @{listOfOppsHaveMultiItemsOnNS}
                     IF    '${oppColOnSSTable}' == '${opp}'
                          ${isOppInListOfOppsHaveMultiItemsOnNS}    Set Variable    ${True}
                          BREAK
                     END
                 END

                 IF    '${isOppInListOfOppsHaveMultiItemsOnNS}' == '${False}'
                      IF    '${pnColOnReportTable}' != '${pnColOnSSTable}'
                          ${result}   Set Variable      ${False}
                          ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                          ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                          Write Excel Cell    row_num=${nextRow}    col_num=1    value=PART NUMBER
                          Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                          Write Excel Cell    row_num=${nextRow}    col_num=3    value=${pnColOnReportTable}
                          Write Excel Cell    row_num=${nextRow}    col_num=4    value=${pnColOnSSTable}
                          Save Excel Document    ${RESULT_FILE_PATH}
                      END
                      IF    '${qtyColOnReportTable}' != '${qtyColOnSSTable}'
                          ${result}   Set Variable      ${False}
                          ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                          ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                          Write Excel Cell    row_num=${nextRow}    col_num=1    value=QTY PER YR
                          Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                          Write Excel Cell    row_num=${nextRow}    col_num=3    value=${qtyColOnReportTable}
                          Write Excel Cell    row_num=${nextRow}    col_num=4    value=${qtyColOnSSTable}
                          Save Excel Document    ${RESULT_FILE_PATH}
                      END
                 END
                 IF    '${bizDevSupportColOnSSTable}' == '- None -'
                      ${bizDevSupportColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 ${isBizDevSupportContainsComma}  Set Variable    ${False}
                 ${isBizDevSupportContainsComma}  Evaluate   "," in """${bizDevSupportColOnSSTable}"""
                 IF    '${isBizDevSupportContainsComma}' == '${True}'
                        ${strArrTemp}   Split String    ${bizDevSupportColOnSSTable}
                        ${bizDevSupportColOnSSTable}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
                        ${bizDevSupportColOnSSTable}  Remove String    ${bizDevSupportColOnSSTable}     ,
                        ${bizDevSupportColOnSSTable}    Set Variable    ${bizDevSupportColOnSSTable.strip()}
                 END
                 ${bizDevSupportColOnReportTable}       Set Variable    ${bizDevSupportColOnReportTable.strip()}
                 IF    '${bizDevSupportColOnReportTable}' != '${bizDevSupportColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=BIZ DEV SUPPORT
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${bizDevSupportColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${bizDevSupportColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${diffProjectTotal}    Evaluate    abs(${projectTotalColOnReportTable}-${projectTotalColOnSSTable})
                 IF    '${diffProjectTotal}' > '1'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=PROJECT TOTAL
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${projectTotalColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${projectTotalColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${probColOnSSTable}  Remove String    ${probColOnSSTable}  %
                 ${probColOnSSTable}  Convert To Number    ${probColOnSSTable}
                 ${probColOnReportTable}    Evaluate    ${probColOnReportTable}*100
                 ${diffProb}    Evaluate    abs(${probColOnReportTable}-${probColOnSSTable})
                 IF    '${diffProb}' >= '1'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=PROB
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${probColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${probColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${oppStageColOnReportTable}' != '${oppStageColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=CURRENT OPP STAGE
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppStageColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppStageColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${oppCategoryColOnReportTable}' != '${oppCategoryColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP CATEGORY
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppCategoryColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppCategoryColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${expSampleShipColOnSSTable}' != 'None'
                      ${expSampleShipColOnSSTable}        Convert Date    ${expSampleShipColOnSSTable}         date_format=%m/%d/%Y
                      ${expSampleShipColOnSSTable}        Convert Date    ${expSampleShipColOnSSTable}         result_format=%m/%d/%Y
                 END
                 IF    '${expSampleShipColOnReportTable}' != 'None'
                      ${expSampleShipColOnReportTable}    Convert Date    ${expSampleShipColOnReportTable}     result_format=%m/%d/%Y
                 END
                 IF    '${expSampleShipColOnReportTable}' != '${expSampleShipColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=EXP SAMPLE SHIP
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${expSampleShipColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${expSampleShipColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${expQualApprovedColOnSSTable}' != 'None'
                      ${expQualApprovedColOnSSTable}        Convert Date    ${expQualApprovedColOnSSTable}         date_format=%m/%d/%Y
                      ${expQualApprovedColOnSSTable}        Convert Date    ${expQualApprovedColOnSSTable}         result_format=%m/%d/%Y
                 END
                 IF    '${expQualApprovedColOnReportTable}' != 'None'
                      ${expQualApprovedColOnReportTable}    Convert Date    ${expQualApprovedColOnReportTable}     result_format=%m/%d/%Y
                 END
                 IF    '${expQualApprovedColOnReportTable}' != '${expQualApprovedColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=EXP QUAL APP'D
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${expQualApprovedColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${expQualApprovedColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${expDWDateColOnSSTable}' != 'None'
                      ${expDWDateColOnSSTable}        Convert Date    ${expDWDateColOnSSTable}         date_format=%m/%d/%Y
                      ${expDWDateColOnSSTable}        Convert Date    ${expDWDateColOnSSTable}         result_format=%m/%d/%Y
                 END
                 IF    '${expDWDateColOnReportTable}' != 'None'
                      ${expDWDateColOnReportTable}    Convert Date    ${expDWDateColOnReportTable}     result_format=%m/%d/%Y
                 END
                 IF    '${expDWDateColOnReportTable}' != '${expDWDateColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=EXP DW DATE
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${expDWDateColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${expDWDateColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${1PPODateColOnSSTable}' != 'None'
                      ${1PPODateColOnSSTable}        Convert Date    ${1PPODateColOnSSTable}         date_format=%m/%d/%Y
                      ${1PPODateColOnSSTable}        Convert Date    ${1PPODateColOnSSTable}         result_format=%m/%d/%Y
                 END
                 IF    '${1PPODateColOnReportTable}' != 'None'
                      ${1PPODateColOnReportTable}    Convert Date    ${1PPODateColOnReportTable}     result_format=%m/%d/%Y
                 END
                 IF    '${1PPODateColOnReportTable}' != '${1PPODateColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=1PPO DATE
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${1PPODateColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${1PPODateColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${DWDateColOnSSTable}' != 'None'
                      ${DWDateColOnSSTable}        Convert Date    ${DWDateColOnSSTable}         date_format=%m/%d/%Y
                      ${DWDateColOnSSTable}        Convert Date    ${DWDateColOnSSTable}         result_format=%m/%d/%Y
                 END
                 IF    '${DWDateColOnReportTable}' != 'None'
                      ${DWDateColOnReportTable}    Convert Date    ${DWDateColOnReportTable}     result_format=%m/%d/%Y
                 END
                 IF    '${DWDateColOnReportTable}' != '${DWDateColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=DW DATE
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${DWDateColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${DWDateColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${DWColOnSSTable}' == '- None -'
                      ${DWColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${DWColOnReportTable}' == 'None'
                      ${DWColOnReportTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${DWColOnReportTable}' != '${DWColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=DESIGN WIN
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${DWColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${DWColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${customerPNColOnSSTable}' == '- None -'
                      ${customerPNColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${customerPNColOnReportTable}' == 'None'
                      ${customerPNColOnReportTable}     Set Variable    ${EMPTY}
                 END

                 IF    '${customerPNColOnReportTable}' != '${customerPNColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=CUSTOMER PN
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${customerPNColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${customerPNColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${subSegmentColOnSSTable}' == '- None -'
                      ${subSegmentColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${subSegmentColOnReportTable}' == 'None'
                      ${subSegmentColOnReportTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${subSegmentColOnReportTable}' != '${subSegmentColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SUB-SEGMENT
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${subSegmentColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${subSegmentColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${programColOnReportTable}     Convert To String    ${programColOnReportTable}
                 ${programColOnSSTable}     Convert To String    ${programColOnSSTable}
                 ${programColOnReportTable}    Remove String    ${programColOnReportTable}   '  "
                 ${programColOnSSTable}    Remove String    ${programColOnSSTable}   '  "

                 IF    '${programColOnSSTable}' == 'None' or '${programColOnSSTable}' == '- None -'
                      ${programColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${programColOnReportTable}' == 'None'
                      ${programColOnReportTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${programColOnReportTable}' != '${programColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP PG NAME
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${programColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${programColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${applicationColOnSSTable}' == '- None -'
                      ${applicationColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${applicationColOnReportTable}' == 'None'
                      ${applicationColOnReportTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${applicationColOnReportTable}' != '${applicationColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=APPLICATION
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${applicationColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${applicationColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 IF    '${functionColOnSSTable}' == '- None -'
                      ${functionColOnSSTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${functionColOnReportTable}' == 'None'
                      ${functionColOnReportTable}     Set Variable    ${EMPTY}
                 END
                 IF    '${functionColOnReportTable}' != '${functionColOnSSTable}'
                      ${result}   Set Variable      ${False}
                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=FUNCTION
                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${functionColOnReportTable}
                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${functionColOnSSTable}
                      Save Excel Document    ${RESULT_FILE_PATH}
                 END
                 ${rowIndexOnReportTableTemp}   Set Variable    ${rowIndexOnReportTable}
                 BREAK
            END
        END
        ${previousOpp}      Set Variable    ${oppColOnSSTable}

    END
    Close All Excel Documents
    
    [Return]    ${result}

Create Table From The SS Of Master Opp Report On NS
    [Arguments]     ${ssFilePath}
    @{table}    Create List

    File Should Exist    ${ssFilePath}
    Open Excel Document    ${ssFilePath}    MasterOppSource
    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssFilePath}

    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
        ${oppColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=1
        ${trackedOppColOnSS}                         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
        ${oppLinkToColOnSS}                          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=3
        ${oemGroupColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=5
        ${samColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=6
        ${saleRepColOnSS}                            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=7
        ${tmColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=8
        ${oppDiscoveryPersonColOnSS}                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
        ${bizDevSupportColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=10
        ${pnColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
        ${qtyColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=12
        ${projectTotalColOnSS}                       Read Excel Cell    row_num=${rowIndexOnSS}    col_num=13
        ${probColOnSS}                               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=14
        ${currentOppStageColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=15
        ${oppCategoryColOnSS}                        Read Excel Cell    row_num=${rowIndexOnSS}    col_num=17
        ${expSampleShipColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
        ${expQualApprovedColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=19
        ${expDWDateColOnSS}                          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=20
        ${1PPODateColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=21
        ${DWDateColOnSS}                             Read Excel Cell    row_num=${rowIndexOnSS}    col_num=22
        ${DWColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=23
        ${customerPNColOnSS}                         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=24
        ${subSegmentColOnSS}                         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=25
        ${programColOnSS}                            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=26
        ${applicationColOnSS}                        Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
        ${functionColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=28
        ${rowOnTable}   Create List
        ...             ${oppColOnSS}
        ...             ${trackedOppColOnSS}
        ...             ${oppLinkToColOnSS}
        ...             ${oemGroupColOnSS}
        ...             ${samColOnSS}
        ...             ${saleRepColOnSS}
        ...             ${tmColOnSS}
        ...             ${oppDiscoveryPersonColOnSS}
        ...             ${bizDevSupportColOnSS}
        ...             ${pnColOnSS}
        ...             ${qtyColOnSS}
        ...             ${projectTotalColOnSS}
        ...             ${probColOnSS}
        ...             ${currentOppStageColOnSS}
        ...             ${oppCategoryColOnSS}
        ...             ${expSampleShipColOnSS}
        ...             ${expQualApprovedColOnSS}
        ...             ${expDWDateColOnSS}
        ...             ${1PPODateColOnSS}
        ...             ${DWDateColOnSS}
        ...             ${DWColOnSS}
        ...             ${customerPNColOnSS}
        ...             ${subSegmentColOnSS}
        ...             ${programColOnSS}
        ...             ${applicationColOnSS}
        ...             ${functionColOnSS}
        Append To List    ${table}   ${rowOnTable}
        ${rowOnTable}   Remove Values From List    ${rowOnTable}
    END
    Close All Excel Documents

    [Return]    ${table}

Create Table For Master Opp Report
    [Arguments]     ${reportFilePath}
    @{table}    Create List

    File Should Exist    ${reportFilePath}
    Open Excel Document    ${reportFilePath}    MasterOppReport
    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}

    FOR    ${rowIndexOnReport}    IN RANGE    5    ${numOfRowsOnReport}+1
        ${oppColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
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
        ${expQualApprovedColOnReport}                    Read Excel Cell    row_num=${rowIndexOnReport}    col_num=20
        ${expDWDateColOnReport}                          Read Excel Cell    row_num=${rowIndexOnReport}    col_num=21
        ${1PPODateColOnReport}                           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=22
        ${DWDateColOnReport}                             Read Excel Cell    row_num=${rowIndexOnReport}    col_num=23
        ${DWColOnReport}                                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=24
        ${customerPNColOnReport}                         Read Excel Cell    row_num=${rowIndexOnReport}    col_num=41
        ${subSegmentColOnReport}                         Read Excel Cell    row_num=${rowIndexOnReport}    col_num=42
        ${programColOnReport}                            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=43
        ${applicationColOnReport}                        Read Excel Cell    row_num=${rowIndexOnReport}    col_num=44
        ${functionColOnReport}                           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=45
        ${rowOnTable}   Create List
        ...             ${oppColOnReport}
        ...             ${trackedOppColOnReport}
        ...             ${oppLinkToColOnReport}
        ...             ${oemGroupColOnReport}
        ...             ${samColOnReport}
        ...             ${saleRepColOnReport}
        ...             ${tmColOnReport}
        ...             ${oppDiscoveryPersonColOnReport}
        ...             ${bizDevSupportColOnReport}
        ...             ${pnColOnReport}
        ...             ${qtyColOnReport}
        ...             ${projectTotalColOnReport}
        ...             ${probColOnReport}
        ...             ${currentOppStageColOnReport}
        ...             ${oppCategoryColOnReport}
        ...             ${expSampleShipColOnReport}
        ...             ${expQualApprovedColOnReport}
        ...             ${expDWDateColOnReport}
        ...             ${1PPODateColOnReport}
        ...             ${DWDateColOnReport}
        ...             ${DWColOnReport}
        ...             ${customerPNColOnReport}
        ...             ${subSegmentColOnReport}
        ...             ${programColOnReport}
        ...             ${applicationColOnReport}
        ...             ${functionColOnReport}
        Append To List    ${table}   ${rowOnTable}
        ${rowOnTable}   Remove Values From List    ${rowOnTable}
    END
    Close All Excel Documents

    [Return]    ${table}

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
             Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP
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
    ${numOfOppsOnNS}        Get Length    ${listOfOppsOnNS}

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
        ${opp}  Read Excel Cell    ${rowIndex}    1
        IF    '${opp}' != '${EMPTY}'
             Append To List    ${listOfOpps}     ${opp}
        END
    END
    Close All Excel Documents
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Sort List    ${listOfOpps}
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
    Close All Excel Documents
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Sort List    ${listOfOpps}
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
    Sort List    ${listOfOpps}
    [Return]    ${listOfOpps}

Get List Of Opps Have Multi Items From The SS Of Master Opp Report On NS
    [Arguments]     ${ssFilePath}
    @{listOfOpps}   Create List

    Open Excel Document    ${ssFilePath}    MasterOppSource
    ${numOfRows}    Get Number Of Rows In Excel    ${ssFilePath}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRows}
        ${currentOpp}   Read Excel Cell    ${rowIndex}    1
        ${nextRow}      Evaluate    ${rowIndex}+1
        ${nextOpp}      Read Excel Cell    ${nextRow}    1
        IF    '${nextOpp}' == '${currentOpp}'
             Append To List    ${listOfOpps}    ${currentOpp}
        END
    END
    Close All Excel Documents
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Sort List    ${listOfOpps}
    [Return]    ${listOfOpps}




