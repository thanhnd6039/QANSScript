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
${masterOPPReportResultFilePath}                     ${RESULT_DIR}\\MasterOppReport\\MasterOppReportResult.xlsx
${ssMasterOPPFilePath}                               ${DOWNLOAD_DIR}\\testMasterOpportunity.xlsx
${masterOPPReportFilePath}                           ${DOWNLOAD_DIR}\\Opportunity Report V3.xlsx
${posOfOPPColOnSSMasterOPP}                          2
${posOfOPPColOnMasterOPPReport}                      1

*** Keywords ***
Initialize Suite
    Log To Console    Initialize Suite

Write The Test Result Of Master OPP Report To Excel
    [Arguments]     ${itemNeedToCheck}     ${opp}     ${valueOnMasterOPPReport}   ${valueOnSSMasterOPP}
    File Should Exist      path=${masterOPPReportResultFilePath}
    Open Excel Document    filename=${masterOPPReportResultFilePath}    doc_id=MasterOPPReportResult
    Switch Current Excel Document    doc_id=MasterOPPReportResult
    ${latestRow}   Get Number Of Rows In Excel    ${masterOPPReportResultFilePath}
    ${nextRow}    Evaluate    ${latestRow}+1
    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${itemNeedToCheck}
    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${opp}
    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${valueOnMasterOPPReport}
    Write Excel Cell    row_num=${nextRow}    col_num=4    value=${valueOnSSMasterOPP}
    Save Excel Document    ${masterOPPReportResultFilePath}
    Close Current Excel Document

Check The Data Of OPP
    [Arguments]     ${nameOfCol}
    ${result}   Set Variable    ${True}
    @{listOfOPPsFromSSMasterOPP}        Create List
    @{listOfOPPsFromMasterOPPReport}    Create List

    IF    '${nameOfCol}' == 'OPP'
         ${listOfOPPsFromSSMasterOPP}        Get List Of Opps From The SS Master Opp
         ${listOfOPPsFromMasterOPPReport}    Get List Of Opps From The Master Opp Report
         ${numOfRowsOnSSMasterOPP}           Get Length    ${listOfOPPsFromSSMasterOPP}
         ${numOfRowsOnMasterOPPReport}       Get Length    ${listOfOPPsFromMasterOPPReport}
         ${startIndexForMasterOPPReport}     Set Variable    0
         FOR    ${rowIndexOnSSMasterOPP}    IN RANGE    0    ${numOfRowsOnSSMasterOPP}
            ${isOPPInMasterOPPReport}   Set Variable    ${False}
            FOR    ${rowIndexOnMasterOPPReport}    IN RANGE    ${startIndexForMasterOPPReport}    ${numOfRowsOnMasterOPPReport}
                IF    '${listOfOPPsFromSSMasterOPP[${rowIndexOnSSMasterOPP}]}' == '${listOfOPPsFromMasterOPPReport[${rowIndexOnMasterOPPReport}]}'
                     ${isOPPInMasterOPPReport}          Set Variable    ${True}
                     ${startIndexForMasterOPPReport}    Evaluate    ${startIndexForMasterOPPReport}+1
                     BREAK
                END                 
            END
            IF    '${isOPPInMasterOPPReport}' == '${False}'
                 ${result}   Set Variable    ${False}
                 Write The Test Result Of Master OPP Report To Excel    itemNeedToCheck=OPP    opp=${EMPTY}    valueOnMasterOPPReport=${EMPTY}    valueOnSSMasterOPP=${listOfOPPsFromSSMasterOPP[${rowIndexOnSSMasterOPP}]}
            END
         END
    ELSE IF  '${nameOfCol}' == 'LINE ID'
        Log To Console    continue
    ELSE
        Log To Console    Invalid
    END
    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The data is different betwween the Master OPP Report and NS
    END
    Close All Excel Documents

Get List Of Opps From The SS Master Opp
    [Arguments]
    @{listOfOpps}   Create List

    File Should Exist    path=${ssMasterOPPFilePath}
    Open Excel Document    filename=${ssMasterOPPFilePath}    doc_id=SSMasterOPP
    ${numOfRows}    Get Number Of Rows In Excel    ${ssMasterOPPFilePath}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRows}+1
        ${oppCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOPPColOnSSMasterOPP}
        IF    '${oppCol}' != '${EMPTY}'
             Append To List    ${listOfOpps}     ${oppCol}
        END
    END
    Close All Excel Documents
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Sort List    ${listOfOpps}
    [Return]    ${listOfOpps}

Get List Of Opps From The Master Opp Report
    @{listOfOpps}   Create List

    File Should Exist      path=${masterOPPReportFilePath}
    Open Excel Document    filename=${masterOPPReportFilePath}    doc_id=MasterOPPReport
    ${numOfRows}           Get Number Of Rows In Excel      ${masterOPPReportFilePath}

    FOR    ${rowIndex}    IN RANGE    5    ${numOfRows}+1
        ${oppCol}   Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOPPColOnMasterOPPReport}
        IF    '${oppCol}' != '${EMPTY}'
             Append To List    ${listOfOpps}     ${oppCol}
        END
    END

    Close All Excel Documents
    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
    Sort List    ${listOfOpps}

    [Return]    ${listOfOpps}


#Check The REV Data
#    [Arguments]     ${masterOPPFilePath}  ${salesDashboardByPNFilePath}  ${ssMasterOPPFilePath}  ${year}  ${quarter}
#    Create Source Table To Verify REV For Each Quarter    ${ssMasterOPPFilePath}    ${salesDashboardByPNFilePath}    ${year}    ${quarter}

#    File Should Exist    ${masterOPPReportFilePath}
#    Open Excel Document    ${masterOPPReportFilePath}    doc_id=MasterOPPReport
#
#    File Should Exist    ${salesDashboardByPNReportFilePath}
#    Open Excel Document    ${salesDashboardByPNReportFilePath}    doc_id=SalesDashboardByPNReport

#    File Should Exist    ${ssMasterOPPFilePath}
#    Open Excel Document    ${ssMasterOPPFilePath}    doc_id=SSMasterOPP
#
#    Switch Current Excel Document    doc_id=SSMasterOPP
#    ${numOfRowsOnSSMasterOPP}   Get Number Of Rows In Excel    ${ssMasterOPPFilePath}


#    Switch Current Excel Document    doc_id=MasterOPPReport
#    ${yearStrOnMasterOPPReport}           Get Substring    ${year}    2  4
#    ${searchStrREVColOnMasterOPPReport}   Set Variable    ${yearStrOnMasterOPPReport}-Q${quarter} REV
#    ${startRowOnMasterOPPReport}          Convert To Number    4
#    ${posOfColREVOnMasterOPPReport}       Get Position Of Column    ${masterOPPReportFilePath}    ${startRowOnMasterOPPReport}    ${searchStrREVColOnMasterOPPReport}
#    ${numOfRowsOnMasterOPPReport}         Get Number Of Rows In Excel    ${masterOPPReportFilePath}
#    ${rowIndexOnSSMasterOPP}              Set Variable    2
#    FOR    ${rowIndexOnMasterOPPReport}    IN RANGE    5    ${numOfRowsOnMasterOPPReport}+1
#        ${oppColOnMasterOPPReport}          Read Excel Cell    row_num=${rowIndexOnMasterOPPReport}    col_num=1
#        ${revColOnMasterOPPReport}          Read Excel Cell    row_num=${rowIndexOnMasterOPPReport}    col_num=${posOfColREVOnMasterOPPReport}
#        ${isMapRevColOnMasterOPPReport}     Set Variable    ${EMPTY}
#        Switch Current Excel Document    doc_id=SSMasterOPP
#        ${isMapRevColOnSSMasterOPP}         Read Excel Cell    ${rowIndexOnSSMasterOPP}    4
#        ${isMapRevColOnMasterOPPReport}     Set Variable    ${isMapRevColOnSSMasterOPP}
#        ${rowIndexOnSSMasterOPP}            Evaluate    ${rowIndexOnSSMasterOPP}+1
#
#        IF    '${isMapRevColOnMasterOPPReport}' == 'No' or '${isMapRevColOnMasterOPPReport}' == '${EMPTY}'
#             IF    '${revColOnMasterOPPReport}' != 'None'
#                  Log To Console    OPP:${oppColOnMasterOPPReport},REV:${revColOnMasterOPPReport}
#             END
#        END
#        Switch Current Excel Document    doc_id=MasterOPPReport
#    END


#Create Source Table To Verify REV For Each Quarter
#    [Arguments]     ${ssMasterOPPFilePath}   ${salesDashboardByPN}      ${year}     ${quarter}
#    @{table}        Create List

#    File Should Exist      ${ssMasterOPPFilePath}
#    Open Excel Document    ${ssMasterOPPFilePath}    doc_id=SSMasterOPP
#    ${numOfRowsOnSSMasterOPP}   Get Number Of Rows In Excel    ${ssMasterOPPFilePath}
#
#    File Should Exist      ${salesDashboardByPN}
#    Open Excel Document    ${salesDashboardByPN}    doc_id=SalesDashboardByPN
#    Switch Current Excel Document    doc_id=SalesDashboardByPN
#    ${numOfRowsOnSalesDashboardByPN}    Get Number Of Rows In Excel    ${salesDashboardByPN}
#    ${searchStrForREVColOnSalesDashboardByPN}   Set Variable    ${year}.Q${quarter}
#    ${rowIndexOfHeaderOnSalesDashboardByPN}     Convert To Number    1
#    ${posOfREVColOnSalesDashboardByPN}   Get Position Of Column    ${salesDashboardByPN}    ${rowIndexOfHeaderOnSalesDashboardByPN}    ${searchStrForREVColOnSalesDashboardByPN}
#    IF    ${posOfREVColOnSalesDashboardByPN} == 0
#         Fail   The quarter parameter or year is invalid
#    END
#
#    Switch Current Excel Document    doc_id=SSMasterOPP
#    FOR    ${rowIndexOnSSMasterOPP}    IN RANGE    2    ${numOfRowsOnSSMasterOPP}+1
#        ${oppCol}       Set Variable    ${EMPTY}
#        ${isMapREVCol}  Set Variable    ${EMPTY}
#        ${oemGroupCol}  Set Variable    ${EMPTY}
#        ${pnCol}        Set Variable    ${EMPTY}
#        ${revCol}       Set Variable    ${EMPTY}
#
#        ${oppColOnSSMasterOPP}        Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=2
#        ${oemGroupColOnSSMasterOPP}   Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=3
#        ${pnColOnSSMasterOPP}         Read Excel Cell    row_num=${rowIndexOnSSMasterOPP}    col_num=4
#        ${str1}     Set Variable    ${oemGroupColOnSSMasterOPP}${pnColOnSSMasterOPP}
#
#        Switch Current Excel Document    doc_id=SalesDashboardByPN
#        FOR    ${rowIndexOnSalesDashboardByPN}    IN RANGE    2    ${numOfRowsOnSalesDashboardByPN}
#            ${oemGroupColOnSalesDashboardByPN}  Read Excel Cell    row_num=${rowIndexOnSalesDashboardByPN}    col_num=3
#            ${pnColOnSalesDashboardByPN}        Read Excel Cell    row_num=${rowIndexOnSalesDashboardByPN}    col_num=1
#            ${str2}     Set Variable    ${oemGroupColOnSalesDashboardByPN}${pnColOnSalesDashboardByPN}
#            IF    '${str1}' == '${str2}'
#                 ${revColOnSalesDashboardByPN}    Read Excel Cell    row_num=${rowIndexOnSalesDashboardByPN}    col_num=${posOfREVColOnSalesDashboardByPN}
#                 IF    ${revColOnSalesDashboardByPN} != 0
#                      Log To Console    OEM:${oemGroupColOnSSMasterOPP},PN:${pnColOnSSMasterOPP},REV:${revColOnSalesDashboardByPN}
#                 END
#                 BREAK
#            END
##            IF    '${oemGroupColOnSSMasterOPP}' == '${oemGroupColOnSalesDashboardByPN}' and '${pnColOnSSMasterOPP}' == '${pnColOnSalesDashboardByPN}'
##                 ${revColOnSalesDashboardByPN}    Read Excel Cell    row_num=${rowIndexOnSalesDashboardByPN}    col_num=${posOfREVColOnSalesDashboardByPN}
##                 IF    ${revColOnSalesDashboardByPN} != 0
##                      Log To Console    OEM:${oemGroupColOnSSMasterOPP},PN:${pnColOnSSMasterOPP},REV:${revColOnSalesDashboardByPN}
##                 END
##                 BREAK
##            END
#        END
#
##        ${oppCol}       Set Variable    ${oppColOnSSMasterOPP}
##        ${oemGroupCol}  Set Variable    ${oemGroupColOnSSMasterOPP}
##        ${pnCol}        Set Variable    ${pnColOnSSMasterOPP}
##        ${revCol}       Set Variable    ${revColOnSalesDashboardByPN}
##
##        ${rowOnTable}   Create List
##        ...             ${oppCol}
##        ...             ${oemGroupCol}
##        ...             ${pnCol}
##        ...             ${revCol}
##        Append To List    ${table}  ${rowOnTable}
#        Switch Current Excel Document    doc_id=SSMasterOPP
#    END

#    [Return]    ${table}

#Navigate To Master Opp Report
#    ${configFileObject}     Load Json From File    ${CONFIG_FILE}
#    ${username}             Get Value From Json    ${configFileObject}    $.accounts[0].username
#    ${username}             Set Variable           ${username}[0]
#    ${pass}                 Get Value From Json    ${configFileObject}    $.accounts[0].password
#    ${pass}                 Set Variable           ${pass}[0]
#    ${url}                  Set Variable           http://${username}:${pass}@report/ReportServer/Pages/ReportViewer.aspx?/NetSuite+Reports/Sales/Opportunity+Report&rs:Command=Render
#    Go To    ${url}
#
#Should See The Title Of Master Opp Report
#    [Arguments]     ${title}
#    Wait Until Element Is Visible    ${txtTitleOfMasterOpp}     ${TIMEOUT}
#    Element Text Should Be    ${txtTitleOfMasterOpp}    ${title}
#
#Select All Opp Stages On Master Opp Report
#    Wait Until Element Is Visible    ${lstOppStageFilter}   ${TIMEOUT}
#    Click Element    ${lstOppStageFilter}
#    Wait Until Element Is Visible    ${chkSelectAllOfOppStageOption}    ${TIMEOUT}
#    ${isCheckSelectAll}     Run Keyword And Return Status    Checkbox Should Be Selected    ${chkSelectAllOfOppStageOption}
#    IF    '${isCheckSelectAll}' == '${False}'
#         Click Element    ${chkSelectAllOfOppStageOption}
#    END
#
#Select Opp Stage On Master Opp Report
#    [Arguments]     ${multiOppStageOptions}
#    Wait Until Element Is Visible    ${lstOppStageFilter}       ${TIMEOUT}
#    Click Element    ${lstOppStageFilter}
#    Wait Until Element Is Visible    ${chkSelectAllOfOppStageOption}    ${TIMEOUT}
#    ${isCheckSelectAll}     Run Keyword And Return Status    Checkbox Should Be Selected    ${chkSelectAllOfOppStageOption}
#    IF    '${isCheckSelectAll}' == '${False}'
#        Click Element    ${chkSelectAllOfOppStageOption}
#        Click Element    ${chkSelectAllOfOppStageOption}
#    ELSE
#        Click Element    ${chkSelectAllOfOppStageOption}
#    END
#
#    FOR    ${oppStage}    IN    @{multiOppStageOptions}
#        IF    '${oppStage}' == '0.Identified'
#             Wait Until Element Is Visible    ${chk0_IdentifyOfOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk0_IdentifyOfOppStageOption}
#        END
#        IF    '${oppStage}' == '1.Opp Approved'
#             Wait Until Element Is Visible    ${chk1_OppApprovedOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk1_OppApprovedOppStageOption}
#        END
#        IF    '${oppStage}' == '2.Eval Submitted/Qual in Progress'
#             Wait Until Element Is Visible    ${chk2_EvalSubmittedOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk2_EvalSubmittedOppStageOption}
#        END
#        IF    '${oppStage}' == '3.Qual Issues'
#             Wait Until Element Is Visible    ${chk3_QualIssuesOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk3_QualIssuesOppStageOption}
#        END
#        IF    '${oppStage}' == '4.Qual Approved'
#             Wait Until Element Is Visible    ${chk4_QualApprovedOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk4_QualApprovedOppStageOption}
#        END
#        IF    '${oppStage}' == '5.First - Production PO'
#             Wait Until Element Is Visible    ${chk5_FirstProductionPOOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk5_FirstProductionPOOppStageOption}
#        END
#        IF    '${oppStage}' == '6.Production'
#             Wait Until Element Is Visible    ${chk6_ProductionOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk6_ProductionOppStageOption}
#        END
#        IF    '${oppStage}' == '7.Hold'
#             Wait Until Element Is Visible    ${chk7_HoldOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk7_HoldOppStageOption}
#        END
#        IF    '${oppStage}' == '8.Lost'
#             Wait Until Element Is Visible    ${chk8_LostOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk8_LostOppStageOption}
#        END
#        IF    '${oppStage}' == '9.Cancelled'
#             Wait Until Element Is Visible    ${chk9_CancelledOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk9_CancelledOppStageOption}
#        END
#        IF    '${oppStage}' == '9.Closed'
#             Wait Until Element Is Visible    ${chk9_ClosedOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk9_ClosedOppStageOption}
#        END
#        IF    '${oppStage}' == '9.Opp Disapproved'
#             Wait Until Element Is Visible    ${chk9_OppDisapprovedOppStageOption}      ${TIMEOUT}
#             Click Element    ${chk9_OppDisapprovedOppStageOption}
#        END
#    END
#
#Filter Created Date On Master Opp Report
#    [Arguments]     ${createdFrom}      ${createdTo}
#    IF    '${createdFrom}' == 'NULL'
#         Wait Until Element Is Visible    ${chkNullOfCreatedFromFilter}     ${TIMEOUT}
#         ${isCheckCheckboxNullOfCreatedFromFilter}  Run Keyword And Return Status    Checkbox Should Be Selected    ${chkNullOfCreatedFromFilter}
#         IF    '${isCheckCheckboxNullOfCreatedFromFilter}' == '${False}'
#              Click Element    ${chkNullOfCreatedFromFilter}
#         END
#    END
#
#    IF    '${createdTo}' == 'NULL'
#         Wait Until Element Is Visible    ${chkNullOfCreatedToFilter}   ${TIMEOUT}
#         ${isCheckCheckboxNullOfCreatedToFilter}  Run Keyword And Return Status    Checkbox Should Be Selected    ${chkNullOfCreatedToFilter}
#         IF    '${isCheckCheckboxNullOfCreatedToFilter}' == '${False}'
#              Click Element    ${chkNullOfCreatedToFilter}
#         END
#    END
#
#Navigate To The Save Search Of Master Opp Report On NS
#    ${url}      Set Variable    https://4499123.app.netsuite.com/app/common/search/searchresults.nl?searchid=4002&whence=
#    Login To NS With Account    PRODUCTION
#    Go To    ${url}
#
#Export Excel Data From The Save Search Of Master Opp Report On NS
#    Export SS Data To CSV
#    Sleep    5s
#    ${fullyFileName}    Get Fully File Name From Given Name    MasterOpps    ${DOWNLOAD_DIR}
#    ${csvFilePath}      Set Variable    ${DOWNLOAD_DIR}\\${fullyFileName}
#    ${xlsxFilePath}     Set Variable    ${DOWNLOAD_DIR}\\MasterOppSource.xlsx
#    Convert Csv To Xlsx    ${csvFilePath}    ${xlsxFilePath}
#
#Compare Data Between Master Opp Report And SS On NS
#    [Arguments]     ${reportFilePath}   ${ssFilePath}
#
#    ${result}   Set Variable    ${True}
#
##    ${verifyNumOfOPPs}                          Verify The Number Of Opps On Master Opp Report                          reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
##    ${verifyDocumentNumberOfOPP}                Verify The Document Number Of Opp On Master Opp Report                  reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
##    ${verifyDataOfOPPsWithOnlyOneItem}          Verify The Data Of OPPs With Only One Item On Master Opp Report                     reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
#    ${verifyOPPsHaveMultiItems}                 Verify The Data Of OPPs Have Multi Items On Master Opp Report    reportFilePath=${reportFilePath}    ssFilePath=${ssFilePath}
#
##    IF    '${verifyNumOfOPPs}' == '${False}' or '${verifyDocumentNumberOfOPP}' == '${False}' or '${verifyDataOfOPPsWithOnlyOneItem}' == '${False}' or '${verifyOPPsHaveMultiItems}' == '${False}'
##         ${result}  Set Variable    ${False}
##         Fail   The data betwwen Master Opp Report and NS is difference
##    END
#
#    [Return]    ${result}
#
#Verify The Data Of OPPs Have Multi Items On Master Opp Report
#    [Arguments]     ${reportFilePath}   ${ssFilePath}
#    @{oppsHaveMultiItemsOnNSTable}     Create List
#    @{listOfOppsHaveMultiItems}     Create List
#    ${result}   Set Variable    ${True}
#
#
#    ${oppsHaveMultiItemsOnNSTable}     Get List Of Opps Have Multi Items From The Master Opp Source    ssFilePath=${ssFilePath}
#    FOR    ${itemRow}    IN    @{oppsHaveMultiItemsOnNSTable}
#        Append To List    ${listOfOppsHaveMultiItems}    ${itemRow[0]}
#    END
#    ${listOfOppsHaveMultiItems}   Remove Duplicates    ${listOfOppsHaveMultiItems}
#
#    File Should Exist    ${ssFilePath}
#    Open Excel Document    ${ssFilePath}    doc_id=MasterOppSource
#    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssFilePath}
#    ${numOfRowsOnReportTable}    Get Length    ${oppsHaveMultiItemsOnNSTable}
#
#    File Should Exist    ${RESULT_FILE_PATH}
#    Open Excel Document    ${RESULT_FILE_PATH}    doc_id=MasterOppResult
#
#    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
#            ${isFound}      Set Variable    ${False}
#            Switch Current Excel Document    MasterOppSource
#            ${oppColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=1
#            ${pnColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
#            ${qtyColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=12
#            ${isOppHaveMultiItems}    Set Variable    ${False}
#
#             FOR    ${opp}    IN    @{listOfOppsHaveMultiItems}
#                  IF    '${oppColOnSS}' == '${opp}'
#                       ${isOppHaveMultiItems}    Set Variable    ${True}
#                       BREAK
#                  END
#             END
#             IF    '${isOppHaveMultiItems}' == '${True}'
#                  FOR    ${rowIndexOnReportTable}    IN RANGE    0    ${numOfRowsOnReportTable}
#                       ${oppColOnReportTable}          Set Variable        ${oppsHaveMultiItemsOnNSTable}[${rowIndexOnReportTable}][0]
#                       ${pnColOnReportTable}           Set Variable        ${oppsHaveMultiItemsOnNSTable}[${rowIndexOnReportTable}][1]
#                       ${qtyColOnReportTable}          Set Variable        ${oppsHaveMultiItemsOnNSTable}[${rowIndexOnReportTable}][2]
#
#                       IF    '${oppColOnSS}' == '${oppColOnReportTable}' and '${pnColOnSS}' == '${pnColOnReportTable}' and '${qtyColOnSS}' == '${qtyColOnReportTable}'
#                            ${isFound}  Set Variable    ${True}
#                            BREAK
#                       END
#                  END
#                  IF    '${isFound}' == '${False}'
#                      Switch Current Excel Document    MasterOppResult
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=MULTI ITEMS
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSS}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${EMPTY}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${pnColOnSS}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                  END
#             END
#    END
#
#    [Return]    ${result}
#
#Verify The Data Of OPPs With Only One Item On Master Opp Report
#    [Arguments]     ${reportFilePath}   ${ssFilePath}
#    ${result}                           Set Variable    ${True}
#    @{reportTable}                      Create List
#    @{ssTable}                          Create List
#    @{listOfOppsHaveMultiItemsOnNS}     Create List
#    @{listOfOppsHaveMultiItemsOnlyContainsOppName}  Create List
#
#
#    ${listOfOppsHaveMultiItemsOnNS}     Get List Of Opps Have Multi Items From The Master Opp Source                 ssFilePath=${ssFilePath}
#    FOR    ${itemRow}    IN    @{listOfOppsHaveMultiItemsOnNS}
#        Append To List    ${listOfOppsHaveMultiItemsOnlyContainsOppName}    ${itemRow[0]}
#    END
#    ${listOfOppsHaveMultiItemsOnlyContainsOppName}   Remove Duplicates    ${listOfOppsHaveMultiItemsOnlyContainsOppName}
#
#    ${reportTable}                      Create Table For Master Opp Report                                          ${reportFilePath}
#    ${ssTable}                          Create Table From The SS Of Master Opp Report On NS                         ${ssFilePath}
#
#    ${reportTable}  Sort Table By Column    ${reportTable}    0
#    ${ssTable}      Sort Table By Column    ${ssTable}        0
#    ${numOfRowsOnReportTable}   Get Length    ${reportTable}
#    ${numOfRowsOnSSTable}       Get Length    ${ssTable}
#
#    Open Excel Document    ${RESULT_FILE_PATH}    doc_id=MasterOppResult
#    ${previousOpp}  Set Variable    ${EMPTY}
#    ${rowIndexOnReportTableTemp}     Set Variable    0
#    FOR    ${rowIndexOnSSTable}    IN RANGE    0    ${numOfRowsOnSSTable}
#        ${oppColOnSSTable}          Set Variable        ${ssTable}[${rowIndexOnSSTable}][0]
#        IF    '${oppColOnSSTable}' == '${previousOpp}'
#             Continue For Loop
#        END
#        Log To Console    OPP: ${oppColOnSSTable}
#        FOR    ${rowIndexOnReportTable}    IN RANGE    ${rowIndexOnReportTableTemp}    ${numOfRowsOnReportTable}
#            ${oppColOnReportTable}  Set Variable   ${reportTable}[${rowIndexOnReportTable}][0]
#
#            IF    '${oppColOnReportTable}' == '${oppColOnSSTable}'
#                 ${trackedOppColOnSSTable}                   Set Variable        ${ssTable}[${rowIndexOnSSTable}][1]
#                 ${oppLinkToColOnSSTable}                    Set Variable        ${ssTable}[${rowIndexOnSSTable}][2]
#                 ${oemGroupColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][3]
#                 ${samColOnSSTable}                          Set Variable        ${ssTable}[${rowIndexOnSSTable}][4]
#                 ${saleRepColOnSSTable}                      Set Variable        ${ssTable}[${rowIndexOnSSTable}][5]
#                 ${tmColOnSSTable}                           Set Variable        ${ssTable}[${rowIndexOnSSTable}][6]
#                 ${oppDiscoveryPersonColOnSSTable}           Set Variable        ${ssTable}[${rowIndexOnSSTable}][7]
#                 ${bizDevSupportColOnSSTable}                Set Variable        ${ssTable}[${rowIndexOnSSTable}][8]
#                 ${pnColOnSSTable}                           Set Variable        ${ssTable}[${rowIndexOnSSTable}][9]
#                 ${qtyColOnSSTable}                          Set Variable        ${ssTable}[${rowIndexOnSSTable}][10]
#                 ${projectTotalColOnSSTable}                 Set Variable        ${ssTable}[${rowIndexOnSSTable}][11]
#                 ${probColOnSSTable}                         Set Variable        ${ssTable}[${rowIndexOnSSTable}][12]
#                 ${oppStageColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][13]
#                 ${oppCategoryColOnSSTable}                  Set Variable        ${ssTable}[${rowIndexOnSSTable}][14]
#                 ${expSampleShipColOnSSTable}                Set Variable        ${ssTable}[${rowIndexOnSSTable}][15]
#                 ${expQualApprovedColOnSSTable}              Set Variable        ${ssTable}[${rowIndexOnSSTable}][16]
#                 ${expDWDateColOnSSTable}                    Set Variable        ${ssTable}[${rowIndexOnSSTable}][17]
#                 ${1PPODateColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][18]
#                 ${DWDateColOnSSTable}                       Set Variable        ${ssTable}[${rowIndexOnSSTable}][19]
#                 ${DWColOnSSTable}                           Set Variable        ${ssTable}[${rowIndexOnSSTable}][20]
#                 ${customerPNColOnSSTable}                   Set Variable        ${ssTable}[${rowIndexOnSSTable}][21]
#                 ${subSegmentColOnSSTable}                   Set Variable        ${ssTable}[${rowIndexOnSSTable}][22]
#                 ${programColOnSSTable}                      Set Variable        ${ssTable}[${rowIndexOnSSTable}][23]
#                 ${applicationColOnSSTable}                  Set Variable        ${ssTable}[${rowIndexOnSSTable}][24]
#                 ${functionColOnSSTable}                     Set Variable        ${ssTable}[${rowIndexOnSSTable}][25]
#
#                 ${trackedOppColOnReportTable}                   Set Variable    ${reportTable}[${rowIndexOnReportTable}][1]
#                 ${oppLinkToColOnReportTable}                    Set Variable    ${reportTable}[${rowIndexOnReportTable}][2]
#                 ${oemGroupColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][3]
#                 ${samColOnReportTable}                          Set Variable    ${reportTable}[${rowIndexOnReportTable}][4]
#                 ${saleRepColOnReportTable}                      Set Variable    ${reportTable}[${rowIndexOnReportTable}][5]
#                 ${tmColOnReportTable}                           Set Variable    ${reportTable}[${rowIndexOnReportTable}][6]
#                 ${oppDiscoveryPersonColOnReportTable}           Set Variable    ${reportTable}[${rowIndexOnReportTable}][7]
#                 ${bizDevSupportColOnReportTable}                Set Variable    ${reportTable}[${rowIndexOnReportTable}][8]
#                 ${pnColOnReportTable}                           Set Variable    ${reportTable}[${rowIndexOnReportTable}][9]
#                 ${qtyColOnReportTable}                          Set Variable    ${reportTable}[${rowIndexOnReportTable}][10]
#                 ${projectTotalColOnReportTable}                 Set Variable    ${reportTable}[${rowIndexOnReportTable}][11]
#                 ${probColOnReportTable}                         Set Variable    ${reportTable}[${rowIndexOnReportTable}][12]
#                 ${oppStageColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][13]
#                 ${oppCategoryColOnReportTable}                  Set Variable    ${reportTable}[${rowIndexOnReportTable}][14]
#                 ${expSampleShipColOnReportTable}                Set Variable    ${reportTable}[${rowIndexOnReportTable}][15]
#                 ${expQualApprovedColOnReportTable}              Set Variable    ${reportTable}[${rowIndexOnReportTable}][16]
#                 ${expDWDateColOnReportTable}                    Set Variable    ${reportTable}[${rowIndexOnReportTable}][17]
#                 ${1PPODateColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][18]
#                 ${DWDateColOnReportTable}                       Set Variable    ${reportTable}[${rowIndexOnReportTable}][19]
#                 ${DWColOnReportTable}                           Set Variable    ${reportTable}[${rowIndexOnReportTable}][20]
#                 ${customerPNColOnReportTable}                   Set Variable    ${reportTable}[${rowIndexOnReportTable}][21]
#                 ${subSegmentColOnReportTable}                   Set Variable    ${reportTable}[${rowIndexOnReportTable}][22]
#                 ${programColOnReportTable}                      Set Variable    ${reportTable}[${rowIndexOnReportTable}][23]
#                 ${applicationColOnReportTable}                  Set Variable    ${reportTable}[${rowIndexOnReportTable}][24]
#                 ${functionColOnReportTable}                     Set Variable    ${reportTable}[${rowIndexOnReportTable}][25]
#
#                 IF    '${trackedOppColOnReportTable}' != '${trackedOppColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=TRACKED OPP
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${trackedOppColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${trackedOppColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${oppLinkToColOnReportTable}' != '${oppLinkToColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP LINK TO
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppLinkToColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppLinkToColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${oemGroupColOnReportTable}' != '${oemGroupColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OEM GROUP
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oemGroupColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oemGroupColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${samColOnReportTable}' != '${samColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SAM
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${samColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${samColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${saleRepColOnReportTable}' != '${saleRepColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SALES REP
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${saleRepColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${saleRepColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${tmColOnReportTable}' != '${tmColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=TECHNICAL MARKETING
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${tmColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${tmColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${oppDiscoveryPersonColOnReportTable}' != '${oppDiscoveryPersonColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP DISCOVERY PERSON
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppDiscoveryPersonColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppDiscoveryPersonColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 ${isOppHaveMultiItems}    Set Variable    ${False}
#
#                 FOR    ${opp}    IN    @{listOfOppsHaveMultiItemsOnlyContainsOppName}
#                     IF    '${oppColOnSSTable}' == '${opp}'
#                          ${isOppHaveMultiItems}    Set Variable    ${True}
#                          BREAK
#                     END
#                 END
#
#                 IF    '${isOppHaveMultiItems}' == '${False}'
#                      IF    '${pnColOnReportTable}' != '${pnColOnSSTable}'
#                          ${result}   Set Variable      ${False}
#                          ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                          ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                          Write Excel Cell    row_num=${nextRow}    col_num=1    value=PART NUMBER
#                          Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                          Write Excel Cell    row_num=${nextRow}    col_num=3    value=${pnColOnReportTable}
#                          Write Excel Cell    row_num=${nextRow}    col_num=4    value=${pnColOnSSTable}
#                          Save Excel Document    ${RESULT_FILE_PATH}
#                      END
#                      IF    '${qtyColOnReportTable}' != '${qtyColOnSSTable}'
#                          ${result}   Set Variable      ${False}
#                          ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                          ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                          Write Excel Cell    row_num=${nextRow}    col_num=1    value=QTY PER YR
#                          Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                          Write Excel Cell    row_num=${nextRow}    col_num=3    value=${qtyColOnReportTable}
#                          Write Excel Cell    row_num=${nextRow}    col_num=4    value=${qtyColOnSSTable}
#                          Save Excel Document    ${RESULT_FILE_PATH}
#                      END
#                 END
#
#                 IF    '${bizDevSupportColOnReportTable}' != '${bizDevSupportColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=BIZ DEV SUPPORT
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${bizDevSupportColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${bizDevSupportColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 ${diffProjectTotal}    Evaluate    abs(${projectTotalColOnReportTable}-${projectTotalColOnSSTable})
#                 IF    ${diffProjectTotal} > 1
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=PROJECT TOTAL
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${projectTotalColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${projectTotalColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 ${diffProb}    Evaluate    abs(${probColOnReportTable}-${probColOnSSTable})
#                 IF    ${diffProb} >= 1
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=PROB
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${probColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${probColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${oppStageColOnReportTable}' != '${oppStageColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=CURRENT OPP STAGE
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppStageColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppStageColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${oppCategoryColOnReportTable}' != '${oppCategoryColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP CATEGORY
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${oppCategoryColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppCategoryColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${expSampleShipColOnReportTable}' != '${expSampleShipColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=EXP SAMPLE SHIP
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${expSampleShipColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${expSampleShipColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${expQualApprovedColOnReportTable}' != '${expQualApprovedColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=EXP QUAL APP'D
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${expQualApprovedColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${expQualApprovedColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${expDWDateColOnReportTable}' != '${expDWDateColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=EXP DW DATE
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${expDWDateColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${expDWDateColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${1PPODateColOnReportTable}' != '${1PPODateColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=1PPO DATE
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${1PPODateColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${1PPODateColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${DWDateColOnReportTable}' != '${DWDateColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=DW DATE
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${DWDateColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${DWDateColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${DWColOnReportTable}' != '${DWColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=DESIGN WIN
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${DWColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${DWColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${customerPNColOnReportTable}' != '${customerPNColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=CUSTOMER PN
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${customerPNColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${customerPNColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${subSegmentColOnReportTable}' != '${subSegmentColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=SUB-SEGMENT
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${subSegmentColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${subSegmentColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${programColOnReportTable}' != '${programColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP PG NAME
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${programColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${programColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${applicationColOnReportTable}' != '${applicationColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=APPLICATION
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${applicationColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${applicationColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#
#                 IF    '${functionColOnReportTable}' != '${functionColOnSSTable}'
#                      ${result}   Set Variable      ${False}
#                      ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#                      ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#                      Write Excel Cell    row_num=${nextRow}    col_num=1    value=FUNCTION
#                      Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oppColOnSSTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=3    value=${functionColOnReportTable}
#                      Write Excel Cell    row_num=${nextRow}    col_num=4    value=${functionColOnSSTable}
#                      Save Excel Document    ${RESULT_FILE_PATH}
#                 END
#                 ${rowIndexOnReportTableTemp}   Set Variable    ${rowIndexOnReportTable}
#                 BREAK
#            END
#        END
#        ${previousOpp}      Set Variable    ${oppColOnSSTable}
#
#    END
#    Close All Excel Documents
#
#    [Return]    ${result}
#
#Create Table From The SS Of Master Opp Report On NS
#    [Arguments]     ${ssFilePath}
#    @{table}    Create List
#
#    File Should Exist    ${ssFilePath}
#    Open Excel Document    ${ssFilePath}    MasterOppSource
#    ${numOfRowsOnSS}    Get Number Of Rows In Excel    ${ssFilePath}
#
#    FOR    ${rowIndexOnSS}    IN RANGE    2    ${numOfRowsOnSS}+1
#        ${oppColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=1
#        ${trackedOppColOnSS}                         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=2
#        ${oppLinkToColOnSS}                          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=3
#        IF    '${oppLinkToColOnSS}' == '- None -'
#             ${oppLinkToColOnSS}      Set Variable    ${EMPTY}
#        END
#
#        ${oemGroupColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=5
#        ${isOEMGroupOnSSContainsColon}  Set Variable    ${False}
#        ${isOEMGroupOnSSContainsColon}  Evaluate   ":" in """${oemGroupColOnSS}"""
#        IF    '${isOEMGroupOnSSContainsColon}' == '${True}'
#              ${strArrTemp}   Split String    ${oemGroupColOnSS}  :
#              ${oemGroupColOnSS}    Set Variable    ${strArrTemp}[1]
#              ${oemGroupColOnSS}    Set Variable    ${oemGroupColOnSS.strip()}
#        END
#
#        ${samColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=6
#        ${saleRepColOnSS}                            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=7
#        ${isSaleRepOnSSContainsComma}  Set Variable    ${False}
#        ${isSaleRepOnSSContainsComma}  Evaluate   "," in """${saleRepColOnSS}"""
#        IF    '${isSaleRepOnSSContainsComma}' == '${True}'
#            ${strArrTemp}   Split String    ${saleRepColOnSS}
#            ${saleRepColOnSS}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
#            ${saleRepColOnSS}  Remove String    ${saleRepColOnSS}     ,
#            ${saleRepColOnSS}    Set Variable    ${saleRepColOnSS.strip()}
#        END
#
#        ${tmColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=8
#        IF    '${tmColOnSS}' == '- None -'
#            ${tmColOnSS}     Set Variable    ${EMPTY}
#        END
#        ${isTMOnSSContainsComma}  Set Variable    ${False}
#        ${isTMOnSSContainsComma}  Evaluate   "," in """${tmColOnSS}"""
#        IF    '${isTMOnSSContainsComma}' == '${True}'
#            ${strArrTemp}   Split String    ${tmColOnSS}
#            ${tmColOnSS}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
#            ${tmColOnSS}  Remove String    ${tmColOnSS}     ,
#            ${tmColOnSS}    Set Variable    ${tmColOnSS.strip()}
#        END
#
#        ${oppDiscoveryPersonColOnSS}                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=9
#        IF    '${oppDiscoveryPersonColOnSS}' == '- None -'
#           ${oppDiscoveryPersonColOnSS}     Set Variable    ${EMPTY}
#        END
#        ${isOPPDiscoveryPersonOnSSContainsComma}  Set Variable    ${False}
#        ${isOPPDiscoveryPersonOnSSContainsComma}  Evaluate   "," in """${oppDiscoveryPersonColOnSS}"""
#        IF    '${isOPPDiscoveryPersonOnSSContainsComma}' == '${True}'
#            ${strArrTemp}   Split String    ${oppDiscoveryPersonColOnSS}
#            ${oppDiscoveryPersonColOnSS}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
#            ${oppDiscoveryPersonColOnSS}    Remove String    ${oppDiscoveryPersonColOnSS}     ,
#            ${oppDiscoveryPersonColOnSS}    Set Variable    ${oppDiscoveryPersonColOnSS.strip()}
#        END
#
#        ${bizDevSupportColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=10
#        IF    '${bizDevSupportColOnSS}' == '- None -'
#              ${bizDevSupportColOnSS}     Set Variable    ${EMPTY}
#        END
#        ${isBizDevSupportContainsComma}  Set Variable    ${False}
#        ${isBizDevSupportContainsComma}  Evaluate   "," in """${bizDevSupportColOnSS}"""
#        IF    '${isBizDevSupportContainsComma}' == '${True}'
#                ${strArrTemp}   Split String    ${bizDevSupportColOnSS}
#                ${bizDevSupportColOnSS}      Catenate    ${strArrTemp}[1]    ${strArrTemp}[0]
#                ${bizDevSupportColOnSS}  Remove String    ${bizDevSupportColOnSS}     ,
#                ${bizDevSupportColOnSS}    Set Variable    ${bizDevSupportColOnSS.strip()}
#        END
#
#        ${pnColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=11
#        ${qtyColOnSS}                                Read Excel Cell    row_num=${rowIndexOnSS}    col_num=12
#        ${projectTotalColOnSS}                       Read Excel Cell    row_num=${rowIndexOnSS}    col_num=13
#        ${probColOnSS}                               Read Excel Cell    row_num=${rowIndexOnSS}    col_num=14
#        ${probColOnSS}  Remove String    ${probColOnSS}  %
#        ${probColOnSS}  Convert To Number    ${probColOnSS}
#
#        ${currentOppStageColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=15
#        ${oppCategoryColOnSS}                        Read Excel Cell    row_num=${rowIndexOnSS}    col_num=17
#        ${expSampleShipColOnSS}                      Read Excel Cell    row_num=${rowIndexOnSS}    col_num=18
#        IF    '${expSampleShipColOnSS}' != 'None'
#              ${expSampleShipColOnSS}        Convert Date    ${expSampleShipColOnSS}         date_format=%m/%d/%Y
#              ${expSampleShipColOnSS}        Convert Date    ${expSampleShipColOnSS}         result_format=%m/%d/%Y
#        END
#
#        ${expQualApprovedColOnSS}                    Read Excel Cell    row_num=${rowIndexOnSS}    col_num=19
#        IF    '${expQualApprovedColOnSS}' != 'None'
#              ${expQualApprovedColOnSS}        Convert Date    ${expQualApprovedColOnSS}         date_format=%m/%d/%Y
#              ${expQualApprovedColOnSS}        Convert Date    ${expQualApprovedColOnSS}         result_format=%m/%d/%Y
#        END
#
#        ${expDWDateColOnSS}                          Read Excel Cell    row_num=${rowIndexOnSS}    col_num=20
#        IF    '${expDWDateColOnSS}' != 'None'
#              ${expDWDateColOnSS}        Convert Date    ${expDWDateColOnSS}         date_format=%m/%d/%Y
#              ${expDWDateColOnSS}        Convert Date    ${expDWDateColOnSS}         result_format=%m/%d/%Y
#        END
#
#        ${1PPODateColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=21
#        IF    '${1PPODateColOnSS}' != 'None'
#              ${1PPODateColOnSS}        Convert Date    ${1PPODateColOnSS}         date_format=%m/%d/%Y
#              ${1PPODateColOnSS}        Convert Date    ${1PPODateColOnSS}         result_format=%m/%d/%Y
#        END
#
#        ${DWDateColOnSS}                             Read Excel Cell    row_num=${rowIndexOnSS}    col_num=22
#        IF    '${DWDateColOnSS}' != 'None'
#              ${DWDateColOnSS}        Convert Date    ${DWDateColOnSS}         date_format=%m/%d/%Y
#              ${DWDateColOnSS}        Convert Date    ${DWDateColOnSS}         result_format=%m/%d/%Y
#        END
#
#        ${DWColOnSS}                                 Read Excel Cell    row_num=${rowIndexOnSS}    col_num=23
#        IF    '${DWColOnSS}' == '- None -'
#              ${DWColOnSS}     Set Variable    ${EMPTY}
#        END
#
#        ${customerPNColOnSS}                         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=24
#        IF    '${customerPNColOnSS}' == '- None -'
#             ${customerPNColOnSS}     Set Variable    ${EMPTY}
#        END
#
#        ${subSegmentColOnSS}                         Read Excel Cell    row_num=${rowIndexOnSS}    col_num=25
#        IF    '${subSegmentColOnSS}' == '- None -'
#              ${subSegmentColOnSS}     Set Variable    ${EMPTY}
#        END
#        ${programColOnSS}                            Read Excel Cell    row_num=${rowIndexOnSS}    col_num=26
#        ${programColOnSS}     Convert To String    ${programColOnSS}
#         ${programColOnSS}    Remove String    ${programColOnSS}   '  "
#         IF    '${programColOnSS}' == 'None' or '${programColOnSS}' == '- None -'
#              ${programColOnSS}     Set Variable    ${EMPTY}
#         END
#
#        ${applicationColOnSS}                        Read Excel Cell    row_num=${rowIndexOnSS}    col_num=27
#        IF    '${applicationColOnSS}' == '- None -'
#              ${applicationColOnSS}     Set Variable    ${EMPTY}
#        END
#
#        ${functionColOnSS}                           Read Excel Cell    row_num=${rowIndexOnSS}    col_num=28
#        IF    '${functionColOnSS}' == '- None -'
#              ${functionColOnSS}     Set Variable    ${EMPTY}
#        END
#
#        ${rowOnTable}   Create List
#        ...             ${oppColOnSS}
#        ...             ${trackedOppColOnSS}
#        ...             ${oppLinkToColOnSS}
#        ...             ${oemGroupColOnSS}
#        ...             ${samColOnSS}
#        ...             ${saleRepColOnSS}
#        ...             ${tmColOnSS}
#        ...             ${oppDiscoveryPersonColOnSS}
#        ...             ${bizDevSupportColOnSS}
#        ...             ${pnColOnSS}
#        ...             ${qtyColOnSS}
#        ...             ${projectTotalColOnSS}
#        ...             ${probColOnSS}
#        ...             ${currentOppStageColOnSS}
#        ...             ${oppCategoryColOnSS}
#        ...             ${expSampleShipColOnSS}
#        ...             ${expQualApprovedColOnSS}
#        ...             ${expDWDateColOnSS}
#        ...             ${1PPODateColOnSS}
#        ...             ${DWDateColOnSS}
#        ...             ${DWColOnSS}
#        ...             ${customerPNColOnSS}
#        ...             ${subSegmentColOnSS}
#        ...             ${programColOnSS}
#        ...             ${applicationColOnSS}
#        ...             ${functionColOnSS}
#        Append To List    ${table}   ${rowOnTable}
#        ${rowOnTable}   Remove Values From List    ${rowOnTable}
#    END
#    Close All Excel Documents
#
#    [Return]    ${table}
#
#Create Table For Master Opp Report
#    [Arguments]     ${reportFilePath}
#    @{table}    Create List
#
#    File Should Exist    ${reportFilePath}
#    Open Excel Document    ${reportFilePath}    MasterOppReport
#    ${numOfRowsOnReport}    Get Number Of Rows In Excel    ${reportFilePath}
#
#    FOR    ${rowIndexOnReport}    IN RANGE    5    ${numOfRowsOnReport}+1
#        ${oppColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=1
#        ${trackedOppColOnReport}                         Read Excel Cell    row_num=${rowIndexOnReport}    col_num=2
#        ${oppLinkToColOnReport}                          Read Excel Cell    row_num=${rowIndexOnReport}    col_num=3
#        IF    '${oppLinkToColOnReport}' == 'None'
#             ${oppLinkToColOnReport}     Set Variable     ${EMPTY}
#        END
#        ${oemGroupColOnReport}                           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=5
#        ${samColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=7
#        ${saleRepColOnReport}                            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=8
#        ${tmColOnReport}                                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=9
#        IF    '${tmColOnReport}' == 'None'
#             ${tmColOnReport}   Set Variable   ${EMPTY}
#        END
#        IF    '${tmColOnReport}' != '${EMPTY}'
#            ${tmColOnReport}    Set Variable    ${tmColOnReport.strip()}
#        END
#
#        ${oppDiscoveryPersonColOnReport}                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=10
#        IF    '${oppDiscoveryPersonColOnReport}' == 'None'
#             ${oppDiscoveryPersonColOnReport}   Set Variable   ${EMPTY}
#        END
#        IF    '${oppDiscoveryPersonColOnReport}' != '${EMPTY}'
#             ${oppDiscoveryPersonColOnReport}   Set Variable   ${oppDiscoveryPersonColOnReport.strip()}
#        END
#
#        ${bizDevSupportColOnReport}                      Read Excel Cell    row_num=${rowIndexOnReport}    col_num=11
#        IF    '${bizDevSupportColOnReport}' == 'None'
#             ${bizDevSupportColOnReport}       Set Variable    ${EMPTY}
#        END
#        IF    '${bizDevSupportColOnReport}' != '${EMPTY}'
#             ${bizDevSupportColOnReport}       Set Variable    ${bizDevSupportColOnReport.strip()}
#        END
#
#        ${pnColOnReport}                                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=12
#        ${qtyColOnReport}                                Read Excel Cell    row_num=${rowIndexOnReport}    col_num=13
#        ${projectTotalColOnReport}                       Read Excel Cell    row_num=${rowIndexOnReport}    col_num=14
#        ${probColOnReport}                               Read Excel Cell    row_num=${rowIndexOnReport}    col_num=15
#        ${probColOnReport}    Evaluate    ${probColOnReport}*100
#
#        ${currentOppStageColOnReport}                    Read Excel Cell    row_num=${rowIndexOnReport}    col_num=16
#        ${oppCategoryColOnReport}                        Read Excel Cell    row_num=${rowIndexOnReport}    col_num=18
#        ${expSampleShipColOnReport}                      Read Excel Cell    row_num=${rowIndexOnReport}    col_num=19
#        IF    '${expSampleShipColOnReport}' != 'None'
#           ${expSampleShipColOnReport}    Convert Date    ${expSampleShipColOnReport}     result_format=%m/%d/%Y
#        END
#
#        ${expQualApprovedColOnReport}                    Read Excel Cell    row_num=${rowIndexOnReport}    col_num=20
#        IF    '${expQualApprovedColOnReport}' != 'None'
#            ${expQualApprovedColOnReport}    Convert Date    ${expQualApprovedColOnReport}     result_format=%m/%d/%Y
#        END
#
#        ${expDWDateColOnReport}                          Read Excel Cell    row_num=${rowIndexOnReport}    col_num=21
#        IF    '${expDWDateColOnReport}' != 'None'
#             ${expDWDateColOnReport}    Convert Date    ${expDWDateColOnReport}     result_format=%m/%d/%Y
#        END
#
#        ${1PPODateColOnReport}                           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=22
#        IF    '${1PPODateColOnReport}' != 'None'
#              ${1PPODateColOnReport}    Convert Date    ${1PPODateColOnReport}     result_format=%m/%d/%Y
#        END
#
#        ${DWDateColOnReport}                             Read Excel Cell    row_num=${rowIndexOnReport}    col_num=23
#        IF    '${DWDateColOnReport}' != 'None'
#              ${DWDateColOnReport}    Convert Date    ${DWDateColOnReport}     result_format=%m/%d/%Y
#        END
#
#        ${DWColOnReport}                                 Read Excel Cell    row_num=${rowIndexOnReport}    col_num=24
#        IF    '${DWColOnReport}' == 'None'
#            ${DWColOnReport}     Set Variable    ${EMPTY}
#        END
#
#        ${customerPNColOnReport}                         Read Excel Cell    row_num=${rowIndexOnReport}    col_num=41
#        IF    '${customerPNColOnReport}' == 'None'
#              ${customerPNColOnReport}     Set Variable    ${EMPTY}
#         END
#
#        ${subSegmentColOnReport}                         Read Excel Cell    row_num=${rowIndexOnReport}    col_num=42
#        IF    '${subSegmentColOnReport}' == 'None'
#              ${subSegmentColOnReport}     Set Variable    ${EMPTY}
#        END
#
#        ${programColOnReport}                            Read Excel Cell    row_num=${rowIndexOnReport}    col_num=43
#        ${programColOnReport}     Convert To String    ${programColOnReport}
#         ${programColOnReport}    Remove String    ${programColOnReport}   '  "
#         IF    '${programColOnReport}' == 'None'
#              ${programColOnReport}     Set Variable    ${EMPTY}
#         END
#
#        ${applicationColOnReport}                        Read Excel Cell    row_num=${rowIndexOnReport}    col_num=44
#        IF    '${applicationColOnReport}' == 'None'
#              ${applicationColOnReport}     Set Variable    ${EMPTY}
#        END
#
#        ${functionColOnReport}                           Read Excel Cell    row_num=${rowIndexOnReport}    col_num=45
#        IF    '${functionColOnReport}' == 'None'
#              ${functionColOnReport}     Set Variable    ${EMPTY}
#        END
#
#        ${rowOnTable}   Create List
#        ...             ${oppColOnReport}
#        ...             ${trackedOppColOnReport}
#        ...             ${oppLinkToColOnReport}
#        ...             ${oemGroupColOnReport}
#        ...             ${samColOnReport}
#        ...             ${saleRepColOnReport}
#        ...             ${tmColOnReport}
#        ...             ${oppDiscoveryPersonColOnReport}
#        ...             ${bizDevSupportColOnReport}
#        ...             ${pnColOnReport}
#        ...             ${qtyColOnReport}
#        ...             ${projectTotalColOnReport}
#        ...             ${probColOnReport}
#        ...             ${currentOppStageColOnReport}
#        ...             ${oppCategoryColOnReport}
#        ...             ${expSampleShipColOnReport}
#        ...             ${expQualApprovedColOnReport}
#        ...             ${expDWDateColOnReport}
#        ...             ${1PPODateColOnReport}
#        ...             ${DWDateColOnReport}
#        ...             ${DWColOnReport}
#        ...             ${customerPNColOnReport}
#        ...             ${subSegmentColOnReport}
#        ...             ${programColOnReport}
#        ...             ${applicationColOnReport}
#        ...             ${functionColOnReport}
#        Append To List    ${table}   ${rowOnTable}
#        ${rowOnTable}   Remove Values From List    ${rowOnTable}
#    END
#    Close All Excel Documents
#
#    [Return]    ${table}
#
#Verify The Document Number Of Opp On Master Opp Report
#    [Arguments]     ${reportFilePath}   ${ssFilePath}
#    ${result}   Set Variable    ${True}
#    @{listOfOppsOnReport}   Create List
#    @{listOfOppsOnNS}       Create List
#
#    ${listOfOppsOnReport}   Get List Of Opps From The Master Opp Report    ${reportFilePath}
#    ${listOfOppsOnNS}       Get List Of Opps From The SS Of Master Opp Report On NS    ${ssFilePath}
#
#    File Should Exist    ${RESULT_FILE_PATH}
#    Open Excel Document    ${RESULT_FILE_PATH}    MasterOppResult
#
#    FOR    ${oppOnNS}    IN    @{listOfOppsOnNS}
#        ${posOfOppInReport}     Set Variable    0
#        ${numOfOppsOnReport}    Get Length    ${listOfOppsOnReport}
#
#        FOR    ${oppOnReport}    IN    @{listOfOppsOnReport}
#            IF    '${oppOnReport}' == '${oppOnNS}'
#                 Remove From List    list_=${listOfOppsOnReport}    index=${posOfOppInReport}
#                 BREAK
#            END
#            ${posOfOppInReport}     Evaluate    ${posOfOppInReport}+1
#        END
#        IF    ${posOfOppInReport} == ${numOfOppsOnReport}
#             ${result}   Set Variable    ${False}
#             ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#             ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#             Write Excel Cell    row_num=${nextRow}    col_num=1    value=OPP
#             Write Excel Cell    row_num=${nextRow}    col_num=2    value=${EMPTY}
#             Write Excel Cell    row_num=${nextRow}    col_num=3    value=${EMPTY}
#             Write Excel Cell    row_num=${nextRow}    col_num=4    value=${oppOnNS}
#             Save Excel Document    ${RESULT_FILE_PATH}
#        END
#    END
#    Close All Excel Documents
#    [Return]    ${result}
#
#Verify The Number Of Opps On Master Opp Report
#    [Arguments]     ${reportFilePath}   ${ssFilePath}
#    ${result}   Set Variable    ${True}
#    @{listOfOppsOnReport}   Create List
#    @{listOfOppsOnNS}       Create List
#
#    ${listOfOppsOnReport}   Get List Of Opps From The Master Opp Report                ${reportFilePath}
#    ${listOfOppsOnNS}       Get List Of Opps From The SS Of Master Opp Report On NS    ${ssFilePath}
#    ${numOfOppsOnReport}    Get Length    ${listOfOppsOnReport}
#    ${numOfOppsOnNS}        Get Length    ${listOfOppsOnNS}
#
#    IF    ${numOfOppsOnReport} != ${numOfOppsOnNS}
#         ${result}      Set Variable    ${False}
#         File Should Exist    ${RESULT_FILE_PATH}
#         Open Excel Document    ${RESULT_FILE_PATH}    MasterOppResult
#         ${latestRowInResultFile}   Get Number Of Rows In Excel    ${RESULT_FILE_PATH}
#         ${nextRow}     Evaluate    ${latestRowInResultFile}+1
#         Write Excel Cell    row_num=${nextRow}    col_num=1    value=Number of Opps
#         Write Excel Cell    row_num=${nextRow}    col_num=2    value=${EMPTY}
#         Write Excel Cell    row_num=${nextRow}    col_num=3    value=${numOfOppsOnReport}
#         Write Excel Cell    row_num=${nextRow}    col_num=4    value=${numOfOppsOnNS}
#         Save Excel Document    ${RESULT_FILE_PATH}
#    END
#
#    Close All Excel Documents
#    [Return]    ${result}
#




#Get List Of Opps Have Multi Items From The Master Opp Report
#    [Arguments]     ${reportFilePath}
#    @{listOfOpps}   Create List
#
#    Open Excel Document    ${reportFilePath}    MasterOppReport
#    ${numOfRows}    Get Number Of Rows In Excel    ${reportFilePath}
#    FOR    ${rowIndex}    IN RANGE    5    ${numOfRows}+1
#        ${currentOpp}   Read Excel Cell    ${rowIndex}    1
#        ${nextRow}  Evaluate    ${rowIndex}+1
#        ${nextOpp}      Read Excel Cell    ${nextRow}    1
#        IF    '${nextOpp}' == '${currentOpp}'
#             Append To List    ${listOfOpps}    ${currentOpp}
#        END
#    END
#    Close All Excel Documents
#    ${listOfOpps}   Remove Duplicates    ${listOfOpps}
#    Sort List    ${listOfOpps}
#    [Return]    ${listOfOpps}
#
#Get List Of Opps Have Multi Items From The Master Opp Source
#    [Arguments]     ${ssFilePath}
#    @{table}     Create List
#    ${sumQty}    Set Variable    0
#
#    File Should Exist    ${ssFilePath}
#    Open Excel Document    ${ssFilePath}    MasterOppSource
#
#    ${numOfRows}    Get Number Of Rows In Excel    ${ssFilePath}
#    ${isOPPHaveMultilItems}     Set Variable    ${False}
#
#    FOR    ${rowIndex}    IN RANGE    2    ${numOfRows}+1
#        ${oppCol}    Read Excel Cell    ${rowIndex}    1
#        ${pnCol}     Read Excel Cell    ${rowIndex}    11
#        ${qtyCol}    Read Excel Cell    ${rowIndex}    12
#
#        ${sumQty}      Evaluate    ${sumQty}+${qtyCol}
#
#        IF    ${rowIndex} < ${numOfRows}
#            ${nextRow}         Evaluate           ${rowIndex}+1
#            ${nextOppCol}      Read Excel Cell    ${nextRow}    1
#            ${nextPNCol}       Read Excel Cell    ${nextRow}    11
#
#            IF    '${oppCol}' == '${nextOppCol}'
#                ${isOPPHaveMultilItems}     Set Variable    ${True}
#            END
#
#            IF    '${pnCol}' == '${nextPNCol}'
#                 Continue For Loop
#            END
#        END
#
#        IF    '${isOPPHaveMultilItems}' == '${True}'
#             ${rowOnTable}   Create List
#             ...             ${oppCol}
#             ...             ${pnCol}
#             ...             ${sumQty}
#             Append To List    ${table}   ${rowOnTable}
#
#        END
#        ${sumQty}    Set Variable    0
#        IF    '${oppCol}' != '${nextOppCol}'
#             ${isOPPHaveMultilItems}     Set Variable    ${False}
#        END
#
#    END
#    Close All Excel Documents
#
#    [Return]    ${table}
#
#Write The Master Opp Table To Excel
#    [Arguments]     ${table}    ${outputFilePath}
#
#    ${numOfRowsOnTable}     Get Length    ${table}
#    Open Excel Document    ${outputFilePath}    OutputFile
#    FOR    ${rowIndexOnTable}    IN RANGE    0    ${numOfRowsOnTable}
#        ${oppColOnTable}    Set Variable    ${table}[${rowIndexOnTable}][0]
#        ${rowIndexTemp}    Evaluate    ${rowIndexOnTable}+2
#        Write Excel Cell    row_num=${rowIndexTemp}    col_num=1    value=${oppColOnTable}
#        Save Excel Document    ${outputFilePath}
#    END
#    Close All Excel Documents


