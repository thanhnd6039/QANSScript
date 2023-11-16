*** Settings ***
Resource    ../CommonPage.robot
Resource    ../NS/LoginPage.robot
Resource    ../NS/SaveSearchPage.robot

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
    The Title Of Save Search Should Contain    Master Opps





