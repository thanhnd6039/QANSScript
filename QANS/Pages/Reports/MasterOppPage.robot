*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${chkNullCreatedFrom}       //*[@id='ReportViewerControl_ctl04_ctl05_cbNull']
${txtTitleOfMasterOpp}      //*[contains(text(),'Master')]
${lstOppStageFilter}              //*[@id='ReportViewerControl_ctl04_ctl29_txtValue']
${chkSelectAllOnOppStageFilter}     //*[@id='ReportViewerControl_ctl04_ctl29_divDropDown_ctl00']

*** Keywords ***
Navigate To Master Opp Report
    ${configFileObject}     Load Json From File    ${CONFIG_FILE}
    ${username}             Get Value From Json    ${configFileObject}    $.accounts[0].username
    ${username}             Set Variable           ${username}[0]
    ${pass}                 Get Value From Json    ${configFileObject}    $.accounts[0].password
    ${pass}                 Set Variable           ${pass}[0]
    ${url}                  Set Variable           http://${username}:${pass}@report/ReportServer/Pages/ReportViewer.aspx?/NetSuite+Reports/Sales/Opportunity+Report&rs:Command=Render
    Go To    ${url}
    Wait Until Page Load Completed
    Wait Until Element Is Visible    ${txtTitleOfMasterOpp}     ${TIMEOUT}
    Element Text Should Be    ${txtTitleOfMasterOpp}    Master Opportunity Report

Select Opp Stage On Master Opp Report
    [Arguments]     ${oppStage}
    IF    '${oppStage}' == '0.Identified'
         Log To Console    0.Identify
    END
#    Wait Until Element Is Visible    ${lstOppStageFilter}
#    Click Element    ${lstOppStageFilter}
#    Wait Until Element Is Visible    ${chkSelectAllOnOppStageFilter}
#    Click Element    ${chkSelectAllOnOppStageFilter}
#    Click Element    ${chkSelectAllOnOppStageFilter}





