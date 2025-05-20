*** Settings ***
Library     SeleniumLibrary
Library     JSONLibrary
Library     ExcelLibrary
Library     Collections
Library     String
Library     XML
Library     DateTime
Library    Dialogs

Library     ../Libs/CExcel.py
Library     ../Libs/COTP.py
Library     ../Libs/CBrowser.py
Library     ../Libs/CDateTime.py

Resource    UtilityPage.robot

*** Variables ***
${CONFIG_DIR}       C:\\RobotFramework\\Config
${TIMEOUT}          60s
${OUTPUT_DIR}       C:\\RobotFramework\\Output
${TEST_DATA_DIR}    C:\\RobotFramework\\TestData
${RESULT_DIR}       C:\\RobotFramework\\Results
${TEST_DATA_FOR_MARGIN_FILE}    ${EXECDIR}\\Resources\\TestData\\TestDataForMarginReport.xlsx

${btnExport}        //*[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink']
${btnViewReport}    //*[@id='ReportViewerControl_ctl04_ctl00']
${lstExcel}         //*/div/a[@title='Excel']

*** Keywords ***
Setup
    [Arguments]     ${browser}
    IF    '${browser}' == 'chrome'
         ${options}    Get Chrome Options
    ELSE IF  '${browser}' == 'firefox'
         ${options}     Get Firefox Options
    ELSE
        Fail    The Browser parameter ${browser} is invalid
    END
    
    Open Browser    browser=${browser}     options=${options}
    Maximize Browser Window

TearDown
    Close Browser

Wait Until Page Load Completed
    FOR    ${count}    IN RANGE    1    601
        ${stage}    Execute Javascript      return document.readyState
        Exit For Loop If    '${stage}' == 'complete'
        Sleep    1s
        IF    ${count} == 300
             Fail   Page is hang or crashed
        END
    END

Navigate To Report
    [Arguments]     ${configFileName}
    ${configFileObject}     Load Json From File    file_name=${CONFIG_DIR}\\SGConfig.json
    ${url}  Get Value From Json    json_object=${configFileObject}    json_path=$.url
    ${url}  Set Variable    ${url[0]}
    Go To    url=${url}
    Wait Until Element Is Visible    locator=${btnViewReport}   timeout=${TIMEOUT}
    Wait Until Element Is Enabled    locator=${btnViewReport}   timeout=${TIMEOUT}

Click On Button View Report
    Wait Until Element Is Visible    ${btnViewReport}   ${TIMEOUT}
    Click Element    ${btnViewReport}

Export Report To
    [Arguments]     ${option}
    ${exportOptionXpath}    Set Variable    //*/div/a[@title='${option}']
    Wait Until Element Is Visible    locator=${btnExport}   timeout=${TIMEOUT}
    Click Element    locator=${btnExport}
    Wait Until Element Is Visible    locator=${exportOptionXpath}   timeout=${TIMEOUT}
    Click Element    ${exportOptionXpath}

Open New Tab
    Execute Javascript      window.open('https://www.google.com')

Remove All Files In Specified Directory
    [Arguments]     ${dirPath}
    @{fileNames}    List Files In Directory    ${dirPath}
    FOR    ${fileName}    IN    @{fileNames}
        Remove File    ${dirPath}\\${fileName}
    END

Write Table To Excel
    [Arguments]     ${filePath}     ${listNameOfCols}     ${table}   ${hasHeader}=${True}

    File Should Exist      path=${filePath}
    Open Excel Document    filename=${filePath}    doc_id=Table
#    Add Header to File
    IF    '${hasHeader}' == '${True}'
         ${count}    Set Variable    1
        FOR    ${nameOfCol}    IN    @{listNameOfCols}
            Write Excel Cell    row_num=1    col_num=${count}     value=${nameOfCol}
            ${count}    Evaluate    ${count}+1
        END
    END

    ${latestRow}   Get Number Of Rows In Excel    ${filePath}
    ${nextRow}     Evaluate    ${latestRow}+1
#    Add Raw Data to File
    FOR    ${rawData}    IN    @{table}
        ${count}        Set Variable    0
        ${colIndex}     Set Variable    1
        FOR    ${nameOfCol}    IN    @{listNameOfCols}
            Write Excel Cell    row_num=${nextRow}    col_num=${colIndex}    value=${rawData[${count}]}
            ${count}    Evaluate    ${count}+1
            ${colIndex}     Evaluate    ${colIndex}+1
        END
        ${nextRow}  Evaluate    ${nextRow}+1
    END
    Save Excel Document    ${filePath}
    Close Current Excel Document

Get Fully File Name From Given Name
    [Arguments]     ${givenName}    ${dirPath}
    ${fullyFileName}    Set Variable    ${EMPTY}
    @{files}    List Files In Directory    ${dirPath}
    FOR    ${file}    IN    @{files}
        ${contains}     Evaluate    "${givenName}" in """${file}"""
        IF    '${contains}' == '${True}'
             ${fullyFileName}   Set Variable    ${file}
             Exit For Loop
        END
    END
    [Return]    ${fullyFileName}










    

