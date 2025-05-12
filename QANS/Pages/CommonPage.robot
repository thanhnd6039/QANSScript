*** Settings ***
Library     SeleniumLibrary
Library     JSONLibrary
Library     ExcelLibrary
Library     Collections
Library     String
Library     XML
Library     DateTime

Library     ../Libs/CExcel.py
Library     ../Libs/COTP.py
Library     ../Libs/CBrowser.py
Library     ../Libs/CDateTime.py
Library    Dialogs


Resource    UtilityPage.robot



*** Variables ***
${CONFIG_DIR}       C:\\RobotFramework\\Config
${CONFIG_FILE}      C:\\RobotFramework\\Config\\Config.json
${CHROMEDRIVER_PATH}    C:\\RobotFramework\\Drivers\\chromedriver.exe
${TIMEOUT}          60s
${OUTPUT_DIR}       ${EXECDIR}
${RESULT_DIR}       C:\\RobotFramework\\Results
${TEST_DATA_FOR_MARGIN_FILE}    ${EXECDIR}\\Resources\\TestData\\TestDataForMarginReport.xlsx

${btnExport}        //*[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink']
${btnViewReport}    //*[@id='ReportViewerControl_ctl04_ctl00']
${lstExcel}         //*/div/a[@title='Excel']



*** Keywords ***
Setup
    [Arguments]     ${browser}
    ${chromeOptions}      Get Chrome Options
    Open Browser    browser=${browser}      options=${chromeOptions}
    Maximize Browser Window

TearDown
    Close Browser

Wait Until Page Load Completed
    FOR    ${count}    IN RANGE    1    61
        ${stage}    Execute Javascript      return document.readyState
        Exit For Loop If    '${stage}' == 'complete'
        Sleep    1s
        IF    ${count} == 60
             Fail   Page is hang or crashed
        END
    END

Navigate To Report
    [Arguments]     ${browser}   ${configFileName}
    ${configFileObject}     Load Json From File    file_name=${CONFIG_DIR}\\SGConfig.json
    ${url}  Get Value From Json    json_object=${configFileObject}    json_path=$.url
    ${url}  Set Variable    ${url[0]}
    ${options}    Get Firefox Options
    Open Browser    url=${url}  browser=${browser}   options=${options}
    Maximize Browser Window
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

Remove All Files in Specified Directory
    [Arguments]     ${dirPath}
    @{fileNames}    List Files In Directory    ${dirPath}
    FOR    ${fileName}    IN    @{fileNames}
        Remove File    ${dirPath}${fileName}
    END

Write Table To Excel
    [Arguments]     ${filePath}     ${listNameOfCols}     ${table}

    File Should Exist      path=${filePath}
    Open Excel Document    filename=${filePath}    doc_id=Table
#    Add Header to File
    ${count}    Set Variable    1
    FOR    ${nameOfCol}    IN    @{listNameOfCols}
        Write Excel Cell    row_num=1    col_num=${count}     value=${nameOfCol}
        ${count}    Evaluate    ${count}+1
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











    

