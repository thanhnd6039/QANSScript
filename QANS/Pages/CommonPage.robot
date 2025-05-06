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

Resource    UtilityPage.robot


*** Variables ***
${CONFIG_FILE}      C:\\RobotFramework\\Config\\Config.json
${TIMEOUT}          60s
${DOWNLOAD_DIR}     C:\\RobotFramework\\Downloads
${OUTPUT_DIR}       C:\\RobotFramework
${RESULT_DIR}       C:\\RobotFramework\\Results
${btnViewReport}    //*[@id='ReportViewerControl_ctl04_ctl00']
${iconExportDataReport}   //*[@id='ReportViewerControl_ctl05_ctl04_ctl00_ButtonImg']

${TEST_DATA_FOR_MARGIN_FILE}    ${EXECDIR}\\Resources\\TestData\\TestDataForMarginReport.xlsx


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

Click On Button View Report
    Wait Until Element Is Visible    ${btnViewReport}   ${TIMEOUT}
    Click Element    ${btnViewReport}

Export Report Data To
    [Arguments]     ${option}
    ${exportOptionXpath}    Set Variable    //*/div/a[@title='${option}']
    Wait Until Element Is Visible    ${iconExportDataReport}      ${TIMEOUT}
    Click Element    ${iconExportDataReport}
    Wait Until Element Is Visible    ${exportOptionXpath}   ${TIMEOUT}
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







    

