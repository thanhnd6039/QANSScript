*** Settings ***
# Suite Setup     Setup Test Environment For Margin Report    browser=firefox
Resource    ../../Pages/Reports/MarginPage.robot

*** Test Cases ***
Verify Revenue QTY on Margin Report
    [Tags]  Margin_0001
    [Documentation]     Verify the QTY data of Revenue on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}     doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=REVENUE    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=REV QTY
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}    Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "REVENUE-QTY" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Revenue QTY data is different between Margin report and SS Revenue Cost Dump
             BREAK
        END
    END
    
    Close All Excel Documents


