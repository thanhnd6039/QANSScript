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
        END
    END   
    Close All Excel Documents

Verify Revenue Amount on Margin Report
    [Tags]  Margin_0002
    [Documentation]     Verify the Amount data of Revenue on Margin report
    
    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=REVENUE    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=COGS AMOUNT
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "REVENUE-AMOUNT" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Revenue Amount data is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Revenue Cost on Margin Report
    [Tags]  Margin_0003
    [Documentation]     Verify the Cost data of Revenue on Margin report
    
    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=REVENUE    attribute=COST    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=COGS AMOUNT
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "REVENUE-COST" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Revenue Cost data is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Backlog QTY on Margin Report
    [Tags]  Margin_0004
    [Documentation]     Verify the QTY data of Backlog on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}    sheetName=Backlog
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=BACKLOG    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL QTY
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG-QTY" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Backlog QTY data is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Backlog Amount on Margin Report
    [Tags]  Margin_0005
    [Documentation]     Verify the Amount data of Backlog on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}   sheetName=Backlog
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=BACKLOG    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL AMOUNT
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG-AMOUNT" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Backlog Amount data is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Backlog Cost on Margin Report
    [Tags]  Margin_0006
    [Documentation]     Verify the Cost data of Backlog on Margin report
    
    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}   sheetName=Backlog
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=BACKLOG    attribute=COST    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL COST
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG-COST" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Backlog Cost data is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Backlog Forecast QTY on Margin Report
    [Tags]  Margin_0007
    [Documentation]     Verify the QTY data of Backlog Forecast on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}    sheetName=Backlog Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=BACKLOG FORECAST    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL FC QTY
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG FORECAST-QTY" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Backlog Forecast QTY data is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Backlog Forecast Amount on Margin Report
    [Tags]  Margin_0008
    [Documentation]     Verify the Amount data of Backlog Forecast on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}   sheetName=Backlog Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=BACKLOG FORECAST    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL FC AMOUNT
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG FORECAST-AMOUNT" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   The Backlog Forecast Amount data is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents