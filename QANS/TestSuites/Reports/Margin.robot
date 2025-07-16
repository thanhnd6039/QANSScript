*** Settings ***
# Suite Setup     Setup Test Environment For Margin Report    browser=firefox
Resource    ../../Pages/Reports/MarginPage.robot
Library    OperatingSystem

*** Test Cases ***
Verify QTY of Revenue on Margin Report
    [Tags]  Margin_0001
    [Documentation]     Verify QTY of Revenue on Margin report

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
             Fail   QTY of Revenue is different between Margin report and SS Revenue Cost Dump            
        END
    END   
    Close All Excel Documents

Verify Amount of Revenue on Margin Report
    [Tags]  Margin_0002
    [Documentation]     Verify Amount of Revenue on Margin report
    
    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=REVENUE    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=REV AMOUNT
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "REVENUE-AMOUNT" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Amount of Revenue is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Cost of Revenue on Margin Report
    [Tags]  Margin_0003
    [Documentation]     Verify Cost of Revenue on Margin report
    
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
             Fail   Cost of Revenue is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify QTY of Backlog on Margin Report
    [Tags]  Margin_0004
    [Documentation]     Verify QTY of Backlog on Margin report

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
             Fail   QTY of Backlog is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Amount of Backlog on Margin Report
    [Tags]  Margin_0005
    [Documentation]     Verify Amount of Backlog on Margin report

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
             Fail   Amount of Backlog is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Cost of Backlog on Margin Report
    [Tags]  Margin_0006
    [Documentation]     Verify Cost of Backlog on Margin report
    
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
             Fail   Cost of Backlog is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify QTY of Backlog Forecast on Margin Report
    [Tags]  Margin_0007
    [Documentation]     Verify QTY of Backlog Forecast on Margin report

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
             Fail   QTY of Backlog Forecast is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Amount of Backlog Forecast on Margin Report
    [Tags]  Margin_0008
    [Documentation]     Verify Amount of Backlog Forecast on Margin report

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
             Fail   Amount of Backlog Forecast is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Cost of Backlog Forecast on Margin Report
    [Tags]  Margin_0009
    [Documentation]     Verify Cost of Backlog Forecast on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}   sheetName=Backlog Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=BACKLOG FORECAST    attribute=COST    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL FC COST
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG FORECAST-COST" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Cost of Backlog Forecast is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify QTY of Customer Forecast on Margin Report
    [Tags]  Margin_0010
    [Documentation]     Verify QTY of Customer Forecast on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}    sheetName=Customer Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=CUSTOMER FORECAST    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=FORECAST QTY
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "CUSTOMER FORECAST-QTY" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   QTY of Customer Forecast is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Amount of Customer Forecast on Margin Report
    [Tags]  Margin_0011
    [Documentation]     Verify Amount of Customer Forecast on Margin report
   
    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}   sheetName=Customer Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=CUSTOMER FORECAST    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=FORECAST
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "CUSTOMER FORECAST-AMOUNT" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Amount of Customer Forecast is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Cost of Customer Forecast on Margin Report
    [Tags]  Margin_0012
    [Documentation]     Verify Cost of Customer Forecast on Margin report

    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE_PATH}    doc_id=TestDataForMargin
    ${numOfRowsOnTestDataForMargin}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE_PATH}   sheetName=Customer Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForMargin}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=CUSTOMER FORECAST    attribute=COST    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=FORECAST COST
    END
    Open Excel Document    filename=${MARGIN_RESULT_FILE_PATH}    doc_id=MarginResult
    Switch Current Excel Document    doc_id=MarginResult
    ${numOfRowsOnMarginResult}  Get Number Of Rows In Excel    filePath=${MARGIN_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnMarginResult}+1
        ${transTypeColOnMarginResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "CUSTOMER FORECAST-COST" in """${transTypeColOnMarginResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Cost of Customer Forecast is different between Margin report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents