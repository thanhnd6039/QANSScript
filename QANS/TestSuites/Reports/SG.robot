*** Settings ***
Suite Setup     Setup Test Environment For SG Report    browser=firefox
Resource    ../../Pages/Reports/SGPage.robot

*** Test Cases ***
Verify QTY of Revenue on SG Report
    [Tags]  SG_0001
    [Documentation]     Verify QTY of Revenue on SG report
    
    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}     doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=REVENUE    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=REV QTY
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}    Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "REVENUE-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   QTY of Revenue is different between SG report and SS Revenue Cost Dump            
        END
    END   
    Close All Excel Documents

Verify Amount of Revenue on SG Report
    [Tags]  SG_0002
    [Documentation]     Verify Amount of Revenue on SG report
    
    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=REVENUE    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=REV AMOUNT
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "REVENUE-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Amount of Revenue is different between SG report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify QTY of Backlog on SG Report
    [Tags]  SG_0003
    [Documentation]     Verify QTY of Backlog on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Backlog
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL QTY
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   QTY of Backlog is different between SG report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify Amount of Backlog on SG Report
    [Tags]  SG_0004
    [Documentation]     Verify Amount of Backlog on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}   sheetName=Backlog
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL AMOUNT
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Amount of Backlog is different between SG report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify QTY of Backlog Forecast on SG Report
    [Tags]  SG_0005
    [Documentation]     Verify QTY of Backlog Forecast on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Backlog Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG FORECAST    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL FC QTY
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG FORECAST-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   QTY of Backlog Forecast is different between SG report and SS Revenue Cost Dump             
        END
    END
    Close All Excel Documents

Verify Amount of Backlog Forecast on SG Report
    [Tags]  SG_0006
    [Documentation]     Verify Amount of Backlog Forecast on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Backlog Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG FORECAST    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL FC AMOUNT
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BACKLOG FORECAST-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Amount of Backlog Forecast is different between SG report and SS Revenue Cost Dump            
        END
    END
    Close All Excel Documents

Verify QTY of Customer Forecast on SG Report
    [Tags]  SG_0007
    [Documentation]     Verify QTY of Customer Forecast on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Customer Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=CUSTOMER FORECAST    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=FORECAST QTY
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "CUSTOMER FORECAST-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   QTY of Customer Forecast is different between SG report and SS Revenue Cost Dump             
        END
    END
    Close All Excel Documents

Verify Amount of Customer Forecast on SG Report
    [Tags]  SG_0008
    [Documentation]     Verify Amount of Customer Forecast on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Customer Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=CUSTOMER FORECAST    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=FORECAST
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "CUSTOMER FORECAST-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Amount of Customer Forecast is different between SG report and SS Revenue Cost Dump             
        END
    END
    Close All Excel Documents

Verify QTY of Budget on SG Report
    [Tags]  SG_0009
    [Documentation]     Verify QTY of Budget on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Budget
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS Approved SF    transType=BUDGET    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSApprovedSF=QTY
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BUDGET-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   QTY of Budget is different between SG report and SS Approved SF            
        END
    END
    Close All Excel Documents

Verify Amount of Budget on SG Report
    [Tags]  SG_0010
    [Documentation]     Verify Amount of Budget on SG report

    File Should Exist      path=${TEST_DATA_FOR_SG_FILE_PATH}
    Open Excel Document    filename=${TEST_DATA_FOR_SG_FILE_PATH}    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_SG_FILE_PATH}    sheetName=Budget
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS Approved SF    transType=BUDGET    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSApprovedSF=Sales FC
    END
    Open Excel Document    filename=${SG_RESULT_FILE_PATH}    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${SG_RESULT_FILE_PATH}
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "BUDGET-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Close All Excel Documents
             Fail   Amount of Budget is different between SG report and SS Approved SF             
        END
    END
    Close All Excel Documents

