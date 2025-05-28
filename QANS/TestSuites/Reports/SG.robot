*** Settings ***
Suite Setup     Setup Test Environment For SG Report    browser=firefox
Resource    ../../Pages/Reports/SGPage.robot

*** Test Cases ***
Verify Revenue QTY on SG Report
    [Tags]  SG_0001
    [Documentation]     Verify the QTY data of Revenue on SG report
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
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContain}    Evaluate    "REVENUE-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContain}' == '${True}'
             Fail   The Revenue QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Revenue Amount on SG Report
    [Tags]  SG_0002
    [Documentation]     Verify the Amount data of Revenue on SG report
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
             Fail   The Revenue Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog QTY on SG Report
    [Tags]  SG_0003
    [Documentation]     Verify the QTY data of Backlog on SG report
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
             Fail   The Backlog QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog Amount on SG Report
    [Tags]  SG_0004
    [Documentation]     Verify the Amount data of Backlog on SG report
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
             Fail   The Backlog Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog Forecast QTY on SG Report
    [Tags]  SG_0005
    [Documentation]     Verify the QTY data of Backlog Forecast on SG report
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
             Fail   The Backlog Forecast QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog Forecast Amount on SG Report
    [Tags]  SG_0006
    [Documentation]     Verify the Amount data of Backlog Forecast on SG report
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
             Fail   The Backlog Forecast Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Customer Forecast QTY on SG Report
    [Tags]  SG_0007
    [Documentation]     Verify the QTY data of Customer Forecast on SG report
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
             Fail   The Customer Forecast QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Customer Forecast Amount on SG Report
    [Tags]  SG_0008
    [Documentation]     Verify the Amount data of Customer Forecast on SG report
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
             Fail   The Customer Forecast Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Budget QTY on SG Report
    [Tags]  SG_0009
    [Documentation]     Verify the QTY data of Budget on SG report
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
             Fail   The Budget QTY data is different between SG report and SS Approved SF
             BREAK
        END
    END
    Close All Excel Documents

Verify Budget Amount on SG Report
    [Tags]  SG_0010
    [Documentation]     Verify the Amount data of Budget on SG report
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
             Fail   The Budget Amount data is different between SG report and SS Approved SF
             BREAK
        END
    END
    Close All Excel Documents

