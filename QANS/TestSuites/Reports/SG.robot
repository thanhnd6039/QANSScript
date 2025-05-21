*** Settings ***
Suite Setup     Setup Test Environment For SG Report    browser=firefox
Resource    ../../Pages/Reports/SGPageV2.robot

*** Test Cases ***
Verify Revenue QTY on SG Report
    [Tags]  SG_0001
    [Documentation]     Verify the QTY data of Revenue on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=REVENUE    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=REV QTY
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "REVENUE-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Revenue QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Revenue Amount on SG Report
    [Tags]  SG_0002
    [Documentation]     Verify the Amount data of Revenue on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=REVENUE    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=REV AMOUNT
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "REVENUE-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Revenue Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog QTY on SG Report
    [Tags]  SG_0003
    [Documentation]     Verify the QTY data of Backlog on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Backlog
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL QTY
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "BACKLOG-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Backlog QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog Amount on SG Report
    [Tags]  SG_0004
    [Documentation]     Verify the Amount data of Backlog on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Backlog
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL AMOUNT
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "BACKLOG-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Backlog Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog Forecast QTY on SG Report
    [Tags]  SG_0005
    [Documentation]     Verify the QTY data of Backlog Forecast on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Backlog Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG FORECAST    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL FC QTY
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "BACKLOG FORECAST-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Backlog Forecast QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Backlog Forecast Amount on SG Report
    [Tags]  SG_0006
    [Documentation]     Verify the Amount data of Backlog Forecast on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Backlog Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=BACKLOG FORECAST    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=BL FC AMOUNT
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "BACKLOG FORECAST-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Backlog Forecast Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Customer Forecast QTY on SG Report
    [Tags]  SG_0007
    [Documentation]     Verify the QTY data of Customer Forecast on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Customer Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=CUSTOMER FORECAST    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=FORECAST QTY
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "CUSTOMER FORECAST-QTY" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Customer Forecast QTY data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Customer Forecast Amount on SG Report
    [Tags]  SG_0008
    [Documentation]     Verify the Amount data of Customer Forecast on SG report
    File Should Exist      path=${TEST_DATA_DIR}\\TestDataForSG.xlsx
    Open Excel Document    filename=${TEST_DATA_DIR}\\TestDataForSG.xlsx    doc_id=TestDataForSG
    ${numOfRowsOnTestDataForSG}    Get Number Of Rows In Excel    filePath=${TEST_DATA_DIR}\\TestDataForSG.xlsx    sheetName=Customer Forecast
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnTestDataForSG}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between SG And SS RCD    transType=CUSTOMER FORECAST    attribute=AMOUNT    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=FORECAST
    END
    Open Excel Document    filename=${OUTPUT_DIR}\\SGResult.xlsx    doc_id=SGResult
    Switch Current Excel Document    doc_id=SGResult
    ${numOfRowsOnSGResult}  Get Number Of Rows In Excel    filePath=${OUTPUT_DIR}\\SGResult.xlsx
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRowsOnSGResult}+1
        ${transTypeColOnSGResult}   Read Excel Cell    row_num=${rowIndex}    col_num=2
        ${transTypesColIsContainRevenueQTY}    Evaluate    "CUSTOMER FORECAST-AMOUNT" in """${transTypeColOnSGResult}"""
        IF    '${transTypesColIsContainRevenueQTY}' == '${True}'
             Fail   The Customer Forecast Amount data is different between SG report and SS Revenue Cost Dump
             BREAK
        END
    END
    Close All Excel Documents

Verify Budget QTY on SG Report
    [Tags]  SG_0009
    [Documentation]     Verify the QTY data of Budget on SG report
    Comparing Data For Every PN Between SG And SS Approved SF    transType=BUDGET    attribute=QTY    year=2025    quarter=2    nameOfColOnSSApprovedSF=QTY


