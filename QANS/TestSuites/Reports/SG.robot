*** Settings ***
Suite Setup     Setup Test Environment For SG Report    browser=firefox
Resource    ../../Pages/Reports/SGPageV2.robot

*** Test Cases ***
Verify Revenue QTY on SG Report
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





    




