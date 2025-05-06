*** Settings ***
Resource    ../../Pages/Reports/MarginPageV2.robot


*** Test Cases ***
Verify Revenue QTY on Margin report
    File Should Exist      path=${TEST_DATA_FOR_MARGIN_FILE}
    Open Excel Document    filename=${TEST_DATA_FOR_MARGIN_FILE}    doc_id=TestDataForMargin
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${TEST_DATA_FOR_MARGIN_FILE}    sheetName=Revenue
    FOR    ${rowIndex}    IN RANGE    2    ${numOfRows}+1
        ${year}     Read Excel Cell    row_num=${rowIndex}    col_num=1
        ${quarter}  Read Excel Cell    row_num=${rowIndex}    col_num=2
        Comparing Data For Every PN Between Margin And SS RCD    transType=REVENUE    attribute=QTY    year=${year}    quarter=${quarter}    nameOfColOnSSRCD=REV QTY
    END
    Close Current Excel Document


