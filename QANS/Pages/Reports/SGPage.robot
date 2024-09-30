*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${testResultOfSGReportByOEMGroupFilePath}   C:\\RobotFramework\\Results\\SGReport\\SGReportResultByOEMGroup.xlsx
${testResultOfSGReportByPNFilePath}         C:\\RobotFramework\\Results\\SGReport\\SGReportResultByPN.xlsx

*** Keywords ***
Convert SS RCD To Pivot And Export To Excel
    [Arguments]     ${ssRCDFilePath}    ${ssRCDForPivotFilePath}    ${year}     ${quarter}

    @{listParentClass}  Create List     COMPONENTS      MEM     STORAGE     NI ITEMS
    ${startRow}     Set Variable    2

    ${quarter}  Set Variable    Q${quarter}-${year}
    File Should Exist    ${ssRCDFilePath}
    Open Excel Document    ${ssRCDFilePath}    doc_id=SSRCD
    ${numOfRowsOnSSRCD}  Get Number Of Rows In Excel    ${ssRCDFilePath}
#    File Should Exist    ${ssRCDForPivotFilePath}
#    Open Excel Document    ${ssRCDForPivotFilePath}    doc_id=SSRCDForPivot
#
#    Switch Current Excel Document    doc_id=SSRCD
    FOR    ${rowIndexOnSSRCD}    IN RANGE    ${startRow}    ${numOfRowsOnSSRCD}+1
        ${oemGroupCol}            Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
        ${parentClassCol}         Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=9
        ${pnCol}                  Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=11
        ${revQtyCol}              Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=29

        ${sumREVQty}    Set Variable    ${revQtyCol}
        Log To Console    OEM:${oemGroupCol};PN:${pnCol};REVQTY11:${sumREVQty}
        IF    '${parentClassCol}' in ${listParentClass}
            FOR    ${rowIndexTemp}    IN RANGE    ${startRow}+1    ${numOfRowsOnSSRCD}+1
                      ${idTemp}                     Read Excel Cell    row_num=${rowIndexTemp}       col_num=1
                      ${oemGroupColTemp}            Read Excel Cell    row_num=${rowIndexTemp}       col_num=2
                      ${pnColTemp}                  Read Excel Cell    row_num=${rowIndexTemp}       col_num=11
                      ${quarterColTemp}             Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
                      ${revQtyColTemp}              Read Excel Cell    row_num=${rowIndexTemp}       col_num=29
                      IF    '${oemGroupColTemp}' == '${oemGroupCol}' and '${pnColTemp}' == '${pnCol}' and '${quarterColTemp}' == '${quarter}'
                           Log To Console    OEM:${oemGroupCol};PN:${pnCol};REVQTY:${revQtyColTemp}; Quarter:${quarter};quarterColTemp: ${quarterColTemp}; ID:${idTemp}
                           ${sumREVQty}     Evaluate    ${sumREVQty}+${revQtyColTemp}
                      END
            END
        ELSE
           Continue For Loop
        END
#        Log To Console    OEM:${oemGroupCol};PN:${pnCol};REVQTY:${sumREVQty}

    END
    Close All Excel Documents


Write The Test Result Of SG Report By OEM Group To Excel
    [Arguments]     ${oemGroup}     ${valueOnSGReport}   ${valueOnNS}
    File Should Exist    ${testResultOfSGReportByOEMGroupFilePath}
    Open Excel Document    ${testResultOfSGReportByOEMGroupFilePath}    doc_id=SGReportResult
    ${latestRowInSGReportResultFile}   Get Number Of Rows In Excel    ${testResultOfSGReportByOEMGroupFilePath}
    ${nextRow}    Evaluate    ${latestRowInSGReportResultFile}+1    
    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${oemGroup}
    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${valueOnSGReport}
    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${valueOnNS}
    Save Excel Document    ${testResultOfSGReportByOEMGroupFilePath}
    Close Current Excel Document

Check Data For Every Quarter By OEM Group
    [Arguments]     ${sgFilePath}   ${ssRCDFilePath}     ${year}     ${quarter}     ${attribute}    ${valueType}
    @{listOfOEMGroupOnSSRCD}    Create List
    @{listOfOEMGroupOnSG}       Create List
#    ${listOfOEMGroupOnSSRCD}    Get List Of OEM Groups From SS RCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarter}    attribute=${attribute}
    ${listOfOEMGroupOnSG}       Get List Of OEM Groups From SG        sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}   attribute=${attribute}   valueType=${valueType}

#    ${quarterOnSSRCD}  Set Variable    Q${quarter}-${year}
#    FOR    ${oemGroupOnSSRCD}    IN    @{listOfOEMGroupOnSSRCD}
#        ${hasOEMGroupOnSG}  Set Variable    ${False}
#        ${valueOnSSRCD}    Get Value By OEM Group From SSRCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarterOnSSRCD}    oemGroup=${oemGroupOnSSRCD}    valueType=REV
#        ${oemGroupOnSSRCD}  Convert To Upper Case    ${oemGroupOnSSRCD}
#        FOR    ${oemGroupOnSG}    IN    @{listOfOEMGroupOnSG}
#            IF    '${oemGroupOnSSRCD}' == '${oemGroupOnSG}'
#                 ${valueOnSG}   Get Value By OEM Group From SG    sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}    oemGroup=${oemGroupOnSG}    valueType=REV
#                 ${hasOEMGroupOnSG}     Set Variable    ${True}
#                 ${valueOnSG}   Remove String    ${valueOnSG}   $   ,
#                 ${valueOnSSRCD}    Convert To Integer    ${valueOnSSRCD}
#                 ${valueOnSG}       Convert To Integer    ${valueOnSG}
#                 ${diff}    Evaluate    abs(${valueOnSSRCD}-${valueOnSG})
#                 IF    ${diff} > 3
#                      Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSSRCD}    valueOnSGReport=${valueOnSG}    valueOnNS=${valueOnSSRCD}
#                 END
#                 BREAK
#            END
#        END
#        IF    '${hasOEMGroupOnSG}' == '${False}'
#             Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSSRCD}    valueOnSGReport=${EMPTY}    valueOnNS=${valueOnSSRCD}
#        END
#    END
#
#    FOR    ${oemGroupOnSG}    IN    @{listOfOEMGroupOnSG}
#        ${hasOEMGroupOnSSRCD}   Set Variable    ${False}
#        FOR    ${oemGroupOnSSRCD}    IN    @{listOfOEMGroupOnSSRCD}
#            ${oemGroupOnSSRCD}  Convert To Upper Case    ${oemGroupOnSSRCD}
#            IF    '${oemGroupOnSG}' == '${oemGroupOnSSRCD}'
#                 ${hasOEMGroupOnSSRCD}  Set Variable    ${True}
#                 BREAK
#            END
#        END
#        IF    '${hasOEMGroupOnSSRCD}' == '${False}'
#             ${valueOnSG}   Get Value By OEM Group From SG    sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}    oemGroup=${oemGroupOnSG}    valueType=REV
#             Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSG}    valueOnSGReport=${valueOnSG}    valueOnNS=${EMPTY}
#        END
#    END

Get Value By OEM Group From SG
    [Arguments]     ${sgFilePath}    ${year}     ${quarter}   ${oemGroup}    ${valueType}
    ${value}    Set Variable    0

    File Should Exist    ${sgFilePath}
    Open Excel Document    filename=${sgFilePath}    doc_id=SG
    ${numOfRowsOnSG}     Get Number Of Rows In Excel    ${sgFilePath}

    IF    '${valueType}' == 'REV'
         ${searchStr}    Set Variable    ${year}.Q${quarter} R
    END
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfCol}  Get Position Of Column    ${sgFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    IF    '${valueType}' == 'REV'
         ${posOfCol}  Evaluate    ${posOfCol}+2
    END
    
    FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
        ${valCol}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfCol}
        IF    '${oemGroup}' == '${oemGroupCol}'
             ${value}   Set Variable    ${valCol}
        END
    END

    Close All Excel Documents

    [Return]    ${value}

Get Value By OEM Group From SSRCD
    [Arguments]     ${ssRCDFilePath}    ${year}     ${quarter}   ${oemGroup}    ${valueType}
    ${value}    Set Variable    0

    File Should Exist    ${ssRCDFilePath}
    Open Excel Document    filename=${ssRCDFilePath}    doc_id=SSRCD
    ${numOfRowsOnSSRCD}     Get Number Of Rows In Excel    ${ssRCDFilePath}

    FOR    ${rowIndexOnSSRCD}    IN RANGE    2    ${numOfRowsOnSSRCD}+1
        ${oemGroupCol}         Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
        ${parentClassCol}      Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=9
        ${yearCol}             Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=17
        ${quarterCol}          Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18

        IF    '${valueType}' == 'REV'
             ${valCol}           Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=30            
        END

        IF    '${oemGroup}' == '${oemGroupCol}' and '${year}' == '${yearCol}' and '${quarter}' == '${quarterCol}'
             IF    '${parentClassCol}' == 'COMPONENTS' or '${parentClassCol}' == 'MEM' or '${parentClassCol}' == 'STORAGE' or '${parentClassCol}' == 'NI ITEMS'
                  ${value}   Evaluate    ${value}+${valCol}
             END
        END
    END

    Close All Excel Documents

    [Return]    ${value}

Check Parameter Year
    [Arguments]     ${year}
    ${result}   Set Variable    ${True}

    [Return]    ${result}

Get List Of OEM Groups From SG
    [Arguments]     ${sgFilePath}    ${year}     ${quarter}     ${attribute}    ${valueType}
    @{listOfOEMGroup}   Create List

    ${year}     Convert To Number    ${year}
    ${year}     Convert To Integer    ${year}
    ${currentYear}  Get Current Year
    IF    ${year} < 0         
         Fail   The parameter year ${year} is invalid. It must be Integer number. Please contact with Admin!
    END

    ${minYear}  Set Variable    2018
    ${maxYear}  Evaluate    ${currentYear}+1   
    IF    ${year} < ${minYear} or ${year} > ${maxYear}
         Fail   The parameter year ${year} is invalid. The range of parameter year is between ${minYear} and ${maxYear}. Please contact with Admin!
    END

    ${quarter}  Convert To Number    ${quarter}
    ${quarter}  Convert To Integer    ${quarter}
    IF    ${quarter} < 0        
         Fail   The parameter quarter ${quarter} is invalid. It must be Interger number. Please contact with Admin!
    END
    IF    ${quarter} < 1 or ${quarter} > 4
         Fail   The parameter quarter ${quarter} is invalid. The range of quarter is between 1 and 4. Please contact with Admin!
    END

#    File Should Exist    ${sgFilePath}
#    Open Excel Document    filename=${sgFilePath}    doc_id=SG
#    ${numOfRowsOnSG}     Get Number Of Rows In Excel    ${sgFilePath}
#    ${currentYear}  Get Current Year

#    ${searchStr}    Set Variable    ${year}.Q${quarter} R
#    ${rowIndexForSearchStr}     Convert To Number    3
#    ${posOfREVCol}  Get Position Of Column    ${sgFilePath}    ${rowIndexForSearchStr}    ${searchStr}
#    ${posOfREVCol}  Evaluate    ${posOfREVCol}+2
#
#    FOR    ${rownIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
#        ${oemGRoupCol}  Read Excel Cell    row_num=${rownIndexOnSG}    col_num=2
#        ${revCol}       Read Excel Cell    row_num=${rownIndexOnSG}    col_num=${posOfREVCol}
#        IF    '${oemGRoupCol}' != 'None' and '${revCol}' != '${EMPTY}'
#             Append To List    ${listOfOEMGroup}    ${oemGRoupCol}
#        END
#    END
#    Close All Excel Documents
#    [Return]    ${listOfOEMGroup}
    
Get List Of OEM Groups From SS RCD
    [Arguments]     ${ssRCDFilePath}    ${year}     ${quarter}   ${attribute}
    @{listOfOEMGroup}   Create List

    ${quarter}  Set Variable    Q${quarter}-${year}
    File Should Exist    ${ssRCDFilePath}
    Open Excel Document    filename=${ssRCDFilePath}    doc_id=SSRCD
    ${numOfRowsOnSSRCD}     Get Number Of Rows In Excel    ${ssRCDFilePath}

    FOR    ${rowIndexOnSSRCD}    IN RANGE    2    ${numOfRowsOnSSRCD}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
        ${parrentClassCol}  Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=9
        ${yearCol}          Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=17
        ${quarterCol}       Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
        IF    '${attribute}' == 'AMOUNT'
             ${attributeCol}     Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=30
        ELSE IF  '${attribute}' == 'QTY'
             ${attributeCol}     Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=29
        END

        IF    '${year}' == '${yearCol}' and '${quarter}' == '${quarterCol}' and '${attributeCol}' != '0'
            IF    '${parrentClassCol}' == 'COMPONENTS' or '${parrentClassCol}' == 'MEM' or '${parrentClassCol}' == 'STORAGE' or '${parrentClassCol}' == 'NI ITEMS'
                 Append To List    ${listOfOEMGroup}     ${oemGroupCol} 
            END                      
        END
    END
    ${listOfOEMGroup}     Remove Duplicates    ${listOfOEMGroup}
    Close All Excel Documents

    [Return]    ${listOfOEMGroup}

