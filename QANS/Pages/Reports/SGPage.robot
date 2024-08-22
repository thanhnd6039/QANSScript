*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${testResultOfSGReportByOEMGroupFilePath}   C:\\RobotFramework\\Results\\SGReport\\SGReportResultByOEMGroup.xlsx
${testResultOfSGReportByPNFilePath}         C:\\RobotFramework\\Results\\SGReport\\SGReportResultByPN.xlsx

*** Keywords ***
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
    [Arguments]     ${sgFilePath}   ${ssRCDFilePath}     ${year}     ${quarter}
    @{listOfOEMGroupOnSSRCD}    Create List
    @{listOfOEMGroupOnSG}       Create List
    ${listOfOEMGroupOnSSRCD}    Get List Of OEM Groups By Amount From SS RCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarter}
    ${listOfOEMGroupOnSG}       Get List Of OEM Groups By Amount From SG        sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}

    ${quarterOnSSRCD}  Set Variable    Q${quarter}-${year}
    FOR    ${oemGroupOnSSRCD}    IN    @{listOfOEMGroupOnSSRCD}
        ${hasOEMGroupOnSG}  Set Variable    ${False}
        ${valueOnSSRCD}    Get Value By OEM Group From SSRCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarterOnSSRCD}    oemGroup=${oemGroupOnSSRCD}    valueType=REV
        FOR    ${oemGroupOnSG}    IN    @{listOfOEMGroupOnSG}
            IF    '${oemGroupOnSSRCD}' == '${oemGroupOnSG}'
                 ${valueOnSG}   Get Value By OEM Group From SG    sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}    oemGroup=${oemGroupOnSG}    valueType=REV
                 ${hasOEMGroupOnSG}     Set Variable    ${True}
#                 ${valueOnSSRCD}   Evaluate  "%.0f" % ${valueOnSSRCD}
                 ${valueOnSG}   Remove String    ${valueOnSG}   $   ,
                 ${valueOnSSRCD}    Convert To Integer    ${valueOnSSRCD}
                 ${valueOnSG}       Convert To Integer    ${valueOnSG}
                 Log To Console    valueOnSSRCD: ${valueOnSSRCD};valueOnSG: ${valueOnSG}
                 IF    ${valueOnSSRCD} == ${valueOnSG}
                      Log To Console    OK
#                      Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSSRCD}    valueOnSGReport=${valueOnSG}    valueOnNS=${valueOnSSRCD}
                 END
                 BREAK
            END            
        END
        IF    '${hasOEMGroupOnSG}' == '${False}'
             Write The Test Result Of SG Report By OEM Group To Excel    oemGroup=${oemGroupOnSSRCD}    valueOnSGReport=${EMPTY}    valueOnNS=${valueOnSSRCD}
        END
    END

#    FOR    ${oemGroupOnSG}    IN    @{listOfOEMGroupOnSG}
#        ${hasOEMGroupOnSSRCD}   Set Variable    ${False}
#        FOR    ${oemGroupOnSSRCD}    IN    @{listOfOEMGroupOnSSRCD}
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
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
        ${yearCol}          Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=17
        ${quarterCol}       Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
        IF    '${valueType}' == 'REV'
             ${valCol}           Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=30            
        END

        IF    '${oemGroup}' == '${oemGroupCol}' and '${year}' == '${yearCol}' and '${quarter}' == '${quarterCol}'
             ${value}   Evaluate    ${value}+${valCol}
        END
    END

    Close All Excel Documents

    [Return]    ${value}


Get List Of OEM Groups By Amount From SG
    [Arguments]     ${sgFilePath}    ${year}     ${quarter}
    @{listOfOEMGroup}   Create List
    
    File Should Exist    ${sgFilePath}
    Open Excel Document    filename=${sgFilePath}    doc_id=SG
    ${numOfRowsOnSG}     Get Number Of Rows In Excel    ${sgFilePath}
    ${searchStr}    Set Variable    ${year}.Q${quarter} R
    ${rowIndexForSearchStr}     Convert To Number    3
    ${posOfREVCol}  Get Position Of Column    ${sgFilePath}    ${rowIndexForSearchStr}    ${searchStr}
    ${posOfREVCol}  Evaluate    ${posOfREVCol}+2

    FOR    ${rownIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
        ${oemGRoupCol}  Read Excel Cell    row_num=${rownIndexOnSG}    col_num=2
        ${revCol}       Read Excel Cell    row_num=${rownIndexOnSG}    col_num=${posOfREVCol}
        IF    '${oemGRoupCol}' != 'None' and '${revCol}' != '${EMPTY}'
             Append To List    ${listOfOEMGroup}    ${oemGRoupCol}
        END
    END
    Close All Excel Documents
    [Return]    ${listOfOEMGroup}
    
Get List Of OEM Groups By Amount From SS RCD
    [Arguments]     ${ssRCDFilePath}    ${year}     ${quarter}
    @{listOfOEMGroup}   Create List

    ${quarter}  Set Variable    Q${quarter}-${year}
    File Should Exist    ${ssRCDFilePath}
    Open Excel Document    filename=${ssRCDFilePath}    doc_id=SSRCD
    ${numOfRowsOnSSRCD}     Get Number Of Rows In Excel    ${ssRCDFilePath}

    FOR    ${rowIndexOnSSRCD}    IN RANGE    2    ${numOfRowsOnSSRCD}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
        ${yearCol}          Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=17
        ${quarterCol}       Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
        ${revAmountCol}     Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=30
        IF    '${year}' == '${yearCol}' and '${quarter}' == '${quarterCol}' and '${revAmountCol}' != '0'
            Append To List    ${listOfOEMGroup}     ${oemGroupCol}           
        END
    END
    ${listOfOEMGroup}     Remove Duplicates    ${listOfOEMGroup}
    Close All Excel Documents

    [Return]    ${listOfOEMGroup}

