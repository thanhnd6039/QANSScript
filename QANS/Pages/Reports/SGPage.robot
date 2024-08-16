*** Settings ***
Resource    ../CommonPage.robot

*** Keywords ***
Check Data REV For Every Quarter By OEM Group
    [Arguments]     ${sgFilePath}   ${ssRCDFilePath}     ${year}     ${quarter}
    @{listOfOEMGroupOnSSRCD}    Create List
    @{listOfOEMGroupOnSG}       Create List
    ${listOfOEMGroupOnSSRCD}    Get List Of OEM Group By Amount From SS RCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarter}
#    ${listOfOEMGroupOnSG}       Get List Of OEM Group By Amount From SG    sgFilePath=${sgFilePath}    year=${year}    quarter=${quarter}

    FOR    ${oemGroupOnSSRCD}    IN    @{listOfOEMGroupOnSSRCD}
        ${value}    Get Value By OEM Group From SSRCD    ssRCDFilePath=${ssRCDFilePath}    year=${year}    quarter=${quarter}    oemGroup=${oemGroupOnSSRCD}
        Log To Console    OEM Group: ${oemGroupOnSSRCD}; VALUE: ${value}

    END

Get Value By OEM Group From SSRCD
    [Arguments]     ${ssRCDFilePath}    ${year}     ${quarter}   ${oemGroup}
    ${value}    Set Variable    0

    File Should Exist    ${ssRCDFilePath}
    Open Excel Document    filename=${ssRCDFilePath}    doc_id=SSRCD
    ${numOfRowsOnSSRCD}     Get Number Of Rows In Excel    ${ssRCDFilePath}

    FOR    ${rowIndexOnSSRCD}    IN RANGE    2    ${numOfRowsOnSSRCD}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=2
        ${yearCol}          Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=17
        ${quarterCol}       Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=18
        ${amountCol}        Read Excel Cell    row_num=${rowIndexOnSSRCD}    col_num=30
        IF    '${oemGroup}' == '${oemGroup}' and '${year}' == '${yearCol}' and '${quarter}' == '${quarterCol}'
             ${value}   Evaluate    ${value}+${amountCol}
             Log To Console    Value11: ${value}
        END
    END
    Log To Console    END
    Close All Excel Documents

    [Return]    ${value}


Get List Of OEM Group By Amount From SG
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
    
Get List Of OEM Group By Amount From SS RCD
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

