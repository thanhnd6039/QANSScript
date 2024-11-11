*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${wowChangeResultFilePath}        C:\\RobotFramework\\Results\\WoWChangeReportResult.xlsx
${wowChangeFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${SGFilePath}                     C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx

*** Keywords ***
Check Data For The OEM East Table
    [Arguments]     ${wowChangeFilePath}  ${SGFilePath}   ${posOfColOnWoWChange}    ${posOfColOnSG}     ${nameOfCol}
    ${result}   Set Variable    ${True}
#    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
#    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange

    File Should Exist      path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG



#   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChange}    IN RANGE    2    7
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        Switch Current Excel Document     doc_id=SG
        ${numOfRowsOnSG}    Get Number Of Rows In Excel    ${SGFilePath}
        FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
            ${oemGroupColOnSG}       Read Excel Cell    row_num=${rowIndexOnSG}    col_num=1
            ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfColOnSGWeeklyActionDB}
            IF    '${dataColOnSGWeeklyActionDB}' == 'None'
                 Continue For Loop
            END
            ${dataColOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDB}
            IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSGWeeklyActionDB}'
                 IF    ${dataColOnWoWChange} != ${dataColOnSGWeeklyActionDB}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupColOnWoWChange}    valueOnWoWChange=${dataColOnWoWChange}    valueOnSGWeeklyActionDB=${dataColOnSGWeeklyActionDB}
                 END
                 BREAK
            END
        END
    END

#   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDB
    ${totalOnSGWeeklyActionDB}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
        ${mainSalesRepColOnSGWeeklyActionDB}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
        ${dataColOnSGWeeklyActionDB}              Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}
        IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMEastTable}
             ${totalOnSGWeeklyActionDB}     Evaluate    ${totalOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
        END
    END
    ${totalOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${totalOnSGWeeklyActionDB}
    Switch Current Excel Document    doc_id=WoWChange
    ${totalOnWoWchange}   Read Excel Cell    row_num=8    col_num=${posOfColOnWoWChange}
    IF    ${totalOnWoWchange} != ${totalOnSGWeeklyActionDB}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${totalOnWoWchange}    ${totalOnSGWeeklyActionDB}
    END

 #  Verify the OTHERS data
    Switch Current Excel Document    doc_id=SGWeeklyActionDB
    ${othersOnSGWeeklyActionDB}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
        ${oemGroupColOnSGWeeklyActionDB}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDB}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
        ${dataColOnSGWeeklyActionDB}              Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}
        IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMEastTable}
             IF    '${oemGroupColOnSGWeeklyActionDB}' not in ${listOfOEMGroupShownInOEMEastTable}
                  ${othersOnSGWeeklyActionDB}     Evaluate    ${othersOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
             END
        END
    END
    ${othersOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${othersOnSGWeeklyActionDB}
    Switch Current Excel Document    doc_id=WoWChange
    ${othersOnWoWChange}   Read Excel Cell    row_num=7    col_num=${posOfColOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSGWeeklyActionDB}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East OTHERS    ${othersOnWoWChange}    ${othersOnSGWeeklyActionDB}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the OEM East table between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents
    