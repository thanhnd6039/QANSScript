*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${wowChangeResultFilePath}        C:\\RobotFramework\\Results\\WoWChangeReportResult.xlsx
${wowChangeFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${sgWeeklyActionDBPreQFilePath}   C:\\RobotFramework\\Downloads\\SG Weekly Action DB Pre Quarter.xlsx
${sgWeeklyActionDBCurQFilePath}   C:\\RobotFramework\\Downloads\\SG Weekly Action DB Current Quarter.xlsx

*** Keywords ***
Write The Test Result Of WoW Change Report To Excel
    [Arguments]     ${item}     ${oemGroup}     ${valueOnWoWChange}   ${valueOnSGWeeklyActionDB}
    File Should Exist    path=${wowChangeResultFilePath}
    Open Excel Document    filename=${wowChangeResultFilePath}    doc_id=WoWChangeReportResult
    Switch Current Excel Document    doc_id=WoWChangeReportResult
    ${latestRowInWoWchangeResult}   Get Number Of Rows In Excel    ${wowChangeResultFilePath}
    ${nextRow}    Evaluate    ${latestRowInWoWchangeResult}+1
    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${item}
    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oemGroup}
    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${valueOnWoWChange}
    Write Excel Cell    row_num=${nextRow}    col_num=4    value=${valueOnSGWeeklyActionDB}
    Save Excel Document    ${wowChangeResultFilePath}
    Close Current Excel Document

#Check Data For The Strategic Table
#    [Arguments]     ${wowChangeReportFilePath}  ${sgWeeklyActionDBReportFilePath}   ${posOfColOnWoWChangeReport}    ${posOfColOnSGWeeklyActionDBReport}     ${nameOfCol}
#    ${result}   Set Variable    ${True}
#    ${totalDataOnSGWeeklyActionDBReport}    Set Variable    0
#
#    File Should Exist    ${wowChangeReportFilePath}
#    Open Excel Document    ${wowChangeReportFilePath}    doc_id=WoWChangeReport
#
#    File Should Exist    ${sgWeeklyActionDBReportFilePath}
#    Open Excel Document    ${sgWeeklyActionDBReportFilePath}    doc_id=SGWeeklyActionDBReport
#
##   Verify the data for each OEM Group
#    FOR    ${rowIndexOnWoWChangeReport}    IN RANGE    2    6
#        Switch Current Excel Document    doc_id=WoWChangeReport
#        ${oemGroupColOnWoWChangeReport}      Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=1
#        ${dataColOnWoWChangeReport}          Read Excel Cell    row_num=${rowIndexOnWoWChangeReport}    col_num=${posOfColOnWoWChangeReport}
#        Switch Current Excel Document    doc_id=SGWeeklyActionDBReport
#        ${numOfRowsOnSGWeeklyActionDBReport}    Get Number Of Rows In Excel    ${sgWeeklyActionDBReportFilePath}
#        FOR    ${rowIndexOnSGWeeklyActionDBReport}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDBReport}+1
#            ${oemGroupColOnSGWeeklyActionDBReport}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=1
#            ${dataColOnSGWeeklyActionDBReport}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDBReport}    col_num=${posOfColOnSGWeeklyActionDBReport}
#            ${dataColOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDBReport}
#            IF    '${oemGroupColOnWoWChangeReport}' == '${oemGroupColOnSGWeeklyActionDBReport}'
#                 ${totalDataOnSGWeeklyActionDBReport}    Evaluate    ${totalDataOnSGWeeklyActionDBReport}+${dataColOnSGWeeklyActionDBReport}
#                 IF    ${dataColOnWoWChangeReport} != ${dataColOnSGWeeklyActionDBReport}
#                      ${result}     Set Variable    ${False}
#                      Write The Test Result Of WoW Change Report To Excel    ${nameOfCol} for Strategic table    ${oemGroupColOnWoWChangeReport}    ${dataColOnWoWChangeReport}    ${dataColOnSGWeeklyActionDBReport}
#                 END
#                 BREAK
#            END
#        END
#    END
#
#    ${totalDataOnSGWeeklyActionDBReport}   Evaluate  "%.2f" % ${totalDataOnSGWeeklyActionDBReport}
#    Switch Current Excel Document    doc_id=WoWChangeReport
#    ${totalDataOnWoWchangeReport}   Read Excel Cell    row_num=6    col_num=${posOfColOnWoWChangeReport}
#    IF    ${totalDataOnWoWchangeReport} != ${totalDataOnSGWeeklyActionDBReport}
#         ${result}     Set Variable    ${False}
#         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    Strategic Total    ${totalDataOnWoWchangeReport}    ${totalDataOnSGWeeklyActionDBReport}
#    END
#
#    IF    '${result}' == '${False}'
#         Close All Excel Documents
#         Fail   The ${nameOfCol} data between the WoW Change Report and SG Weekly Action Report is different
#    END
#    Close All Excel Documents

Get List Of Sales Member In OEM East Table
    @{listOfSalesMember}    Create List
    Append To List    ${listOfSalesMember}      Chris Seitz
    Append To List    ${listOfSalesMember}      Daniel Schmidt
    Append To List    ${listOfSalesMember}      Eli Tiomkin
    Append To List    ${listOfSalesMember}      Michael Pauser

    [Return]    ${listOfSalesMember}

Get List Of Sales Member In OEM West Table
    @{listOfSalesMember}    Create List
    Append To List    ${listOfSalesMember}      Amy Duong
    Append To List    ${listOfSalesMember}      Caden Douglas
    Append To List    ${listOfSalesMember}      Michael Nilsson
    Append To List    ${listOfSalesMember}      Tiger Wang
    Append To List    ${listOfSalesMember}      Yoda Yasunobu
    Append To List    ${listOfSalesMember}      Nicole Lau
    Append To List    ${listOfSalesMember}      Huan Tran
    Append To List    ${listOfSalesMember}      Dave Beasley
    Append To List    ${listOfSalesMember}      Cameron Sinclair

    [Return]    ${listOfSalesMember}

Get List Of OEM Group Shown In OEM East Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      MELLANOX GROUP
    Append To List    ${listOfOEMGroup}      NOKIA/ALCATEL LUCENT WORLDWIDE
    Append To List    ${listOfOEMGroup}      CURTISS WRIGHT GROUP
    Append To List    ${listOfOEMGroup}      JUNIPER NETWORKS
    Append To List    ${listOfOEMGroup}      ERICSSON WORLDWIDE

    [Return]    ${listOfOEMGroup}

Get List Of OEM Group Shown In OEM West Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      PALO ALTO NETWORKS
    Append To List    ${listOfOEMGroup}      ARISTA
    Append To List    ${listOfOEMGroup}      SCHWEITZER ENGINEERING LABORATORIES (SEL)
    Append To List    ${listOfOEMGroup}      KINEMETRICS INC.
    Append To List    ${listOfOEMGroup}      PANASONIC AVIONICS
    Append To List    ${listOfOEMGroup}      ZTE KANGXUN TELECOM CO. LTD.

    [Return]    ${listOfOEMGroup}

Check Data For The OEM East Table
    [Arguments]     ${wowChangeFilePath}  ${sgWeeklyActionDBFilePath}   ${posOfColOnWoWChange}    ${posOfColOnSGWeeklyActionDB}     ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${listOfSalesMemberInOEMEastTable}   Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange

    File Should Exist      path=${sgWeeklyActionDBFilePath}
    Open Excel Document    filename=${sgWeeklyActionDBFilePath}    doc_id=SGWeeklyActionDB

#   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChange}    IN RANGE    2    7
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        Switch Current Excel Document     doc_id=SGWeeklyActionDB
        ${numOfRowsOnSGWeeklyActionDB}    Get Number Of Rows In Excel    ${sgWeeklyActionDBFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
            ${oemGroupColOnSGWeeklyActionDB}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
            ${dataColOnSGWeeklyActionDB}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}
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

Check Data For The OEM West Table
    [Arguments]     ${wowChangeFilePath}  ${sgWeeklyActionDBFilePath}   ${posOfColOnWoWChange}    ${posOfColOnSGWeeklyActionDB}     ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange

    File Should Exist      path=${sgWeeklyActionDBFilePath}
    Open Excel Document    filename=${sgWeeklyActionDBFilePath}    doc_id=SGWeeklyActionDB

#   Verify the data for each OEM Group
    ${totalOnSGWeeklyActionDB}    Set Variable    0
    FOR    ${rowIndexOnWoWChange}    IN RANGE    11    17
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        Switch Current Excel Document    doc_id=SGWeeklyActionDB
        ${numOfRowsOnSGWeeklyActionDB}    Get Number Of Rows In Excel    ${sgWeeklyActionDBFilePath}
        FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
            ${oemGroupColOnSGWeeklyActionDB}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
            ${dataColOnSGWeeklyActionDB}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}

            IF    '${dataColOnSGWeeklyActionDB}' == 'None'
                 ${dataColOnSGWeeklyActionDB}     Set Variable    0
            END
            ${dataColOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDB}

            IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSGWeeklyActionDB}'
#                 ${totalOnSGWeeklyActionDB}   Evaluate    ${totalOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
                 IF    ${dataColOnWoWChange} != ${dataColOnSGWeeklyActionDB}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${dataColOnWoWChange}    ${dataColOnSGWeeklyActionDB}
                 END
                 BREAK
            END
        END
    END

 #  Verify the Others data
    Switch Current Excel Document    doc_id=SGWeeklyActionDB
    ${othersOnSGWeeklyActionDB}                      Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
        ${oemGroupColOnSGWeeklyActionDB}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDB}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
        ${dataColOnSGWeeklyActionDB}              Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}
        IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMWestTable} and '${oemGroupColOnSGWeeklyActionDB}' not in ${listOfOEMGroupShownInOEMWestTable}
             ${othersOnSGWeeklyActionDB}     Evaluate    ${othersOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
        END
    END

    ${othersOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${othersOnSGWeeklyActionDB}
    Switch Current Excel Document    doc_id=WoWChange
    ${othersOnWoWChange}   Read Excel Cell    row_num=17    col_num=${posOfColOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSGWeeklyActionDB}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West OTHERS    ${othersOnWoWChange}    ${othersOnSGWeeklyActionDB}
    END

    #   Verify the Total data
    Switch Current Excel Document    doc_id=SGWeeklyActionDB
    ${totalOnSGWeeklyActionDB}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
        ${mainSalesRepColOnSGWeeklyActionDB}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
        ${dataColOnSGWeeklyActionDB}              Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}
        IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMWestTable}
             ${totalOnSGWeeklyActionDB}     Evaluate    ${totalOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
        END
    END
    ${totalOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${totalOnSGWeeklyActionDB}
    Switch Current Excel Document    doc_id=WoWChange
    ${totalOnWoWchange}   Read Excel Cell    row_num=18    col_num=${posOfColOnWoWChange}
    IF    ${totalOnWoWchange} != ${totalOnSGWeeklyActionDB}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${totalOnWoWchange}    ${totalOnSGWeeklyActionDB}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} column for the OEM West table between the WoW Change Report and SG Weekly Action Report is different
    END
    Close All Excel Documents

Check The Commit Or Comment Data
    [Arguments]     ${wowChangeFilePath}  ${wowChangeOnVDCFilePath}   ${posOfColOnWoWChange}    ${nameOfCol}   ${table}
    ${result}   Set Variable    ${True}

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}         doc_id=WoWChange

    File Should Exist      path=${wowChangeOnVDCFilePath}
    Open Excel Document    filename=${wowChangeOnVDCFilePath}    doc_id=WoWChangeOnVDC

    IF    '${table}' == 'OEM East'
        FOR    ${rowIndexOnWoWChange}    IN RANGE    2    9
            Switch Current Excel Document    doc_id=WoWChange
            ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
            IF    '${nameOfCol}' == 'TW Commit'
                 IF    '${dataColOnWoWChange}' == 'None'
                      ${dataColOnWoWChange}     Set Variable    ${EMPTY}
                 END
                 IF    '${dataColOnWoWChange}' != '${EMPTY}'
                      ${result}  Set Variable    ${False}
                      IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East OTHERS    ${dataColOnWoWChange}    ${EMPTY}
                      ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${dataColOnWoWChange}    ${EMPTY}
                      ELSE
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${dataColOnWoWChange}    ${EMPTY}
                      END
                 END
                 Continue For Loop
            END
            Switch Current Excel Document    doc_id=WoWChangeOnVDC
            FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    2   9
                ${oemGroupColOnWoWChangeOnVDC}           Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=1
                IF    '${nameOfCol}' == 'LW Commit'
                     ${posOfColOnWoWChangeOnVDC}         Evaluate    ${posOfColOnWoWChange}+1
                     ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfColOnWoWChangeOnVDC}
                ELSE IF  '${nameOfCol}' == 'Comments'
                     ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfColOnWoWChange}
                ELSE
                   Close All Excel Documents
                   Fail  The name of column ${nameOfCol} is invalid
                END

                IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnWoWChangeOnVDC}'
                     IF    '${dataColOnWoWChange}' == 'None'
                          ${dataColOnWoWChange}  Set Variable    ${EMPTY}
                     END
                     IF    '${dataColOnWoWChangeOnVDC}' == 'None'
                          ${dataColOnWoWChangeOnVDC}  Set Variable    ${EMPTY}
                     END

                     IF   '${dataColOnWoWChange}' != '${dataColOnWoWChangeOnVDC}'
                         ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East OTHERS    ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                         ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total     ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                         END
                     END
                     BREAK
                END
            END
        END
    END

    IF    '${table}' == 'OEM West'
        FOR    ${rowIndexOnWoWChange}    IN RANGE    11    19
            Switch Current Excel Document    doc_id=WoWChange
            ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
            IF    '${nameOfCol}' == 'TW Commit'
                 IF    '${dataColOnWoWChange}' == 'None'
                      ${dataColOnWoWChange}     Set Variable    ${EMPTY}
                 END
                 IF    '${dataColOnWoWChange}' != '${EMPTY}'
                      ${result}  Set Variable    ${False}
                      IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West OTHERS    ${dataColOnWoWChange}    ${EMPTY}
                      ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total     ${dataColOnWoWChange}    ${EMPTY}
                      ELSE
                           Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${dataColOnWoWChange}    ${EMPTY}
                      END
                 END
                 Continue For Loop
            END
            Switch Current Excel Document    doc_id=WoWChangeOnVDC
            FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    11   19
                ${oemGroupColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=1
                IF    '${nameOfCol}' == 'LW Commit'
                     ${posOfColOnWoWChangeOnVDC}   Evaluate    ${posOfColOnWoWChange}+1
                     ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfColOnWoWChangeOnVDC}
                ELSE IF  '${nameOfCol}' == 'Comments'
                     ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfColOnWoWChange}
                ELSE
                   Close All Excel Documents
                   Fail  The name of column ${nameOfCol} is invalid
                END

                IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnWoWChangeOnVDC}'
                     IF    '${dataColOnWoWChange}' == 'None'
                          ${dataColOnWoWChange}  Set Variable    ${EMPTY}
                     END
                     IF    '${dataColOnWoWChangeOnVDC}' == 'None'
                          ${dataColOnWoWChangeOnVDC}  Set Variable    ${EMPTY}
                     END
                     IF   '${dataColOnWoWChange}' != '${dataColOnWoWChangeOnVDC}'
                         ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West OTHERS    ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                         ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                         END
                     END
                     BREAK
                END
            END
        END
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} column for the ${table} table between the WoW Change Report and WoW Change Report On VDC is different
    END
    Close All Excel Documents

Check The WoW Data
    [Arguments]     ${wowChangeFilePath}  ${wowChangeOnVDCFilePath}     ${table}    ${posOfColOnWoWChange}    ${nameOfCol}

    ${result}   Set Variable    ${True}

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}         doc_id=WoWChange

    File Should Exist      path=${wowChangeOnVDCFilePath}
    Open Excel Document    filename=${wowChangeOnVDCFilePath}    doc_id=WoWChangeOnVDC

    IF    '${table}' == 'OEM East'
        FOR    ${rowIndexOnWoWChange}    IN RANGE    2    9
            Switch Current Excel Document    doc_id=WoWChange
            ${oemGroupColOnWoWChange}           Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            ${wowColOnWoWChange}                Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
            ${posOfDataColOnWoWChange}          Evaluate    ${posOfColOnWoWChange}-1
            ${dataColOnWoWChange}               Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfDataColOnWoWChange}
            Switch Current Excel Document    doc_id=WoWChangeOnVDC
            FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    2   9
                ${oemGroupColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnWoWChangeOnVDC}'
                    ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfDataColOnWoWChange}
                    ${wowData}  Evaluate    ${dataColOnWoWChange}-${dataColOnWoWChangeOnVDC}
                    ${wowColOnWoWChange}   Evaluate  "%.2f" % ${wowColOnWoWChange}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChange}' == '-0.00'
                         ${wowColOnWoWChange}     Set Variable    0.00
                    END
                    IF    '${wowColOnWoWChange}' != '${wowData}'
                        ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East OTHERS    ${wowColOnWoWChange}    ${wowData}
                         ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total    ${wowColOnWoWChange}    ${wowData}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${wowColOnWoWChange}    ${wowData}
                         END
                    END
                    BREAK
                END
            END
        END
    END

    IF    '${table}' == 'OEM West'
        FOR    ${rowIndexOnWoWChange}    IN RANGE    11    19
            Switch Current Excel Document    doc_id=WoWChange
            ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            ${wowColOnWoWChange}               Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
            ${posOfDataColOnWoWChange}         Evaluate    ${posOfColOnWoWChange}-1
            ${dataColOnWoWChange}              Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfDataColOnWoWChange}
            Switch Current Excel Document    doc_id=WoWChangeOnVDC
            FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    11   19
                ${oemGroupColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=1
                IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnWoWChangeOnVDC}'
                    ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfDataColOnWoWChange}
                    ${wowData}  Evaluate    ${dataColOnWoWChange}-${dataColOnWoWChangeOnVDC}
                    ${wowColOnWoWChange}   Evaluate  "%.2f" % ${wowColOnWoWChange}
                    ${wowData}   Evaluate  "%.2f" % ${wowData}
                    IF    '${wowColOnWoWChange}' == '-0.00'
                         ${wowColOnWoWChange}     Set Variable    0.00
                    END
                    IF    '${wowColOnWoWChange}' != '${wowData}'
                        ${result}  Set Variable    ${False}
                         IF    '${oemGroupColOnWoWChange}' == 'OTHERS'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West OTHERS    ${wowColOnWoWChange}    ${wowData}
                         ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total    ${wowColOnWoWChange}    ${wowData}
                         ELSE
                              Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${wowColOnWoWChange}    ${wowData}
                         END
                    END
                    BREAK
                END
            END
        END
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END
    Close All Excel Documents

Check The GAP Data
    [Arguments]     ${wowChangeOnVDCFilePath}    ${sgWeeklyActionDBFilePath}     ${wowChangeFilePath}    ${table}    ${posOfColOnWoWChange}    ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    File Should Exist      path=${wowChangeOnVDCFilePath}
    Open Excel Document    filename=${wowChangeOnVDCFilePath}      doc_id=WoWChangeOnVDC

    File Should Exist      path=${sgWeeklyActionDBFilePath}
    Open Excel Document    filename=${sgWeeklyActionDBFilePath}    doc_id=SGWeeklyActionDB

    Switch Current Excel Document    doc_id=SGWeeklyActionDB
    ${numOfRowsOnSGWeeklyActionDB}    Get Number Of Rows In Excel    ${sgWeeklyActionDBFilePath}

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange

    IF    '${table}' == 'OEM East'
        FOR    ${rowIndexOnWoWChange}    IN RANGE    2    9
            Switch Current Excel Document    doc_id=WoWChange
            ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            ${gapColOnWoWChange}               Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
            ${gapColOnWoWChange}   Evaluate  "%.2f" % ${gapColOnWoWChange}

            ${los}      Set Variable    0
            Switch Current Excel Document    doc_id=SGWeeklyActionDB
            IF  '${oemGroupColOnWoWChange}' == 'OTHERS'
                 FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
                     ${oemGroupColOnSGWeeklyActionDB}        Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
                     ${mainSalesRepColOnSGWeeklyActionDB}    Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
                     ${losColOnSGWeeklyActionDB}             Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=7
                     IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMEastTable}
                          IF    '${oemGroupColOnSGWeeklyActionDB}' not in ${listOfOEMGroupShownInOEMEastTable}
                                ${los}   Evaluate    ${los}+${losColOnSGWeeklyActionDB}
                          END
                     END
                 END
            ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
                    ${oemGroupColOnSGWeeklyActionDB}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
                    ${mainSalesRepColOnSGWeeklyActionDB}  Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
                    ${losColOnSGWeeklyActionDB}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=7
                    IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMEastTable}
                          ${los}   Evaluate    ${los}+${losColOnSGWeeklyActionDB}
                    END
                END
            ELSE
                FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
                     ${oemGroupColOnSGWeeklyActionDB}     Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
                     ${losColOnSGWeeklyActionDB}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=7
                     IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSGWeeklyActionDB}'
                         ${los}     Set Variable    ${losColOnSGWeeklyActionDB}
                         BREAK
                     END
                 END
            END

            Switch Current Excel Document    doc_id=WoWChangeOnVDC
            ${commit}   Set Variable    0
            FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    2    9
                ${oemGroupColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=1
                ${twCommitColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=5

                IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnWoWChangeOnVDC}'
                    ${commit}   Set Variable    ${twCommitColOnWoWChangeOnVDC}
                    BREAK
                END
            END
            ${gapByFormular}    Evaluate    ${los}-${commit}
            ${gapByFormular}   Evaluate  "%.2f" % ${gapByFormular}

            IF    '${gapColOnWoWChange}' != '${gapByFormular}'
                  ${result}   Set Variable    ${False}
                  IF    '${oemGroupColOnWoWChange}' == 'Total'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East Total     ${gapColOnWoWChange}    ${gapByFormular}
                  ELSE IF  '${oemGroupColOnWoWChange}' == 'OTHERS'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM East OTHERS     ${gapColOnWoWChange}    ${gapByFormular}
                  ELSE
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${gapColOnWoWChange}    ${gapByFormular}
                  END
            END
        END
    END

    IF    '${table}' == 'OEM West'
        FOR    ${rowIndexOnWoWChange}    IN RANGE    11    19
            Switch Current Excel Document    doc_id=WoWChange
            ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
            ${gapColOnWoWChange}               Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
            ${gapColOnWoWChange}   Evaluate  "%.2f" % ${gapColOnWoWChange}

            ${los}      Set Variable    0
            Switch Current Excel Document    doc_id=SGWeeklyActionDB
            IF  '${oemGroupColOnWoWChange}' == 'OTHERS'
                 FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
                     ${oemGroupColOnSGWeeklyActionDB}        Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
                     ${mainSalesRepColOnSGWeeklyActionDB}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
                     ${losColOnSGWeeklyActionDB}                   Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=7
                     IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMWestTable} and '${oemGroupColOnSGWeeklyActionDB}' not in ${listOfOEMGroupShownInOEMWestTable}
                            ${los}   Evaluate    ${los}+${losColOnSGWeeklyActionDB}
                     END

                 END
            ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
                    ${oemGroupColOnSGWeeklyActionDB}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
                    ${mainSalesRepColOnSGWeeklyActionDB}   Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
                    ${losColOnSGWeeklyActionDB}            Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=7

                    IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMWestTable}
                          ${los}   Evaluate    ${los}+${losColOnSGWeeklyActionDB}
                    END
                END
            ELSE
                FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
                     ${oemGroupColOnSGWeeklyActionDB}     Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
                     ${losColOnSGWeeklyActionDB}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=7
                     IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSGWeeklyActionDB}'
                         ${los}     Set Variable    ${losColOnSGWeeklyActionDB}
                         BREAK
                     END
                 END
            END

            Switch Current Excel Document    doc_id=WoWChangeOnVDC
            ${commit}   Set Variable    0
            FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    11    19
                ${oemGroupColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=1
                ${twCommitColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=5

                IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnWoWChangeOnVDC}'
                    ${commit}   Set Variable    ${twCommitColOnWoWChangeOnVDC}
                    BREAK
                END
            END
            ${gapByFormular}    Evaluate    ${los}-${commit}
            ${gapByFormular}   Evaluate  "%.2f" % ${gapByFormular}

            IF    '${gapColOnWoWChange}' != '${gapByFormular}'
                  ${result}   Set Variable    ${False}
                  IF    '${oemGroupColOnWoWChange}' == 'Total'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West Total      ${gapColOnWoWChange}    ${gapByFormular}
                  ELSE IF  '${oemGroupColOnWoWChange}' == 'OTHERS'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    OEM West OTHERS     ${gapColOnWoWChange}    ${gapByFormular}
                  ELSE
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${gapColOnWoWChange}    ${gapByFormular}
                  END
            END
        END
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END
    Close All Excel Documents