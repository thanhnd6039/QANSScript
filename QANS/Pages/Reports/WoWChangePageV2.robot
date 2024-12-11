*** Settings ***
Resource    ../CommonPage.robot

*** Variables ***
${wowChangeResultFilePath}        C:\\RobotFramework\\Results\\WoWChangeReportResult.xlsx
${wowChangeFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${SGFilePath}                     C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx
${SGWeeklyActionDBFilePath}       C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB.xlsx

*** Keywords ***
Write The Test Result Of WoW Change Report To Excel
    [Arguments]     ${item}     ${oemGroup}     ${valueOnWoWChange}   ${valueOnSG}
    File Should Exist    path=${wowChangeResultFilePath}
    Open Excel Document    filename=${wowChangeResultFilePath}    doc_id=WoWChangeReportResult
    Switch Current Excel Document    doc_id=WoWChangeReportResult
    ${latestRowInWoWchangeResult}   Get Number Of Rows In Excel    ${wowChangeResultFilePath}
    ${nextRow}    Evaluate    ${latestRowInWoWchangeResult}+1
    Write Excel Cell    row_num=${nextRow}    col_num=1    value=${item}
    Write Excel Cell    row_num=${nextRow}    col_num=2    value=${oemGroup}
    Write Excel Cell    row_num=${nextRow}    col_num=3    value=${valueOnWoWChange}
    Write Excel Cell    row_num=${nextRow}    col_num=4    value=${valueOnSG}
    Save Excel Document    ${wowChangeResultFilePath}
    Close Current Excel Document
    
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

Check The Ship, Backlog, LOS Data
    [Arguments]     ${table}    ${nameOfCol}    ${posOfColOnWoWChange}    ${posOfRColOnSG}=0      ${posOfBColOnSG}=0

    ${result}   Set Variable    ${True}
    ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    0
    ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    0
    ${othersRowIndexOnWoWChange}             Set Variable    0
    ${totalRowIndexOnWoWChange}              Set Variable    0

    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange

    File Should Exist      path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG
    Switch Current Excel Document     doc_id=SG
    ${numOfRowsOnSG}    Get Number Of Rows In Excel    ${SGFilePath}
    IF    '${table}' == 'OEM East'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    2
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    7
        ${othersRowIndexOnWoWChange}             Set Variable    7
        ${totalRowIndexOnWoWChange}              Set Variable    8
    ELSE IF     '${table}' == 'OEM West'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    11
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    17
        ${othersRowIndexOnWoWChange}             Set Variable    17
        ${totalRowIndexOnWoWChange}              Set Variable    18
    ELSE
        Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END
    Switch Current Excel Document    doc_id=WoWChange
   #   Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChange}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}    ${endRowIndexForOEMGroupOnWoWChange}
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        Switch Current Excel Document     doc_id=SG
        FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
            ${oemGroupColOnSG}       Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
            IF    '${oemGroupColOnSG}' == 'None'
                 Continue For Loop
            END
            IF    '${nameOfCol}' == 'Ships' or '${nameOfCol}' == 'Pre Q Ships'
                 ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
            ELSE IF  '${nameOfCol}' == 'Backlog'
                 ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
            ELSE IF  '${nameOfCol}' == 'LOS'
                 ${dataRColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
                 ${dataBolOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
                 IF    '${dataRColOnSG}' == 'None'
                     ${dataRColOnSG}   Set Variable    0
                 END
                 IF    '${dataBolOnSG}' == 'None'
                     ${dataBolOnSG}    Set Variable    0
                 END
                 ${dataColOnSG}     Evaluate    ${dataRColOnSG}+${dataBolOnSG}
            ELSE
                 Fail   The nameOfCol parameter ${nameOfCol} is invalid. Please contact with the Administrator for supporting
            END

            IF    '${dataColOnSG}' == 'None'
                 ${dataColOnSG}     Set Variable    0
            END
            ${dataColOnSG}   Evaluate  "%.2f" % ${dataColOnSG}
            IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSG}'
                 IF    ${dataColOnWoWChange} != ${dataColOnSG}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupColOnWoWChange}    valueOnWoWChange=${dataColOnWoWChange}    valueOnSG=${dataColOnSG}
                 END
                 BREAK
            END
        END
    END
    #   Verify the Total data
    Switch Current Excel Document    doc_id=SG
    ${totalOnSG}  Set Variable    0
    FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
        ${mainSalesRepColOnSG}      Read Excel Cell    row_num=${rowIndexOnSG}    col_num=3
        IF    '${nameOfCol}' == 'Ships' or '${nameOfCol}' == 'Pre Q Ships'
             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
        ELSE IF  '${nameOfCol}' == 'Backlog'
             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
        ELSE IF  '${nameOfCol}' == 'LOS'
             ${dataRColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
             ${dataBolOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
             IF    '${dataRColOnSG}' == 'None'
                  ${dataRColOnSG}   Set Variable    0
             END
             IF    '${dataBolOnSG}' == 'None'
                  ${dataBolOnSG}    Set Variable    0
             END
             ${dataColOnSG}     Evaluate    ${dataRColOnSG}+${dataBolOnSG}
        ELSE
             Fail   The nameOfCol parameter ${nameOfCol} is invalid. Please contact with the Administrator for supporting
        END

        IF    '${dataColOnSG}' == 'None'
               ${dataColOnSG}     Set Variable    0
        END

        IF    '${table}' == 'OEM East'
             IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
                   ${totalOnSG}     Evaluate    ${totalOnSG}+${dataColOnSG}
             END
        ELSE IF   '${table}' == 'OEM West'
            IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
                   ${totalOnSG}     Evaluate    ${totalOnSG}+${dataColOnSG}
            END
        ELSE
            Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
        END
    END
    ${totalOnSG}   Evaluate  "%.2f" % ${totalOnSG}
    Switch Current Excel Document    doc_id=WoWChange
    ${totalOnWoWchange}   Read Excel Cell    row_num=${totalRowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
    IF    ${totalOnWoWchange} != ${totalOnSG}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total    ${totalOnWoWchange}    ${totalOnSG}
    END
    #  Verify the OTHERS data
    Switch Current Excel Document    doc_id=SG
    ${othersOnSG}  Set Variable    0
    FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
        ${oemGroupColOnSG}          Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
        ${mainSalesRepColOnSG}      Read Excel Cell    row_num=${rowIndexOnSG}    col_num=3
        IF    '${nameOfCol}' == 'Ships' or '${nameOfCol}' == 'Pre Q Ships'
             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
        ELSE IF  '${nameOfCol}' == 'Backlog'
             ${dataColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
        ELSE IF  '${nameOfCol}' == 'LOS'
             ${dataRColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
             ${dataBolOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}
             IF    '${dataRColOnSG}' == 'None'
                  ${dataRColOnSG}   Set Variable    0
             END
             IF    '${dataBolOnSG}' == 'None'
                  ${dataBolOnSG}    Set Variable    0
             END
             ${dataColOnSG}     Evaluate    ${dataRColOnSG}+${dataBolOnSG}
        ELSE
             Fail   The nameOfCol parameter ${nameOfCol} is invalid. Please contact with the Administrator for supporting
        END

        IF    '${dataColOnSG}' == 'None'
               ${dataColOnSG}     Set Variable    0
        END

        IF    '${table}' == 'OEM East'
             IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
                IF    '${oemGroupColOnSG}' not in ${listOfOEMGroupShownInOEMEastTable}
                   ${othersOnSG}     Evaluate    ${othersOnSG}+${dataColOnSG}
                END
             END
        ELSE IF   '${table}' == 'OEM West'
            IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
                IF    '${oemGroupColOnSG}' not in ${listOfOEMGroupShownInOEMWestTable}
                   ${othersOnSG}     Evaluate    ${othersOnSG}+${dataColOnSG}
                END
            END
        ELSE
            Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
        END

    END
    ${othersOnSG}   Evaluate  "%.2f" % ${othersOnSG}
    Switch Current Excel Document    doc_id=WoWChange
    ${othersOnWoWChange}   Read Excel Cell    row_num=${othersRowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSG}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${othersOnWoWChange}    ${othersOnSG}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table between the WoW Change Report and SG Report is different
    END
    Close All Excel Documents

Check The Budget Data
    [Arguments]     ${table}    ${nameOfCol}    ${posOfColOnWoWChange}    ${posOfColOnSGWeeklyActionDB}

    ${result}   Set Variable    ${True}
    ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    0
    ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    0
    ${othersRowIndexOnWoWChange}             Set Variable    0
    ${totalRowIndexOnWoWChange}              Set Variable    0

    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange

    File Should Exist      path=${SGWeeklyActionDBFilePath}
    Open Excel Document    filename=${SGWeeklyActionDBFilePath}    doc_id=SGWeeklyActionDB
    Switch Current Excel Document     doc_id=SGWeeklyActionDB
    ${numOfRowsOnSGWeeklyActionDB}    Get Number Of Rows In Excel    ${SGWeeklyActionDBFilePath}
    IF    '${table}' == 'OEM East'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    2
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    7
        ${othersRowIndexOnWoWChange}             Set Variable    7
        ${totalRowIndexOnWoWChange}              Set Variable    8
    ELSE IF     '${table}' == 'OEM West'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    11
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    17
        ${othersRowIndexOnWoWChange}             Set Variable    17
        ${totalRowIndexOnWoWChange}              Set Variable    18
    ELSE
        Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    Switch Current Excel Document    doc_id=WoWChange
    # Verify the data for each OEM Group
    FOR    ${rowIndexOnWoWChange}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}    ${endRowIndexForOEMGroupOnWoWChange}
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}      Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${dataColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        Switch Current Excel Document     doc_id=SGWeeklyActionDB
        FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
            ${oemGroupColOnSGWeeklyActionDB}       Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
            IF    '${oemGroupColOnSGWeeklyActionDB}' == 'Total'
                 BREAK
            END
            ${dataColOnSGWeeklyActionDB}           Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}

            IF    '${dataColOnSGWeeklyActionDB}' == 'None'
                 ${dataColOnSGWeeklyActionDB}     Set Variable    0
            END
            ${dataColOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${dataColOnSGWeeklyActionDB}
            IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSGWeeklyActionDB}'
                 IF    ${dataColOnWoWChange} != ${dataColOnSGWeeklyActionDB}
                      ${result}     Set Variable    ${False}
                      Write The Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupColOnWoWChange}    valueOnWoWChange=${dataColOnWoWChange}    valueOnSG=${dataColOnSGWeeklyActionDB}
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

        IF    '${dataColOnSGWeeklyActionDB}' == 'None'
               ${dataColOnSGWeeklyActionDB}     Set Variable    0
        END

        IF    '${table}' == 'OEM East'
             IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMEastTable}
                   ${totalOnSGWeeklyActionDB}     Evaluate    ${totalOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
             END
        ELSE IF   '${table}' == 'OEM West'
             IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMWestTable}
                   ${totalOnSGWeeklyActionDB}     Evaluate    ${totalOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
             END
        ELSE
            Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
        END
    END
    ${totalOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${totalOnSGWeeklyActionDB}
    Switch Current Excel Document    doc_id=WoWChange
    ${totalOnWoWchange}   Read Excel Cell    row_num=${totalRowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
    IF    ${totalOnWoWchange} != ${totalOnSGWeeklyActionDB}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total    ${totalOnWoWchange}    ${totalOnSGWeeklyActionDB}
    END
    #  Verify the OTHERS data
    Switch Current Excel Document    doc_id=SGWeeklyActionDB
    ${othersOnSGWeeklyActionDB}  Set Variable    0
    FOR    ${rowIndexOnSGWeeklyActionDB}    IN RANGE    4    ${numOfRowsOnSGWeeklyActionDB}+1
        ${oemGroupColOnSGWeeklyActionDB}          Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=1
        ${mainSalesRepColOnSGWeeklyActionDB}      Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=2
        ${dataColOnSGWeeklyActionDB}              Read Excel Cell    row_num=${rowIndexOnSGWeeklyActionDB}    col_num=${posOfColOnSGWeeklyActionDB}

        IF    '${dataColOnSGWeeklyActionDB}' == 'None'
               ${dataColOnSGWeeklyActionDB}     Set Variable    0
        END

        IF    '${table}' == 'OEM East'
             IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMEastTable}
                 IF    '${oemGroupColOnSGWeeklyActionDB}' not in ${listOfOEMGroupShownInOEMEastTable}
                      ${othersOnSGWeeklyActionDB}     Evaluate    ${othersOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
                 END
             END
        ELSE IF   '${table}' == 'OEM West'
             IF    '${mainSalesRepColOnSGWeeklyActionDB}' in ${listOfSalesMemberInOEMWestTable}
                 IF    '${oemGroupColOnSGWeeklyActionDB}' not in ${listOfOEMGroupShownInOEMWestTable}
                      ${othersOnSGWeeklyActionDB}     Evaluate    ${othersOnSGWeeklyActionDB}+${dataColOnSGWeeklyActionDB}
                 END
             END
        ELSE
            Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
        END


    END
    ${othersOnSGWeeklyActionDB}   Evaluate  "%.2f" % ${othersOnSGWeeklyActionDB}
    Switch Current Excel Document    doc_id=WoWChange
    ${othersOnWoWChange}   Read Excel Cell    row_num=${othersRowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSGWeeklyActionDB}
         ${result}     Set Variable    ${False}
         Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${othersOnWoWChange}    ${othersOnSGWeeklyActionDB}
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table between the WoW Change Report and SG Report is different
    END
    Close All Excel Documents

Check The Commit Or Comment Data
    [Arguments]     ${table}    ${nameOfCol}    ${posOfColOnWoWChange}
    ${result}   Set Variable    ${True}
    ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    0
    ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    0


    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}         doc_id=WoWChange

    File Should Exist      path=${wowChangeOnVDCFilePath}
    Open Excel Document    filename=${wowChangeOnVDCFilePath}    doc_id=WoWChangeOnVDC

    IF    '${table}' == 'OEM East'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    2
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    9
    ELSE IF     '${table}' == 'OEM West'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    11
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    19
    ELSE
        Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    FOR    ${rowIndexOnWoWChange}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}    ${endRowIndexForOEMGroupOnWoWChange}
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
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${dataColOnWoWChange}    ${EMPTY}
                  ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total    ${dataColOnWoWChange}    ${EMPTY}
                  ELSE
                       Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${dataColOnWoWChange}    ${EMPTY}
                  END
             END
             Continue For Loop
        END
        Switch Current Excel Document    doc_id=WoWChangeOnVDC
        FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}   ${endRowIndexForOEMGroupOnWoWChange}
            ${oemGroupColOnWoWChangeOnVDC}           Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=1
            IF    '${nameOfCol}' == 'LW Commit'
                 ${posOfColOnWoWChangeOnVDC}         Evaluate    ${posOfColOnWoWChange}+1
                 ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfColOnWoWChangeOnVDC}
            ELSE IF  '${nameOfCol}' == 'Comments'
                 ${dataColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndexOnWoWChangeOnVDC}    col_num=${posOfColOnWoWChange}
            ELSE
               Close All Excel Documents
               Fail  The name of column ${nameOfCol} is invalid. Please contact with the Administrator for supporting
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
                          Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                     ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
                          Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total     ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                     ELSE
                          Write The Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${oemGroupColOnWoWChange}    ${dataColOnWoWChange}    ${dataColOnWoWChangeOnVDC}
                     END
                 END
                 BREAK
            END
        END
    END

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} column for the ${table} table between the WoW Change Report and WoW Change Report On VDC is different
    END
    Close All Excel Documents

Check The WoW Data
    [Arguments]     ${table}    ${nameOfCol}    ${posOfColOnWoWChange}

    ${result}   Set Variable    ${True}

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}         doc_id=WoWChange

    File Should Exist      path=${wowChangeOnVDCFilePath}
    Open Excel Document    filename=${wowChangeOnVDCFilePath}    doc_id=WoWChangeOnVDC

    IF    '${table}' == 'OEM East'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    2
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    9
    ELSE IF     '${table}' == 'OEM West'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    11
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    19
    ELSE
        Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    FOR    ${rowIndexOnWoWChange}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}    ${endRowIndexForOEMGroupOnWoWChange}
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}           Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${wowColOnWoWChange}                Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        ${posOfDataColOnWoWChange}          Evaluate    ${posOfColOnWoWChange}-1
        ${dataColOnWoWChange}               Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfDataColOnWoWChange}
        Switch Current Excel Document    doc_id=WoWChangeOnVDC
        FOR    ${rowIndexOnWoWChangeOnVDC}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}   ${endRowIndexForOEMGroupOnWoWChange}
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

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END
    Close All Excel Documents

Check The GAP Data
    [Arguments]     ${table}    ${nameOfCol}    ${posOfColOnWoWChange}      ${posOfRColOnSG}      ${posOfBColOnSG}
    ${result}   Set Variable    ${True}
    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange

    File Should Exist      path=${wowChangeOnVDCFilePath}
    Open Excel Document    filename=${wowChangeOnVDCFilePath}      doc_id=WoWChangeOnVDC

    File Should Exist      path=${SGFilePath}
    Open Excel Document    filename=${SGFilePath}    doc_id=SG
    Switch Current Excel Document    doc_id=SG
    ${numOfRowsOnSG}    Get Number Of Rows In Excel    ${SGFilePath}

    IF    '${table}' == 'OEM East'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    2
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    9
    ELSE IF     '${table}' == 'OEM West'
        ${startRowIndexForOEMGroupOnWoWChange}   Set Variable    11
        ${endRowIndexForOEMGroupOnWoWChange}     Set Variable    19
    ELSE
        Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    FOR    ${rowIndexOnWoWChange}    IN RANGE    ${startRowIndexForOEMGroupOnWoWChange}    ${endRowIndexForOEMGroupOnWoWChange}
        Switch Current Excel Document    doc_id=WoWChange
        ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=1
        ${gapColOnWoWChange}               Read Excel Cell    row_num=${rowIndexOnWoWChange}    col_num=${posOfColOnWoWChange}
        ${gapColOnWoWChange}   Evaluate  "%.2f" % ${gapColOnWoWChange}

        ${los}      Set Variable    0
        Switch Current Excel Document    doc_id=SG
        IF  '${oemGroupColOnWoWChange}' == 'OTHERS'
             FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
                 ${oemGroupColOnSG}        Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
                 ${mainSalesRepColOnSG}    Read Excel Cell    row_num=${rowIndexOnSG}    col_num=3
                 ${dataRColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
                 ${dataBColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}

                 IF    '${dataRColOnSG}' == 'None'
                      ${dataRColOnSG}   Set Variable    0
                 END
                 IF    '${dataBColOnSG}' == 'None'
                      ${dataBColOnSG}   Set Variable    0
                 END
                 ${dataLOSOnSG}     Evaluate    ${dataRColOnSG}+${dataBColOnSG}

                 IF    '${table}' == 'OEM East'
                      IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
                          IF    '${oemGroupColOnSG}' not in ${listOfOEMGroupShownInOEMEastTable}
                                ${los}   Evaluate    ${los}+${dataLOSOnSG}
                          END
                      END
                 ELSE IF    ${table} == 'OEM West'
                    IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
                          IF    '${oemGroupColOnSG}' not in ${listOfOEMGroupShownInOEMWestTable}
                                ${los}   Evaluate    ${los}+${dataLOSOnSG}
                          END
                    END
                 ELSE
                    Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
                 END

             END
        ELSE IF  '${oemGroupColOnWoWChange}' == 'Total'
            FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
                ${oemGroupColOnSG}        Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
                ${mainSalesRepColOnSG}    Read Excel Cell    row_num=${rowIndexOnSG}    col_num=3
                ${dataRColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
                ${dataBColOnSG}           Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}

                IF    '${dataRColOnSG}' == 'None'
                      ${dataRColOnSG}   Set Variable    0
                 END
                 IF    '${dataBColOnSG}' == 'None'
                      ${dataBColOnSG}   Set Variable    0
                 END
                 ${dataLOSOnSG}     Evaluate    ${dataRColOnSG}+${dataBColOnSG}

                 IF    '${table}' == 'OEM East'
                      IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
                           ${los}   Evaluate    ${los}+${dataLOSOnSG}
                      END
                 ELSE IF    '${table}' == 'OEM West'
                      IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMWestTable}
                           ${los}   Evaluate    ${los}+${dataLOSOnSG}
                      END
                 ELSE
                     Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
                 END

            END
        ELSE
            FOR    ${rowIndexOnSG}    IN RANGE    6    ${numOfRowsOnSG}+1
                 ${oemGroupColOnSG}     Read Excel Cell    row_num=${rowIndexOnSG}    col_num=2
                 ${dataRColOnSG}        Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfRColOnSG}
                 ${dataBColOnSG}        Read Excel Cell    row_num=${rowIndexOnSG}    col_num=${posOfBColOnSG}

                 IF    '${dataRColOnSG}' == 'None'
                      ${dataRColOnSG}   Set Variable    0
                 END
                 IF    '${dataBColOnSG}' == 'None'
                      ${dataBColOnSG}   Set Variable    0
                 END
                 ${dataLOSOnSG}     Evaluate    ${dataRColOnSG}+${dataBColOnSG}

                 IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSG}'
                     ${los}     Set Variable    ${dataLOSOnSG}
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

    IF    '${result}' == '${False}'
         Close All Excel Documents
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END
    Close All Excel Documents

