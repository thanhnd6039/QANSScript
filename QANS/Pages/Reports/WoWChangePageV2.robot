*** Settings ***
Resource    ../CommonPage.robot
Resource    SGPage.robot

*** Variables ***
${wowChangeResultFilePath}        C:\\RobotFramework\\Results\\WoWChangeReportResult.xlsx
${wowChangeFilePath}              C:\\RobotFramework\\Downloads\\Wow Change [Current Week].xlsx
${wowChangeOnVDCFilePath}         C:\\RobotFramework\\Downloads\\Wow Change [Current Week] On VDC.xlsx
${SGFilePath}                     C:\\RobotFramework\\Downloads\\Sales Gap Report NS With SO Forecast.xlsx
${SGWeeklyActionDBFilePath}       C:\\RobotFramework\\Downloads\\SalesGap Weekly Actions Gap Week DB.xlsx

${posOfOEMGroupColOnWoWChange}              1

*** Keywords ***
Get WoW By Formular By OEM GRoup
    [Arguments]     ${table}    ${nameOfCol}    ${oemGroup}
    ${wowByFormula}           Set Variable    0
    ${valueOnWoWChange}       Set Variable    0
    ${valueOnWoWChangeOnVDC}  Set Variable    0


    ${tableOnWoWChange}     Create Table On WoW Change    table=${table}    nameOfCol=${nameOfCol}
    ${tableOnWoWChangeOnVDC}    Create Table On WoW Change On VDC    table=${table}    nameOfCol=${nameOfCol}

    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[0]}
        IF    '${oemGroupOnWoWChange}' == '${oemGroup}'
            ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[1]}
            BREAK
        END
    END

    FOR    ${rowOnWoWChangeOnVDC}    IN    @{tableOnWoWChangeOnVDC}
        ${oemGroupOnWoWChangeOnVDC}  Set Variable    ${rowOnWoWChangeOnVDC[0]}
        IF    '${oemGroupOnWoWChangeOnVDC}' == '${oemGroup}'
            ${valueOnWoWChangeOnVDC}     Set Variable    ${rowOnWoWChangeOnVDC[1]}
            BREAK
        END
    END

    ${wowByFormula}     Evaluate    ${valueOnWoWChange}-${valueOnWoWChangeOnVDC}
    [Return]    ${wowByFormula}

Check WoW On WoW Change
    [Arguments]     ${table}    ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${tableOnWoWChange}                    Create Table On WoW Change    table=${table}    nameOfCol=${nameOfCol}
    ${wowByFormula}     Set Variable    0

    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[0]}
        ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[1]}
        IF    '${nameOfCol}' == 'WoW Of Ships'
             ${wowByFormula}         Get WoW By Formular By OEM GRoup    table=${table}    nameOfCol=Ships    oemGroup=${oemGroupOnWoWChange}
        ELSE
             ${wowByFormula}         Get WoW By Formular By OEM GRoup    table=${table}    nameOfCol=LOS    oemGroup=${oemGroupOnWoWChange}
        END

        ${valueOnWoWChange}      Evaluate  "%.2f" % ${valueOnWoWChange}
        ${wowByFormula}          Evaluate  "%.2f" % ${wowByFormula}

        IF    '${valueOnWoWChange}' != '${wowByFormula}'
              ${result}     Set Variable    ${False}
              IF    '${oemGroupOnWoWChange}' == 'Total'
                   Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${table} Total    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${wowByFormula}
              ELSE IF   '${oemGroupOnWoWChange}' == 'OTHERS'
                   Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${table} OTHERS    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${wowByFormula}
              ELSE
                   Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupOnWoWChange}    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${wowByFormula}
              END
        END
    END
    IF    '${result}' == '${False}'
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END

Check TW Commit On WoW Change
    [Arguments]     ${table}    ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${tableOnWoWChange}     Create Table On WoW Change    table=${table}    nameOfCol=${nameOfCol}
    
    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[0]}
        ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[1]}
        IF    '${valueOnWoWChange}' != '${EMPTY}'
             ${result}     Set Variable    ${False}
             IF    '${oemGroupOnWoWChange}' == 'Total'
                  Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${table} Total    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${EMPTY}
             ELSE IF   '${oemGroupOnWoWChange}' == 'OTHERS'
                  Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${table} OTHERS    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${EMPTY}
             ELSE
                  Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupOnWoWChange}    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${EMPTY}
             END
        END
    END
    IF    '${result}' == '${False}'
         Fail   The ${nameOfCol} data for the ${table} table is wrong
    END
Check LW Commit, Comment On WoW Change
    [Arguments]     ${table}    ${nameOfCol}
    ${result}   Set Variable    ${True}
    ${tableOnWoWChange}     Create Table On WoW Change    table=${table}    nameOfCol=${nameOfCol}
    ${tableOnWoWChangeOnVDC}    Create Table On WoW Change On VDC    table=${table}    nameOfCol=${nameOfCol}
    FOR    ${rowOnWoWChange}    IN    @{tableOnWoWChange}
        ${oemGroupOnWoWChange}  Set Variable    ${rowOnWoWChange[0]}
        ${valueOnWoWChange}     Set Variable    ${rowOnWoWChange[1]}
        FOR    ${rowOnWoWChangeOnVDC}    IN    @{tableOnWoWChangeOnVDC}
            ${oemGroupOnWoWChangeOnVDC}  Set Variable    ${rowOnWoWChangeOnVDC[0]}
            ${valueOnWoWChangeOnVDC}     Set Variable    ${rowOnWoWChangeOnVDC[1]}
            IF    '${oemGroupOnWoWChange}' == '${oemGroupOnWoWChangeOnVDC}'
                 IF    '${valueOnWoWChange}' != '${valueOnWoWChangeOnVDC}'
                      ${result}     Set Variable    ${False}
                      IF    '${oemGroupOnWoWChange}' == 'Total'
                           Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${table} Total    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${valueOnWoWChangeOnVDC}
                      ELSE IF   '${oemGroupOnWoWChange}' == 'OTHERS'
                           Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${table} OTHERS    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${valueOnWoWChangeOnVDC}
                      ELSE
                           Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupOnWoWChange}    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${valueOnWoWChangeOnVDC}
                      END
                 END
                 BREAK
            END
        END
    END
    IF    '${result}' == '${False}'
         Fail   The ${nameOfCol} data for the ${table} table is different between the WoW Change Report and WoW Change on VDC
    END

Create Table On WoW Change On VDC
    [Arguments]     ${table}    ${nameOfCol}

    @{tableOnWoWChangeOnVDC}     Create List
    ${result}       Set Variable    ${True}
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0

    IF    '${table}' != 'OEM East' and '${table}' != 'OEM West + Channel'
         Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    ${startRow}     Get Start Row On WoW Change    table=${table}
    ${endRow}       Get End Row On WoW Change    table=${table}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    IF    '${nameOfCol}' == 'LW Commit'
         ${posOfValueCol}    Set Variable    5
    ELSE IF   '${nameOfCol}' == 'WoW Of Ships'
         ${posOfValueCol}    Set Variable    7
    ELSE IF   '${nameOfCol}' == 'WoW Of LOS'
         ${posOfValueCol}    Set Variable    10
    ELSE
        ${posOfValueCol}     Get Position Of Column On WoW Change    table=${table}    nameOfCol=${nameOfCol}
    END

    File Should Exist      path=${wowChangeOnVDCFilePath}
    Open Excel Document    filename=${wowChangeOnVDCFilePath}           doc_id=WoWChangeOnVDC
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${wowChangeOnVDCFilePath}
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${totalRow}+1
        ${oemGroupColOnWoWChangeOnVDC}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnWoWChange}
        ${valueColOnWoWChangeOnVDC}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        ${rowOnTable}   Create List
        ...             ${oemGroupColOnWoWChangeOnVDC}
        ...             ${valueColOnWoWChangeOnVDC}
        Append To List    ${tableOnWoWChangeOnVDC}   ${rowOnTable}
    END

    Close Current Excel Document
    [Return]    ${tableOnWoWChangeOnVDC}

Create Table On WoW Change
    [Arguments]     ${table}    ${nameOfCol}

    @{tableOnWoWChange}     Create List
    ${result}       Set Variable    ${True}
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0

    IF    '${table}' != 'OEM East' and '${table}' != 'OEM West + Channel'
         Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    ${startRow}     Get Start Row On WoW Change    table=${table}
    ${endRow}       Get End Row On WoW Change    table=${table}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    IF    '${nameOfCol}' == 'LW Commit'
         ${posOfValueCol}    Set Variable    4
    ELSE IF  '${nameOfCol}' == 'TW Commit'
         ${posOfValueCol}    Set Variable    5
    ELSE IF  '${nameOfCol}' == 'WoW Of Ships'
         ${posOfValueCol}    Set Variable    7
    ELSE IF  '${nameOfCol}' == 'WoW Of LOS'
         ${posOfValueCol}    Set Variable    10
    ELSE
        ${posOfValueCol}     Get Position Of Column On WoW Change    table=${table}    nameOfCol=${nameOfCol}
    END

    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${wowChangeFilePath}
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${totalRow}+1
        ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnWoWChange}
        ${valueColOnWoWChange}             Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        ${rowOnTable}   Create List
        ...             ${oemGroupColOnWoWChange}
        ...             ${valueColOnWoWChange}
        Append To List    ${tableOnWoWChange}   ${rowOnTable}
    END

    Close Current Excel Document
    [Return]    ${tableOnWoWChange}

Check BGT, Ship, Backlog, LOS On WoW Change
    [Arguments]     ${table}    ${nameOfCol}     ${transType}   ${attribute}   ${year}     ${quarter}

    ${result}       Set Variable    ${True}
    ${startRow}     Set Variable    0
    ${endRow}       Set Variable    0
    ${othersRow}    Set Variable    0
    ${totalRow}     Set Variable    0   
    ${sumOfValueOfOEMGroup}     Set Variable    0

    IF    '${table}' != 'OEM East' and '${table}' != 'OEM West + Channel'
         Fail    The table parameter ${table} is invalid. Please contact with the Administrator for supporting
    END

    ${startRow}     Get Start Row On WoW Change    table=${table}
    ${endRow}       Get End Row On WoW Change    table=${table}
    ${othersRow}    Evaluate    ${endRow}+1
    ${totalRow}     Evaluate    ${endRow}+2

    IF    '${nameOfCol}' == 'Pre Q Ships'
         ${posOfValueCol}    Set Variable    2
    ELSE IF  '${nameOfCol}' == 'Current Q Budget'
         ${posOfValueCol}    Set Variable    3
    ELSE IF  '${nameOfCol}' == 'Current Q Budget'
         ${posOfValueCol}    Set Variable    4
    ELSE
        ${posOfValueCol}     Get Position Of Column On WoW Change    table=${table}    nameOfCol=${nameOfCol}
    END

    IF    '${posOfValueCol}' == '0'
         Fail   Not found the position of column
    END

    ${listOfSalesMemberInOEMEastTable}       Get List Of Sales Member In OEM East Table
    ${listOfOEMGroupShownInOEMEastTable}     Get List Of OEM Group Shown In OEM East Table
    ${listOfSalesMemberInOEMWestTable}       Get List Of Sales Member In OEM West Table
    ${listOfOEMGroupShownInOEMWestTable}     Get List Of OEM Group Shown In OEM West Table
    ${tableOnSG}    Create Table For SG Report    transType=${transType}    attribute=${attribute}    year=${year}    quarter=${quarter}
    
    File Should Exist      path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}           doc_id=WoWChange
    ${numOfRows}  Get Number Of Rows In Excel    filePath=${wowChangeFilePath}
   #   Verify the data for each OEM Group
    FOR    ${rowIndex}    IN RANGE    ${startRow}    ${endRow}+1
        ${oemGroupColOnWoWChange}          Read Excel Cell    row_num=${rowIndex}    col_num=${posOfOEMGroupColOnWoWChange}
        ${valueOnWoWChange}                Read Excel Cell    row_num=${rowIndex}    col_num=${posOfValueCol}
        ${valueOnSG}   Set Variable     ${EMPTY}
        FOR    ${rowData}    IN    @{tableOnSG}
            ${oemGroupColOnSG}  Set Variable    ${rowData[0]}
            IF    '${oemGroupColOnWoWChange}' == '${oemGroupColOnSG}'
                 ${valueOnSG}   Set Variable    ${rowData[2]}
                 BREAK
            END
        END
        ${valueOnWoWChange}      Evaluate  "%.2f" % ${valueOnWoWChange}
        ${valueOnSG}              Evaluate  "%.2f" % ${valueOnSG}
        IF    ${valueOnWoWChange} != ${valueOnSG}
              ${result}     Set Variable    ${False}
              Write Test Result Of WoW Change Report To Excel    item=${nameOfCol}    oemGroup=${oemGroupColOnWoWChange}    valueOnWoWChange=${valueOnWoWChange}    valueOnSG=${valueOnSG}
        END
        ${sumOfValueOfOEMGroup}     Evaluate    ${sumOfValueOfOEMGroup}+${valueOnSG}
    END
    #   Verify the Total data
    ${totalOnSG}    Set Variable    0
    ${valueOnSG}    Set Variable    0
    FOR    ${rawData}    IN    @{tableOnSG}
        ${mainSalesRepColOnSG}  Set Variable    ${rawData[1]}
        IF    '${table}' == 'OEM East'
             IF    '${mainSalesRepColOnSG}' in ${listOfSalesMemberInOEMEastTable}
                ${valueOnSG}    Set Variable    ${rawData[2]}
                ${totalOnSG}    Evaluate    ${totalOnSG}+${valueOnSG}
             END
        END
    END

    ${totalOnWoWchange}   Read Excel Cell    row_num=${totalRow}    col_num=${posOfValueCol}
    ${totalOnSG}   Evaluate  "%.2f" % ${totalOnSG}
    ${totalOnWoWchange}   Evaluate  "%.2f" % ${totalOnWoWchange}
    IF    ${totalOnWoWchange} != ${totalOnSG}
         ${result}     Set Variable    ${False}
         Write Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} Total    ${totalOnWoWchange}    ${totalOnSG}
    END
    #  Verify the OTHERS data
    ${othersOnSG}   Evaluate    ${totalOnSG}-${sumOfValueOfOEMGroup}
    ${othersOnWoWChange}   Read Excel Cell    row_num=${othersRow}    col_num=${posOfValueCol}
    ${othersOnSG}   Evaluate  "%.2f" % ${othersOnSG}
    ${othersOnWoWChange}   Evaluate  "%.2f" % ${othersOnWoWChange}
    IF    ${othersOnWoWChange} != ${othersOnSG}
         ${result}     Set Variable    ${False}
         Write Test Result Of WoW Change Report To Excel    ${nameOfCol}    ${table} OTHERS    ${othersOnWoWChange}    ${othersOnSG}
    END

    IF    '${result}' == '${False}'
         Close Current Excel Document
         Fail   The ${nameOfCol} data for the ${table} table is different between the WoW Change Report and SG Report
    END
    Close Current Excel Document
    
Write Test Result Of WoW Change Report To Excel
    [Arguments]     ${item}     ${oemGroup}     ${valueOnWoWChange}   ${valueOnSG}
    File Should Exist      path=${wowChangeResultFilePath}
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
    Append To List    ${listOfSalesMember}      Clint Stalker

    [Return]    ${listOfSalesMember}

Get List Of OEM Group Shown In OEM East Table
    @{listOfOEMGroup}   Create List
    Append To List    ${listOfOEMGroup}      NVIDIA/MELLANOX
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
    Append To List    ${listOfOEMGroup}      PANASONIC AVIONICS
    Append To List    ${listOfOEMGroup}      ZTE KANGXUN TELECOM CO. LTD.
    Append To List    ${listOfOEMGroup}      NATIONAL INSTRUMENTS

    [Return]    ${listOfOEMGroup}


Get Row Index For Search Col
    [Arguments]     ${table}
    ${rowIndexForSearchCol}     Set Variable    0

    File Should Exist    path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${wowChangeFilePath}
    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}      Read Excel Cell    row_num=${rowIndex}    col_num=1
        IF    '${oemGroupCol}' == '${table}'
             ${rowIndexForSearchCol}    Set Variable    ${rowIndex}
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${rowIndexForSearchCol}

Get Position Of Column On WoW Change
    [Arguments]     ${table}    ${nameOfCol}
    ${pos}  Set Variable    0
    ${rowIndexForSearchCol}     Set Variable    0

    ${rowIndexForSearchCol}     Get Row Index For Search Col    ${table}
    ${pos}      Get Position Of Column    filePath=${wowChangeFilePath}    rowIndex=${rowIndexForSearchCol}    searchStr=${nameOfCol}

    [Return]    ${pos}


Get Start Row On WoW Change
    [Arguments]     ${table}

    ${startRow}     Set Variable    0
    ${posOfCol}     Get Position Of Column On WoW Change    table=${table}   nameOfCol=${table}

    File Should Exist    path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${wowChangeFilePath}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == '${table}'
             ${startRow}    Evaluate    ${rowIndex}+1
             BREAK
        END
    END

    Close Current Excel Document
    [Return]    ${startRow}

Get End Row On WoW Change
    [Arguments]     ${table}

    ${endRow}   Set Variable    0
    ${count}    Set Variable    0

    ${posOfCol}     Get Position Of Column On WoW Change    table=${table}   nameOfCol=${table}

    File Should Exist    path=${wowChangeFilePath}
    Open Excel Document    filename=${wowChangeFilePath}    doc_id=WoWChange
    ${numOfRows}    Get Number Of Rows In Excel    filePath=${wowChangeFilePath}

    FOR    ${rowIndex}    IN RANGE    1    ${numOfRows}+1
        ${oemGroupCol}     Read Excel Cell    row_num=${rowIndex}    col_num=${posOfCol}
        IF    '${oemGroupCol}' == 'OTHERS'
             ${count}   Evaluate    ${count}+1
             IF    '${table}' == 'OEM East'
                  IF    '${count}' == '1'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             ELSE IF     '${table}' == 'OEM West + Channel'
                  IF    '${count}' == '2'
                       ${endRow}    Evaluate    ${rowIndex}-1
                       BREAK
                  END
             END
        END
    END

    Close Current Excel Document
    [Return]    ${endRow}





